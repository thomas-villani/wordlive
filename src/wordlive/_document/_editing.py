"""Text editing: prepend/append/delete, find/replace, field update."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .. import _com, _findreplace
from .._anchors import (
    Heading,
    _utf16_len,
    _within_table,
)
from ..exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    ReplaceVerificationError,
)

if TYPE_CHECKING:
    from .._anchors import Anchor

from ._core import DocumentCore


class EditingMixin(DocumentCore):
    """Text editing: prepend/append/delete, find/replace, field update."""

    def _scope_range(self, scope: Anchor | None) -> tuple[Any, int]:
        """Return (COM Range, absolute_start_offset) for a find/replace scope.

        Headings expand to their *section* (body under the heading); other
        anchor kinds use their own range. `None` means the whole document.
        """
        with _com.translate_com_errors():
            if scope is None:
                rng = self._doc.Content
            elif isinstance(scope, Heading):
                rng = scope.section_range()
            else:
                rng = scope.com
            return rng, int(rng.Start)

    def _scope_segments(self, scope: Anchor | None) -> list[tuple[int, str]]:
        """Split a find/replace scope into segments with an exact text↔position map.

        `Range.Text` string offsets line up 1:1 with Word document positions
        *within* a single body run or a single table cell, but NOT across table
        structure: once a range spans cells, `len(Range.Text) != End - Start`, so
        matching on a whole-document `.Text` and adding `base + offset` silently
        drifts into a neighbouring cell. Segmenting at table-cell boundaries keeps
        every segment's offsets exact — contiguous non-table paragraphs form one
        body segment (so cross-paragraph matches still work) and each table cell
        is its own segment. Each tuple is `(base_position, text)`; a match at
        `m.start` inside a segment maps back to the absolute `base + m.start`.

        Segments come out in document order (ascending base), so the matches the
        callers build from them preserve the original document ordering.
        """
        with _com.translate_com_errors():
            rng, base = self._scope_range(scope)
            # Fast path: a scope with no table structure maps 1:1 already, so one
            # segment over the whole range reproduces the original behavior (and
            # avoids per-paragraph reads the test fake doesn't model). Only ranges
            # that actually span a table need the boundary-aware walk below.
            try:
                has_table = int(rng.Tables.Count) > 0
            except (TypeError, ValueError, AttributeError):
                has_table = False
            if not has_table:
                return [(base, str(rng.Text or ""))]
            doc_com = self._doc  # the raw COM document (see _scope_range)
            segments: list[tuple[int, str]] = []
            seg_key: object | None = None
            seg_start: int | None = None
            seg_end: int | None = None

            def flush() -> None:
                nonlocal seg_key, seg_start, seg_end
                if seg_start is not None and seg_end is not None and seg_end > seg_start:
                    text = str(doc_com.Range(seg_start, seg_end).Text or "")
                    if seg_key != "body":
                        # A cell's text ends with CR + the cell mark (`\r\x07`),
                        # which together occupy a single document position — so
                        # `len(text)` runs one past `End - Start`, and a match at
                        # the cell's tail would map its end *past* the cell into
                        # the next one (the cause of the old `'Opus\r\x072'`
                        # boundary error). Drop those trailing markers; the
                        # remaining content stays 1:1 with document positions.
                        text = text.rstrip("\r\n\x07")
                    segments.append((seg_start, text))
                seg_key = seg_start = seg_end = None

            for para in rng.Paragraphs:
                pr = para.Range
                ps, pe = int(pr.Start), int(pr.End)
                if _within_table(doc_com, ps, pe):
                    # Key by the containing cell so two adjacent cells never share
                    # a segment (a range spanning them would break the 1:1 map).
                    try:
                        key: object = int(pr.Cells(1).Range.Start)
                    except Exception:
                        key = ps  # defensive: give this paragraph its own segment
                else:
                    key = "body"
                if key != seg_key:
                    flush()
                    seg_key, seg_start = key, ps
                seg_end = pe
            flush()
        return segments

    def find(
        self,
        text: str,
        *,
        scope: Anchor | None = None,
        mode: str = "fuzzy",
    ) -> list[dict[str, Any]]:
        """Locate every occurrence of `text` within `scope` (or the whole doc).

        `mode` selects the matcher:

        - `fuzzy` (default): whitespace- and Unicode-normalized (NFKC, smart
          quotes, dashes, NBSP) — forgiving of cosmetic drift.
        - `literal`: exact substring, no folding.
        - `regex`: `text` is a Python regular expression.

        Returns a list of `{anchor_id, start, end, text}` where offsets are
        absolute document positions and `text` is the actual original substring
        (not the normalized form).

        `anchor_id` for each match is `range:START-END`, which resolves through
        `anchor_by_id` to a `RangeAnchor` — so a hit can be fed straight back
        into `replace --anchor-id` or `comments.add`. The offsets are live,
        though, so use them before further edits shift the document.

        Matches are located per *segment* (contiguous body text or a single table
        cell) so the returned offsets stay exact even inside tables; see
        `_scope_segments`.
        """
        segments = self._scope_segments(scope)
        results: list[dict[str, Any]] = []
        for base, haystack in segments:
            for m in _findreplace.find_matches(haystack, text, mode=mode):
                results.append(
                    {
                        "anchor_id": f"range:{base + m.start}-{base + m.end}",
                        "start": base + m.start,
                        "end": base + m.end,
                        "text": m.text,
                    }
                )
        return results

    def find_replace(
        self,
        find: str,
        replace: str,
        *,
        scope: Anchor | None = None,
        all: bool = False,
        occurrence: int | None = None,
        mode: str = "fuzzy",
        required: bool = True,
    ) -> list[dict[str, Any]]:
        """Plain-text replace. See `find()` for matching semantics.

        Args:
            find: the text to look for.
            replace: the replacement text. In `regex` mode it may carry
                backreferences (`\\1`) that expand per match.
            scope: optional anchor to restrict the search to. Headings expand
                to their body section.
            all: replace every match.
            occurrence: 1-based index — replace only the Nth match.
            mode: `fuzzy` (default) / `literal` / `regex` — see `find()`.
            required: when `False`, zero matches returns `[]` instead of raising.
                Used by idempotent batch autofixes (the linter) where an earlier
                fix may already have removed an overlapping match in the same pass.

        Raises:
            AnchorNotFoundError: zero matches and `required` (uses `kind='find'`).
            AmbiguousMatchError: more than one match and neither `all` nor
                `occurrence` was given.

        Returns the list of replacements actually applied, each
        `{anchor_id, start, end, text}` in their pre-replacement coordinates.

        Matching is segment-aware (see `_scope_segments`), so a match inside a
        table cell resolves to the right cell rather than drifting into its
        neighbour. As a backstop, each write is verified against the located text
        and raises `ReplaceVerificationError` rather than overwriting the wrong
        span.
        """
        # Only `regex` needs the replacement at match time (per-hit backreference
        # expansion); `fuzzy`/`literal` apply the single `replace` to every match.
        repl_template = replace if mode == "regex" else None
        segments = self._scope_segments(scope)
        match_payloads: list[dict[str, Any]] = [
            {
                "anchor_id": f"range:{base + m.start}-{base + m.end}",
                "start": base + m.start,
                "end": base + m.end,
                "text": m.text,
                # Private to the write loop; stripped from the returned payloads.
                "_replacement": m.replacement if m.replacement is not None else replace,
            }
            for base, haystack in segments
            for m in _findreplace.find_matches(haystack, find, mode=mode, replacement=repl_template)
        ]
        if not match_payloads:
            if not required:
                return []
            raise AnchorNotFoundError("find", find)

        if occurrence is not None:
            if occurrence < 1 or occurrence > len(match_payloads):
                raise AnchorNotFoundError("find", f"{find} (occurrence {occurrence})")
            to_apply = [match_payloads[occurrence - 1]]
        elif all:
            to_apply = match_payloads
        elif len(match_payloads) == 1:
            to_apply = match_payloads
        else:
            raise AmbiguousMatchError(find, match_payloads)

        with _com.translate_com_errors():
            # Word's final paragraph mark is undeletable; a range whose End reaches
            # Content.End straddles it and raises COM 0x80020009. Clamp the write
            # target (not the returned payload, which promises pre-edit offsets).
            doc_end = int(self._doc.Content.End)
            # Apply in reverse so earlier offsets don't shift.
            for m in reversed(to_apply):
                start, end = m["start"], min(m["end"], doc_end - 1)
                if end <= start:
                    # Clamped away to nothing (match was only the trailing mark).
                    continue
                target = self._doc.Range(start, end)
                # Verify the resolved span before writing. An empty resolved text
                # means we can't check (the fake COM, or a genuinely empty range)
                # — proceed. A non-empty mismatch means the offset map drifted
                # (table position divergence): refuse rather than corrupt.
                resolved = str(target.Text or "")
                if resolved and not _findreplace.normalized_equal(resolved, m["text"]):
                    raise ReplaceVerificationError(
                        find, m["text"], resolved, anchor_id=m["anchor_id"]
                    )
                target.Text = m["_replacement"]
        # Drop the internal replacement key; the documented payload is 4 keys.
        return [{k: v for k, v in m.items() if k != "_replacement"} for m in to_apply]

    def prepend(self, text: str) -> None:
        """Prepend `text` to the very start of the document, inline (no new paragraph).

        The mirror of [`append`][wordlive.Document.append]: `text` lands before
        the document's first character, joining the opening paragraph. Embed
        `\\r` / `\\n` for your own paragraph breaks; reach for
        [`prepend_paragraph`][wordlive.Document.prepend_paragraph] when you want
        `text` to *become* a new first paragraph. Wrap in `doc.edit(...)` for
        atomic undo. Not idempotent — each call adds more text.
        """
        with _com.translate_com_errors():
            self._doc.Content.InsertBefore(text)

    def prepend_paragraph(self, text: str, *, style: str | None = None) -> None:
        """Prepend `text` as a new paragraph at the very start of the document.

        The mirror of [`append_paragraph`][wordlive.Document.append_paragraph]
        — for a title, a banner, or a disclaimer above everything else. `text`
        may contain `\\r` / `\\n` to prepend several paragraphs at once. If
        `style` is given it must name a style defined in the document, otherwise
        `StyleNotFoundError` is raised before any text is inserted. Wrap in
        `doc.edit(...)` for atomic undo. Not idempotent.

        Equivalent to `insert_paragraph_before(text, style=style)` on the
        document's first paragraph.
        """
        style_obj = self.styles[style] if style is not None else None  # validate early
        with _com.translate_com_errors():
            doc_com = self._doc
            # The start has no terminal-mark complication: write "<text><break>"
            # at offset 0 so `text` becomes a new first paragraph.
            insert_rng = doc_com.Range(0, 0)
            insert_rng.Text = text + "\r"
            if style_obj is not None:
                # Word counts UTF-16 code units; len() under-counts surrogates.
                styled = doc_com.Range(0, _utf16_len(text))
                styled.Style = style_obj.com

    def append(self, text: str) -> None:
        """Append `text` to the very end of the document, inline (no new paragraph).

        The high-level form of the old `doc.com.Content.InsertAfter(...)` escape
        hatch: `text` lands immediately after the document's last character,
        continuing the final paragraph. Embed `\\r` / `\\n` to introduce your
        own paragraph breaks; reach for
        [`append_paragraph`][wordlive.Document.append_paragraph] when you want
        `text` to *become* a new paragraph. Wrap in `doc.edit(...)` for atomic
        undo. Not idempotent — each call adds more text.
        """
        with _com.translate_com_errors():
            self._doc.Content.InsertAfter(text)

    def append_paragraph(self, text: str, *, style: str | None = None) -> None:
        """Append `text` as a new paragraph at the very end of the document.

        The polite, high-level "end of doc" helper — there is no named anchor
        for the position past the last paragraph, so this is how you add a
        closing note, drop in a generated summary, or build a document from the
        bottom up. `text` may contain `\\r` / `\\n` to append several paragraphs
        at once. If `style` is given it must name a style defined in the
        document, otherwise `StyleNotFoundError` is raised before any text is
        inserted. Wrap in `doc.edit(...)` for atomic undo. Not idempotent —
        each call adds another paragraph.

        Equivalent to calling `insert_paragraph_after(text, style=style)` on the
        document's last paragraph, without having to locate it first.
        """
        style_obj = self.styles[style] if style is not None else None  # validate early
        with _com.translate_com_errors():
            doc_com = self._doc
            doc_end = int(doc_com.Content.End)
            # Same trick as Anchor.insert_paragraph_after's terminal branch:
            # write "<break><text>" just before the final paragraph mark so
            # `text` becomes a new final paragraph (the original mark closes
            # it). Writing at Range(doc_end, doc_end) — past the final mark —
            # is a "value out of range" COM error.
            anchor_pos = max(0, doc_end - 1)
            insert_rng = doc_com.Range(anchor_pos, anchor_pos)
            insert_rng.Text = "\r" + text
            if style_obj is not None:
                # Word counts UTF-16 code units; len() under-counts surrogate
                # pairs and would leave the tail of astral text unstyled.
                text_start = anchor_pos + 1
                styled = doc_com.Range(text_start, text_start + _utf16_len(text))
                styled.Style = style_obj.com

    def delete_paragraph(self, anchor: str | Anchor) -> None:
        """Delete the paragraph(s) at `anchor` — text *and* the trailing mark.

        `anchor` is an anchor id (`para:N`, `heading:N`) or an `Anchor`; the
        whole paragraph is removed, mark included, so the surrounding text closes
        up with no empty line left behind (the gap `set_text("")` would leave).
        A range anchor that spans several paragraphs removes all of them.

        Word keeps a mandatory empty paragraph at the very end of the document:
        deleting the *last* paragraph clears its content but leaves that final
        mark (its range otherwise straddles the undeletable terminal mark and
        raises COM `0x80020009`). Wrap in `doc.edit(...)` for atomic undo.
        """
        obj = self.anchor_by_id(anchor) if isinstance(anchor, str) else anchor
        with _com.translate_com_errors():
            rng = obj.com
            start, end = int(rng.Start), int(rng.End)
            doc_end = int(self._doc.Content.End)
            # Never let the range reach Word's undeletable final paragraph mark.
            end = min(end, doc_end - 1)
            if end <= start:
                return
            self._doc.Range(start, end).Delete()

    def update_fields(self) -> None:
        """Refresh the document's fields — recompute every `{ PAGE }`, `{ REF }`, etc.

        Fields (page numbers, cross-references, dates, a TOC) cache their last
        rendered value; after edits that change them, this recomputes the
        document's main-story fields via `Fields.Update()`. The clean "make the
        numbers right again" verb — pair it with
        [`insert_field`][wordlive.Anchor.insert_field]. A
        [`snapshot`][wordlive.Document.snapshot] also forces repagination, so
        `{ PAGE }`/`{ NUMPAGES }` in headers and footers settle without this.
        Wrap in `doc.edit(...)` for atomic undo.

        Scope is the main text story; refreshing fields that live only in
        headers/footers or other stories is deferred.
        """
        with _com.translate_com_errors():
            self._doc.Fields.Update()
