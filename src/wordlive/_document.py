"""Document wrapper + DocumentCollection."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from typing import TYPE_CHECKING, Any

from . import _com, _findreplace
from ._anchors import (
    BookmarkCollection,
    ContentControlCollection,
    EndAnchor,
    Heading,
    HeadingCollection,
    Paragraph,
    ParagraphCollection,
    RangeAnchor,
    StartAnchor,
    _IndexedHeading,
    _utf16_len,
    paragraph_text,
)
from ._comments import CommentCollection
from ._edit import EditScope
from ._lists import ListCollection
from ._sections import SectionCollection
from ._selection import Selection
from ._styles import StyleCollection
from ._tables import TableCollection
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    DocumentNotFoundError,
)

if TYPE_CHECKING:
    from ._anchors import Anchor
    from ._app import Word


class Document:
    """Wraps a Word Document COM object."""

    def __init__(self, word: Word, doc: Any) -> None:
        self._word = word
        self._doc = doc

    @property
    def com(self) -> Any:
        return self._doc

    @property
    def name(self) -> str:
        with _com.translate_com_errors():
            return str(self._doc.Name)

    @property
    def path(self) -> str:
        with _com.translate_com_errors():
            return str(self._doc.FullName)

    @property
    def bookmarks(self) -> BookmarkCollection:
        return BookmarkCollection(self)

    @property
    def content_controls(self) -> ContentControlCollection:
        return ContentControlCollection(self)

    @property
    def styles(self) -> StyleCollection:
        return StyleCollection(self)

    @property
    def tables(self) -> TableCollection:
        """Iterable, indexable view over the document's tables.

        Index by 1-based position (`doc.tables[1]`) or `Title`
        (`doc.tables["Budget"]`). Cells are anchors: `doc.tables[1].cell(2, 3)`
        — or `doc.anchor_by_id("table:1:2:3")` — returns a `Cell` that works
        with `set_text`, `apply_style`, and `format_paragraph`.
        """
        return TableCollection(self)

    @property
    def headings(self) -> HeadingCollection:
        """Iterable view over the document's headings.

        Symmetric with `bookmarks`, `content_controls`, and `styles`. Index by
        visible text (`doc.headings["Risks"]`) or 1-based paragraph position
        (`doc.headings[3]`). `Document.heading(name)` remains as sugar for
        `self.headings[name]`.
        """
        return HeadingCollection(self)

    @property
    def paragraphs(self) -> ParagraphCollection:
        """Indexable, iterable view over *every* paragraph (not just headings).

        Index by 1-based position (`doc.paragraphs[2]`) to get a `Paragraph`
        anchor (`para:N`) that works with `set_text`, `apply_style`,
        `format_paragraph`, and the list verbs. `doc.paragraphs.list()` emits
        offsets, so a body paragraph can be turned into a `range:START-END`
        target for a mid-paragraph insertion. `para:N` shares its index space
        with `heading:N`.
        """
        return ParagraphCollection(self)

    @property
    def lists(self) -> ListCollection:
        """Read-only, iterable view over the document's bullet / numbered lists.

        Index a list by 1-based position (`doc.lists[2]`) to get a
        [`RangeAnchor`][wordlive.RangeAnchor] over its range, so every list verb
        (`apply_list`, `restart_numbering`, …) is available on it. List
        formatting itself is applied through any anchor's `apply_list(...)`.
        """
        return ListCollection(self)

    @property
    def sections(self) -> SectionCollection:
        """Indexable view over the document's sections, headers, and footers.

        `doc.sections[1].header()` / `.footer()` return `HeaderFooter` anchors
        (addressed `header:S:WHICH` / `footer:S:WHICH`) that work with
        `set_text` / `apply_style` like any other anchor.
        """
        return SectionCollection(self)

    @property
    def comments(self) -> CommentCollection:
        """Iterable, indexable view over the document's review comments.

        `doc.comments.add(anchor, text, author=...)` attaches a comment to any
        anchor's range without changing the text — the polite, side-channel way
        to flag something. Index existing comments by 1-based position
        (`doc.comments[2]`) to `resolve()` or `delete()` them.
        """
        return CommentCollection(self)

    @property
    def selection(self) -> Selection:
        return self._word.selection

    @property
    def start(self) -> StartAnchor:
        """An anchor at the very start of the document — the prepend target.

        The mirror of [`end`][wordlive.Document.end]. `doc.start` (anchor id
        `start`, also `anchor_by_id("start")`) names the position before the
        first paragraph; its insert verbs all prepend —
        `doc.start.insert_paragraph_after(text)` adds a new first paragraph
        (delegating to [`prepend_paragraph`][wordlive.Document.prepend_paragraph])
        and `insert_after(text)` prepends inline (delegating to
        [`prepend`][wordlive.Document.prepend]). The CLI reaches it too:
        `wordlive insert --anchor-id start --text "…"`.
        """
        return StartAnchor(self)

    @property
    def end(self) -> EndAnchor:
        """An anchor at the very end of the document — the append target.

        `doc.end` (anchor id `end`, also `anchor_by_id("end")`) names the one
        position no content names: past the last paragraph. Its insert verbs
        all append — `doc.end.insert_paragraph_after(text)` adds a new final
        paragraph (delegating to [`append_paragraph`][wordlive.Document.append_paragraph]),
        `insert_after(text)` appends inline (delegating to
        [`append`][wordlive.Document.append]), and `insert_image(...)` drops a
        picture at the end. Because it resolves through `anchor_by_id`, the CLI
        reaches it too: `wordlive insert --anchor-id end --text "…"`.
        """
        return EndAnchor(self)

    @property
    def track_changes(self) -> bool:
        """Whether Word's Track Changes is currently on for this document."""
        with _com.translate_com_errors():
            return bool(self._doc.TrackRevisions)

    @track_changes.setter
    def track_changes(self, value: bool) -> None:
        with _com.translate_com_errors():
            self._doc.TrackRevisions = bool(value)

    @contextmanager
    def tracked_changes(self) -> Iterator[None]:
        """Turn on Track Changes for the duration of the block, then restore it.

        Every mutation made inside the scope is recorded as a tracked revision
        the user can accept or reject — "make this edit *visibly*." The prior
        `TrackRevisions` setting is restored on exit, so the scope stays polite
        even when the user had tracking off.

        Pairs with `edit()` for an atomic, visibly-tracked batch:

            with doc.tracked_changes(), doc.edit("Suggest rewordings"):
                doc.find_replace("utilise", "use", all=True)
        """
        with _com.translate_com_errors():
            previous = bool(self._doc.TrackRevisions)
            self._doc.TrackRevisions = True
        try:
            yield
        finally:
            with _com.translate_com_errors():
                self._doc.TrackRevisions = previous

    def heading(self, name: str) -> Heading:
        # Lazy lookup — Heading.__init__ doesn't hit COM. _range() validates.
        return Heading(self, name)

    def range(self, start: int, end: int) -> RangeAnchor:
        """Return a `RangeAnchor` over the absolute offsets `[start, end)`.

        Offsets are UTF-16 code units — the coordinates Word uses and that
        `find()` emits as `range:START-END`. Lazy: the offsets aren't validated
        against the document until the anchor is used.
        """
        return RangeAnchor(self, start, end)

    def anchor_by_id(self, anchor_id: str) -> Anchor:
        """Resolve an `anchor_id` string into an Anchor.

        Recognised forms:
          - `start`            — the position before the first paragraph (the prepend target)
          - `end`              — the position past the last paragraph (the append target)
          - `heading:N`        — Nth paragraph in the document (1-based, must be a heading)
          - `para:N`           — Nth paragraph (1-based, any paragraph; same index space as `heading:N`)
          - `bookmark:NAME`    — bookmark by name
          - `cc:NAME`          — content control by Title (or Tag)
          - `table:N:R:C`      — cell at 1-based (row, column) of the Nth table
          - `range:START-END`  — arbitrary character span (the form `find()` emits)
          - `header:S:WHICH`   — the WHICH header of section S (WHICH = primary/first/even)
          - `footer:S:WHICH`   — the WHICH footer of section S

        The bare `table:N` form is not an anchor (a whole table is a collection,
        not a single range) — use `doc.tables[N]` instead.

        Raises `AnchorNotFoundError` for unknown schemes or missing anchors.
        """
        if anchor_id == "start":
            # Bare keyword (no `kind:value` form) for the document-start
            # position. See `Document.start`.
            return self.start
        if anchor_id == "end":
            # Bare keyword for the document-end position. See `Document.end`.
            return self.end
        if not isinstance(anchor_id, str) or ":" not in anchor_id:
            raise AnchorNotFoundError("anchor", anchor_id)
        kind, _, value = anchor_id.partition(":")
        if kind == "heading":
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError("heading", anchor_id) from e
            return _IndexedHeading(self, idx)
        if kind == "para":
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError("paragraph", anchor_id) from e
            # Lazy, like heading:N — a bad index raises AnchorNotFoundError on use.
            return Paragraph(self, idx)
        if kind == "bookmark":
            return self.bookmarks[value]
        if kind == "cc":
            return self.content_controls[value]
        if kind == "table":
            parts = value.split(":")
            if len(parts) != 3:
                # `table:N` (whole table) isn't a single-range anchor.
                raise AnchorNotFoundError("table cell", anchor_id)
            try:
                t, r, c = (int(p) for p in parts)
            except ValueError as e:
                raise AnchorNotFoundError("table cell", anchor_id) from e
            return self.tables[t].cell(r, c)
        if kind == "range":
            start_str, sep, end_str = value.partition("-")
            if not sep:
                raise AnchorNotFoundError("range", anchor_id)
            try:
                start, end = int(start_str), int(end_str)
            except ValueError as e:
                raise AnchorNotFoundError("range", anchor_id) from e
            try:
                return self.range(start, end)
            except ValueError as e:
                raise AnchorNotFoundError("range", anchor_id) from e
        if kind in ("header", "footer"):
            parts = value.split(":")
            if len(parts) != 2:
                raise AnchorNotFoundError(kind, anchor_id)
            section_str, which = parts
            try:
                section_index = int(section_str)
            except ValueError as e:
                raise AnchorNotFoundError(kind, anchor_id) from e
            try:
                section = self.sections[section_index]
            except AnchorNotFoundError as e:
                raise AnchorNotFoundError(kind, anchor_id) from e
            try:
                if kind == "footer":
                    return section.footer(which)
                return section.header(which)
            except ValueError as e:
                # Unknown WHICH (primary/first/even) — surface as a missing anchor.
                raise AnchorNotFoundError(kind, anchor_id) from e
        raise AnchorNotFoundError("anchor", anchor_id)

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

    def find(
        self,
        text: str,
        *,
        scope: Anchor | None = None,
    ) -> list[dict[str, Any]]:
        """Locate every fuzzy occurrence of `text` within `scope` (or the whole doc).

        Matching is whitespace- and Unicode-normalized (NFKC, smart quotes,
        dashes, NBSP). Returns a list of `{anchor_id, start, end, text}` where
        offsets are absolute document positions and `text` is the actual
        original substring (not the normalized form).

        `anchor_id` for each match is `range:START-END`, which resolves through
        `anchor_by_id` to a `RangeAnchor` — so a hit can be fed straight back
        into `replace --anchor-id` or `comments.add`. The offsets are live,
        though, so use them before further edits shift the document.
        """
        with _com.translate_com_errors():
            rng, base = self._scope_range(scope)
            haystack = str(rng.Text or "")

        matches = _findreplace.find_matches(haystack, text)
        return [
            {
                "anchor_id": f"range:{base + m.start}-{base + m.end}",
                "start": base + m.start,
                "end": base + m.end,
                "text": m.text,
            }
            for m in matches
        ]

    def find_replace(
        self,
        find: str,
        replace: str,
        *,
        scope: Anchor | None = None,
        all: bool = False,
        occurrence: int | None = None,
    ) -> list[dict[str, Any]]:
        """Fuzzy plain-text replace. See `find()` for matching semantics.

        Args:
            find: the text to look for (fuzzy-matched).
            replace: the replacement text.
            scope: optional anchor to restrict the search to. Headings expand
                to their body section.
            all: replace every match.
            occurrence: 1-based index — replace only the Nth match.

        Raises:
            AnchorNotFoundError: zero matches (uses `kind='find'`).
            AmbiguousMatchError: more than one match and neither `all` nor
                `occurrence` was given.

        Returns the list of replacements actually applied, each
        `{anchor_id, start, end, text}` in their pre-replacement coordinates.
        """
        with _com.translate_com_errors():
            rng, base = self._scope_range(scope)
            haystack = str(rng.Text or "")

        matches = _findreplace.find_matches(haystack, find)
        if not matches:
            raise AnchorNotFoundError("find", find)

        match_payloads = [
            {
                "anchor_id": f"range:{base + m.start}-{base + m.end}",
                "start": base + m.start,
                "end": base + m.end,
                "text": m.text,
            }
            for m in matches
        ]

        if occurrence is not None:
            if occurrence < 1 or occurrence > len(matches):
                raise AnchorNotFoundError("find", f"{find} (occurrence {occurrence})")
            to_apply = [match_payloads[occurrence - 1]]
        elif all:
            to_apply = match_payloads
        elif len(matches) == 1:
            to_apply = match_payloads
        else:
            raise AmbiguousMatchError(find, match_payloads)

        with _com.translate_com_errors():
            # Apply in reverse so earlier offsets don't shift.
            for m in reversed(to_apply):
                target = self._doc.Range(m["start"], m["end"])
                target.Text = replace
        return to_apply

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

    def outline(self) -> list[dict[str, Any]]:
        """Return all heading paragraphs as `[{level, text, anchor_id}, ...]`."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    continue
                if level >= 10:
                    continue
                out.append(
                    {
                        "level": level,
                        "text": paragraph_text(para),
                        "anchor_id": f"heading:{idx}",
                    }
                )
        return out

    @contextmanager
    def edit(self, label: str) -> Iterator[EditScope]:
        """Open an atomic-undo / Selection-preserving edit scope.

        ```
        with doc.edit("Update address"):
            doc.bookmarks["Address"].set_text("…")
        ```
        """
        scope = EditScope(self._word, label)
        with scope:
            yield scope

    def go_to(self, anchor: Anchor, scroll: bool = True) -> None:
        """Move the user's Selection to the given anchor (rare — most ops preserve it).

        Does NOT open an `UndoRecord` — cursor moves don't belong on the user's
        undo stack. If you want the move to ride along with a batch of edits,
        call this inside a `doc.edit(...)` scope and the surrounding
        `UndoRecord` will still group everything together.
        """
        with _com.translate_com_errors():
            rng = anchor.com
            collapsed = self._doc.Range(int(rng.Start), int(rng.Start))
            collapsed.Select()
            if scroll:
                try:
                    self._word.com.ActiveWindow.ScrollIntoView(collapsed)
                except Exception:
                    pass


class DocumentCollection:
    """Indexable view over open documents."""

    def __init__(self, word: Word) -> None:
        self._word = word

    @property
    def _com_collection(self) -> Any:
        return self._word.com.Documents

    @property
    def active(self) -> Document:
        with _com.translate_com_errors():
            try:
                doc = self._word.com.ActiveDocument
            except Exception as e:
                raise DocumentNotFoundError("<active>") from e
        return Document(self._word, doc)

    def __getitem__(self, name: str) -> Document:
        with _com.translate_com_errors():
            for doc in self._com_collection:
                if str(doc.Name) == name:
                    return Document(self._word, doc)
        raise DocumentNotFoundError(name)

    def __iter__(self) -> Iterator[Document]:
        with _com.translate_com_errors():
            docs = list(self._com_collection)
        for d in docs:
            yield Document(self._word, d)

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._com_collection.Count)

    def list(self) -> list[dict[str, Any]]:
        """`[{name, path, is_active}, ...]` — used by `wordlive status`."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            active_name: str | None
            try:
                active_name = str(self._word.com.ActiveDocument.Name)
            except Exception:
                active_name = None
            for doc in self._com_collection:
                name = str(doc.Name)
                out.append(
                    {
                        "name": name,
                        "path": str(doc.FullName),
                        "is_active": name == active_name,
                    }
                )
        return out
