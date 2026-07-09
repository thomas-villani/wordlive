"""Insert text, paragraphs, blocks, sections, tables, breaks, fields, notes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .. import _com
from ..constants import (
    WdCollapseDirection,
    WdFieldType,
)
from ..exceptions import OpError

if TYPE_CHECKING:
    pass

from ._helpers import (
    _BREAK_TYPES,
    _CC_LIST_TYPES,
    _CC_TYPE_NAMES,
    _FIELD_TYPES,
    _coerce_named,
    _final_paragraph_empty,
    _markdown_segments,
    _normalize_cc_items,
    _normalize_table_data,
    _utf16_len,
    _validate_table_data,
    _within_table,
)

if TYPE_CHECKING:
    from ._content_controls import ContentControl
    from ._range import RangeAnchor

from ._anchor_core import AnchorCore


class AnchorInsertMixin(AnchorCore):
    """Insert text, paragraphs, blocks, sections, tables, breaks, fields, notes."""

    def insert_before(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._range()
            insert_rng = self._doc.com.Range(rng.Start, rng.Start)
            insert_rng.Text = text

    def insert_after(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._range()
            insert_rng = self._doc.com.Range(rng.End, rng.End)
            insert_rng.Text = text

    def insert_paragraph_before(self, text: str, style: str | None = None) -> None:
        """Insert a new paragraph immediately before this anchor's range.

        If `style` is given it must name a style defined in the document;
        otherwise `StyleNotFoundError` is raised before any text is inserted.
        """
        style_obj = self._doc.styles[style] if style is not None else None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            start = int(self._range().Start)
            insert_rng = doc_com.Range(start, start)
            insert_rng.Text = text + "\r"
            if style_obj is not None:
                # Word measures Range offsets in UTF-16 code units; Python's
                # len() under-counts surrogate pairs and leaves the tail unstyled.
                styled = doc_com.Range(start, start + _utf16_len(text))
                styled.Style = style_obj.com

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        """Insert a new paragraph immediately after this anchor's range.

        If `style` is given it must name a style defined in the document;
        otherwise `StyleNotFoundError` is raised before any text is inserted.

        When the anchor is (or ends at) the document's final paragraph there is
        no position *after* the terminal paragraph mark to write to — Word
        rejects `Range(end, end)` there with a "value out of range" COM error.
        In that case the new paragraph is split in just before the final mark
        instead, so appending to the end of a document — the common
        "build from scratch" case, where the only paragraph *is* the last one —
        just works.
        """
        style_obj = self._doc.styles[style] if style is not None else None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            end = int(self._range().End)
            doc_end = int(doc_com.Content.End)
            if end >= doc_end:
                # Anchor ends at the final paragraph mark. Insert "<break><text>"
                # just before that mark: the leading break terminates the
                # anchor's paragraph and `text` becomes a new final paragraph
                # (the original final mark now closes it).
                anchor_pos = max(0, doc_end - 1)
                insert_rng = doc_com.Range(anchor_pos, anchor_pos)
                insert_rng.Text = "\r" + text
                text_start = anchor_pos + 1
            else:
                insert_rng = doc_com.Range(end, end)
                insert_rng.Text = text + "\r"
                text_start = end
            if style_obj is not None:
                # Word measures Range offsets in UTF-16 code units; Python's
                # len() under-counts surrogate pairs and leaves the tail unstyled.
                styled = doc_com.Range(text_start, text_start + _utf16_len(text))
                styled.Style = style_obj.com

    def insert_block(self, items: list[Any], *, where: str = "after") -> RangeAnchor:
        """Insert a contiguous run of styled paragraphs at this anchor, atomically.

        The multi-paragraph counterpart to `insert_paragraph_after` — drop a
        whole styled section (a feature list, a set of bullets, a heading plus
        its body) in **one** op, in natural reading order. Inserting paragraphs
        one at a time forces a reverse-order dance to dodge positional-anchor
        renumbering; this places them all at a single point so order is just the
        order of `items`.

        Each item is one paragraph, given as either a plain string or a dict:

        - ``"some text"`` — sugar for ``{"text": "some text"}``.
        - ``{"text": "**Bold lead** — rest", "style": "List Bullet"}`` — `text`
          carries the tiny inline markdown (`**bold**`, `*italic*`,
          `***both***`, and `` `code` `` for a monospace run; escape a literal
          delimiter with a backslash, ``\\*`` / ``\\```), and `style` names the
          paragraph style.
        - ``{"runs": [{"text": "Bold lead", "bold": true}, {"text": " — rest"}],
          "style": "List Bullet"}`` — the structured form: each run is
          ``{text, bold?, italic?, underline?, code?, style?}`` (a per-run character
          style). Use it when markup is ambiguous or you need a run `style`.

        Returns a [`RangeAnchor`][wordlive.RangeAnchor] spanning the inserted
        block (`range:START-END`), so a follow-up op can target the whole run —
        e.g. `apply_list` it into a bulleted section, or comment on it. `where`
        is ``"after"`` (default) or ``"before"`` this anchor's range. Resolves
        every paragraph/run style up front, so an unknown style name raises
        `StyleNotFoundError` before any text is inserted. Wrap in `doc.edit(...)`
        for atomic undo. Raises `OpError` for a malformed `items` payload.
        """
        from .._runs import CODE_FONT, normalize_block_items, runs_to_text

        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        norm = normalize_block_items(items)
        # Resolve every paragraph + run style before touching the document, so a
        # bad name fails the whole block cleanly rather than leaving a partial,
        # half-styled run behind.
        para_styles = [self._doc.styles[s] if s else None for _, s in norm]
        run_styles: dict[str, Any] = {}
        for runs, _ in norm:
            for r in runs:
                if r.style and r.style not in run_styles:
                    run_styles[r.style] = self._doc.styles[r.style]
        para_texts = [runs_to_text(runs) for runs, _ in norm]
        joined = "\r".join(para_texts)
        with _com.translate_com_errors():
            doc_com = self._doc.com
            if where == "before":
                start = int(self._range().Start)
                doc_com.Range(start, start).Text = joined + "\r"
            else:
                end = int(self._range().End)
                doc_end = int(doc_com.Content.End)
                final_mark = max(0, doc_end - 1)
                if end >= final_mark:
                    # We're at the document's terminal paragraph mark — the one
                    # position you can't write *past* (`doc.end` resolves here,
                    # as does an anchor ending at the last paragraph). The old
                    # `end >= doc_end` guard missed `doc.end` (whose range ends at
                    # doc_end - 1), so the block was written *before* the final
                    # mark and merged into a non-empty last paragraph — stealing
                    # its style too. Decide by whether that final paragraph holds
                    # text, so the block neither merges nor leaves a stray empty:
                    if _final_paragraph_empty(doc_com, final_mark):
                        # Reuse the empty final paragraph as the block's last one:
                        # write without a trailing break so the existing mark
                        # closes it (no leftover empty paragraph).
                        doc_com.Range(final_mark, final_mark).Text = joined
                        start = final_mark
                    else:
                        # Open a fresh paragraph after the final one: the leading
                        # break terminates it and the block becomes new trailing
                        # paragraphs (the original final mark closes the last one).
                        doc_com.Range(final_mark, final_mark).Text = "\r" + joined
                        start = final_mark + 1
                else:
                    doc_com.Range(end, end).Text = joined + "\r"
                    start = end
            span_end = start + _utf16_len(joined)
            # Paragraph styling and run formatting both preserve text length, so
            # offsets stay valid throughout: walk the block by deterministic
            # UTF-16 offset (Word counts code units) rather than re-querying
            # Paragraphs() after each mutation.
            off = start
            for (runs, _), style_obj, ptext in zip(norm, para_styles, para_texts, strict=True):
                plen = _utf16_len(ptext)
                if style_obj is not None:
                    doc_com.Range(off, off + plen).Paragraphs(1).Range.Style = style_obj.com
                roff = off
                for run in runs:
                    rlen = _utf16_len(run.text)
                    if run.formatted():
                        sub = doc_com.Range(roff, roff + rlen)
                        if run.bold is not None:
                            sub.Bold = bool(run.bold)
                        if run.italic is not None:
                            sub.Italic = bool(run.italic)
                        if run.underline is not None:
                            sub.Underline = 1 if run.underline else 0
                        if run.code:
                            # Direct font, not a character style — `**bold**`
                            # sets Font.Bold rather than applying `Strong`, and a
                            # code span follows suit. Set before `style` so an
                            # explicit character style still wins.
                            sub.Font.Name = CODE_FONT
                        if run.style:
                            sub.Style = run_styles[run.style].com
                    roff += rlen
                off += plen + 1  # + the paragraph mark (CR)
        from ._range import RangeAnchor  # lazy: _range imports Anchor

        return RangeAnchor(self._doc, start, span_end)

    def insert_section(
        self, heading: str, body: Any, *, level: int = 1, where: str = "after"
    ) -> RangeAnchor:
        """Insert a heading plus its body in one atomic op.

        The opinionated common case over `insert_block`: a single
        ``Heading {level}`` paragraph followed by `body`, placed in reading
        order at one point. `heading` carries the same inline markdown a block
        item's `text` does (`**bold**`, `*italic*`); `body` is the `insert_block`
        items shape — a list of plain strings or ``{text|runs, style?}`` dicts
        (a bare string is sugar for a one-paragraph body). `level` is 1–9 and
        selects the built-in ``Heading {level}`` style (validated before any
        mutation; an absent style raises `StyleNotFoundError` via `insert_block`).

        Returns the section's spanning [`RangeAnchor`][wordlive.RangeAnchor]
        (`range:START-END`). Wrap in `doc.edit(...)` for atomic undo.
        """
        if not isinstance(level, int) or isinstance(level, bool) or not 1 <= level <= 9:
            raise ValueError(f"level must be an integer 1–9; got {level!r}")
        if isinstance(body, str):
            body = [body]
        if not isinstance(body, list):
            raise OpError(
                f"insert_section body must be a string or list; got {type(body).__name__}"
            )
        items = [{"text": heading, "style": f"Heading {level}"}, *body]
        return self.insert_block(items, where=where)

    def insert_markdown(self, md: str, *, where: str = "after") -> RangeAnchor:
        """Insert a constrained-Markdown block as real Word structure, atomically.

        Maps a deliberately tiny block dialect (see `_markdown`) to paragraphs,
        headings, and lists: ``#``/``##``/``###`` → `Heading 1/2/3`, ``-``/``*``
        → a bulleted list, ``1.`` → a numbered list, blank-line-separated text →
        `Normal` paragraphs, with inline ``**bold**``/``*italic*`` spans honoured.
        It is **a subset, not CommonMark** — no code fences, nested lists, block
        quotes, or tables in v1; anything unrecognised is literal paragraph text.

        The whole block is one `insert_block` (one contiguous write); each
        same-kind list run is then `apply_list`-ed over its own span, so a
        numbered list reads 1..N. `where` is ``"after"`` (default) or ``"before"``
        this anchor's range. Returns the [`RangeAnchor`][wordlive.RangeAnchor]
        spanning everything inserted. Raises `OpError` for empty markdown.
        """
        from .._markdown import parse_markdown
        from .._runs import normalize_block_items, runs_to_text

        blocks = parse_markdown(md)
        if not blocks:
            raise OpError("insert_markdown requires non-empty markdown")
        # Flatten every block into ONE insert_block (a single contiguous write —
        # chaining separate inserts would land each list before the previous
        # block's paragraph mark and merge them). Record which paragraph runs are
        # lists so we can apply_list over their spans afterwards.
        segments = _markdown_segments(blocks)
        items: list[dict[str, Any]] = []
        list_groups: list[tuple[int, int, str]] = []  # (first_para, last_para, list_type)
        for seg_items, list_type in segments:
            start_idx = len(items)
            items.extend(seg_items)
            if list_type is not None:
                list_groups.append((start_idx, len(items) - 1, list_type))
        rng = self.insert_block(items, where=where)
        if not list_groups:
            return rng
        # Recompute each paragraph's offset exactly as insert_block walks them
        # (UTF-16 text length + one CR each, from the block's start), so a list
        # group's span can be addressed without re-querying the document.
        texts = [runs_to_text(runs) for runs, _ in normalize_block_items(items)]
        offsets: list[int] = []
        off = rng.start
        for t in texts:
            offsets.append(off)
            off += _utf16_len(t) + 1
        for first, last, list_type in list_groups:
            from ._range import RangeAnchor  # lazy: _range imports Anchor

            span = RangeAnchor(self._doc, offsets[first], offsets[last] + _utf16_len(texts[last]))
            span.apply_list(list_type)
        return rng

    def insert_table(
        self,
        rows: int | None = None,
        cols: int | None = None,
        *,
        where: str = "after",
        style: str | None = None,
        data: list[Any] | None = None,
        header: bool = False,
    ) -> Any:
        """Create a `rows` × `cols` table at this anchor and return it.

        The structural counterpart to `insert_image` — it *creates* new
        document structure rather than editing existing structure. Returns the
        new [`Table`][wordlive.Table] wrapper so create → fill → read closes on
        one object; the table's 1-based document index is on `.index`.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range —
        so `doc.headings["Pricing"].insert_table(...)` drops a table just under
        a heading, and `doc.end.insert_table(...)` (i.e.
        [`Document.add_table`][wordlive.Document.add_table]) appends one.

        `style` names a table style defined in the document (e.g. ``"Table
        Grid"``); an unknown name raises `StyleNotFoundError` before anything is
        inserted. `style=None` applies the built-in ``"Table Grid"`` when it's
        available, so a table has visible borders by default rather than the
        invisible cell gridlines of a styleless table.

        `data` populates the cells at creation and can be given two ways:

        - a **row-major 2-D list** (``[[r1c1, r1c2], …]``); or
        - **records** — a list of dicts (``[{"Item": "Travel", "Cost": "$400"},
          …]``), where the first record's keys become a header row and each
          dict a body row (so `header` is forced on). The natural shape for
          tabular data an LLM already has as rows of objects.

        When `data` is given, `rows`/`cols` are **optional** — they're inferred
        from the data's shape — so the common case is just
        ``end.insert_table(data=…)``. Pass them explicitly to pad the grid
        larger than the data; `data` is validated against the final `rows` ×
        `cols` up front (`OpError` on overflow) and a short payload leaves the
        trailing cells empty. Filling at creation keeps the whole grid in one
        atomic undo and beats a `set_cell` storm. With no `data`, both `rows`
        and `cols` are required.

        `header=True` bolds the first row as a header (records imply it). Wrap
        in `doc.edit(...)` for atomic undo. Raises `ValueError` for an unknown
        `where` and `OpError` for a non-positive `rows`/`cols`, a missing
        dimension with no data to infer it from, or a bad `data` shape.
        """
        from .._tables import Table, index_of

        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        # Normalise data first so rows/cols can be inferred from its shape.
        grid: list[list[Any]] | None = None
        if data is not None:
            grid, header_from_data = _normalize_table_data(data)
            header = header or header_from_data
            if rows is None:
                rows = len(grid)
            if cols is None:
                cols = max((len(r) for r in grid), default=0)
        if rows is None or cols is None:
            raise OpError("insert_table needs rows and cols, or a data payload to infer them from")
        if isinstance(rows, bool) or not isinstance(rows, int) or rows < 1:
            raise OpError(f"table rows must be a positive integer; got {rows!r}")
        if isinstance(cols, bool) or not isinstance(cols, int) or cols < 1:
            raise OpError(f"table cols must be a positive integer; got {cols!r}")
        if grid is not None:
            _validate_table_data(grid, rows, cols)
        # Resolve the style up-front so a bad name fails before any mutation.
        if style is not None:
            style_obj = self._doc.styles[style]  # StyleNotFoundError (exit 2) if missing
        elif "Table Grid" in self._doc.styles:
            style_obj = self._doc.styles["Table Grid"]
        else:
            style_obj = None
        # New cells inherit the *paragraph* style at the insertion point — drop a
        # table right after a `Heading 2` and Word makes every cell Heading 2,
        # which renders as large heading text and pollutes the navigation
        # outline. Reset the cells to the body default (`Normal`) so a table
        # looks like a table regardless of where it was anchored. The table
        # `style` above (borders etc.) and `header` bolding still apply on top.
        normal_obj = self._doc.styles["Normal"] if "Normal" in self._doc.styles else None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            # Word's final paragraph mark is undeletable and Tables.Add needs a
            # paragraph *after* the insertion point to anchor the table; at/after
            # that mark there is none, so the add raises COM 0x80020009. Push a
            # trailing paragraph first so the table lands before it (a document
            # can't end with a table anyway — Word keeps a paragraph after one).
            doc_end = int(doc_com.Content.End)
            if pos >= doc_end - 1:
                pos = max(0, doc_end - 1)
                doc_com.Range(pos, pos).Text = "\r"
            # Word merges two tables that touch with no paragraph mark between
            # them, so a table appended at the end (or dropped next to another)
            # would silently fuse into its neighbour. Push a separator paragraph
            # onto whichever side abuts an existing table; untouched insertions
            # into ordinary text get no stray paragraph.
            if _within_table(doc_com, pos - 1, pos):
                doc_com.Range(pos, pos).Text = "\r"
                pos += 1
            if _within_table(doc_com, pos, pos + 1):
                doc_com.Range(pos, pos).Text = "\r"
            insert_rng = doc_com.Range(pos, pos)
            table_com = doc_com.Tables.Add(insert_rng, rows, cols)
            if style_obj is not None:
                table_com.Style = style_obj.com
            if normal_obj is not None:
                # Per-cell rather than table_com.Range.Style: a paragraph style
                # set on the whole table range can bleed onto the paragraph that
                # follows the table; the cell loop is contained and explicit.
                normal_com = normal_obj.com
                for r in range(1, rows + 1):
                    for c in range(1, cols + 1):
                        table_com.Cell(r, c).Range.Style = normal_com
            if grid:
                for r, row in enumerate(grid, start=1):
                    for c, val in enumerate(row, start=1):
                        table_com.Cell(r, c).Range.Text = str(val)
            if header:
                table_com.Rows(1).Range.Bold = True
            index = index_of(self._doc.com, table_com)
        return Table(self._doc, table_com, index)

    def insert_break(self, kind: str = "page", *, where: str = "after") -> None:
        """Insert a page, column, or section break at this anchor.

        The explicit one-off break — the clean alternative to appending a
        paragraph whose text is a literal form-feed. `kind` is one of:

        - ``"page"`` (default) — a manual page break (the 90% case).
        - ``"column"`` — a column break (multi-column layouts).
        - ``"section_next"`` — a section break that starts the new section on
          the next page.
        - ``"section_continuous"`` — a section break with no page break, so the
          new section flows on the same page.

        Section breaks pair with [`Document.sections`][wordlive.Document.sections]:
        each new section gets its own headers/footers and page setup. To make a
        *style* (e.g. every `Heading 1`) open a new page without a stray break
        character, prefer
        [`format_paragraph(page_break_before=True)`][wordlive.Anchor.format_paragraph]
        instead — it survives reflow.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Raises `ValueError` for an
        unknown `kind` or `where`.
        """
        if kind not in _BREAK_TYPES:
            raise ValueError(f"unknown break kind {kind!r}; expected one of {sorted(_BREAK_TYPES)}")
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        break_type = _BREAK_TYPES[kind]
        # A section break creates a *new* paragraph to carry the break, and that
        # paragraph inherits the anchor's style — drop one before a `Heading 1`
        # and Word makes the break paragraph a heading, leaving a spurious empty
        # entry in the navigation outline / TOC. Reset it to `Normal` so the break
        # is invisible to the outline. (Page/column breaks are an in-paragraph
        # character and create no such paragraph, so they need no reset.)
        is_section = kind in ("section_next", "section_continuous")
        normal_obj = (
            self._doc.styles["Normal"] if is_section and "Normal" in self._doc.styles else None
        )
        with _com.translate_com_errors():
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            insert_rng = self._doc.com.Range(pos, pos)
            insert_rng.InsertBreak(Type=int(break_type))
            if normal_obj is not None:
                # The break now occupies the position we inserted at; the
                # paragraph containing `pos` is the break paragraph.
                break_para = self._doc.com.Range(pos, pos).Paragraphs(1)
                break_para.Range.Style = normal_obj.com

    def insert_field(self, kind: str, *, text: str | None = None, where: str = "after") -> None:
        """Insert a Word field at this anchor — a self-updating value, not literal text.

        A field shows a computed value Word keeps current: a page number, the
        page count, today's date, the file name, a document property. The named
        kinds are:

        - ``"page"`` — the current page number (`{ PAGE }`).
        - ``"numpages"`` — the total page count (`{ NUMPAGES }`); pair with
          ``"page"`` for "Page X of Y".
        - ``"date"`` / ``"time"`` — the current date / time.
        - ``"filename"`` — the document's file name.
        - ``"author"`` / ``"title"`` — document-property fields.

        For anything else, ``kind="field"`` is the escape hatch: pass the raw
        field code as `text` (e.g.
        ``insert_field("field", text="REF myBookmark \\\\h")``) and Word inserts an
        empty field carrying that code.

        Page numbers belong in a header or footer — because a `HeaderFooter`
        *is* an anchor, ``doc.sections[1].footer().insert_field("page")`` works,
        and [`HeaderFooter.insert_page_number()`][wordlive.HeaderFooter] is the
        sugar for it. Newly inserted fields render once; call
        [`Document.update_fields()`][wordlive.Document] (or take a `snapshot`,
        which repaginates) to refresh them after later edits.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Bad input raises `OpError`. Wrap in `doc.edit(...)` for atomic undo.
        """
        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            wd_type = _coerce_named(kind, _FIELD_TYPES, "field kind")
            if wd_type == int(WdFieldType.EMPTY) and not text:
                raise ValueError(
                    'field kind "field" requires the raw field code via text= (e.g. text="PAGE")'
                )
            with _com.translate_com_errors():
                # Collapse a *duplicate* of the anchor's own range, so the field
                # lands in the same story — critical for header/footer anchors,
                # whose offsets are not main-document positions (a `doc.Range`
                # there would target the body instead).
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                # Positional args: the Type=/Text= keywords are dropped under
                # pywin32 late binding (same gotcha as TabStops.Add / Footnotes).
                if text is not None:
                    insert_rng.Fields.Add(insert_rng, wd_type, text)
                else:
                    insert_rng.Fields.Add(insert_rng, wd_type)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_footnote(self, text: str, *, where: str = "after") -> Any:
        """Insert a footnote at this anchor and return it as a `Footnote` anchor.

        A footnote drops a reference mark in the main text and puts `text` in the
        note body at the bottom of the page; Word auto-numbers the mark. The
        returned [`Footnote`][wordlive.Footnote] is addressed `footnote:N`, so
        `note.set_text(...)` edits the body and `note.delete()` removes the mark
        and body together. Discover existing footnotes with
        [`doc.footnotes`][wordlive.Document.footnotes].

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range —
        the side the reference mark lands on. Wrap in `doc.edit(...)` for atomic
        undo. Bad input raises `OpError`.
        """
        return self._insert_note("Footnotes", "footnote", text, where=where)

    def insert_endnote(self, text: str, *, where: str = "after") -> Any:
        """Insert an endnote at this anchor and return it as an `Endnote` anchor.

        The endnote mirror of [`insert_footnote`][wordlive.Anchor.insert_footnote]:
        the reference mark lands in the main text and `text` collects at the end
        of the document (or section). The returned
        [`Endnote`][wordlive.Endnote] is addressed `endnote:N`; discover existing
        endnotes with [`doc.endnotes`][wordlive.Document.endnotes].

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        return self._insert_note("Endnotes", "endnote", text, where=where)

    def _insert_note(self, attr: str, scheme: str, text: str, *, where: str) -> Any:
        """Shared footnote/endnote insertion (`attr` is the COM collection name)."""
        from .._notes import Endnote, Footnote, index_of_note

        cls = Footnote if scheme == "footnote" else Endnote
        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            with _com.translate_com_errors():
                rng = self._range()
                # A note's reference mark always lands in the main text story, so
                # a plain document Range at the anchor's edge is correct (unlike
                # insert_field, which can target a footer's own story).
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                coll = getattr(self._doc.com, attr)
                # Positional args: an empty Reference auto-numbers, and the
                # Reference=/Text= keywords are dropped under pywin32 late binding
                # (same gotcha as Fields.Add / TabStops.Add).
                coll.Add(insert_rng, "", text)
                index = index_of_note(coll, pos)
            return cls(self._doc, index)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_content_control(
        self,
        kind: str = "rich_text",
        *,
        title: str | None = None,
        tag: str | None = None,
        items: list[Any] | None = None,
        where: str = "wrap",
        lock_contents: bool = False,
        lock_control: bool = False,
    ) -> ContentControl:
        """Wrap (or insert) a content control at this anchor and return it.

        Content controls are the building blocks of form-like documents — a
        labelled region the user fills in. `kind` selects the type:
        ``"rich_text"`` (the default — formatted text), ``"text"`` (plain text),
        ``"picture"``, ``"combo_box"`` / ``"dropdown"`` (a pick list — pass
        `items`), ``"date"``, ``"checkbox"`` (Word 2013+), ``"building_block"``,
        ``"group"``, or ``"repeating_section"`` (Word 2013+).

        `where` is ``"wrap"`` (default — the control surrounds this anchor's
        existing range, so a `range:START-END` from `find` wraps that phrase) or
        ``"before"`` / ``"after"`` (insert a fresh empty control at the anchor's
        start / end). Set `title` and/or `tag` to give the control a name: a
        titled control is addressable later as `cc:TITLE` (falling back to the
        tag) and shows a label in Word's UI. `items` populates a combo box or
        dropdown — each is a string, or a `{"text": ..., "value": ...}` dict.
        `lock_contents` stops the user editing the value; `lock_control` stops
        them deleting the control.

        Returns the new [`ContentControl`][wordlive.ContentControl] (usable even
        when unnamed — it caches the live control). Wrap in `doc.edit(...)` for
        atomic undo. Bad input raises `OpError`.
        """
        try:
            if where not in ("wrap", "before", "after"):
                raise ValueError(f"where must be 'wrap', 'before', or 'after'; got {where!r}")
            try:
                cc_type = _CC_TYPE_NAMES[str(kind).lower()]
            except (KeyError, AttributeError) as e:
                raise ValueError(
                    f"unknown content control kind {kind!r}; "
                    f"expected one of {sorted(_CC_TYPE_NAMES)}"
                ) from e
            if items is not None and cc_type not in _CC_LIST_TYPES:
                raise ValueError("items is only valid for a 'combo_box' or 'dropdown' control")
            with _com.translate_com_errors():
                target = self._range().Duplicate
                if where != "wrap":
                    target.Collapse(
                        int(
                            WdCollapseDirection.START
                            if where == "before"
                            else WdCollapseDirection.END
                        )
                    )
                # Positional args (Type, Range) for late-binding safety.
                cc = self._doc.com.ContentControls.Add(int(cc_type), target)
                if title is not None:
                    cc.Title = str(title)
                if tag is not None:
                    cc.Tag = str(tag)
                if lock_contents:
                    cc.LockContents = True
                if lock_control:
                    cc.LockContentControl = True
                if items:
                    entries = cc.DropdownListEntries
                    for entry_text, value in _normalize_cc_items(items):
                        entries.Add(entry_text, value)  # positional (Text, Value)
            name = title or tag or ""
            from ._content_controls import ContentControl  # lazy: _content_controls imports Anchor

            return ContentControl(self._doc, name, com=cc)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
