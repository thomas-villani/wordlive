"""Anchor types — semantic handles for ranges inside a Word document.

Anchors target a `Range`, never the live `Selection`. Each public mutation
goes through the COM error translator. Operations are intentionally small;
they compose with `Document.edit()` for atomic-undo behaviour.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com, _images, _lists
from .constants import (
    MsoTriState,
    WdBreakType,
    WdInformation,
    WdNumberType,
    WdParagraphAlignment,
    WdWrapType,
)
from .exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from pathlib import Path

    from ._document import Document
    from ._snapshot import Snapshot


_ALIGNMENT_NAMES = {
    "left": WdParagraphAlignment.LEFT,
    "center": WdParagraphAlignment.CENTER,
    "centre": WdParagraphAlignment.CENTER,
    "right": WdParagraphAlignment.RIGHT,
    "justify": WdParagraphAlignment.JUSTIFY,
}


def _utf16_len(s: str) -> int:
    """Length of `s` in UTF-16 code units — Word's native character count.

    Python's `len()` counts code points, so astral-plane characters (emoji,
    historic scripts) count as 1. Word counts UTF-16 code units, so the same
    character counts as 2. Use this whenever the result is fed back into a
    Word `Range(start, end)` after a `Range.Text = ...` assignment.
    """
    return len(s.encode("utf-16-le")) // 2


def range_text(rng: Any) -> str:
    """Read a COM range's text with inline shapes surfaced as ``[image]`` tokens.

    Word represents each inline shape (embedded picture / OLE object) as a single
    placeholder character in the text stream. That character is *not* a reserved
    control code — it varies by build and is indistinguishable by value from real
    text (a forward slash, on some Word versions) — so a naive string replace
    would clobber genuine characters. Instead we locate the shapes via the
    ``InlineShapes`` collection and swap only the character at each shape's own
    position, leaving real text untouched. A range with no inline shapes returns
    its raw text unchanged.
    """
    raw = str(rng.Text or "")
    try:
        shapes = rng.InlineShapes
        count = int(shapes.Count)
        if count <= 0:
            return raw
        base = int(rng.Start)
        offsets = sorted({int(shapes.Item(i).Range.Start) - base for i in range(1, count + 1)})
    except Exception:
        # If the shape geometry can't be read, fall back to the raw text rather
        # than risk mangling it — a phantom char is better than a crash.
        return raw
    chars = list(raw)
    for off in reversed(offsets):
        if 0 <= off < len(chars):
            chars[off] = "[image]"
    return "".join(chars)


def _coerce_alignment(value: Any) -> int:
    if isinstance(value, WdParagraphAlignment):
        return int(value)
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    if isinstance(value, str):
        try:
            return int(_ALIGNMENT_NAMES[value.lower()])
        except KeyError:
            raise ValueError(
                f"unknown alignment {value!r}; expected one of {sorted(set(_ALIGNMENT_NAMES))}"
            ) from None
    raise TypeError(
        f"alignment must be WdParagraphAlignment, int, or str; got {type(value).__name__}"
    )


# Floating wrap keywords -> WdWrapType. "inline" and "auto" are handled
# specially by insert_image and are not in this map.
_WRAP_NAMES: dict[str, WdWrapType] = {
    "square": WdWrapType.SQUARE,
    "tight": WdWrapType.TIGHT,
    "through": WdWrapType.THROUGH,
    "top-bottom": WdWrapType.TOP_BOTTOM,
    "front": WdWrapType.FRONT,
    "behind": WdWrapType.BEHIND,
}
_WRAP_VALUES: frozenset[str] = frozenset({"inline", "auto", *_WRAP_NAMES})


# Break keywords -> WdBreakType. `insert_break(kind=...)` accepts exactly these.
_BREAK_TYPES: dict[str, WdBreakType] = {
    "page": WdBreakType.PAGE,
    "column": WdBreakType.COLUMN,
    "section_next": WdBreakType.SECTION_NEXT_PAGE,
    "section_continuous": WdBreakType.SECTION_CONTINUOUS,
}


def _resolve_wrap(wrap: str, inline_shape: Any, insert_rng: Any) -> WdWrapType:
    """Resolve a wrap keyword to a concrete `WdWrapType` for a floating shape.

    `"auto"` picks Square when the image is at most half the section's usable
    text width (`PageWidth - LeftMargin - RightMargin`), else top-and-bottom.
    """
    if wrap != "auto":
        return _WRAP_NAMES[wrap]
    ps = insert_rng.PageSetup
    usable = float(ps.PageWidth) - float(ps.LeftMargin) - float(ps.RightMargin)
    if float(inline_shape.Width) <= usable / 2:
        return WdWrapType.SQUARE
    return WdWrapType.TOP_BOTTOM


def _validate_table_data(data: Any, rows: int, cols: int) -> None:
    """Check a row-major `data` payload fits a `rows` × `cols` grid.

    Raised before any COM call so a bad shape is a clean `OpError` (exit 1)
    rather than a "subscript out of range" deep inside Word. Underfilling is
    allowed — fewer rows, or short rows — and leaves the trailing cells empty
    (matching `add_row`'s leniency); only *overflowing* the declared grid is an
    error, since that's the case that would otherwise blow up mid-insert.
    """
    if not isinstance(data, list):
        raise OpError(f"table data must be a list of rows; got {type(data).__name__}")
    if len(data) > rows:
        raise OpError(f"table data has {len(data)} rows but the table has only {rows}")
    for i, row in enumerate(data, start=1):
        if not isinstance(row, list):
            raise OpError(f"table data row {i} must be a list; got {type(row).__name__}")
        if len(row) > cols:
            raise OpError(
                f"table data row {i} has {len(row)} cells but the table has only {cols} column(s)"
            )


def _within_table(doc_com: Any, start: int, end: int) -> bool:
    """Whether the `[start, end)` span sits inside a table.

    Used to detect when a new table's insertion point abuts an existing one —
    Word silently *merges* two tables with no paragraph mark between them, so
    `insert_table` drops a separator paragraph on any abutting side. A negative
    `start` (before the document) or a probe Word rejects reads as "not in a
    table".
    """
    if start < 0:
        return False
    try:
        return bool(doc_com.Range(start, end).Information(int(WdInformation.WITH_IN_TABLE)))
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Base
# ---------------------------------------------------------------------------


class Anchor(ABC):
    """Abstract base — subclasses know how to materialise their COM Range.

    Concrete subclasses must implement `_range()` and `set_text()`. Other
    operations (`text`, `insert_before`, `insert_after`, `delete`,
    `apply_style`, `format_paragraph`) are derived and inherited as-is.
    """

    kind: str = "anchor"
    name: str = ""

    def __init__(self, doc: Document, name: str) -> None:
        self._doc = doc
        self.name = name

    @property
    def com(self) -> Any:
        """Raw COM range. Subclasses override."""
        return self._range()

    @abstractmethod
    def _range(self) -> Any:
        """Return the COM Range that this anchor refers to. Must be overridden."""

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return range_text(self._range())

    @property
    @abstractmethod
    def anchor_id(self) -> str:
        """Stable string identifier for this anchor (e.g. `bookmark:Address`).

        Each anchor kind has its own scheme (`bookmark:`, `cc:`, `heading:`),
        so subclasses must declare theirs explicitly — no useful default
        exists at this level.
        """

    @abstractmethod
    def set_text(self, text: str) -> None:
        """Replace the anchor's text in place. Must be overridden."""

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

    def insert_image(
        self,
        image: str | Path | bytes,
        *,
        wrap: str,
        where: str = "after",
        block: bool = False,
        width: float | None = None,
        height: float | None = None,
        alt_text: str | None = None,
        lock_aspect: bool = True,
    ) -> None:
        """Insert an image at this anchor (atomic-undo when inside `doc.edit()`).

        `image` is a file path, raw image bytes, or a base64 string — a `str`
        is treated as a path when it names an existing file, otherwise as
        base64. Word embeds the picture (`SaveWithDocument=True`) and
        auto-detects its natural size, so `width`/`height` (points) are optional
        overrides. `alt_text` sets the image's accessibility text.

        `wrap` is required — there is no default — so layout intent is always
        explicit:

        - ``"inline"`` keeps the image in the text flow (an `InlineShape`).
        - ``"auto"`` floats it: Square when its width is at most half the
          section's usable text width, else top-and-bottom.
        - ``"square" | "tight" | "through" | "top-bottom" | "front" | "behind"``
          floats it with that wrap type.

        `where` is ``"after"`` (default) or ``"before"`` the anchor's range.

        `block` places the image in its own new paragraph (reset to ``Normal``)
        rather than embedding it in the anchor's text run — so
        ``heading.insert_image(..., wrap="inline", where="before", block=True)``
        drops the image on its own line *above* the heading instead of joining
        the heading text. Without it, an inline image anchored at a heading lands
        mid-run and the heading text trails it on the same line.

        Raises `ImageSourceError` for a missing/unreadable/invalid image and
        `ValueError` for an unknown `wrap` or `where`.
        """
        if wrap not in _WRAP_VALUES:
            raise ValueError(f"unknown wrap {wrap!r}; expected one of {sorted(_WRAP_VALUES)}")
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        # New paragraphs inherit the anchor's style — a block image above a
        # heading would otherwise become a heading-styled (and outline-polluting)
        # paragraph. Reset it to the body default, like insert_table does.
        normal_obj = self._doc.styles["Normal"] if block and "Normal" in self._doc.styles else None
        with _images.image_on_disk(image) as disk_path:
            with _com.translate_com_errors():
                doc_com = self._doc.com
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                if block:
                    # Open a fresh paragraph at the insertion point and target it,
                    # so the image sits on its own line instead of in the run.
                    doc_com.Range(pos, pos).Text = "\r"
                    if normal_obj is not None:
                        doc_com.Range(pos, pos).Paragraphs(1).Range.Style = normal_obj.com
                insert_rng = doc_com.Range(pos, pos)
                ish = insert_rng.InlineShapes.AddPicture(
                    FileName=disk_path,
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Range=insert_rng,
                )
                ish.LockAspectRatio = int(MsoTriState.TRUE if lock_aspect else MsoTriState.FALSE)
                if width is not None:
                    ish.Width = float(width)
                if height is not None:
                    ish.Height = float(height)
                if alt_text is not None:
                    ish.AlternativeText = alt_text
                if wrap == "inline":
                    return
                wrap_type = _resolve_wrap(wrap, ish, insert_rng)
                shape = ish.ConvertToShape()
                shape.WrapFormat.Type = int(wrap_type)
                if alt_text is not None:
                    # AlternativeText doesn't always survive the conversion.
                    shape.AlternativeText = alt_text

    def insert_table(
        self,
        rows: int,
        cols: int,
        *,
        where: str = "after",
        style: str | None = None,
        data: list[list[Any]] | None = None,
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

        `data` populates the cells at creation: a row-major 2-D list
        (``[[r1c1, r1c2], …]``), validated against `rows` × `cols` up front
        (`OpError` on overflow). A short or partial `data` leaves the remaining
        cells empty. Filling at creation keeps the whole grid in one atomic
        undo and beats a `set_cell` storm.

        `header=True` bolds the first row as a header. Wrap in `doc.edit(...)`
        for atomic undo. Raises `ValueError` for an unknown `where` and
        `OpError` for a non-positive `rows`/`cols` or a bad `data` shape.
        """
        from ._tables import Table, index_of

        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        if isinstance(rows, bool) or not isinstance(rows, int) or rows < 1:
            raise OpError(f"table rows must be a positive integer; got {rows!r}")
        if isinstance(cols, bool) or not isinstance(cols, int) or cols < 1:
            raise OpError(f"table cols must be a positive integer; got {cols!r}")
        if data is not None:
            _validate_table_data(data, rows, cols)
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
            if data:
                for r, row in enumerate(data, start=1):
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

    def snapshot(self, out: str | Path | None = None, *, dpi: int = 150) -> list[Snapshot]:
        """Render the page(s) this anchor sits on to PNG — let a model *see* it.

        A heading expands to its whole section; any other anchor renders the
        page(s) its range spans. Returns a list of
        [`Snapshot`][wordlive.Snapshot] (one per page); pass `out` to also write
        the image(s) to disk. Sugar for
        [`Document.snapshot_anchor`][wordlive.Document.snapshot_anchor]; see it
        for the full semantics. Requires the `snapshot` extra (PyMuPDF).
        """
        return self._doc.snapshot_anchor(self, out, dpi=dpi)

    def delete(self) -> None:
        with _com.translate_com_errors():
            self._range().Delete()

    def apply_style(self, name: str) -> None:
        """Apply the named paragraph or character style to this anchor's range.

        Word selects paragraph- vs. character-style behaviour from the style's
        own `Type`; we don't model that distinction. Raises `StyleNotFoundError`
        if the style isn't defined in the document.
        """
        style = self._doc.styles[name]  # raises StyleNotFoundError if missing
        with _com.translate_com_errors():
            self._range().Style = style.com

    def format_paragraph(
        self,
        *,
        alignment: Any = None,
        left_indent: float | None = None,
        right_indent: float | None = None,
        first_line_indent: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        page_break_before: bool | None = None,
    ) -> None:
        """Set paragraph-formatting properties on this anchor's range.

        All kwargs are optional; only the ones explicitly passed are written.
        Indent and spacing values are in points (Word's native unit for
        `ParagraphFormat.LeftIndent` etc.). `alignment` accepts a
        `WdParagraphAlignment` enum, its int value, or a string
        (`"left"`/`"center"`/`"right"`/`"justify"`).

        `page_break_before=True` forces the paragraph to begin on a new page —
        the *clean* way to page-break (e.g. apply it to every `Heading 1`): it's
        a paragraph property that survives reflow and leaves no stray break
        character, unlike [`insert_break`][wordlive.Anchor.insert_break].
        `False` clears the property.
        """
        with _com.translate_com_errors():
            pf = self._range().ParagraphFormat
            if alignment is not None:
                pf.Alignment = _coerce_alignment(alignment)
            if left_indent is not None:
                pf.LeftIndent = float(left_indent)
            if right_indent is not None:
                pf.RightIndent = float(right_indent)
            if first_line_indent is not None:
                pf.FirstLineIndent = float(first_line_indent)
            if space_before is not None:
                pf.SpaceBefore = float(space_before)
            if space_after is not None:
                pf.SpaceAfter = float(space_after)
            if page_break_before is not None:
                pf.PageBreakBefore = bool(page_break_before)

    def apply_list(self, list_type: str = "bulleted", *, continue_previous: bool = False) -> None:
        """Turn this anchor's paragraphs into a list.

        `list_type` is `"bulleted"`, `"numbered"`, or `"outline"` (the three
        `ListGalleries`). By default numbering starts fresh at 1; pass
        `continue_previous=True` to continue from a list immediately above.
        Raises `ValueError` for an unknown `list_type`.
        """
        gallery_type = _lists.gallery_for(list_type)  # ValueError before any mutation
        with _com.translate_com_errors():
            _lists.apply_list_template(
                self._range(), gallery_type, continue_previous=continue_previous
            )

    def remove_list(self) -> None:
        """Strip list formatting (bullets / numbers) from this anchor's paragraphs."""
        with _com.translate_com_errors():
            self._range().ListFormat.RemoveNumbers(NumberType=int(WdNumberType.ALL_NUMBERS))

    def list_info(self) -> dict[str, Any]:
        """Describe the list this anchor sits in: `{type, level, number, string}`.

        `type` is `"none"` when there's no list formatting, otherwise one of
        `"bulleted"`, `"numbered"`, `"outline"`, `"number-only"`, or `"mixed"`.
        `number` is the first paragraph's value, `string` its rendered marker.
        """
        with _com.translate_com_errors():
            return _lists.read_list_info(self._range())

    def restart_numbering(self) -> None:
        """Restart this list's numbering at 1.

        Re-applies the range's current list template with "continue previous"
        off. Raises `ValueError` if the range isn't part of a list.
        """
        with _com.translate_com_errors():
            _lists.restart_numbering(self._range())

    def indent_list(self) -> None:
        """Demote this list item one level (e.g. level 1 -> 2)."""
        with _com.translate_com_errors():
            self._range().ListFormat.ListIndent()

    def outdent_list(self) -> None:
        """Promote this list item one level (e.g. level 2 -> 1)."""
        with _com.translate_com_errors():
            self._range().ListFormat.ListOutdent()

    def __repr__(self) -> str:
        return f"<{type(self).__name__} {self.name!r}>"


# ---------------------------------------------------------------------------
# Arbitrary ranges
# ---------------------------------------------------------------------------


class RangeAnchor(Anchor):
    """An anchor over an arbitrary character range — `doc.range(start, end)`.

    Unlike bookmarks/headings/cells, a range anchor names nothing in the
    document: it's a pair of absolute character offsets (UTF-16 code units, the
    same coordinates Word's `Document.Range(start, end)` uses and that
    `Document.find()` emits as `range:START-END`). It's the generic target when
    no named anchor exists — feed a `find()` hit straight into a `replace`, or
    drop a comment on an offset span.

    The anchor is ephemeral: offsets resolve live against the document on each
    access, so an edit elsewhere that shifts the text can leave it pointing at
    the wrong span. Resolve, act, discard. `set_text` keeps the anchor's own
    `end` in sync with the replacement so chained ops on the same instance stay
    consistent.
    """

    kind = "range"

    def __init__(self, doc: Document, start: int, end: int) -> None:
        start = int(start)
        end = int(end)
        if start < 0 or end < start:
            raise ValueError(f"invalid range offsets: start={start}, end={end}")
        super().__init__(doc, name=f"range:{start}-{end}")
        self._start = start
        self._end = end

    @property
    def start(self) -> int:
        return self._start

    @property
    def end(self) -> int:
        return self._end

    @property
    def anchor_id(self) -> str:
        return f"range:{self._start}-{self._end}"

    def _range(self) -> Any:
        return self._doc.com.Range(self._start, self._end)

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._doc.com.Range(self._start, self._end)
            rng.Text = text
        # A Range.Text assignment resizes the span; keep our end in sync so a
        # follow-up read/op on the same anchor sees the replacement rather than
        # the stale coordinates. Word counts UTF-16 code units, not code points.
        self._end = self._start + _utf16_len(text)


# ---------------------------------------------------------------------------
# Start / end of document
# ---------------------------------------------------------------------------


class StartAnchor(Anchor):
    """A zero-width anchor at the very start of the document body — `doc.start`.

    The mirror of [`EndAnchor`][wordlive.EndAnchor]: the insertion point before
    the first paragraph. `doc.start` returns it and `anchor_by_id("start")`
    resolves it, so "prepend to the document" composes with the usual verbs and
    the CLI `--anchor-id` plumbing.

    Only the *prepend* direction is meaningful at a single start-point, so every
    insert verb lands text at the start: `insert_paragraph_before` /
    `insert_paragraph_after` add a new first paragraph (delegating to
    [`Document.prepend_paragraph`][wordlive.Document.prepend_paragraph]), and
    `insert_before` / `insert_after` / `set_text` prepend inline (delegating to
    [`Document.prepend`][wordlive.Document.prepend]). `text` is always empty and
    `delete()` is a no-op. `insert_image` and `apply_style` are inherited: they
    resolve to the collapsed start position.
    """

    kind = "start"

    def __init__(self, doc: Document) -> None:
        super().__init__(doc, name="start")

    @property
    def anchor_id(self) -> str:
        return "start"

    def _range(self) -> Any:
        # Collapsed at offset 0 — the position Document.prepend* writes to.
        return self._doc.com.Range(0, 0)

    def set_text(self, text: str) -> None:
        # Nothing to replace at the start-point — prepend instead.
        self._doc.prepend(text)

    def insert_after(self, text: str) -> None:
        self._doc.prepend(text)

    def insert_before(self, text: str) -> None:
        # A single start-point has no distinct "after"; prepending is the only
        # sensible reading, and it keeps `--anchor-id start` honest either way.
        self._doc.prepend(text)

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        self._doc.prepend_paragraph(text, style=style)

    def insert_paragraph_before(self, text: str, style: str | None = None) -> None:
        self._doc.prepend_paragraph(text, style=style)


class EndAnchor(Anchor):
    """A zero-width anchor at the very end of the document body — `doc.end`.

    The one position no content names: the insertion point past the last
    paragraph. `doc.end` returns it and `anchor_by_id("end")` resolves it, so
    "append to the document" composes with the same verbs and the same CLI
    `--anchor-id` plumbing as every other anchor — no `.com` drop needed.

    Only the *append* direction is meaningful at a single end-point, so every
    insert verb lands text at the end: `insert_paragraph_after` /
    `insert_paragraph_before` add a new final paragraph (delegating to
    [`Document.append_paragraph`][wordlive.Document.append_paragraph]), and
    `insert_after` / `insert_before` / `set_text` append inline (delegating to
    [`Document.append`][wordlive.Document.append]). `text` is always empty and
    `delete()` is a no-op — there is no content here to read or remove.
    `insert_image` and `apply_style` are inherited: they resolve to the
    collapsed end position, so an image lands at the end and a style falls on
    the final paragraph.
    """

    kind = "end"

    def __init__(self, doc: Document) -> None:
        super().__init__(doc, name="end")

    @property
    def anchor_id(self) -> str:
        return "end"

    def _range(self) -> Any:
        # Collapsed just before the final paragraph mark — the position
        # Document.append* writes to, and a safe target for the inherited verbs
        # (a zero-width span reads "" and deletes nothing).
        with _com.translate_com_errors():
            end = int(self._doc.com.Content.End)
        pos = max(0, end - 1)
        return self._doc.com.Range(pos, pos)

    def set_text(self, text: str) -> None:
        # Nothing to replace at the end-point — append instead.
        self._doc.append(text)

    def insert_after(self, text: str) -> None:
        self._doc.append(text)

    def insert_before(self, text: str) -> None:
        # A single end-point has no distinct "before"; appending is the only
        # sensible reading, and it keeps `--anchor-id end` honest either way.
        self._doc.append(text)

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        self._doc.append_paragraph(text, style=style)

    def insert_paragraph_before(self, text: str, style: str | None = None) -> None:
        self._doc.append_paragraph(text, style=style)


# ---------------------------------------------------------------------------
# Paragraphs
# ---------------------------------------------------------------------------


class Paragraph(Anchor):
    """A paragraph located by 1-based index over `doc.Paragraphs`.

    `para:N` addresses *any* paragraph — body text, headings, list items alike.
    `heading:N` is the same index space narrowed to heading paragraphs, so
    `para:5` and `heading:5` resolve to the same paragraph when paragraph 5 is a
    heading. A `Paragraph` inherits every anchor verb (`set_text`, `apply_style`,
    `format_paragraph`, `apply_list`, `insert_paragraph_before/after`, …).
    """

    kind = "paragraph"

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"para:{index}")
        self._index = index

    @property
    def anchor_id(self) -> str:
        return f"para:{self._index}"

    @property
    def index(self) -> int:
        return self._index

    def _paragraph(self) -> Any:
        for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
            if idx == self._index:
                # Keep .name informative for repr / error messages.
                self.name = paragraph_text(para) or self.name
                return para
        raise AnchorNotFoundError("paragraph", f"para:{self._index}")

    @property
    def level(self) -> int:
        with _com.translate_com_errors():
            return int(self._paragraph().OutlineLevel)

    @property
    def is_heading(self) -> bool:
        return self.level < 10

    def _range(self) -> Any:
        return self._paragraph().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return paragraph_text(self._paragraph())

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            para_range = self._paragraph().Range
            start = int(para_range.Start)
            end = int(para_range.End)
            # Preserve the trailing paragraph mark so the paragraph isn't merged
            # with the next one (same approach as Heading.set_text).
            inner = self._doc.com.Range(start, max(start, end - 1))
            inner.Text = text


class ParagraphCollection:
    """Indexable, iterable view over every paragraph in the document.

    Unlike `headings`, this includes body paragraphs and list items, not just
    heading paragraphs. Index by 1-based position (`doc.paragraphs[2]`); iterate
    for a `Paragraph` per paragraph. `list()` emits each paragraph's `start` /
    `end` offsets, so a body paragraph can be turned into a `range:START-END`
    insertion point for mid-paragraph edits.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def _count(self) -> int:
        with _com.translate_com_errors():
            return sum(1 for _ in self._doc.com.Paragraphs)

    def __len__(self) -> int:
        return self._count()

    def __getitem__(self, index: int) -> Paragraph:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"paragraph index must be int, got {type(index).__name__}")
        if index < 1 or index > self._count():
            raise AnchorNotFoundError("paragraph", f"para:{index}")
        return Paragraph(self._doc, index)

    def __iter__(self) -> Iterator[Paragraph]:
        with _com.translate_com_errors():
            count = sum(1 for _ in self._doc.com.Paragraphs)
        for idx in range(1, count + 1):
            yield Paragraph(self._doc, idx)

    def at(self, offset: int) -> Paragraph | None:
        """Return the paragraph whose range contains `offset`, or None.

        Used to map a character offset (e.g. the cursor position) back to a
        `para:N` anchor.
        """
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                rng = para.Range
                if int(rng.Start) <= offset < int(rng.End):
                    return Paragraph(self._doc, idx)
        return None

    def list(self) -> list[dict[str, Any]]:
        """Every paragraph as `[{index, anchor_id, level, is_heading, start, end, text}, ...]`."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    level = 10
                rng = para.Range
                out.append(
                    {
                        "index": idx,
                        "anchor_id": f"para:{idx}",
                        "level": level,
                        "is_heading": level < 10,
                        "start": int(rng.Start),
                        "end": int(rng.End),
                        "text": paragraph_text(para),
                    }
                )
        return out


# ---------------------------------------------------------------------------
# Bookmarks
# ---------------------------------------------------------------------------


class Bookmark(Anchor):
    kind = "bookmark"

    @property
    def anchor_id(self) -> str:
        return f"bookmark:{self.name}"

    def _range(self) -> Any:
        doc_com = self._doc.com
        if not doc_com.Bookmarks.Exists(self.name):
            raise AnchorNotFoundError("bookmark", self.name)
        return doc_com.Bookmarks(self.name).Range

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            doc_com = self._doc.com
            if not doc_com.Bookmarks.Exists(self.name):
                raise AnchorNotFoundError("bookmark", self.name)
            rng = doc_com.Bookmarks(self.name).Range
            start = int(rng.Start)
            rng.Text = text
            # Setting Range.Text deletes the bookmark; re-add covering the new content.
            # Word measures Range offsets in UTF-16 code units, not Python code points.
            new_end = start + _utf16_len(text)
            new_rng = doc_com.Range(start, new_end)
            doc_com.Bookmarks.Add(Name=self.name, Range=new_rng)


def _is_user_bookmark(name: str) -> bool:
    """Word auto-creates internal bookmarks for TOC entries, cross-references,
    and form-field anchors — all of them named with a leading underscore. Those
    are noise for the user-facing `list()` / iteration paths; agents addressing
    them by exact name (via `bookmarks[name]`) still work.
    """
    return not name.startswith("_")


class BookmarkCollection:
    """Indexable view over a document's bookmarks.

    `list()` and iteration return only user-visible bookmarks. Word's hidden
    bookmarks (`_Toc...`, `_Ref...`, etc.) are filtered out by default; address
    them by their exact name through `bookmarks[name]` if you need them.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __getitem__(self, name: str) -> Bookmark:
        with _com.translate_com_errors():
            if not self._doc.com.Bookmarks.Exists(name):
                raise AnchorNotFoundError("bookmark", name)
        return Bookmark(self._doc, name)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        with _com.translate_com_errors():
            return bool(self._doc.com.Bookmarks.Exists(name))

    def list(self, *, include_hidden: bool = False) -> list[str]:
        """Names of every user-visible bookmark in document order.

        Set `include_hidden=True` to also return Word's internal bookmarks
        (TOC entries, cross-references, etc.) whose names start with `_`.
        """
        with _com.translate_com_errors():
            names = [str(bm.Name) for bm in self._doc.com.Bookmarks]
        if include_hidden:
            return names
        return [n for n in names if _is_user_bookmark(n)]

    def __iter__(self) -> Iterator[Bookmark]:
        for name in self.list():
            yield Bookmark(self._doc, name)


# ---------------------------------------------------------------------------
# Content controls
# ---------------------------------------------------------------------------


def _cc_by_name(doc_com: Any, name: str) -> Any | None:
    """Find a content control by its Title (Tag falls back). Returns None if missing.

    Reject empty `name` explicitly — many content controls have neither a
    Title nor a Tag, and the naive `cc.Title or "" == ""` test would match
    the first such control. Callers asking for `""` get `None` instead.
    """
    if not name:
        return None
    for cc in doc_com.ContentControls:
        if str(cc.Title or "") == name or str(cc.Tag or "") == name:
            return cc
    return None


class ContentControl(Anchor):
    kind = "content_control"

    @property
    def anchor_id(self) -> str:
        return f"cc:{self.name}"

    def _cc(self) -> Any:
        cc = _cc_by_name(self._doc.com, self.name)
        if cc is None:
            raise AnchorNotFoundError("content_control", self.name)
        return cc

    def _range(self) -> Any:
        return self._cc().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            cc = self._cc()
            return range_text(cc.Range)

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            cc = self._cc()
            cc.Range.Text = text


class ContentControlCollection:
    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __getitem__(self, name: str) -> ContentControl:
        with _com.translate_com_errors():
            if _cc_by_name(self._doc.com, name) is None:
                raise AnchorNotFoundError("content_control", name)
        return ContentControl(self._doc, name)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        with _com.translate_com_errors():
            return _cc_by_name(self._doc.com, name) is not None

    def list(self) -> list[str]:
        with _com.translate_com_errors():
            names: list[str] = []
            for cc in self._doc.com.ContentControls:
                names.append(str(cc.Title or cc.Tag or ""))
            return names

    def __iter__(self) -> Iterator[ContentControl]:
        for name in self.list():
            if name:
                yield ContentControl(self._doc, name)


# ---------------------------------------------------------------------------
# Headings
# ---------------------------------------------------------------------------


def paragraph_text(para: Any) -> str:
    """Heading text minus the trailing paragraph mark, inline shapes tokenized."""
    return range_text(para.Range).rstrip("\r\n\x07")


def _find_heading_paragraph(doc_com: Any, name: str) -> tuple[Any, int] | None:
    """Locate a heading paragraph by visible text. Returns (Paragraph, 1-based index)."""
    for idx, para in enumerate(doc_com.Paragraphs, start=1):
        try:
            level = int(para.OutlineLevel)
        except Exception:
            continue
        if level >= 10:  # WdOutlineLevel: 1-9 are headings; 10 is body text
            continue
        if paragraph_text(para) == name:
            return para, idx
    return None


def _section_range(doc_com: Any, target_para: Any, target_level: int) -> Any:
    """COM Range from the end of `target_para` to the next paragraph whose
    OutlineLevel is a heading and `<= target_level` — or to the end of the
    document's last paragraph if no such boundary exists.
    """
    paragraphs = list(doc_com.Paragraphs)
    target_start = int(target_para.Range.Start)

    idx: int | None = None
    for i, p in enumerate(paragraphs):
        try:
            if int(p.Range.Start) == target_start:
                idx = i
                break
        except Exception:
            continue
    if idx is None:
        end = int(target_para.Range.End)
        return doc_com.Range(end, end)

    section_start = int(target_para.Range.End)
    section_end: int | None = None
    for p in paragraphs[idx + 1 :]:
        try:
            lvl = int(p.OutlineLevel)
        except Exception:
            continue
        if lvl < 10 and lvl <= target_level:
            section_end = int(p.Range.Start)
            break
    if section_end is None:
        try:
            section_end = int(paragraphs[-1].Range.End)
        except Exception:
            section_end = section_start
    return doc_com.Range(section_start, section_end)


class Heading(Anchor):
    kind = "heading"

    def _paragraph(self) -> Any:
        found = _find_heading_paragraph(self._doc.com, self.name)
        if found is None:
            raise AnchorNotFoundError("heading", self.name)
        return found[0]

    def _paragraph_and_index(self) -> tuple[Any, int]:
        """Default lookup goes by visible text; subclasses can override."""
        found = _find_heading_paragraph(self._doc.com, self.name)
        if found is None:
            raise AnchorNotFoundError("heading", self.name)
        return found

    @property
    def anchor_id(self) -> str:
        with _com.translate_com_errors():
            _, idx = self._paragraph_and_index()
        return f"heading:{idx}"

    @property
    def level(self) -> int:
        with _com.translate_com_errors():
            return int(self._paragraph().OutlineLevel)

    def section_range(self) -> Any:
        """COM Range covering the body under this heading.

        Spans from the end of the heading paragraph to the start of the next
        heading whose level is `<=` this one's (or to the end of the document
        if no such heading exists). Excludes the heading paragraph itself.
        """
        with _com.translate_com_errors():
            para = self._paragraph()
            level = int(para.OutlineLevel)
            return _section_range(self._doc.com, para, level)

    def section_text(self) -> str:
        """Plain text of the body under this heading."""
        with _com.translate_com_errors():
            return str(self.section_range().Text or "")

    def _range(self) -> Any:
        return self._paragraph().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return paragraph_text(self._paragraph())

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            para_range = self._paragraph().Range
            start = int(para_range.Start)
            end = int(para_range.End)
            # Preserve the trailing paragraph mark.
            inner = self._doc.com.Range(start, max(start, end - 1))
            inner.Text = text

    # insert_paragraph_after / insert_paragraph_before are inherited from Anchor;
    # for a Heading, _range() is the heading paragraph, so "after" lands a new
    # paragraph just below the heading (the original v0 behaviour).


class HeadingCollection:
    """Iterable, indexable view over a document's headings.

    Symmetric with `BookmarkCollection` and `ContentControlCollection`:

        for h in doc.headings:           # iteration → Heading per heading paragraph
            ...
        doc.headings["Risks"]            # by visible text
        doc.headings[3]                  # by 1-based paragraph index
        "Risks" in doc.headings          # membership
        doc.headings.list()              # same shape as doc.outline()
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __getitem__(self, key: str | int) -> Heading:
        if isinstance(key, bool):
            # bool is a subclass of int; reject before the int branch matches.
            raise TypeError(f"heading key must be str or int, got {type(key).__name__}")
        if isinstance(key, int):
            return _IndexedHeading(self._doc, key)
        if isinstance(key, str):
            with _com.translate_com_errors():
                if _find_heading_paragraph(self._doc.com, key) is None:
                    raise AnchorNotFoundError("heading", key)
            return Heading(self._doc, key)
        raise TypeError(f"heading key must be str or int, got {type(key).__name__}")

    def __contains__(self, key: object) -> bool:
        if isinstance(key, bool):
            return False
        if isinstance(key, int):
            # 1-based paragraph index must reference an actual heading paragraph.
            with _com.translate_com_errors():
                for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                    if idx != key:
                        continue
                    try:
                        lvl = int(para.OutlineLevel)
                    except Exception:
                        return False
                    return lvl < 10
            return False
        if not isinstance(key, str):
            return False
        with _com.translate_com_errors():
            return _find_heading_paragraph(self._doc.com, key) is not None

    def list(self) -> list[dict[str, Any]]:
        """Same shape as `Document.outline()` — `[{level, text, anchor_id}, ...]`."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
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

    def __iter__(self) -> Iterator[Heading]:
        for entry in self.list():
            # Each entry's anchor_id is `heading:N`; index-based heading
            # disambiguates duplicate visible text.
            idx = int(entry["anchor_id"].split(":", 1)[1])
            yield _IndexedHeading(self._doc, idx)


class _IndexedHeading(Heading):
    """A Heading located by 1-based paragraph index — used by anchor_by_id('heading:N').

    Disambiguates duplicate heading text. The display name is set to the resolved
    heading text the first time `_paragraph()` succeeds so error messages and
    `.name` reads stay informative.
    """

    def __init__(self, doc: Document, paragraph_index: int) -> None:
        super().__init__(doc, name=f"heading:{paragraph_index}")
        self._paragraph_index = paragraph_index

    @property
    def anchor_id(self) -> str:
        return f"heading:{self._paragraph_index}"

    def _paragraph(self) -> Any:
        for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
            if idx != self._paragraph_index:
                continue
            try:
                level = int(para.OutlineLevel)
            except Exception:
                break
            if level >= 10:
                break
            self.name = paragraph_text(para) or self.name
            return para
        raise AnchorNotFoundError("heading", f"heading:{self._paragraph_index}")

    def _paragraph_and_index(self) -> tuple[Any, int]:
        return self._paragraph(), self._paragraph_index
