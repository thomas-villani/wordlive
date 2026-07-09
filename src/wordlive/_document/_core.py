"""Document spine: construction, the COM handle, the edit scope, anchor
resolution, and the collection accessors the feature mixins build on."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from typing import TYPE_CHECKING, Any, cast

from .. import _com
from .._anchors import (
    Bookmark,
    BookmarkCollection,
    ChartCollection,
    ContentControlCollection,
    EndAnchor,
    EquationCollection,
    Heading,
    HeadingCollection,
    ImageCollection,
    Paragraph,
    ParagraphCollection,
    RangeAnchor,
    ShapeCollection,
    StartAnchor,
    TextBoxCollection,
    _IndexedHeading,
    _pin_name_for,
)
from .._comments import CommentCollection
from .._edit import EditScope
from .._fields import FieldCollection
from .._hyperlinks import HyperlinkCollection
from .._lists import ListCollection
from .._notes import EndnoteCollection, FootnoteCollection
from .._properties import PropertyCollection
from .._revisions import RevisionCollection
from .._sections import SectionCollection
from .._selection import Selection
from .._sources import SourceCollection
from .._styles import StyleCollection
from .._tables import TableCollection
from .._themes import DocumentTheme
from .._variables import VariableCollection
from ..exceptions import (
    AnchorNotFoundError,
    OpError,
)

if TYPE_CHECKING:
    from .._anchors import Anchor
    from .._app import Word

if TYPE_CHECKING:
    from . import Document


def _markup_flag(markup: str) -> bool:
    """Coerce a snapshot `markup` argument (`"none"` / `"all"`) to a bool."""
    value = str(markup).lower()
    if value in ("none", "off", "false"):
        return False
    if value in ("all", "on", "true"):
        return True
    raise OpError(f"markup must be 'none' or 'all', got {markup!r}")


def _resolve_level_band(levels: int | tuple[int, int] | None) -> tuple[int, int]:
    """Normalise a `pin_outline` `levels` argument into an inclusive `(lo, hi)` band.

    `None` -> all heading levels (1–9); an `int` n -> `1..n`; a `(lo, hi)` pair ->
    that inclusive band. Raises `OpError` on a bad shape or out-of-range value.
    """
    if levels is None:
        return 1, 9
    if isinstance(levels, bool):  # bool is an int subclass — reject before the int branch
        raise OpError(f"levels must be an int or (lo, hi) tuple, not {levels!r}")
    if isinstance(levels, int):
        if not 1 <= levels <= 9:
            raise OpError(f"levels must be between 1 and 9, got {levels}")
        return 1, levels
    if isinstance(levels, (tuple, list)) and len(levels) == 2:
        lo, hi = int(levels[0]), int(levels[1])
        if not (1 <= lo <= hi <= 9):
            raise OpError(f"invalid level band: {levels!r} (expected 1 <= lo <= hi <= 9)")
        return lo, hi
    raise OpError(f"levels must be an int or (lo, hi) tuple, got {levels!r}")


@dataclass(frozen=True)
class WatermarkInfo:
    """The text watermark stamped behind a document's pages (`doc.watermark()`).

    `text` is the watermark's text (Word stamps the same text into every
    section's header story, so this is the common value); `sections` lists the
    1-based section indices that carry it. The read mirror of
    [`set_watermark`][wordlive.Document.set_watermark] /
    [`remove_watermark`][wordlive.Document.remove_watermark].
    """

    text: str
    sections: list[int]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class DocumentCore:
    """Core state and primitives shared by every `Document` feature mixin."""

    @property
    def _as_document(self) -> Document:
        """`self`, narrowed to the concrete `Document`.

        The feature mixins are only ever mixed into `Document`, but a type
        checker only sees the mixin. Collaborators are annotated on `Document`,
        and mixins call across to one another, so both need the concrete type.
        Runtime-free: `cast` returns its argument unchanged.
        """
        return cast("Document", self)

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
    def saved(self) -> bool:
        """Whether the document has no unsaved changes (Word's `Document.Saved`).

        `True` right after a save; `False` once an edit dirties it. A brand-new,
        never-saved document reads `False` until its first save. This is the same
        flag `wordlive status` reports per open document.
        """
        with _com.translate_com_errors():
            return bool(self._doc.Saved)

    @property
    def bookmarks(self) -> BookmarkCollection:
        return BookmarkCollection(self._as_document)

    @property
    def content_controls(self) -> ContentControlCollection:
        return ContentControlCollection(self._as_document)

    @property
    def sources(self) -> SourceCollection:
        """The document's bibliography sources — add and look up by tag.

        See `SourceCollection`: `doc.sources.add(...)`
        registers a source, `doc.sources["Smith2020"]` looks one up, and the
        collection is iterable and `in`-testable by tag. Cite a source with
        [`Anchor.insert_citation`][wordlive.Anchor.insert_citation] and list the
        cited ones with [`Document.add_bibliography`][wordlive.Document.add_bibliography].
        """
        return SourceCollection(self._as_document)

    @property
    def styles(self) -> StyleCollection:
        return StyleCollection(self._as_document)

    @property
    def theme(self) -> DocumentTheme:
        """The document's theme — the document-wide brand primitive.

        See [`DocumentTheme`][wordlive.DocumentTheme]: `doc.theme.apply("Facet")`
        swaps the whole theme, `doc.theme.set_colors(accent1="#1A73E8")` /
        `doc.theme.set_fonts(major="Arial")` set brand colours/fonts, and
        `doc.theme.colors` / `doc.theme.to_dict()` read it back. Wrap mutations in
        `doc.edit(...)` for atomic undo.
        """
        return DocumentTheme(self._as_document)

    @property
    def tables(self) -> TableCollection:
        """Iterable, indexable view over the document's tables.

        Index by 1-based position (`doc.tables[1]`) or `Title`
        (`doc.tables["Budget"]`). Cells are anchors: `doc.tables[1].cell(2, 3)`
        — or `doc.anchor_by_id("table:1:2:3")` — returns a `Cell` that works
        with `set_text`, `apply_style`, and `format_paragraph`.
        """
        return TableCollection(self._as_document)

    @property
    def headings(self) -> HeadingCollection:
        """Iterable view over the document's headings.

        Symmetric with `bookmarks`, `content_controls`, and `styles`. Index by
        visible text (`doc.headings["Risks"]`) or 1-based paragraph position
        (`doc.headings[3]`). `Document.heading(name)` remains as sugar for
        `self.headings[name]`.
        """
        return HeadingCollection(self._as_document)

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
        return ParagraphCollection(self._as_document)

    @property
    def lists(self) -> ListCollection:
        """Read-only, iterable view over the document's bullet / numbered lists.

        Index a list by 1-based position (`doc.lists[2]`) to get a
        [`RangeAnchor`][wordlive.RangeAnchor] over its range, so every list verb
        (`apply_list`, `restart_numbering`, …) is available on it. List
        formatting itself is applied through any anchor's `apply_list(...)`.
        """
        return ListCollection(self._as_document)

    @property
    def images(self) -> ImageCollection:
        """Read-only, iterable view over the document's embedded images (`doc.images`).

        Index an image by 1-based position (`doc.images[2]`) to get an
        [`ImageAnchor`][wordlive.ImageAnchor] (`image:N`), then `read_image()`
        for its raw bytes + MIME type — the path for handing an embedded picture
        to a vision model. `list()` summarises each image (MIME, size, alt text,
        and the `para:N` it's anchored in). The write mirror is any anchor's
        [`insert_image`][wordlive.Anchor.insert_image].
        """
        return ImageCollection(self._as_document)

    @property
    def equations(self) -> EquationCollection:
        """Read-only, iterable view over the document's equations (`doc.equations`).

        Index an equation by 1-based position (`doc.equations[2]`) to get an
        [`EquationAnchor`][wordlive.EquationAnchor] (`equation:N`), then `mathml`
        / `linear` to read it. `list()` summarises each equation (type, a linear
        preview, and the `para:N` it sits in). The write mirror is any anchor's
        [`insert_equation`][wordlive.Anchor.insert_equation].
        """
        return EquationCollection(self._as_document)

    @property
    def charts(self) -> ChartCollection:
        """Read-only, iterable view over the document's charts (`doc.charts`).

        Index a chart by 1-based position (`doc.charts[2]`) to get a
        [`ChartAnchor`][wordlive.ChartAnchor] (`chart:N`); `chart_type` / `title`
        read its metadata. `list()` summarises each chart (kind, title, and the
        `para:N` it sits in). Charts are inserted with their data link broken
        (static data), so reading the series back is deferred — this view is
        metadata only. The write mirror is any anchor's
        [`insert_chart`][wordlive.Anchor.insert_chart].
        """
        return ChartCollection(self._as_document)

    @property
    def shapes(self) -> ShapeCollection:
        """Iterable view over the document's floating shapes (`doc.shapes`).

        Index a shape by 1-based position (`doc.shapes[2]`) to get a
        [`ShapeAnchor`][wordlive.ShapeAnchor] (`shape:N`) — a text box, a floating
        image, or WordArt — then restyle it in place (`set_wrap` / `set_position`
        / `set_size` / `format` / `replace_image`). `list()` summarises each shape
        (kind, size, wrap, the `para:N` it's anchored in). Header-story watermarks
        are excluded; positions follow document order and renumber as shapes come
        and go. The write mirror is any anchor's
        [`insert_text_box`][wordlive.Anchor.insert_text_box] / a floating
        [`insert_image`][wordlive.Anchor.insert_image].
        """
        return ShapeCollection(self._as_document)

    @property
    def text_boxes(self) -> TextBoxCollection:
        """The text boxes among `doc.shapes` — the ``shape_type == "text_box"`` subset.

        A discovery filter, not a second id space: `doc.text_boxes[1]` returns a
        [`ShapeAnchor`][wordlive.ShapeAnchor] that keeps its canonical `shape:N`
        id (its position among *all* floating shapes). Created by any anchor's
        [`insert_text_box`][wordlive.Anchor.insert_text_box].
        """
        return TextBoxCollection(self._as_document)

    @property
    def sections(self) -> SectionCollection:
        """Indexable view over the document's sections, headers, and footers.

        `doc.sections[1].header()` / `.footer()` return `HeaderFooter` anchors
        (addressed `header:S:WHICH` / `footer:S:WHICH`) that work with
        `set_text` / `apply_style` like any other anchor.
        """
        return SectionCollection(self._as_document)

    @property
    def comments(self) -> CommentCollection:
        """Iterable, indexable view over the document's review comments.

        `doc.comments.add(anchor, text, author=...)` attaches a comment to any
        anchor's range without changing the text — the polite, side-channel way
        to flag something. Index existing comments by 1-based position
        (`doc.comments[2]`) to `resolve()` or `delete()` them.
        """
        return CommentCollection(self._as_document)

    @property
    def revisions(self) -> RevisionCollection:
        """Iterable view over the document's tracked changes (`doc.revisions`).

        When Track Changes is on, every edit is a `Revision` the user can accept
        or reject. `doc.revisions.list()` reports each as
        `{index, type, author, text, anchor_id, start, end, date}` — the
        *structured* way to see what tracked edits a batch recorded (the visual
        way is [`snapshot(markup="all")`][wordlive.Document.snapshot]). Index by
        1-based position (`doc.revisions[2]`); `type` is `"insert"` / `"delete"`
        / `"format"` / … .

        Resolve them too: `doc.revisions[2].accept()` / `.reject()` for one, or
        [`accept_all`][wordlive.RevisionCollection.accept_all] /
        [`reject_all`][wordlive.RevisionCollection.reject_all]
        (`within=anchor` to scope to one section/range) for many. For a read that
        separates the inserted from the deleted runs of a just-edited range, use
        [`Anchor.text_final`][wordlive.Anchor.text_final] /
        [`text_original`][wordlive.Anchor.text_original] /
        [`revision_segments`][wordlive.Anchor.revision_segments]. Writing tracked
        changes is [`tracked_changes()`][wordlive.Document.tracked_changes].
        """
        return RevisionCollection(self._as_document)

    @property
    def footnotes(self) -> FootnoteCollection:
        """Read-only, iterable view over the document's footnotes (`doc.footnotes`).

        Index a footnote by 1-based position (`doc.footnotes[2]`) to get a
        [`Footnote`][wordlive.Footnote] anchor (`footnote:N`) whose `set_text` /
        `delete` edit the note. `list()` summarises each note (number, body
        text, and the `para:N` it's anchored at). Create one with
        [`Anchor.insert_footnote`][wordlive.Anchor.insert_footnote].
        """
        return FootnoteCollection(self._as_document)

    @property
    def endnotes(self) -> EndnoteCollection:
        """Read-only, iterable view over the document's endnotes (`doc.endnotes`).

        The endnote mirror of [`footnotes`][wordlive.Document.footnotes]; notes
        are addressed `endnote:N`. Create one with
        [`Anchor.insert_endnote`][wordlive.Anchor.insert_endnote].
        """
        return EndnoteCollection(self._as_document)

    @property
    def hyperlinks(self) -> HyperlinkCollection:
        """Read-only, iterable view over the document's hyperlinks (`doc.hyperlinks`).

        The read mirror of [`Anchor.link_to`][wordlive.Anchor.link_to]: index a
        link by 1-based position (`doc.hyperlinks[2]`) to get a
        [`Hyperlink`][wordlive.Hyperlink], or `list()` to summarise each —
        visible text, external `address` or internal `sub_address` bookmark,
        screen tip, and the `range:START-END` / `para:N` it sits in.
        """
        return HyperlinkCollection(self._as_document)

    @property
    def fields(self) -> FieldCollection:
        """Read-only, iterable view over the document's fields (`doc.fields`).

        The read mirror of [`Anchor.insert_field`][wordlive.Anchor.insert_field]:
        index a field by 1-based position (`doc.fields[2]`) to get a
        [`Field`][wordlive.Field], or `list()` to summarise each — its `kind`
        (the code's leading keyword, `PAGE` / `REF` / `TOC` / …), raw `code`,
        rendered `result`, and the `range:START-END` / `para:N` it sits in.
        Refresh stale results with [`update_fields`][wordlive.Document.update_fields].
        """
        return FieldCollection(self._as_document)

    @property
    def properties(self) -> PropertyCollection:
        """Read/write view over the document's built-in and custom properties (metadata).

        `doc.properties.read()` returns `{"builtin": {…}, "custom": {…}}` — the
        Title / Author / Keywords / … bag plus any custom name/value pairs.
        `doc.properties.set("Title", "…")` writes a built-in property;
        `set(name, value, custom=True)` writes (creating if needed) a custom one.
        Wrap writes in `doc.edit(...)` for atomic undo.
        """
        return PropertyCollection(self._as_document)

    @property
    def variables(self) -> VariableCollection:
        """Read/write view over the document's variables (`doc.variables`).

        Document variables are invisible named string storage — the backing store
        for `{ DOCVARIABLE name }` fields. `doc.variables.list()` returns
        `{name: value}`; `set(name, value)` / `delete(name)` manage them. Wrap
        writes in `doc.edit(...)` for atomic undo.
        """
        return VariableCollection(self._as_document)

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
        return StartAnchor(self._as_document)

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
        return EndAnchor(self._as_document)

    def heading(self, name: str) -> Heading:
        # Lazy lookup — Heading.__init__ doesn't hit COM. _range() validates.
        return Heading(self._as_document, name)

    def range(self, start: int, end: int) -> RangeAnchor:
        """Return a `RangeAnchor` over the absolute offsets `[start, end)`.

        Offsets are UTF-16 code units — the coordinates Word uses and that
        `find()` emits as `range:START-END`. Lazy: the offsets aren't validated
        against the document until the anchor is used.
        """
        return RangeAnchor(self._as_document, start, end)

    def anchor_by_id(self, anchor_id: str) -> Anchor:
        """Resolve an `anchor_id` string into an Anchor.

        Recognised forms:
          - `start`            — the position before the first paragraph (the prepend target)
          - `end`              — the position past the last paragraph (the append target)
          - `heading:N`        — Nth paragraph in the document (1-based, must be a heading)
          - `para:N`           — Nth paragraph (1-based, any paragraph; same index space as `heading:N`)
          - `bookmark:NAME`    — bookmark by name
          - `pin:CODE`         — a durable handle minted by `pin` / `stamp` / `pin_outline`
          - `cc:NAME`          — content control by Title (or Tag)
          - `footnote:N`       — Nth footnote (1-based), resolving to its note body
          - `endnote:N`        — Nth endnote (1-based), resolving to its note body
          - `image:N`          — Nth embedded image (1-based, Word's InlineShapes order)
          - `equation:N`       — Nth equation (1-based, Word's OMaths order)
          - `chart:N`          — Nth chart (1-based, document order over chart inline shapes)
          - `shape:N`          — Nth floating shape (1-based, document order: text box / image / WordArt)
          - `textbox:N`        — Nth text box (alias onto its canonical `shape:N`)
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
            return _IndexedHeading(self._as_document, idx)
        if kind == "para":
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError("paragraph", anchor_id) from e
            # Lazy, like heading:N — a bad index raises AnchorNotFoundError on use.
            return Paragraph(self._as_document, idx)
        if kind == "bookmark":
            return self.bookmarks[value]
        if kind == "pin":
            name = _pin_name_for(value)
            with _com.translate_com_errors():
                if not self._doc.Bookmarks.Exists(name):
                    # A vanished pin (its content was deleted) correctly misses.
                    raise AnchorNotFoundError("pin", anchor_id)
            return Bookmark.pin(self._as_document, value)
        if kind == "cc":
            return self.content_controls[value]
        if kind in ("footnote", "endnote"):
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError(kind, anchor_id) from e
            coll = self.footnotes if kind == "footnote" else self.endnotes
            try:
                return coll[idx]
            except AnchorNotFoundError as e:
                raise AnchorNotFoundError(kind, anchor_id) from e
        if kind == "image":
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError("image", anchor_id) from e
            try:
                return self.images[idx]
            except AnchorNotFoundError as e:
                raise AnchorNotFoundError("image", anchor_id) from e
        if kind == "equation":
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError("equation", anchor_id) from e
            try:
                return self.equations[idx]
            except AnchorNotFoundError as e:
                raise AnchorNotFoundError("equation", anchor_id) from e
        if kind == "chart":
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError("chart", anchor_id) from e
            try:
                return self.charts[idx]
            except AnchorNotFoundError as e:
                raise AnchorNotFoundError("chart", anchor_id) from e
        if kind == "shape":
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError("shape", anchor_id) from e
            try:
                return self.shapes[idx]
            except AnchorNotFoundError as e:
                raise AnchorNotFoundError("shape", anchor_id) from e
        if kind == "textbox":
            # A thin alias onto the text-box subset of shape:N — the returned
            # ShapeAnchor reports its canonical shape:N id, not textbox:N.
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError("text box", anchor_id) from e
            try:
                return self.text_boxes[idx]
            except AnchorNotFoundError as e:
                raise AnchorNotFoundError("text box", anchor_id) from e
        if kind == "table":
            parts = value.split(":")
            if len(parts) != 3:
                # `table:N` (whole table) isn't a single-range anchor.
                raise AnchorNotFoundError("table cell", anchor_id)
            # `table:N:row:R` / `table:N:col:C` address a whole row / column;
            # `table:N:R:C` (all numeric) addresses a single cell.
            selector = parts[1].lower()
            if selector in ("row", "col"):
                try:
                    t, idx = int(parts[0]), int(parts[2])
                except ValueError as e:
                    kind_label = "table row" if selector == "row" else "table column"
                    raise AnchorNotFoundError(kind_label, anchor_id) from e
                if selector == "row":
                    return self.tables[t].row(idx)
                return self.tables[t].column(idx)
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
        raise AnchorNotFoundError(
            "anchor",
            anchor_id,
            hint=(
                f"unknown anchor type {kind!r}; expected one of "
                "start/end/heading/para/bookmark/pin/cc/footnote/endnote/image/equation/"
                "chart/shape/textbox/table/range/header/footer"
            ),
        )

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
