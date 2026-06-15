"""Document wrapper + DocumentCollection."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Any

from . import _com, _findreplace, _proofing, _snapshot
from ._anchors import (
    BookmarkCollection,
    ContentControlCollection,
    EndAnchor,
    EquationCollection,
    Heading,
    HeadingCollection,
    ImageCollection,
    Paragraph,
    ParagraphCollection,
    RangeAnchor,
    StartAnchor,
    _IndexedHeading,
    _utf16_len,
    _within_table,
    paragraph_text,
)
from ._comments import CommentCollection
from ._edit import EditScope
from ._fields import FieldCollection
from ._hyperlinks import HyperlinkCollection
from ._lists import ListCollection
from ._notes import EndnoteCollection, FootnoteCollection
from ._properties import PropertyCollection
from ._revisions import RevisionCollection
from ._sections import SectionCollection
from ._selection import Selection
from ._snapshot import Snapshot
from ._sources import SourceCollection
from ._styles import StyleCollection
from ._tables import Table, TableCollection
from ._variables import VariableCollection
from .constants import WdInformation, WdSaveFormat, WdStatistic
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    DocumentNotFoundError,
    OpError,
    ReplaceVerificationError,
)

if TYPE_CHECKING:
    from ._anchors import Anchor
    from ._app import Word


def _markup_flag(markup: str) -> bool:
    """Coerce a snapshot `markup` argument (`"none"` / `"all"`) to a bool."""
    value = str(markup).lower()
    if value in ("none", "off", "false"):
        return False
    if value in ("all", "on", "true"):
        return True
    raise OpError(f"markup must be 'none' or 'all', got {markup!r}")


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
    def saved(self) -> bool:
        """Whether the document has no unsaved changes (Word's `Document.Saved`).

        `True` right after a save; `False` once an edit dirties it. A brand-new,
        never-saved document reads `False` until its first save. This is the same
        flag `wordlive status` reports per open document.
        """
        with _com.translate_com_errors():
            return bool(self._doc.Saved)

    def save(self) -> str:
        """Save the document to its existing file, returning the absolute path.

        Raises [`OpError`][wordlive.OpError] if the document has never
        been saved (it has no path yet) — call [`save_as`][wordlive.Document.save_as]
        first. This is the **ungated** Python-API surface: it writes wherever the
        document already lives. The CLI / MCP `save` verb additionally checks that
        path against the configured save-directory whitelist before calling this.
        """
        with _com.translate_com_errors():
            folder = str(self._doc.Path)
            if not folder:
                raise OpError("document has never been saved; use save_as(path) first")
            self._doc.Save()
            return str(self._doc.FullName)

    def save_as(self, path: str | Path, *, fmt: str = "docx", overwrite: bool = False) -> str:
        """Save the document to `path`, returning the absolute path written.

        `fmt` is `"docx"` (the modern Open XML format). For PDF, use
        [`export_pdf`][wordlive.Document.export_pdf] (it goes through a different
        COM call and takes a page range). By default refuses to clobber an
        existing file — pass `overwrite=True` to allow it. **Ungated** like
        [`save`][wordlive.Document.save]; the CLI / MCP surface whitelists the
        target first.
        """
        target = Path(path).expanduser()
        fmt_norm = str(fmt).lower().lstrip(".")
        if fmt_norm == "pdf":
            raise OpError("save_as does not write PDF; use export_pdf(path) instead")
        if fmt_norm not in ("docx",):
            raise OpError(f"unsupported save format {fmt!r}; supported: docx (PDF via export_pdf)")
        if not overwrite and target.exists():
            raise OpError(
                f"refusing to overwrite existing file {str(target)!r}; pass overwrite=True"
            )
        abspath = str(target.resolve())
        with _com.translate_com_errors():
            self._doc.SaveAs2(FileName=abspath, FileFormat=int(WdSaveFormat.DOCUMENT_DEFAULT))
        return abspath

    def export_pdf(
        self, path: str | Path, *, from_page: int | None = None, to_page: int | None = None
    ) -> str:
        """Export the document (or a page span) to a PDF at `path`; return the path.

        `from_page` / `to_page` are 1-based and inclusive; omit both to export the
        whole document, or give `from_page` alone to export a single page. Goes
        through `Document.ExportAsFixedFormat` (the same engine
        [`snapshot`][wordlive.Document.snapshot] uses), so the PDF is a
        pixel-faithful render — the recommended "hand back a deliverable" path.
        Overwrites an existing file. **Ungated** like [`save`][wordlive.Document.save].
        """
        abspath = str(Path(path).expanduser().resolve())
        _snapshot._export_pdf(self._doc, abspath, from_page=from_page, to_page=to_page)
        return abspath

    @property
    def bookmarks(self) -> BookmarkCollection:
        return BookmarkCollection(self)

    @property
    def content_controls(self) -> ContentControlCollection:
        return ContentControlCollection(self)

    @property
    def sources(self) -> SourceCollection:
        """The document's bibliography sources — add and look up by tag.

        See `SourceCollection`: `doc.sources.add(...)`
        registers a source, `doc.sources["Smith2020"]` looks one up, and the
        collection is iterable and `in`-testable by tag. Cite a source with
        [`Anchor.insert_citation`][wordlive.Anchor.insert_citation] and list the
        cited ones with [`Document.add_bibliography`][wordlive.Document.add_bibliography].
        """
        return SourceCollection(self)

    @property
    def bibliography_style(self) -> str:
        """The citation/bibliography style (e.g. ``"APA"``, ``"MLA"``, ``"Chicago"``).

        Read/write. Setting it changes how every citation and the bibliography
        render (refresh them with [`update_fields`][wordlive.Document.update_fields]).
        Word accepts a build-dependent set of identifiers; an unsupported value
        raises [`OpError`][wordlive.OpError].
        """
        with _com.translate_com_errors():
            return str(self._doc.Bibliography.BibliographyStyle)

    @bibliography_style.setter
    def bibliography_style(self, style: str) -> None:
        if not str(style).strip():
            raise OpError("bibliography_style must be a non-empty string")
        with _com.translate_com_errors():
            self._doc.Bibliography.BibliographyStyle = str(style)

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
    def images(self) -> ImageCollection:
        """Read-only, iterable view over the document's embedded images (`doc.images`).

        Index an image by 1-based position (`doc.images[2]`) to get an
        [`ImageAnchor`][wordlive.ImageAnchor] (`image:N`), then `read_image()`
        for its raw bytes + MIME type — the path for handing an embedded picture
        to a vision model. `list()` summarises each image (MIME, size, alt text,
        and the `para:N` it's anchored in). The write mirror is any anchor's
        [`insert_image`][wordlive.Anchor.insert_image].
        """
        return ImageCollection(self)

    @property
    def equations(self) -> EquationCollection:
        """Read-only, iterable view over the document's equations (`doc.equations`).

        Index an equation by 1-based position (`doc.equations[2]`) to get an
        [`EquationAnchor`][wordlive.EquationAnchor] (`equation:N`), then `mathml`
        / `linear` to read it. `list()` summarises each equation (type, a linear
        preview, and the `para:N` it sits in). The write mirror is any anchor's
        [`insert_equation`][wordlive.Anchor.insert_equation].
        """
        return EquationCollection(self)

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
    def revisions(self) -> RevisionCollection:
        """Read-only, iterable view over the document's tracked changes (`doc.revisions`).

        When Track Changes is on, every edit is a `Revision` the user can accept
        or reject. `doc.revisions.list()` reports each as
        `{index, type, author, text, anchor_id, start, end, date}` — the
        *structured* way to see what tracked edits a batch recorded (the visual
        way is [`snapshot(markup="all")`][wordlive.Document.snapshot]). Index by
        1-based position (`doc.revisions[2]`); `type` is `"insert"` / `"delete"`
        / `"format"` / … . Writing tracked changes is
        [`tracked_changes()`][wordlive.Document.tracked_changes].
        """
        return RevisionCollection(self)

    @property
    def footnotes(self) -> FootnoteCollection:
        """Read-only, iterable view over the document's footnotes (`doc.footnotes`).

        Index a footnote by 1-based position (`doc.footnotes[2]`) to get a
        [`Footnote`][wordlive.Footnote] anchor (`footnote:N`) whose `set_text` /
        `delete` edit the note. `list()` summarises each note (number, body
        text, and the `para:N` it's anchored at). Create one with
        [`Anchor.insert_footnote`][wordlive.Anchor.insert_footnote].
        """
        return FootnoteCollection(self)

    @property
    def endnotes(self) -> EndnoteCollection:
        """Read-only, iterable view over the document's endnotes (`doc.endnotes`).

        The endnote mirror of [`footnotes`][wordlive.Document.footnotes]; notes
        are addressed `endnote:N`. Create one with
        [`Anchor.insert_endnote`][wordlive.Anchor.insert_endnote].
        """
        return EndnoteCollection(self)

    @property
    def hyperlinks(self) -> HyperlinkCollection:
        """Read-only, iterable view over the document's hyperlinks (`doc.hyperlinks`).

        The read mirror of [`Anchor.link_to`][wordlive.Anchor.link_to]: index a
        link by 1-based position (`doc.hyperlinks[2]`) to get a
        [`Hyperlink`][wordlive.Hyperlink], or `list()` to summarise each —
        visible text, external `address` or internal `sub_address` bookmark,
        screen tip, and the `range:START-END` / `para:N` it sits in.
        """
        return HyperlinkCollection(self)

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
        return FieldCollection(self)

    @property
    def properties(self) -> PropertyCollection:
        """Read/write view over the document's built-in and custom properties (metadata).

        `doc.properties.read()` returns `{"builtin": {…}, "custom": {…}}` — the
        Title / Author / Keywords / … bag plus any custom name/value pairs.
        `doc.properties.set("Title", "…")` writes a built-in property;
        `set(name, value, custom=True)` writes (creating if needed) a custom one.
        Wrap writes in `doc.edit(...)` for atomic undo.
        """
        return PropertyCollection(self)

    @property
    def variables(self) -> VariableCollection:
        """Read/write view over the document's variables (`doc.variables`).

        Document variables are invisible named string storage — the backing store
        for `{ DOCVARIABLE name }` fields. `doc.variables.list()` returns
        `{name: value}`; `set(name, value)` / `delete(name)` manage them. Wrap
        writes in `doc.edit(...)` for atomic undo.
        """
        return VariableCollection(self)

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

    def add_table(
        self,
        rows: int,
        cols: int,
        *,
        style: str | None = None,
        data: list[list[Any]] | None = None,
        header: bool = False,
    ) -> Table:
        """Append a `rows` × `cols` table at the end of the document and return it.

        The "build a document from the bottom up" helper for tables — the
        counterpart to [`append_paragraph`][wordlive.Document.append_paragraph].
        Sugar for `self.end.insert_table(...)`; see
        [`Anchor.insert_table`][wordlive.Anchor.insert_table] for the full
        semantics of `style` (defaults to the built-in ``"Table Grid"``), `data`
        (row-major fill, validated up front), and `header`. To place a table
        somewhere other than the end, resolve a position anchor and call
        `insert_table` on it directly (e.g.
        `doc.headings["Pricing"].insert_table(3, 2, ...)`). Wrap in
        `doc.edit(...)` for atomic undo.
        """
        return self.end.insert_table(
            rows, cols, where="after", style=style, data=data, header=header
        )

    def add_toc(
        self,
        *,
        levels: tuple[int, int] = (1, 3),
        use_heading_styles: bool = True,
        hyperlinks: bool = True,
    ) -> Any:
        """Insert a table of contents at the very start of the document.

        The "documents want their TOC at the top" helper — sugar for
        `self.start.insert_toc(...)`. See
        [`Anchor.insert_toc`][wordlive.Anchor.insert_toc] for the full semantics
        of `levels` (a ``(upper, lower)`` heading-level pair), `use_heading_styles`,
        and `hyperlinks`, and for the page-number-repagination caveat. Returns
        the new [`Toc`][wordlive.Toc]. Wrap in `doc.edit(...)` for atomic undo.
        """
        return self.start.insert_toc(
            levels=levels, use_heading_styles=use_heading_styles, hyperlinks=hyperlinks
        )

    def add_index(
        self,
        *,
        columns: int = 2,
        run_in: bool = False,
        right_align_page_numbers: bool = False,
    ) -> Any:
        """Insert a back-of-book index at the very end of the document.

        The "indexes live at the back" helper — sugar for
        `self.end.insert_index(...)`. See
        [`Anchor.insert_index`][wordlive.Anchor.insert_index] for the full
        semantics of `columns`, `run_in`, and `right_align_page_numbers`, and for
        the page-number-repagination caveat. Mark entries first with
        [`Anchor.mark_index_entry`][wordlive.Anchor.mark_index_entry]. Returns the
        new [`Index`][wordlive.Index]. Wrap in `doc.edit(...)` for atomic undo.
        """
        return self.end.insert_index(
            columns=columns, run_in=run_in, right_align_page_numbers=right_align_page_numbers
        )

    def add_bibliography(self) -> Any:
        """Insert a bibliography at the very end of the document.

        The "references live at the back" helper — sugar for
        `self.end.insert_bibliography()`. See
        [`Anchor.insert_bibliography`][wordlive.Anchor.insert_bibliography] for the
        field-block / repagination caveat. Add sources with
        [`doc.sources.add`][wordlive.SourceCollection.add] and cite them with
        [`Anchor.insert_citation`][wordlive.Anchor.insert_citation] first. Returns
        the new [`Bibliography`][wordlive.Bibliography]. Wrap in `doc.edit(...)`.
        """
        return self.end.insert_bibliography()

    def add_table_of_authorities(
        self,
        *,
        category: str | int = "all",
        passim: bool = True,
        keep_entry_formatting: bool = True,
        entry_separator: str | None = None,
        page_range_separator: str | None = None,
    ) -> Any:
        """Insert a table of authorities at the very end of the document.

        Sugar for `self.end.insert_table_of_authorities(...)`. See
        [`Anchor.insert_table_of_authorities`][wordlive.Anchor.insert_table_of_authorities]
        for the full semantics of `category`, `passim`, `keep_entry_formatting`,
        and the separators, and for the page-number-repagination caveat. Mark
        citations first with [`Anchor.mark_citation`][wordlive.Anchor.mark_citation].
        Returns the new [`TableOfAuthorities`][wordlive.TableOfAuthorities]. Wrap in
        `doc.edit(...)` for atomic undo.
        """
        return self.end.insert_table_of_authorities(
            category=category,
            passim=passim,
            keep_entry_formatting=keep_entry_formatting,
            entry_separator=entry_separator,
            page_range_separator=page_range_separator,
        )

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
          - `footnote:N`       — Nth footnote (1-based), resolving to its note body
          - `endnote:N`        — Nth endnote (1-based), resolving to its note body
          - `image:N`          — Nth embedded image (1-based, Word's InlineShapes order)
          - `equation:N`       — Nth equation (1-based, Word's OMaths order)
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
        raise AnchorNotFoundError(
            "anchor",
            anchor_id,
            hint=(
                f"unknown anchor type {kind!r}; expected one of "
                "start/end/heading/para/bookmark/cc/footnote/endnote/image/table/range/header/footer"
            ),
        )

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

        Matches are located per *segment* (contiguous body text or a single table
        cell) so the returned offsets stay exact even inside tables; see
        `_scope_segments`.
        """
        segments = self._scope_segments(scope)
        results: list[dict[str, Any]] = []
        for base, haystack in segments:
            for m in _findreplace.find_matches(haystack, text):
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

        Matching is segment-aware (see `_scope_segments`), so a match inside a
        table cell resolves to the right cell rather than drifting into its
        neighbour. As a backstop, each write is verified against the located text
        and raises `ReplaceVerificationError` rather than overwriting the wrong
        span.
        """
        segments = self._scope_segments(scope)
        match_payloads: list[dict[str, Any]] = [
            {
                "anchor_id": f"range:{base + m.start}-{base + m.end}",
                "start": base + m.start,
                "end": base + m.end,
                "text": m.text,
            }
            for base, haystack in segments
            for m in _findreplace.find_matches(haystack, find)
        ]
        if not match_payloads:
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

    def stats(self) -> dict[str, Any]:
        """A one-call summary of the document — the "what am I looking at" read.

        Returns `{pages, words, characters, paragraphs, lines, sections,
        headings, tables, images, equations, comments, revisions, saved}`. The five text
        counts come from Word's own `ComputeStatistics`; the structural counts
        come from wordlive's discovery collections (so they agree with
        `doc.tables` / `doc.images` / `outline` etc.); `saved` is `doc.saved`.

        `pages`/`lines` are print-layout truth, so the document is
        **repaginated first** (content-neutral — selection, scroll, and view are
        untouched), the same guarantee a `snapshot` gives. A pure read; nothing
        is mutated — `Repaginate` flips Word's dirty bit, so the document's
        `Saved` state is snapshotted and restored around it.
        """
        with _com.translate_com_errors(), _com.preserve_saved(self._doc):
            self._doc.Repaginate()
            text_counts = {
                "pages": int(self._doc.ComputeStatistics(int(WdStatistic.PAGES))),
                "words": int(self._doc.ComputeStatistics(int(WdStatistic.WORDS))),
                "characters": int(self._doc.ComputeStatistics(int(WdStatistic.CHARACTERS))),
                "paragraphs": int(self._doc.ComputeStatistics(int(WdStatistic.PARAGRAPHS))),
                "lines": int(self._doc.ComputeStatistics(int(WdStatistic.LINES))),
            }
        return {
            **text_counts,
            "sections": len(self.sections),
            "headings": len(self.outline()),
            "tables": len(self.tables),
            "images": len(self.images),
            "equations": len(self.equations),
            "comments": len(self.comments),
            "revisions": len(self.revisions),
            "saved": self.saved,
        }

    def proofing(self) -> dict[str, Any]:
        """Run Word's proofing tools and report spelling, grammar, and readability.

        Returns `{spelling, grammar, readability}`. `spelling` and `grammar` are
        each `{count, errors}` — the exact error count plus a (capped) list of
        `{text, anchor_id, para}` for the flagged runs, so a `range:START-END`
        can be fed back into `read` or `comments.add`. `readability` is Word's
        readability statistics (Flesch Reading Ease, Flesch-Kincaid Grade Level,
        passive-sentence %, word/sentence averages), snake_cased.

        Heavier than [`stats`][wordlive.Document.stats]: it asks Word to (re)check
        the document. Still a pure read — nothing is mutated. If proofing is
        disabled or the document is protected, the affected section reports a
        `None` count / empty readability rather than failing.
        """
        return _proofing.read_proofing(self)

    def _page_of(self, position: int) -> int:
        """1-based page number that document offset `position` falls on."""
        with _com.translate_com_errors():
            rng = self._doc.Range(int(position), int(position))
            return int(rng.Information(int(WdInformation.ACTIVE_END_PAGE_NUMBER)))

    @staticmethod
    def _resolve_page_arg(pages: int | tuple[int, int] | None) -> tuple[int | None, int | None]:
        """Normalise a `pages` argument into a `(from, to)` 1-based span (or all)."""
        if pages is None:
            return None, None
        if isinstance(pages, bool):  # bool is an int subclass — reject before the int branch
            raise ValueError(f"pages must be an int or (start, end) tuple, not {pages!r}")
        if isinstance(pages, int):
            if pages < 1:
                raise ValueError(f"page number must be >= 1, got {pages}")
            return pages, pages
        if isinstance(pages, (tuple, list)) and len(pages) == 2:
            start, end = int(pages[0]), int(pages[1])
            if start < 1 or end < start:
                raise ValueError(f"invalid page span: {pages!r}")
            return start, end
        raise ValueError(f"pages must be an int or (start, end) tuple, got {pages!r}")

    def _anchor_page_span(self, anchor: Anchor) -> tuple[int, int]:
        """Page span an anchor occupies. Headings expand to their whole section.

        Mirrors `_scope_range`'s heading-means-its-body rule, so a snapshot of a
        `heading:` anchor shows the section a model is editing, not just the
        heading line.
        """
        with _com.translate_com_errors():
            if isinstance(anchor, Heading):
                head = anchor.com
                start, end = int(head.Start), int(anchor.section_range().End)
            else:
                rng = anchor.com
                start, end = int(rng.Start), int(rng.End)
        from_page = self._page_of(start)
        to_page = max(from_page, self._page_of(max(start, end)))
        return from_page, to_page

    def snapshot(
        self,
        out: str | Path | None = None,
        *,
        pages: int | tuple[int, int] | None = None,
        dpi: int = 150,
        max_dim: int | None = None,
        markup: str = "none",
    ) -> list[Snapshot]:
        """Render document page(s) to PNG so a vision model can *see* the layout.

        Word exports a pixel-faithful PDF of the live document and wordlive
        rasterises the requested pages — a true WYSIWYG image (real fonts,
        spacing, page geometry), ideal for iterating on style and formatting.

        `pages` selects what to render: `None` (default) renders every page,
        an `int` a single 1-based page, and a `(start, end)` tuple an inclusive
        span. Returns one [`Snapshot`][wordlive.Snapshot] per page (so a single
        page is a one-element list); read `.png` for the bytes.

        If `out` is given the image is also written there: a single page to `out`
        itself, multiple pages alongside it as `<stem>-p<N><suffix>`.

        `markup` is `"none"` (default — render the final document) or `"all"`
        (render tracked changes and comments as visible revision marks and
        balloons). The marks come from the export, not a view change, so the
        user's on-screen markup setting is left untouched. The structured
        counterpart is [`revisions`][wordlive.Document.revisions].

        `dpi` controls resolution; ~150 reads well for a vision model without
        bloating the image. `max_dim` caps each page's **long edge** in pixels,
        only ever lowering the resolution — the lever for a cheap *whole-document*
        layout check (a vision model is billed on pixel area, so a long-edge cap
        gives a predictable per-page token budget regardless of paper size; ~1000
        keeps a page legible for "did my styling land" at a fraction of the
        tokens). `dpi=72` is a coarser alternative. Read-only — the document and
        the user's cursor are untouched. Requires the `snapshot` extra (PyMuPDF),
        else [`SnapshotError`][wordlive.SnapshotError].
        """
        if max_dim is not None and (isinstance(max_dim, bool) or int(max_dim) < 1):
            raise OpError(f"max_dim must be a positive integer (pixels); got {max_dim!r}")
        from_page, to_page = self._resolve_page_arg(pages)
        rendered = _snapshot.render(
            self._doc,
            from_page=from_page,
            to_page=to_page,
            dpi=dpi,
            max_dim=max_dim,
            markup=_markup_flag(markup),
        )
        return _snapshot.build_snapshots(rendered, out)

    def snapshot_anchor(
        self,
        anchor: Anchor,
        out: str | Path | None = None,
        *,
        dpi: int = 150,
        max_dim: int | None = None,
        markup: str = "none",
    ) -> list[Snapshot]:
        """Render the page(s) an anchor sits on. Backs [`Anchor.snapshot`][wordlive.Anchor.snapshot].

        A `heading:` anchor expands to its whole section (the heading plus the
        body beneath it, up to the next same-or-higher heading); any other
        anchor renders the page(s) its range spans. See
        [`snapshot`][wordlive.Document.snapshot] for `out`/`dpi`/`max_dim`/`markup`
        semantics and the return shape.
        """
        if max_dim is not None and (isinstance(max_dim, bool) or int(max_dim) < 1):
            raise OpError(f"max_dim must be a positive integer (pixels); got {max_dim!r}")
        from_page, to_page = self._anchor_page_span(anchor)
        rendered = _snapshot.render(
            self._doc,
            from_page=from_page,
            to_page=to_page,
            dpi=dpi,
            max_dim=max_dim,
            markup=_markup_flag(markup),
        )
        return _snapshot.build_snapshots(rendered, out)

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
        """`[{name, path, saved, is_active}, ...]` — used by `wordlive status`.

        `name` is the document's window name (e.g. ``Report.docx``, or
        ``Document1`` for one never saved) and is always non-empty so a caller
        can confirm which document it is about to edit. `saved` is whether the
        document has an on-disk location yet; `path` is that full path, or empty
        for an unsaved document. The active document is matched by full path
        (falling back to name), which is robust when several unsaved documents
        share a blank path.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            active_name: str | None
            active_full: str | None
            try:
                active = self._word.com.ActiveDocument
                active_name = str(active.Name)
                active_full = str(active.FullName)
            except Exception:
                active_name = active_full = None
            for doc in self._com_collection:
                name = str(doc.Name or "")
                full = str(doc.FullName or "")
                try:
                    on_disk = bool(str(doc.Path or ""))
                except Exception:
                    on_disk = False
                is_active = (full == active_full) if full and active_full else (name == active_name)
                out.append(
                    {
                        "name": name or full or "Document",
                        "path": full if on_disk else "",
                        "saved": on_disk,
                        "is_active": bool(is_active),
                    }
                )
        return out
