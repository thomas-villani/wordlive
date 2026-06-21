"""Document wrapper + DocumentCollection."""

from __future__ import annotations

import difflib
import secrets
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Any

from . import _checkpoint, _com, _findreplace, _linting, _proofing, _shapes, _snapshot
from ._anchors import (
    _WL_PREFIX,
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
    ShapeAnchor,
    ShapeCollection,
    StartAnchor,
    TextBoxCollection,
    _bookmarks_including_hidden,
    _IndexedHeading,
    _mint_wl_bookmark,
    _new_pin_code,
    _pin_id_for,
    _pin_name_for,
    _utf16_len,
    _validate_pin_slug,
    _within_table,
    paragraph_text,
)
from ._checkpoint import Checkpoint
from ._comments import CommentCollection
from ._edit import EditScope
from ._fields import FieldCollection
from ._format import to_bgr
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
from ._themes import DocumentTheme
from ._variables import VariableCollection
from .constants import (
    MsoPresetTextEffect,
    MsoTriState,
    WdHeaderFooterIndex,
    WdInformation,
    WdRelativeHorizontalPosition,
    WdRelativeVerticalPosition,
    WdSaveFormat,
    WdShapePosition,
    WdStatistic,
    WdWrapSideType,
    WdWrapType,
)
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

    def pin(self, anchor: Anchor | str, name: str | None = None) -> dict[str, Any]:
        """Plant a durable handle on `anchor`'s range and return its `pin:` id.

        The fix for fragile positional ids: `pin("para:7")` mints a hidden
        bookmark over that paragraph's range and hands back a `pin:<code>` anchor
        id that keeps pointing at the same content across later inserts / deletes
        / edits (Word maintains the association natively — that's the durability).
        Resolve it like any anchor — `doc.anchor_by_id("pin:a3f9c2")` — or feed it
        straight into another op. If the pinned content is later deleted the handle
        correctly vanishes (resolving it raises `AnchorNotFoundError`).

        `anchor` is an [`Anchor`][wordlive.Anchor] or an anchor id string. `name`
        optionally gives a readable slug (``budget-intro`` -> ``pin:budget-intro``;
        lowercase words joined by single hyphens); omit it for a random code.
        Re-using a slug moves the handle to the new range (Word's `Bookmarks.Add`
        semantics). Editing *through* the pin (`set_text`) keeps it; rewriting the
        same span via a different anchor's `Range.Text` drops it.

        Returns `{"anchor_id": "pin:…", "pin": "pin:…", "target": <resolved id>}`.
        `stamp` is an alias. Wrap in `doc.edit(...)` for atomic undo — but do not
        call it inside an already-open edit scope (custom undo records don't nest;
        the `exec` batch already owns one). The CLI verb is
        `wordlive pin ANCHOR_ID [--name SLUG]`; the exec op is `pin`.
        """
        resolved = self.anchor_by_id(anchor) if isinstance(anchor, str) else anchor
        target = anchor if isinstance(anchor, str) else resolved.anchor_id
        if name is not None:
            code = _validate_pin_slug(name)
            with _com.translate_com_errors():
                _mint_wl_bookmark(self._doc, resolved.com, code)
        else:
            with _com.translate_com_errors():
                code = _new_pin_code()
                while self._doc.Bookmarks.Exists(_pin_name_for(code)):
                    code = _new_pin_code()
                _mint_wl_bookmark(self._doc, resolved.com, code)
        return {"anchor_id": f"pin:{code}", "pin": f"pin:{code}", "target": target}

    # `stamp` reads better for "stamp a handle on this"; same operation.
    stamp = pin

    def _existing_pin_starts(self) -> dict[int, str]:
        """Map each existing `_wl_` bookmark's range start -> its pin code.

        Backs `pin_outline` idempotency: a heading whose range start already
        carries a wordlive handle reuses it instead of minting a duplicate.
        """
        out: dict[int, str] = {}
        with _com.translate_com_errors():
            for bm in _bookmarks_including_hidden(self._doc):
                nm = str(bm.Name)
                if nm.startswith(_WL_PREFIX):
                    out[int(bm.Range.Start)] = _pin_id_for(nm)
        return out

    def pin_outline(self, *, levels: int | tuple[int, int] | None = None) -> dict[str, str]:
        """Pin every heading at once and return the `{heading_id: pin_id}` map.

        A durable navigation scaffold up front: stamp a handle on each heading so
        an agent can address sections by `pin:` ids that survive the inserts /
        deletes it is about to make, instead of re-reading `outline` after every
        edit. Idempotent — a heading already carrying a wordlive handle reuses it,
        so calling this twice returns the same map (run it once on a stable
        document; the reuse keys on each heading's range start).

        `levels` filters which headings get pinned: `None` (default) pins every
        heading, an `int` n pins levels ``1..n``, and a ``(lo, hi)`` tuple pins
        the inclusive band. Returns an ordered ``{"heading:3": "pin:a3f9c2", …}``.
        Wrap in `doc.edit(...)` for atomic undo. See
        [`pin`][wordlive.Document.pin] for the single-anchor form.
        """
        lo, hi = _resolve_level_band(levels)
        existing = self._existing_pin_starts()
        out: dict[str, str] = {}
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    continue
                if level >= 10 or not (lo <= level <= hi):
                    continue
                rng = para.Range
                start = int(rng.Start)
                code = existing.get(start)
                if code is None:
                    code = _new_pin_code()
                    while self._doc.Bookmarks.Exists(_pin_name_for(code)):
                        code = _new_pin_code()
                    _mint_wl_bookmark(self._doc, rng, code)
                    existing[start] = code
                out[f"heading:{idx}"] = f"pin:{code}"
        return out

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
    def theme(self) -> DocumentTheme:
        """The document's theme — the document-wide brand primitive.

        See [`DocumentTheme`][wordlive.DocumentTheme]: `doc.theme.apply("Facet")`
        swaps the whole theme, `doc.theme.set_colors(accent1="#1A73E8")` /
        `doc.theme.set_fonts(major="Arial")` set brand colours/fonts, and
        `doc.theme.colors` / `doc.theme.to_dict()` read it back. Wrap mutations in
        `doc.edit(...)` for atomic undo.
        """
        return DocumentTheme(self)

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
        return ChartCollection(self)

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
        return ShapeCollection(self)

    @property
    def text_boxes(self) -> TextBoxCollection:
        """The text boxes among `doc.shapes` — the ``shape_type == "text_box"`` subset.

        A discovery filter, not a second id space: `doc.text_boxes[1]` returns a
        [`ShapeAnchor`][wordlive.ShapeAnchor] that keeps its canonical `shape:N`
        id (its position among *all* floating shapes). Created by any anchor's
        [`insert_text_box`][wordlive.Anchor.insert_text_box].
        """
        return TextBoxCollection(self)

    def group_shapes(self, *anchor_ids: str) -> ShapeAnchor:
        """Group two or more floating shapes into a single group shape.

        Each `anchor_id` is a `shape:N` (the members must all be floating shapes).
        Returns the new group's [`ShapeAnchor`][wordlive.ShapeAnchor] (`shape:N`,
        `shape_type == "group"`) — move / size / delete it as one unit, or
        [`ungroup`][wordlive.ShapeAnchor.ungroup] it to get the members back. Word
        requires members to allow overlap, so this enables that first. Wrap in
        `doc.edit(...)` for atomic undo. Raises `OpError` for fewer than two
        shapes or a non-shape anchor.
        """
        if len(anchor_ids) < 2:
            raise OpError("group_shapes needs at least two shape anchors")
        with _com.translate_com_errors():
            coms: list[Any] = []
            for aid in anchor_ids:
                anchor = self.anchor_by_id(aid)
                if not isinstance(anchor, ShapeAnchor):
                    raise OpError(f"{aid!r} is not a shape; group_shapes needs shape:N anchors")
                coms.append(anchor._shape())
            group = _shapes.group_shapes(self.com, coms)
            # Locate the new group by a unique temp name — don't assume "last".
            orig_name = str(group.Name or "")
            probe_name = f"_wl_shape_{secrets.token_hex(8)}"
            group.Name = probe_name
            index = _shapes.index_of_named(self.com, probe_name)
            if orig_name:
                group.Name = orig_name
        return ShapeAnchor(self, index)

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

    # Word's own text-watermark feature names its WordArt shapes with this prefix
    # (e.g. "PowerPlusWaterMarkObject357921"); reusing it means set_watermark
    # replaces a watermark the user added through the ribbon, and remove_watermark
    # finds it — the established convention, not a wordlive marker.
    _WATERMARK_NAME_PREFIX = "PowerPlusWaterMarkObject"

    def set_watermark(
        self,
        text: str,
        *,
        font: str = "Calibri",
        color: str = "#C0C0C0",
        layout: str = "diagonal",
        semitransparent: bool = True,
    ) -> int:
        """Stamp a text watermark (DRAFT / CONFIDENTIAL / …) behind every page.

        Adds a WordArt shape to each section's primary header story — the same
        mechanism (and shape name) as Word's *Design → Watermark → Custom*, so it
        shows behind the body text on every page and replaces any existing text
        watermark. `layout` is ``"diagonal"`` (default, rotated 45°) or
        ``"horizontal"``; `color` is the fill colour (``"#C0C0C0"`` / ``"red"``);
        `semitransparent` washes it out (50% transparency) so body text stays
        readable. Returns the number of sections stamped.

        Any prior watermark is cleared first, so calling it twice doesn't stack.
        Remove one with [`remove_watermark`][wordlive.Document.remove_watermark].
        Wrap in `doc.edit(...)` for atomic undo. Raises `OpError` for a bad
        `layout` or `color`.
        """
        if layout not in ("diagonal", "horizontal"):
            raise OpError(f"watermark layout must be 'diagonal' or 'horizontal'; got {layout!r}")
        try:
            fill_bgr = to_bgr(color)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        rotation = 315.0 if layout == "diagonal" else 0.0
        self.remove_watermark()
        with _com.translate_com_errors():
            sections = self._doc.Sections
            count = int(sections.Count)
            for s in range(1, count + 1):
                section = sections(s)
                header = section.Headers(int(WdHeaderFooterIndex.PRIMARY))
                ps = section.PageSetup
                usable = float(ps.PageWidth) - float(ps.LeftMargin) - float(ps.RightMargin)
                width = max(72.0, usable)
                # Shapes live on the HeaderFooter itself, not its Range (a Range
                # has no .Shapes) — this is the header story Word's own watermark
                # feature draws into.
                shape = header.Shapes.AddTextEffect(
                    PresetTextEffect=int(MsoPresetTextEffect.TEXT_EFFECT1),
                    Text=text,
                    FontName=font,
                    FontSize=1.0,  # WordArt scales to the box; explicit size below
                    FontBold=int(MsoTriState.FALSE),
                    FontItalic=int(MsoTriState.FALSE),
                    Left=0.0,
                    Top=0.0,
                )
                shape.Name = f"{self._WATERMARK_NAME_PREFIX}{s}"
                shape.TextEffect.NormalizedHeight = False
                shape.Line.Visible = int(MsoTriState.FALSE)
                shape.Fill.Visible = int(MsoTriState.TRUE)
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = fill_bgr
                shape.Fill.Transparency = 0.5 if semitransparent else 0.0
                shape.Rotation = rotation
                shape.LockAspectRatio = int(MsoTriState.TRUE)
                shape.Width = width
                shape.Height = width / 5.0
                shape.WrapFormat.AllowOverlap = True
                shape.WrapFormat.Side = int(WdWrapSideType.BOTH)
                shape.WrapFormat.Type = int(WdWrapType.BEHIND)
                shape.RelativeHorizontalPosition = int(WdRelativeHorizontalPosition.MARGIN)
                shape.RelativeVerticalPosition = int(WdRelativeVerticalPosition.MARGIN)
                shape.Left = float(WdShapePosition.CENTER)
                shape.Top = float(WdShapePosition.CENTER)
        return count

    def remove_watermark(self) -> int:
        """Remove any text watermark added by `set_watermark` (or Word's ribbon).

        Deletes every WordArt shape named like Word's watermark object across all
        sections' header stories. Returns the number of shapes removed (0 if there
        was no watermark). Wrap in `doc.edit(...)` for atomic undo.
        """
        removed = 0
        with _com.translate_com_errors():
            sections = self._doc.Sections
            for s in range(1, int(sections.Count) + 1):
                header = sections(s).Headers(int(WdHeaderFooterIndex.PRIMARY))
                shapes = header.Shapes
                # Delete back-to-front: removing a shape renumbers those after it.
                for i in range(int(shapes.Count), 0, -1):
                    shape = shapes(i)
                    if str(shape.Name or "").startswith(self._WATERMARK_NAME_PREFIX):
                        shape.Delete()
                        removed += 1
        return removed

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
        `doc.sources.add` and cite them with
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
        if kind == "pin":
            name = _pin_name_for(value)
            with _com.translate_com_errors():
                if not self._doc.Bookmarks.Exists(name):
                    # A vanished pin (its content was deleted) correctly misses.
                    raise AnchorNotFoundError("pin", anchor_id)
            return Bookmark.pin(self, value)
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

    def outline(self, *, pin: bool = False) -> list[dict[str, Any]]:
        """Return all heading paragraphs as `[{level, text, anchor_id}, ...]`.

        With `pin=True` each row also carries a durable `pin` id and the headings
        are pinned as a side effect (idempotent — see
        [`pin_outline`][wordlive.Document.pin_outline]). This **mutates** the
        document, so it is a Python-API-only convenience; the read surfaces
        (`wordlive read outline`, MCP `word_read outline`) stay pure — pin in bulk
        via `pin_outline` / the `pin_outline` exec op instead.
        """
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
        if pin:
            pinmap = self.pin_outline()
            for row in out:
                handle = pinmap.get(row["anchor_id"])
                if handle:
                    row["pin"] = handle
        return out

    def between(
        self,
        start: str | Anchor,
        end: str | Anchor,
        *,
        inclusive: bool = False,
    ) -> RangeAnchor:
        """Return a `RangeAnchor` spanning the gap between two anchors.

        The "give me the block between these two headings" read. `start` and
        `end` are anchor ids (e.g. ``"heading:1"`` / ``"heading:3"``) or
        `Anchor` objects; the headline use is a pair of `heading:N` ids, but any
        anchors work (bookmarks, paragraphs, ranges).

        With ``inclusive=False`` (default) the span runs from the **end** of
        `start`'s range to the **start** of `end`'s range — the content strictly
        between them, excluding both bounding paragraphs (so two headings yield
        just the body in between). With ``inclusive=True`` it runs from the
        start of `start` to the end of `end`, covering both bounding paragraphs.

        Read ``.text`` on the result for the spanned text, or feed its
        `range:START-END` id into any range-taking op. A pure read (the returned
        offsets are live — use them before further edits shift the document).
        Raises `OpError` if `end` begins before `start`.
        """
        with _com.translate_com_errors():
            s_anchor = self.anchor_by_id(start) if isinstance(start, str) else start
            e_anchor = self.anchor_by_id(end) if isinstance(end, str) else end
            s_rng, e_rng = s_anchor.com, e_anchor.com
            s_start, s_end = int(s_rng.Start), int(s_rng.End)
            e_start, e_end = int(e_rng.Start), int(e_rng.End)
        if e_start < s_start:
            raise OpError(
                f"'between' end anchor ({e_anchor.anchor_id}) begins before start "
                f"anchor ({s_anchor.anchor_id})"
            )
        if inclusive:
            lo, hi = min(s_start, e_start), max(s_end, e_end)
        else:
            # Strictly between: end of `start` to start of `end`. When the anchors
            # abut with no gap, clamp to an empty span at the boundary.
            lo, hi = s_end, max(s_end, e_start)
        return self.range(lo, hi)

    def nearest_heading(
        self,
        where: str | Anchor | int,
        *,
        direction: str = "before",
    ) -> dict[str, Any] | None:
        """The heading nearest to a position, scanning ``before`` or ``after`` it.

        `where` is an anchor id (``"para:12"``), an `Anchor`, or a raw character
        offset (int). `direction` is ``"before"`` (the nearest heading at or
        above the position — i.e. the section the position sits in) or
        ``"after"`` (the next heading past it). Returns an `outline()`-shaped
        row ``{level, text, anchor_id}`` (``anchor_id`` is ``heading:N``), or
        ``None`` if there is no heading in that direction. A pure read.
        """
        if direction not in ("before", "after"):
            raise OpError(f"direction must be 'before' or 'after', got {direction!r}")
        with _com.translate_com_errors():
            if isinstance(where, str):
                offset = int(self.anchor_by_id(where).com.Start)
            elif isinstance(where, int):  # raw character offset
                offset = int(where)
            else:
                offset = int(where.com.Start)
            best: dict[str, Any] | None = None
            for idx, para in enumerate(self._doc.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    continue
                if level >= 10:  # body text, not a heading
                    continue
                h_start = int(para.Range.Start)
                row = {"level": level, "text": paragraph_text(para), "anchor_id": f"heading:{idx}"}
                if direction == "before":
                    if h_start <= offset:
                        best = row  # paragraphs are in order; keep the last one at/above
                    else:
                        break
                elif h_start > offset:  # "after": first heading strictly past the offset
                    best = row
                    break
        return best

    def find_paragraphs(
        self,
        text: str,
        *,
        limit: int = 5,
        min_score: float = 0.6,
    ) -> list[dict[str, Any]]:
        """Fuzzy-rank paragraphs by similarity to `text` (typo/paraphrase tolerant).

        Unlike `find()` (exact substring on normalized text), this scores
        **every paragraph** against `text` with `difflib.SequenceMatcher` over
        the same normalized form (NFKC, smart quotes, dashes, whitespace) — so
        an approximately-remembered paragraph still locates its `para:N`.
        Returns up to `limit` rows, sorted by descending `score`, keeping only
        those with ``score >= min_score``:
        ``[{anchor_id, index, score, text, level, is_heading}, ...]``. An empty
        or whitespace-only query returns ``[]``. A pure read.
        """
        if limit < 1:
            raise OpError(f"limit must be >= 1, got {limit}")
        if not 0.0 <= min_score <= 1.0:
            raise OpError(f"min_score must be in [0, 1], got {min_score}")
        needle = _findreplace._normalize(text).text
        if not needle:
            return []
        scored: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.Paragraphs, start=1):
                raw = paragraph_text(para)
                hay = _findreplace._normalize(raw).text
                if not hay:
                    continue
                score = difflib.SequenceMatcher(None, needle, hay).ratio()
                if score < min_score:
                    continue
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    level = 10
                scored.append(
                    {
                        "anchor_id": f"para:{idx}",
                        "index": idx,
                        "score": round(score, 4),
                        "text": raw,
                        "level": level,
                        "is_heading": level < 10,
                    }
                )
        scored.sort(key=lambda r: r["score"], reverse=True)
        return scored[:limit]

    def stats(self) -> dict[str, Any]:
        """A one-call summary of the document — the "what am I looking at" read.

        Returns `{pages, words, characters, paragraphs, lines, sections,
        headings, tables, images, equations, charts, comments, revisions, saved}`. The five text
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
            "charts": len(self.charts),
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

    def lint(
        self, *, rules: Any = None, within: str | Anchor | None = None
    ) -> list[dict[str, Any]]:
        """Audit the document for publishing-quality defects — a pure read.

        Returns a severity-ranked list of findings, each
        `{rule, kind, severity, anchor_id, message, fixable, fix, observed,
        expected}`. `kind` is `consistency` (a direct override fighting the
        applied style — a `Heading 1` at 15pt), `structural` (an objective layout
        defect — a heading that may dangle at a page foot, a multi-page table with
        no repeating header, a numbered list Word split into independent "1."
        runs), or `policy` (a house-style target — none ship yet). A `fixable`
        finding carries an op-shaped `fix` describing exactly what
        [`regularize`][wordlive.Document.regularize] would change.

        `rules` selects which rules run: `None` is the default set (all
        consistency + structural); a list of rule ids / tags
        (`["headings", "lists"]`) includes only those; `{"exclude": [...]}` runs
        the default set minus the listed ids/tags. `within=anchor` scopes the
        audit to an anchor's range (a heading's section, a `range:`, a table).

        Read-only — selection, scroll, and `Saved` are untouched (the layout
        rules repaginate content-neutrally, like [`stats`][wordlive.Document.stats]).
        """
        return [f.to_dict() for f in _linting.run_lint(self, rules=rules, within=within)]

    def regularize(
        self,
        *,
        rules: Any = None,
        within: str | Anchor | None = None,
        dry_run: bool = False,
    ) -> dict[str, Any]:
        """Apply the fixable [`lint`][wordlive.Document.lint] findings in one
        atomic-undo step. Returns `{applied, skipped, findings}`.

        Each fixable finding's `fix` op(s) run through the batch op loop inside a
        single `doc.edit("Regularize formatting")`, so one Ctrl-Z reverts the
        whole pass and the user's selection/scroll are preserved. The default
        fixes are **targeted and idempotent** — they bring a drifted direct
        override back to its style's value, so running `regularize` twice applies
        nothing the second time. `rules` / `within` are as for `lint`.
        `dry_run=True` plans the fixes (returning them in `findings`) without
        writing.

        Content-changing fixes (deleting stray paragraphs, inserting captions)
        are out of scope here — formatting/structure fixes only. If Track Changes
        is on, the edits are tracked like any other for the user to review.
        """
        return _linting.regularize(self, rules=rules, within=within, dry_run=dry_run)

    def checkpoint(
        self,
        *,
        include: str = "text+style",
        within: str | Anchor | None = None,
    ) -> Checkpoint:
        """Fingerprint the document's structure right now — a pure read.

        Returns an opaque, serialisable [`Checkpoint`][wordlive.Checkpoint] (call
        `.to_json()` to store it). Later, feed it to
        [`changes_since`][wordlive.Document.changes_since] (checkpoint → now) or
        [`diff`][wordlive.Document.diff] (two stored checkpoints) for a structured,
        content-aligned change list — the only reliable way to answer "what
        changed in session" (Word emits no content-change event), and the way an
        agent verifies its own edits landed without re-reading the whole document.

        `include` sets the fingerprint depth: ``"text"`` (cheapest — a restyle is
        invisible), ``"text+style"`` (default — folds the applied paragraph-style
        name in, so a restyle surfaces), or ``"text+format"`` (also hashes each
        paragraph's `format_info`, so a pure direct-formatting edit surfaces as a
        `reformat`). `within=anchor` fingerprints just one section/range.

        Read-only — walks paragraphs like [`outline`][wordlive.Document.outline],
        touching no selection/scroll and leaving `Saved` untouched.
        """
        return _checkpoint.build_checkpoint(self, include=include, within=within)

    def changes_since(self, cp: Checkpoint | str | dict[str, Any]) -> list[dict[str, Any]]:
        """Diff a stored checkpoint against the document **now** — a pure read.

        `cp` is a [`Checkpoint`][wordlive.Checkpoint] (or its `to_json()` string /
        parsed dict, so a token round-tripped through a file works directly).
        Returns the change list described in [`diff`][wordlive.Document.diff]; the
        checkpoint's `include` depth and `within` scope are re-derived so the two
        fingerprints are comparable. An unchanged document returns ``[]`` via the
        `doc_hash` fast-path.
        """
        return _checkpoint.changes_since(self, cp)

    def diff(
        self,
        cp_a: Checkpoint | str | dict[str, Any],
        cp_b: Checkpoint | str | dict[str, Any],
    ) -> list[dict[str, Any]]:
        """Diff two stored checkpoints → a structured, content-aligned change list.

        Each change is one of: ``replace`` (text edit), ``insert``, ``delete``,
        ``restyle`` (same text, paragraph style changed), or ``reformat`` (same
        text+style, direct formatting changed — only with ``include="text+format"``).
        Records carry ``{op, anchor_id, index_before, index_after, text_before,
        text_after, style_before, style_after}`` as applicable; inserts/replaces/
        restyles carry the **current** ``para:N`` (`anchor_id`) so the caller can
        act on the change immediately, while a delete references only the old
        index/text (its anchor is gone).

        Alignment is by paragraph **content**, not index (`para:N` renumbers under
        inserts/deletes). Both checkpoints must share the same `include` depth.
        Move detection is deferred — a cut-paste surfaces as delete+insert. A pure
        read (the tokens carry the data; Word is not touched).
        """
        return _checkpoint.diff_checkpoints(cp_a, cp_b)

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
