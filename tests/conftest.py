"""Test fixtures.

`fake_word` builds a MagicMock that quacks like a Word.Application COM object —
enough surface to exercise the politeness logic, anchors, and CLI shape without
needing a real Word install.

`real_word` is for the @pytest.mark.smoke suite and is skipped if Word can't
be reached.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import Any
from unittest.mock import MagicMock

import pytest

# ---------------------------------------------------------------------------
# Fake COM Application
# ---------------------------------------------------------------------------


class _FakeBookmarkRegistry:
    """Mimics ActiveDocument.Bookmarks: Exists, by-name lookup, Add, iteration.

    Bookmark mocks (and their `.Range`) are cached per name so a test that
    writes via `Bookmarks(name).Range.X = y` can read `X` back through the
    same handle. `Add` and the seeding helper invalidate the cache so the
    range offsets stay consistent with the registered tuple.
    """

    def __init__(self) -> None:
        self._items: dict[str, tuple[int, int]] = {}
        self._cache: dict[str, Any] = {}

    def add(self, name: str, start: int, end: int) -> None:
        self._items[name] = (start, end)
        self._cache.pop(name, None)

    # COM-style methods
    def Exists(self, name: str) -> bool:
        return name in self._items

    def Add(self, Name: str | None = None, Range: Any | None = None) -> None:
        if Name is None or Range is None:
            return
        self._items[Name] = (int(Range.Start), int(Range.End))
        self._cache.pop(Name, None)

    def __call__(self, name: str) -> Any:
        if name in self._cache:
            return self._cache[name]
        s, e = self._items[name]
        bm = MagicMock(name=f"Bookmark[{name}]")
        bm.Name = name
        bm.Range = _make_range(s, e)
        self._cache[name] = bm
        return bm

    def __iter__(self) -> Iterable[Any]:
        return iter([self(name) for name in self._items])


class _FakeContentControls:
    def __init__(self, controls: list[dict[str, Any]]) -> None:
        self._items = []
        for cc in controls:
            mock = MagicMock(name=f"CC[{cc.get('title', '?')}]")
            mock.Title = cc.get("title", "")
            mock.Tag = cc.get("tag", "")
            mock.Range = _make_range(cc.get("start", 0), cc.get("end", 0))
            mock.Range.Text = cc.get("text", "")
            self._items.append(mock)

    def __iter__(self) -> Iterable[Any]:
        return iter(self._items)


class _FakeParagraphs:
    def __init__(self, paragraphs: list[dict[str, Any]]) -> None:
        self._items = []
        for p in paragraphs:
            mock = MagicMock(name=f"Para[{p.get('text', '?')[:20]}]")
            mock.OutlineLevel = p.get("level", 10)
            mock.Range = _make_range(p.get("start", 0), p.get("end", 0))
            mock.Range.Text = p.get("text", "") + "\r"
            # Applied paragraph style name, surfaced by ParagraphCollection.list().
            mock.Range.Style.NameLocal = p.get("style", "Normal")
            self._items.append(mock)

    def __iter__(self) -> Iterable[Any]:
        return iter(self._items)


class _FakeListFormat:
    """Mimics Range.ListFormat: apply/remove templates, in/out-dent, read state.

    `ApplyListTemplate` reads the gallery off the template handed in by
    `_list_application` (1=bullet, 2=number, 3=outline) and updates the
    readable list state to match, so apply -> list_info round-trips. A template
    is `None` until a list is applied, so `restart_numbering` on a plain range
    raises (matching Word).
    """

    _GALLERY_TO_TYPE = {1: 2, 2: 3, 3: 4}  # bullet->BULLET, number->SIMPLE, outline->OUTLINE
    _MARKER = {2: "•", 3: "1.", 4: "1."}

    def __init__(self) -> None:
        self.ListType = 0
        self.ListLevelNumber = 1
        self.ListValue = 0
        self.ListString = ""
        self._gallery: int | None = None
        self._continue = False

    def ApplyListTemplate(
        self, ListTemplate=None, ContinuePreviousList=False, ApplyTo=0, DefaultListBehavior=2, **kw
    ):
        gallery = getattr(ListTemplate, "_gallery", None)
        self._gallery = int(gallery) if gallery is not None else None
        self.ListType = self._GALLERY_TO_TYPE.get(self._gallery, 3)
        self._continue = bool(ContinuePreviousList)
        self.ListValue = 1 if self.ListType in (3, 4) else 0
        self.ListString = self._MARKER.get(self.ListType, "")
        self.ListLevelNumber = 1

    def RemoveNumbers(self, NumberType=3):
        self.ListType = 0
        self.ListLevelNumber = 1
        self.ListValue = 0
        self.ListString = ""
        self._gallery = None

    def ListIndent(self) -> None:
        self.ListLevelNumber = min(9, self.ListLevelNumber + 1)

    def ListOutdent(self) -> None:
        self.ListLevelNumber = max(1, self.ListLevelNumber - 1)

    @property
    def ListTemplate(self):
        if self._gallery is None:
            return None
        t = MagicMock(name="ListTemplate")
        t._gallery = self._gallery
        return t


def _list_application() -> MagicMock:
    """A stand-in Application whose ListGalleries(n).ListTemplates(1) carries gallery n."""
    app = MagicMock(name="ListApplication")

    def galleries(gallery_type):
        gallery = MagicMock(name=f"Gallery[{gallery_type}]")

        def templates(n):
            t = MagicMock(name=f"Template[{gallery_type}:{n}]")
            t._gallery = int(gallery_type)
            return t

        gallery.ListTemplates.side_effect = templates
        return gallery

    app.ListGalleries.side_effect = galleries
    return app


class _FakeWrapFormat:
    """Mimics Shape.WrapFormat — only `.Type` is exercised."""

    def __init__(self) -> None:
        self.Type = 7  # wdWrapInline until ConvertToShape sets a real wrap


class _FakeShape:
    """A floating Shape, as produced by InlineShape.ConvertToShape()."""

    def __init__(self, width: float, height: float, alt_text: str) -> None:
        self.Width = width
        self.Height = height
        self.AlternativeText = alt_text
        self.WrapFormat = _FakeWrapFormat()


class _FakeInlineShape:
    """The InlineShape returned by Range.InlineShapes.AddPicture()."""

    def __init__(self) -> None:
        self.Width = 100.0
        self.Height = 80.0
        self.AlternativeText = ""
        self.LockAspectRatio = -1
        self.converted = None  # the _FakeShape, once ConvertToShape() runs

    def ConvertToShape(self) -> _FakeShape:
        self.converted = _FakeShape(self.Width, self.Height, self.AlternativeText)
        return self.converted


class _FakeInlineShapes:
    """Mimics Range.InlineShapes: AddPicture records its call and returns a shape."""

    def __init__(self) -> None:
        self.shape = _FakeInlineShape()
        self.AddPicture = MagicMock(name="AddPicture", return_value=self.shape)

    @property
    def Count(self) -> int:
        # No picture present unless a test wires one up; range_text() reads this
        # to decide whether to tokenize inline shapes.
        return 0


class _FakeBorders:
    """Mimics `Range.Borders` — a callable vending one stable child per edge index.

    A bare MagicMock would return the *same* `return_value` for `Borders(-1)` and
    `Borders(-3)`, so a per-side assert (top vs. bottom) couldn't tell them apart.
    Memoising one child per index keeps each edge's `LineStyle`/`LineWidth`/`Color`
    distinct, which is what `set_borders` writes.
    """

    def __init__(self) -> None:
        self._edges: dict[int, MagicMock] = {}

    def __call__(self, index: int) -> MagicMock:
        return self._edges.setdefault(int(index), MagicMock(name=f"Border[{index}]"))


def _make_range(start: int, end: int) -> MagicMock:
    rng = MagicMock(name=f"Range[{start},{end}]")
    rng.Start = start
    rng.End = end
    rng.Text = ""
    # List support: a stateful ListFormat plus an Application that vends
    # list-gallery templates, so apply_list / list_info / restart work.
    rng.ListFormat = _FakeListFormat()
    rng.Application = _list_application()
    # Image support: every range can take an inline picture and knows its
    # section's page geometry (Letter defaults) for the wrap="auto" heuristic.
    rng.InlineShapes = _FakeInlineShapes()
    rng.PageSetup = _FakePageSetup()
    # A plain range contains no tables; find/replace's segment fast-path checks
    # Range.Tables.Count to decide whether table-boundary segmentation is needed.
    # Tests that exercise the in-table path set a positive Count on their range.
    rng.Tables = MagicMock(name="RangeTables")
    rng.Tables.Count = 0
    # Snapshot support: Range.Information(wdActiveEndPageNumber=3) reports the
    # page a (collapsed) range sits on. The tiny fake document is all one page,
    # so report page 1 for any query; tests that need a specific page set
    # `rng.Information.return_value` on the cached range themselves.
    rng.Information = MagicMock(name="Information", return_value=1)
    # Border support: one stable child per edge so per-side asserts work.
    rng.Borders = _FakeBorders()
    return rng


_DEFAULT_STYLES: tuple[dict[str, Any], ...] = (
    {"name": "Normal", "type": 1, "builtin": True, "in_use": True},
    {"name": "Body Text", "type": 1, "builtin": True, "in_use": True},
    {"name": "Heading 1", "type": 1, "builtin": True, "in_use": True},
    {"name": "Heading 2", "type": 1, "builtin": True, "in_use": True},
)


class _FakeStyles:
    """Mimics ActiveDocument.Styles: iterable of objects exposing NameLocal/Type/BuiltIn/InUse.

    Note: real Word's `Styles(name)` raises a generic com_error for missing
    names. wordlive validates membership via iteration *first*, so the fake
    only needs to support iteration. We still implement __call__ for the rare
    direct lookup path; it raises KeyError if the style isn't there.
    """

    def __init__(self, styles: tuple[dict[str, Any], ...] | list[dict[str, Any]]) -> None:
        self._items: list[Any] = []
        for s in styles:
            mock = MagicMock(name=f"Style[{s['name']}]")
            mock.NameLocal = s["name"]
            mock.Type = s.get("type", 1)
            mock.BuiltIn = s.get("builtin", True)
            mock.InUse = s.get("in_use", True)
            self._items.append(mock)

    def __iter__(self) -> Iterable[Any]:
        return iter(self._items)

    def __call__(self, name: str) -> Any:
        for s in self._items:
            if s.NameLocal == name:
                return s
        raise KeyError(name)

    def Add(self, Name: str | None = None, Type: int = 1) -> Any:
        """Mimic Styles.Add(Name, Type): append a new style and return its mock.

        `StyleCollection.add` then re-resolves it by name via `Style.com`'s
        direct-then-iterate lookup, so the new style must be iterable here.
        """
        mock = MagicMock(name=f"Style[{Name}]")
        mock.NameLocal = Name
        mock.Type = int(Type)
        mock.BuiltIn = False
        mock.InUse = False
        self._items.append(mock)
        return mock


class _FakeTable:
    """Mimics a Word Table COM object: Rows/Columns counts, Cell(r,c), row add/delete.

    Cells are backed by persistent MagicMock ranges so writes (`Range.Text`,
    `Range.Style`, `Range.ParagraphFormat`) round-trip through the same handle.
    Cell text is seeded with Word's trailing `\\r\\x07` markers so the marker-
    stripping path in `Cell.text` gets exercised.
    """

    def __init__(
        self,
        grid: list[list[str]],
        title: str = "",
        *,
        start: int = 0,
        owner: Any | None = None,
    ) -> None:
        self.Title = title
        self.Style: Any = None
        self._start = start
        self._owner = owner
        self._rows = [[self._mk_cell_range(text) for text in row] for row in grid]

    @staticmethod
    def _mk_cell_range(text: str) -> MagicMock:
        rng = MagicMock(name="CellRange")
        # Word terminates cell text with CR + the cell mark; populated cells
        # (set via Range.Text) carry no markers, so only seed them for non-empty
        # seed text to mirror both states.
        rng.Text = (text + "\r\x07") if text else "\r\x07"
        rng.Start = 0
        # No inline shapes by default, so range_text() reads cell text verbatim.
        ish = MagicMock(name="CellInlineShapes")
        ish.Count = 0
        rng.InlineShapes = ish
        return rng

    @property
    def Range(self) -> Any:
        rng = MagicMock(name="TableRange")
        rng.Start = self._start
        return rng

    @property
    def Rows(self) -> Any:
        return _FakeRows(self._rows, self._mk_cell_range)

    @property
    def Columns(self) -> Any:
        cols = MagicMock(name="Columns")
        cols.Count = len(self._rows[0]) if self._rows else 0
        return cols

    def Cell(self, row: int, col: int) -> Any:
        cell = MagicMock(name=f"Cell[{row},{col}]")
        cell.Range = self._rows[row - 1][col - 1]
        return cell

    def Delete(self) -> None:
        if self._owner is not None:
            self._owner._remove(self)


class _FakeRows:
    """Rows view: Count, Add() (append at end), Rows(i).Delete().

    Shares the parent table's row list so structural edits persist.
    """

    def __init__(self, rows: list[list[Any]], mk: Any) -> None:
        self._rows = rows
        self._mk = mk

    @property
    def Count(self) -> int:
        return len(self._rows)

    def Add(self, BeforeRow: Any | None = None) -> None:
        ncols = len(self._rows[0]) if self._rows else 1
        self._rows.append([self._mk("") for _ in range(ncols)])

    def __call__(self, index: int) -> Any:
        row = MagicMock(name=f"Row[{index}]")
        rows = self._rows
        row.Delete = lambda: rows.__delitem__(index - 1)
        return row


class _FakeTablesCollection:
    """Mimics doc.Tables: Count, 1-based call lookup, iteration, Add, Delete.

    `Add(Range, NumRows, NumColumns)` builds an empty `_FakeTable` seeded with
    the insertion range's `Start` and inserts it in document order (sorted by
    Start), so `index_of` can recover its 1-based position the same way it does
    against real Word.
    """

    def __init__(self, tables: list[_FakeTable]) -> None:
        self._tables = tables
        for t in tables:
            t._owner = self

    @property
    def Count(self) -> int:
        return len(self._tables)

    def __call__(self, index: int) -> _FakeTable:
        return self._tables[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(self._tables)

    def Add(
        self, Range: Any = None, NumRows: int = 1, NumColumns: int = 1, *args: Any, **kwargs: Any
    ) -> _FakeTable:
        start = int(getattr(Range, "Start", 0)) if Range is not None else 0
        grid = [["" for _ in range(int(NumColumns))] for _ in range(int(NumRows))]
        t = _FakeTable(grid, start=start, owner=self)
        self._tables.append(t)
        self._tables.sort(key=lambda x: x._start)
        return t

    def _remove(self, table: _FakeTable) -> None:
        self._tables.remove(table)


class _FakeComment:
    """Mimics a Word Comment: Author, Range (body), Scope (anchored range), Done."""

    def __init__(self, registry: _FakeComments, index: int, scope: Any, text: str) -> None:
        self._registry = registry
        self.Index = index
        self.Author = ""
        body = MagicMock(name=f"CommentBody[{index}]")
        body.Text = text
        self.Range = body
        self.Scope = scope if scope is not None else _make_range(0, 0)
        self.Done = False

    def Delete(self) -> None:
        self._registry._remove(self)


class _FakeComments:
    """Mimics doc.Comments: Count, Add(Range, Text), 1-based call lookup, iteration."""

    def __init__(self) -> None:
        self._items: list[_FakeComment] = []

    @property
    def Count(self) -> int:
        return len(self._items)

    def Add(self, Range: Any = None, Text: str = "") -> _FakeComment:
        c = _FakeComment(self, len(self._items) + 1, Range, str(Text))
        self._items.append(c)
        return c

    def __call__(self, index: int) -> _FakeComment:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))

    def _remove(self, comment: _FakeComment) -> None:
        self._items.remove(comment)
        for i, c in enumerate(self._items, start=1):
            c.Index = i


class _FakeNote:
    """Mimics a Word Footnote/Endnote: a body Range, a Reference range, Delete."""

    def __init__(self, registry: _FakeNotes, ref_start: int, text: str) -> None:
        self._registry = registry
        body = MagicMock(name="NoteBody")
        body.Text = text
        ish = MagicMock(name="NoteInlineShapes")
        ish.Count = 0
        body.InlineShapes = ish
        self.Range = body
        ref = MagicMock(name="NoteReference")
        ref.Start = int(ref_start)
        ref.Delete = MagicMock(name="ReferenceDelete")
        self.Reference = ref

    def Delete(self) -> None:
        self._registry._remove(self)


class _FakeNotes:
    """Mimics doc.Footnotes / doc.Endnotes: Count, Add(Range, Reference, Text), lookup.

    `Add` is a MagicMock (side-effect appends a `_FakeNote`) so tests can assert
    the positional call args — notably the empty `Reference` that auto-numbers.
    """

    def __init__(self) -> None:
        self._items: list[_FakeNote] = []
        self.Add = MagicMock(name="NotesAdd", side_effect=self._add)

    def _add(self, Range: Any = None, Reference: str = "", Text: str = "") -> _FakeNote:
        ref_start = int(getattr(Range, "Start", 0)) if Range is not None else 0
        note = _FakeNote(self, ref_start, str(Text))
        self._items.append(note)
        # Footnotes/Endnotes are ordered by document position, like real Word.
        self._items.sort(key=lambda n: int(n.Reference.Start))
        return note

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeNote:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))

    def _remove(self, note: _FakeNote) -> None:
        self._items.remove(note)


class _FakeTOC:
    """Mimics a Word TableOfContents: a Range plus Update / UpdatePageNumbers."""

    def __init__(self) -> None:
        rng = MagicMock(name="TOCRange")
        rng.Text = "Table of Contents"
        ish = MagicMock(name="TOCInlineShapes")
        ish.Count = 0
        rng.InlineShapes = ish
        self.Range = rng
        self.Update = MagicMock(name="TOCUpdate")
        self.UpdatePageNumbers = MagicMock(name="TOCUpdatePageNumbers")


class _FakeTOCs:
    """Mimics doc.TablesOfContents: Add(...) records its (positional) args, Count, lookup."""

    def __init__(self) -> None:
        self._items: list[_FakeTOC] = []
        # A MagicMock wrapper so tests can assert the positional Add arguments.
        self.Add = MagicMock(name="TOCAdd", side_effect=self._add)

    def _add(self, *args: Any, **kwargs: Any) -> _FakeTOC:
        toc = _FakeTOC()
        self._items.append(toc)
        return toc

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeTOC:
        return self._items[index - 1]


class _FakeWordList:
    """Mimics a Word List COM object: a Range plus a ListParagraphs count."""

    def __init__(self, rng: Any, count: int) -> None:
        self.Range = rng
        lp = MagicMock(name="ListParagraphs")
        lp.Count = count
        self.ListParagraphs = lp


class _FakeLists:
    """Mimics doc.Lists: Count, 1-based call lookup, iteration."""

    def __init__(self, lists: list[_FakeWordList]) -> None:
        self._lists = lists

    @property
    def Count(self) -> int:
        return len(self._lists)

    def __call__(self, index: int) -> _FakeWordList:
        return self._lists[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(self._lists)


_WHICH_INDEX = {"primary": 1, "first": 2, "even": 3}


class _FakeHeaderFooter:
    """Mimics a Word HeaderFooter: a settable Range, Exists, LinkToPrevious."""

    def __init__(self, text: str = "", *, exists: bool = True, linked: bool = False) -> None:
        rng = _make_range(0, len(text))
        rng.Text = text
        self.Range = rng
        self.Exists = exists
        self.LinkToPrevious = linked


class _FakeHeadersFooters:
    """Mimics Section.Headers / Section.Footers: 1/2/3-indexed call lookup."""

    def __init__(self, mapping: dict[int, _FakeHeaderFooter]) -> None:
        self._m = mapping

    def __call__(self, index: int) -> _FakeHeaderFooter:
        return self._m[int(index)]


class _FakeTextColumns:
    """Mimics PageSetup.TextColumns: SetCount records its arg; Count/Spacing settable."""

    def __init__(self) -> None:
        self.Count = 1
        self.Spacing = 36.0
        self.SetCount = MagicMock(name="SetCount")


class _FakePageSetup:
    def __init__(self, **kw: Any) -> None:
        self.Orientation = kw.get("orientation", 0)
        self.TopMargin = kw.get("top", 72.0)
        self.BottomMargin = kw.get("bottom", 72.0)
        self.LeftMargin = kw.get("left", 72.0)
        self.RightMargin = kw.get("right", 72.0)
        self.PageWidth = kw.get("width", 612.0)
        self.PageHeight = kw.get("height", 792.0)
        # Item-4 write surface (set_page_setup): paper size, gutter, columns.
        self.PaperSize = kw.get("paper_size", 2)  # wdPaperLetter
        self.Gutter = kw.get("gutter", 0.0)
        self.TextColumns = _FakeTextColumns()


class _FakeSection:
    def __init__(
        self,
        index: int,
        headers: dict[int, _FakeHeaderFooter],
        footers: dict[int, _FakeHeaderFooter],
        page_setup: _FakePageSetup,
    ) -> None:
        self.Index = index
        self.Headers = _FakeHeadersFooters(headers)
        self.Footers = _FakeHeadersFooters(footers)
        self.PageSetup = page_setup


class _FakeSections:
    """Mimics doc.Sections: Count, 1-based call lookup, iteration."""

    def __init__(self, sections: list[_FakeSection]) -> None:
        self._s = sections

    @property
    def Count(self) -> int:
        return len(self._s)

    def __call__(self, index: int) -> _FakeSection:
        return self._s[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(self._s)


def _build_hf_map(texts: dict[str, str]) -> dict[int, _FakeHeaderFooter]:
    """All three header/footer indices, seeded from a {which: text} dict."""
    return {idx: _FakeHeaderFooter(texts.get(name, "")) for name, idx in _WHICH_INDEX.items()}


def _make_document(
    *,
    name: str = "Test.docx",
    full_name: str = r"C:\test\Test.docx",
    bookmarks: dict[str, tuple[int, int]] | None = None,
    content_controls: list[dict[str, Any]] | None = None,
    paragraphs: list[dict[str, Any]] | None = None,
    content: str = "",
    styles: tuple[dict[str, Any], ...] | list[dict[str, Any]] | None = None,
    tables: list[dict[str, Any]] | None = None,
    lists: list[dict[str, Any]] | None = None,
    sections: list[dict[str, Any]] | None = None,
    footnotes: list[dict[str, Any]] | None = None,
    endnotes: list[dict[str, Any]] | None = None,
) -> MagicMock:
    doc = MagicMock(name=f"Document[{name}]")
    doc.Name = name
    doc.FullName = full_name

    bm_registry = _FakeBookmarkRegistry()
    for bm_name, (s, e) in (bookmarks or {}).items():
        bm_registry.add(bm_name, s, e)
    doc.Bookmarks = bm_registry

    # Cross-reference items: Word returns ordered strings per reference type —
    # bookmark *names* (type 2) and heading *texts* (type 1) — which
    # `insert_cross_reference` indexes into. Footnote/endnote types use the
    # 1-based index directly, so they're not needed here.
    def _xref_items(ref_type: Any) -> tuple[str, ...]:
        rt = int(ref_type)
        if rt == 2:  # wdRefTypeBookmark
            return tuple((bookmarks or {}).keys())
        if rt == 1:  # wdRefTypeHeading
            return tuple(p["text"] for p in (paragraphs or []) if p.get("level", 10) < 10)
        return ()

    doc.GetCrossReferenceItems = MagicMock(side_effect=_xref_items)

    doc.ContentControls = _FakeContentControls(content_controls or [])
    doc.Paragraphs = _FakeParagraphs(paragraphs or [])
    doc.Styles = _FakeStyles(styles if styles is not None else _DEFAULT_STYLES)
    doc.Tables = _FakeTablesCollection(
        [_FakeTable(t["grid"], t.get("title", "")) for t in (tables or [])]
    )
    doc.Comments = _FakeComments()

    # Track Changes flag: a plain settable bool so doc.track_changes /
    # tracked_changes() round-trip through the same handle.
    doc.TrackRevisions = False

    # `Range(start, end)` returns a range whose `.Text` can be set/read by the
    # caller. Cache by (start, end) so a write through one handle is visible on
    # the next lookup of the same span — needed for RangeAnchor round-trips.
    _range_cache: dict[tuple[int, int], MagicMock] = {}

    def _range_factory(start: int, end: int) -> MagicMock:
        key = (int(start), int(end))
        rng = _range_cache.get(key)
        if rng is None:
            rng = _make_range(start, end)
            _range_cache[key] = rng
        return rng

    doc.Range.side_effect = _range_factory

    # `Content` is the full document range; find/replace reads `.Text` from it.
    content_range = _make_range(0, len(content))
    content_range.Text = content
    doc.Content = content_range

    # Lists: build each list's Range through the cached factory so list_info /
    # restart_numbering see the same ListFormat the RangeAnchor resolves to.
    list_objs: list[_FakeWordList] = []
    for spec in lists or []:
        rng = _range_factory(spec["start"], spec["end"])
        lf = rng.ListFormat
        lf.ListType = spec.get("type", 3)
        lf.ListValue = 1 if lf.ListType in (3, 4) else 0
        lf.ListString = {2: "•", 3: "1.", 4: "1."}.get(lf.ListType, "")
        lf._gallery = {2: 1, 3: 2, 4: 3}.get(lf.ListType)
        list_objs.append(_FakeWordList(rng, spec.get("count", 0)))
    doc.Lists = _FakeLists(list_objs)

    # Sections: at least one (real documents always have one).
    section_specs = sections if sections is not None else [{}]
    section_objs: list[_FakeSection] = []
    for i, spec in enumerate(section_specs, start=1):
        section_objs.append(
            _FakeSection(
                i,
                _build_hf_map(spec.get("headers", {})),
                _build_hf_map(spec.get("footers", {})),
                _FakePageSetup(**spec.get("page_setup", {})),
            )
        )
    doc.Sections = _FakeSections(section_objs)

    # Footnotes / endnotes: seed each note's reference at the given offset (so
    # `list()` can map it back to a paragraph). Empty by default.
    def _seed_notes(specs: list[dict[str, Any]] | None) -> _FakeNotes:
        notes = _FakeNotes()
        for spec in specs or []:
            ref_start = int(spec.get("ref_start", 0))
            notes.Add(_range_factory(ref_start, ref_start), "", spec.get("text", ""))
        return notes

    doc.Footnotes = _seed_notes(footnotes)
    doc.Endnotes = _seed_notes(endnotes)
    doc.TablesOfContents = _FakeTOCs()

    return doc


def _make_application(documents: list[MagicMock]) -> MagicMock:
    app = MagicMock(name="Word.Application")
    app.Visible = True

    # Documents collection: indexable + iterable + Count + ActiveDocument
    docs_collection = MagicMock(name="Documents")
    docs_collection.Count = len(documents)
    docs_collection.__iter__ = MagicMock(side_effect=lambda: iter(documents))
    app.Documents = docs_collection
    app.ActiveDocument = documents[0] if documents else MagicMock(name="NoActiveDoc")

    # Selection / view
    app.Selection.Start = 0
    app.Selection.End = 0
    app.Selection.Text = ""
    app.ActiveWindow.VerticalPercentScrolled = 0

    # UndoRecord
    app.UndoRecord.StartCustomRecord = MagicMock(name="StartCustomRecord")
    app.UndoRecord.EndCustomRecord = MagicMock(name="EndCustomRecord")
    return app


@pytest.fixture
def fake_word(monkeypatch: pytest.MonkeyPatch) -> MagicMock:
    """A MagicMock Application with one document, one bookmark, one heading, one CC."""
    doc = _make_document(
        bookmarks={"Address": (13, 24)},
        content_controls=[
            {"title": "Signatory", "tag": "sig", "start": 29, "end": 35, "text": "Jane Doe"}
        ],
        paragraphs=[
            {"level": 1, "text": "Introduction", "start": 0, "end": 13},
            {"level": 10, "text": "Body text here.", "start": 13, "end": 29},
            {"level": 2, "text": "Risks", "start": 29, "end": 35},
        ],
        # Self-consistent text — paragraph offsets above index into this string.
        content="Introduction\rBody text here.\rRisks\r",
        tables=[{"grid": [["A1", "B1"], ["A2", "B2"]], "title": "Grid"}],
        # A 2-item numbered list over the body region (offsets 13–29).
        lists=[{"start": 13, "end": 29, "count": 2, "type": 3}],
        # One section with a primary header and footer seeded for read tests.
        sections=[{"headers": {"primary": "Confidential Draft"}, "footers": {"primary": "Page 1"}}],
        # One footnote whose reference mark sits in the body paragraph (13–29),
        # so footnotes.list() maps it back to para:2.
        footnotes=[{"ref_start": 20, "text": "A seeded footnote."}],
    )
    app = _make_application([doc])

    from wordlive import _com

    monkeypatch.setattr(_com, "get_active_word", lambda: app)
    monkeypatch.setattr(_com, "launch_word", lambda visible=True: app)
    return app


@pytest.fixture
def no_word(monkeypatch: pytest.MonkeyPatch) -> None:
    """Simulate Word not running."""
    from wordlive import _com
    from wordlive.exceptions import WordNotRunningError

    def _raise(*args: Any, **kwargs: Any) -> None:
        raise WordNotRunningError("no running Microsoft Word instance found")

    monkeypatch.setattr(_com, "get_active_word", _raise)
    monkeypatch.setattr(_com, "launch_word", _raise)


@pytest.fixture
def real_word():
    """Smoke fixture: yields a wordlive.Word, or skips if Word isn't reachable."""
    import wordlive
    from wordlive.exceptions import WordNotRunningError

    try:
        ctx = wordlive.attach()
        word = ctx.__enter__()
    except WordNotRunningError as e:
        pytest.skip(f"Word not running: {e}")
        return
    try:
        yield word
    finally:
        ctx.__exit__(None, None, None)
