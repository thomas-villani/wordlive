"""Test fixtures.

`fake_word` builds a MagicMock that quacks like a Word.Application COM object —
enough surface to exercise the politeness logic, anchors, and CLI shape without
needing a real Word install.

`real_word` is for the @pytest.mark.smoke suite and is skipped if Word can't
be reached.
"""

from __future__ import annotations

from typing import Any, Iterable
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
            mock.Range.Text = (p.get("text", "") + "\r")
            self._items.append(mock)

    def __iter__(self) -> Iterable[Any]:
        return iter(self._items)


def _make_range(start: int, end: int) -> MagicMock:
    rng = MagicMock(name=f"Range[{start},{end}]")
    rng.Start = start
    rng.End = end
    rng.Text = ""
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


class _FakeTable:
    """Mimics a Word Table COM object: Rows/Columns counts, Cell(r,c), row add/delete.

    Cells are backed by persistent MagicMock ranges so writes (`Range.Text`,
    `Range.Style`, `Range.ParagraphFormat`) round-trip through the same handle.
    Cell text is seeded with Word's trailing `\\r\\x07` markers so the marker-
    stripping path in `Cell.text` gets exercised.
    """

    def __init__(self, grid: list[list[str]], title: str = "") -> None:
        self.Title = title
        self._rows = [[self._mk_cell_range(text) for text in row] for row in grid]

    @staticmethod
    def _mk_cell_range(text: str) -> MagicMock:
        rng = MagicMock(name="CellRange")
        rng.Text = text + "\r\x07"
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
    """Mimics doc.Tables: Count, 1-based call lookup, iteration."""

    def __init__(self, tables: list[_FakeTable]) -> None:
        self._tables = tables

    @property
    def Count(self) -> int:
        return len(self._tables)

    def __call__(self, index: int) -> _FakeTable:
        return self._tables[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(self._tables)


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
) -> MagicMock:
    doc = MagicMock(name=f"Document[{name}]")
    doc.Name = name
    doc.FullName = full_name

    bm_registry = _FakeBookmarkRegistry()
    for bm_name, (s, e) in (bookmarks or {}).items():
        bm_registry.add(bm_name, s, e)
    doc.Bookmarks = bm_registry

    doc.ContentControls = _FakeContentControls(content_controls or [])
    doc.Paragraphs = _FakeParagraphs(paragraphs or [])
    doc.Styles = _FakeStyles(styles if styles is not None else _DEFAULT_STYLES)
    doc.Tables = _FakeTablesCollection(
        [_FakeTable(t["grid"], t.get("title", "")) for t in (tables or [])]
    )

    # `Range(start, end)` returns a fresh range whose `.Text` can be set/read
    # by the caller; default `_make_range` sets `.Text = ""`.
    doc.Range.side_effect = _make_range

    # `Content` is the full document range; find/replace reads `.Text` from it.
    content_range = _make_range(0, len(content))
    content_range.Text = content
    doc.Content = content_range

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
        content_controls=[{"title": "Signatory", "tag": "sig", "start": 29, "end": 35, "text": "Jane Doe"}],
        paragraphs=[
            {"level": 1, "text": "Introduction", "start": 0, "end": 13},
            {"level": 10, "text": "Body text here.", "start": 13, "end": 29},
            {"level": 2, "text": "Risks", "start": 29, "end": 35},
        ],
        # Self-consistent text — paragraph offsets above index into this string.
        content="Introduction\rBody text here.\rRisks\r",
        tables=[{"grid": [["A1", "B1"], ["A2", "B2"]], "title": "Grid"}],
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
