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
    """Mimics ActiveDocument.Bookmarks: Exists, by-name lookup, Add, iteration."""

    def __init__(self) -> None:
        self._items: dict[str, tuple[int, int]] = {}

    def add(self, name: str, start: int, end: int) -> None:
        self._items[name] = (start, end)

    # COM-style methods
    def Exists(self, name: str) -> bool:
        return name in self._items

    def Add(self, Name: str | None = None, Range: Any | None = None) -> None:
        if Name is None or Range is None:
            return
        self._items[Name] = (int(Range.Start), int(Range.End))

    def __call__(self, name: str) -> Any:
        s, e = self._items[name]
        bm = MagicMock(name=f"Bookmark[{name}]")
        bm.Name = name
        bm.Range = _make_range(s, e)
        return bm

    def __iter__(self) -> Iterable[Any]:
        out = []
        for name, (s, e) in self._items.items():
            bm = MagicMock(name=f"Bookmark[{name}]")
            bm.Name = name
            bm.Range = _make_range(s, e)
            out.append(bm)
        return iter(out)


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


def _make_document(
    *,
    name: str = "Test.docx",
    full_name: str = r"C:\test\Test.docx",
    bookmarks: dict[str, tuple[int, int]] | None = None,
    content_controls: list[dict[str, Any]] | None = None,
    paragraphs: list[dict[str, Any]] | None = None,
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
    doc.Range.side_effect = _make_range
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
        bookmarks={"Address": (10, 25)},
        content_controls=[{"title": "Signatory", "tag": "sig", "start": 30, "end": 40, "text": "Jane Doe"}],
        paragraphs=[
            {"level": 1, "text": "Introduction", "start": 0, "end": 14},
            {"level": 10, "text": "Body text here.", "start": 14, "end": 30},
            {"level": 2, "text": "Risks", "start": 30, "end": 37},
        ],
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
