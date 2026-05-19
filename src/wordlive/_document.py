"""Document wrapper + DocumentCollection."""

from __future__ import annotations

from contextlib import contextmanager
from typing import Any, Iterator, TYPE_CHECKING

from . import _com
from ._anchors import (
    BookmarkCollection,
    ContentControlCollection,
    Heading,
    _paragraph_text,
)
from ._edit import EditScope
from ._selection import Selection
from .exceptions import DocumentNotFoundError

if TYPE_CHECKING:
    from ._app import Word
    from ._anchors import Anchor


class Document:
    """Wraps a Word Document COM object."""

    def __init__(self, word: "Word", doc: Any) -> None:
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
    def selection(self) -> Selection:
        return self._word.selection

    def heading(self, name: str) -> Heading:
        # Lazy lookup — Heading.__init__ doesn't hit COM. _range() validates.
        return Heading(self, name)

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
                        "text": _paragraph_text(para),
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

    def go_to(self, anchor: "Anchor", scroll: bool = True) -> None:
        """Move the user's Selection to the given anchor (rare — most ops preserve it)."""
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

    def __init__(self, word: "Word") -> None:
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
