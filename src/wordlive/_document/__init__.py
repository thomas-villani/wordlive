"""Document wrapper + DocumentCollection.

The `Document` god-class is decomposed into a `DocumentCore` spine plus feature
mixins (`_core`, `_editing`, `_reading`, `_structure`, `_persistence`); this
module assembles them and keeps `DocumentCollection`.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from .. import _com
from ..exceptions import DocumentNotFoundError
from ._core import WatermarkInfo
from ._editing import EditingMixin
from ._persistence import PersistenceMixin
from ._reading import ReadingMixin
from ._structure import StructureMixin

if TYPE_CHECKING:
    from .._app import Word

__all__ = ["Document", "DocumentCollection", "WatermarkInfo"]


class Document(EditingMixin, ReadingMixin, StructureMixin, PersistenceMixin):
    """Wraps a Word Document COM object."""


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
