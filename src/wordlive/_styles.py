"""Document-scoped style enumeration and lookup.

Styles are read-only wrappers — wordlive consumes existing styles, it does not
define or modify them. Membership is checked by iterating `doc.Styles` rather
than calling `doc.Styles(name)` and trapping the COM error, because Word does
not reserve a HRESULT for "style not found" and the generic `pywintypes.com_error`
would be indistinguishable from a real failure.
"""

from __future__ import annotations

from typing import Any, Iterator, TYPE_CHECKING

from . import _com
from .constants import WdStyleType
from .exceptions import StyleNotFoundError

if TYPE_CHECKING:
    from ._document import Document


_STYLE_TYPE_NAMES = {
    int(WdStyleType.PARAGRAPH): "paragraph",
    int(WdStyleType.CHARACTER): "character",
    int(WdStyleType.TABLE): "table",
    int(WdStyleType.LIST): "list",
}


def _style_type_name(value: Any) -> str:
    try:
        return _STYLE_TYPE_NAMES.get(int(value), "unknown")
    except (TypeError, ValueError):
        return "unknown"


class Style:
    """A read-only view onto a single Word style.

    Properties access the COM object lazily; nothing is cached so renames or
    deletions during the session don't return stale data.
    """

    def __init__(self, doc: "Document", name: str) -> None:
        self._doc = doc
        self._name = name

    @property
    def name(self) -> str:
        return self._name

    @property
    def com(self) -> Any:
        """Raw COM Style object. Raises `StyleNotFoundError` if the style is gone.

        Tries direct lookup (`Styles(name)`) first — O(1) on Word's side — and
        falls back to iteration only if that raises. Membership *checking*
        still iterates (Word doesn't reserve an HRESULT for "missing style"
        and a generic com_error would be indistinguishable from a real
        failure), but once the caller has a `Style` instance the name is
        presumed valid and the direct path is safe.
        """
        doc_com = self._doc.com
        with _com.translate_com_errors():
            try:
                return doc_com.Styles(self._name)
            except Exception:
                pass
            for s in doc_com.Styles:
                if str(s.NameLocal) == self._name:
                    return s
        raise StyleNotFoundError(self._name)

    @property
    def type(self) -> str:
        return _style_type_name(self.com.Type)

    @property
    def builtin(self) -> bool:
        return bool(self.com.BuiltIn)

    @property
    def in_use(self) -> bool:
        return bool(self.com.InUse)

    def to_dict(self) -> dict[str, Any]:
        with _com.translate_com_errors():
            com = self.com
            return {
                "name": self._name,
                "type": _style_type_name(com.Type),
                "builtin": bool(com.BuiltIn),
                "in_use": bool(com.InUse),
            }

    def __repr__(self) -> str:
        return f"<Style {self._name!r}>"


class StyleCollection:
    """Indexable, iterable view over a document's styles."""

    def __init__(self, doc: "Document") -> None:
        self._doc = doc

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        with _com.translate_com_errors():
            for s in self._doc.com.Styles:
                if str(s.NameLocal) == name:
                    return True
        return False

    def __getitem__(self, name: str) -> Style:
        if name not in self:
            raise StyleNotFoundError(name)
        return Style(self._doc, name)

    def __iter__(self) -> Iterator[Style]:
        with _com.translate_com_errors():
            names = [str(s.NameLocal) for s in self._doc.com.Styles]
        for n in names:
            yield Style(self._doc, n)

    def list(self) -> list[dict[str, Any]]:
        """All styles as `{name, type, builtin, in_use}` dicts."""
        with _com.translate_com_errors():
            return [
                {
                    "name": str(s.NameLocal),
                    "type": _style_type_name(s.Type),
                    "builtin": bool(s.BuiltIn),
                    "in_use": bool(s.InUse),
                }
                for s in self._doc.com.Styles
            ]
