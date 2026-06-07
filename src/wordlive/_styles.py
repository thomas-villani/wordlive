"""Document-scoped style enumeration, lookup, creation, and modification.

`doc.styles.add(...)` defines a new style; the resulting `Style` is writable —
`style.format_run(...)` / `style.format_paragraph(...)` set its character and
paragraph defaults (reusing the same kwarg vocabulary as the anchor methods),
and `base_style` / `next_paragraph_style` chain styles together. Lookup is still
read-only and membership is checked by iterating `doc.Styles` rather than calling
`doc.Styles(name)` and trapping the COM error, because Word does not reserve a
HRESULT for "style not found" and the generic `pywintypes.com_error` would be
indistinguishable from a real failure.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from .constants import WdStyleType
from .exceptions import OpError, StyleNotFoundError

if TYPE_CHECKING:
    from ._document import Document


_STYLE_TYPE_FROM_NAME: dict[str, WdStyleType] = {
    "paragraph": WdStyleType.PARAGRAPH,
    "character": WdStyleType.CHARACTER,
    "table": WdStyleType.TABLE,
    "list": WdStyleType.LIST,
}


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

    def __init__(self, doc: Document, name: str) -> None:
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

    def format_run(self, **kwargs: Any) -> None:
        """Set this style's character (font) defaults.

        Same kwargs as [`Anchor.format_run`][wordlive.Anchor.format_run] —
        `bold`/`italic`/`underline`/`font`/`size`/`color`/… — minus `highlight`
        (a style's `Font` has no highlight property). Tri-state: only the kwargs
        you pass are written. Bad input raises `OpError`.
        """
        from ._anchors import _apply_font

        font_name = kwargs.pop("font", None)
        if "highlight" in kwargs:
            raise OpError("a style's font has no highlight; set highlight on a range instead")
        try:
            with _com.translate_com_errors():
                _apply_font(self.com.Font, font_name=font_name, **kwargs)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def format_paragraph(self, **kwargs: Any) -> None:
        """Set this style's paragraph defaults.

        Same kwargs as
        [`Anchor.format_paragraph`][wordlive.Anchor.format_paragraph]
        (`alignment`, indents, spacing, `page_break_before`). Tri-state. Bad
        input raises `OpError`.
        """
        from ._anchors import _apply_paragraph_format

        try:
            with _com.translate_com_errors():
                _apply_paragraph_format(self.com.ParagraphFormat, **kwargs)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    @property
    def base_style(self) -> str | None:
        """The name of the style this one inherits from (`None` if unset)."""
        with _com.translate_com_errors():
            try:
                return str(self.com.BaseStyle.NameLocal)
            except Exception:
                return None

    @base_style.setter
    def base_style(self, name: str) -> None:
        base = self._doc.styles[name]  # StyleNotFoundError if missing
        with _com.translate_com_errors():
            self.com.BaseStyle = base.com

    @property
    def next_paragraph_style(self) -> str | None:
        """The name of the style applied to the *next* paragraph (`None` if unset)."""
        with _com.translate_com_errors():
            try:
                return str(self.com.NextParagraphStyle.NameLocal)
            except Exception:
                return None

    @next_paragraph_style.setter
    def next_paragraph_style(self, name: str) -> None:
        nxt = self._doc.styles[name]  # StyleNotFoundError if missing
        with _com.translate_com_errors():
            self.com.NextParagraphStyle = nxt.com

    def __repr__(self) -> str:
        return f"<Style {self._name!r}>"


class StyleCollection:
    """Indexable, iterable view over a document's styles."""

    def __init__(self, doc: Document) -> None:
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

    def add(
        self,
        name: str,
        *,
        type: str = "paragraph",
        based_on: str | None = None,
        next_style: str | None = None,
    ) -> Style:
        """Define a new style and return it as a writable `Style`.

        `type` is `"paragraph"` (default), `"character"`, `"table"`, or `"list"`.
        `based_on` and `next_style` are names of existing styles (the inheritance
        parent and the style applied to the following paragraph). The brand /
        template primitive: define a house style once, then `apply_style` it
        everywhere. Style its defaults via the returned object's
        `format_run(...)` / `format_paragraph(...)`. Bad `type` raises `OpError`;
        an unknown `based_on` / `next_style` raises `StyleNotFoundError`.
        """
        style_type = _STYLE_TYPE_FROM_NAME.get(type)
        if style_type is None:
            raise OpError(
                f"unknown style type {type!r}; expected one of {sorted(_STYLE_TYPE_FROM_NAME)}"
            )
        # Resolve referenced styles up front so a miss fails before mutating.
        base_com = self[based_on].com if based_on is not None else None
        next_com = self[next_style].com if next_style is not None else None
        with _com.translate_com_errors():
            # Positional Name, Type — keywords drop under pywin32 late binding.
            self._doc.com.Styles.Add(name, int(style_type))
            new = Style(self._doc, name)
            if base_com is not None:
                new.com.BaseStyle = base_com
            if next_com is not None:
                new.com.NextParagraphStyle = next_com
        return new
