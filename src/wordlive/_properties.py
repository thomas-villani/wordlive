"""Document properties — the file's metadata (Title, Author, Keywords, …).

Word keeps two property bags on every document: the **built-in** set
(`BuiltInDocumentProperties` — Title, Subject, Author, Keywords, Comments,
Category, Manager, Company, plus read-only stats like the creation date and word
count) and a free-form **custom** set (`CustomDocumentProperties` — any
name/value pair you like, the same bag `{ DOCPROPERTY name }` fields read).

`doc.properties` is the read/write view over both. `read()` returns
`{"builtin": {…}, "custom": {…}}`; `set(name, value)` writes a built-in property
by name, and `set(name, value, custom=True)` writes (creating if needed) a custom
one. The natural shape here is a name→value mapping, not the list-of-dicts the
positional collections use, because properties are addressed by name.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from .constants import MsoDocProperty
from .exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from ._document import Document


def _prop_value(raw: Any) -> Any:
    """A JSON-safe Python value for a property — dates become ISO-8601 strings.

    Word hands back native types (str / int / float / bool) for most properties
    and a COM date for the timestamp ones; `isoformat()` normalises the latter so
    the CLI/MCP can emit it. Anything exotic falls back to `str`.
    """
    if isinstance(raw, (str, int, float, bool)) or raw is None:
        return raw
    iso = getattr(raw, "isoformat", None)
    if callable(iso):
        try:
            return str(iso())
        except Exception:
            return str(raw)
    return str(raw)


def _read_bag(bag: Any) -> dict[str, Any]:
    """Read a `DocumentProperties` COM bag into a `{name: value}` dict.

    A built-in property that has never been set (several of the date/stat ones)
    raises on `.Value` access; those are skipped rather than surfaced as errors,
    so the dict holds exactly the properties that carry a value.
    """
    out: dict[str, Any] = {}
    for prop in bag:
        try:
            name = str(prop.Name)
        except Exception:
            continue
        try:
            out[name] = _prop_value(prop.Value)
        except Exception:
            # Unset built-in (e.g. "Last print date") — no value to report.
            continue
    return out


def _mso_type(value: Any) -> int:
    """Infer the `MsoDocProperty` type for a new custom property from `value`."""
    if isinstance(value, bool):
        return int(MsoDocProperty.BOOLEAN)
    if isinstance(value, int):
        return int(MsoDocProperty.NUMBER)
    if isinstance(value, float):
        return int(MsoDocProperty.FLOAT)
    return int(MsoDocProperty.STRING)


class PropertyCollection:
    """Read/write view over a document's built-in and custom properties.

    `doc.properties.read()` returns `{"builtin": {…}, "custom": {…}}`. Write a
    built-in property with `set("Title", "…")` and a custom one with
    `set("Project", "Apollo", custom=True)` (created if it doesn't exist). The
    built-in **stat** properties (word count, creation date, …) are read-only;
    Word raises if you try to set one, surfaced as an `OpError`.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def builtin(self) -> dict[str, Any]:
        """The built-in properties that carry a value, as `{name: value}`."""
        with _com.translate_com_errors():
            return _read_bag(self._doc.com.BuiltInDocumentProperties)

    def custom(self) -> dict[str, Any]:
        """The custom (user-defined) properties, as `{name: value}`."""
        with _com.translate_com_errors():
            return _read_bag(self._doc.com.CustomDocumentProperties)

    def read(self) -> dict[str, Any]:
        """`{"builtin": {…}, "custom": {…}}` — every property with a value."""
        return {"builtin": self.builtin(), "custom": self.custom()}

    # `list()` mirrors the other collections' read entrypoint name.
    list = read

    def get(self, name: str) -> Any:
        """Look up one property's value by name (built-in first, then custom).

        Raises `AnchorNotFoundError` (kind `"property"`) if no built-in or custom
        property of that name carries a value.
        """
        builtin = self.builtin()
        if name in builtin:
            return builtin[name]
        custom = self.custom()
        if name in custom:
            return custom[name]
        raise AnchorNotFoundError("property", name)

    def set(self, name: str, value: Any, *, custom: bool = False) -> None:
        """Set property `name` to `value` (a built-in by default, or a custom one).

        With `custom=False` (default) this writes a **built-in** property — the
        writable ones are Title, Subject, Author, Keywords, Comments, Category,
        Manager, Company, Content status, and Hyperlink base; the stat/date
        properties are read-only and raise `OpError`. With `custom=True` it sets
        the custom property, creating it if absent (the type is inferred from
        `value`: bool/int/float/str). Wrap in `doc.edit(...)` for atomic undo.
        """
        if custom:
            self._set_custom(name, value)
        else:
            self._set_builtin(name, value)

    def _set_builtin(self, name: str, value: Any) -> None:
        with _com.translate_com_errors():
            bag = self._doc.com.BuiltInDocumentProperties
            try:
                prop = bag(name)
            except Exception as exc:  # unknown built-in name
                raise OpError(
                    f"{name!r} is not a built-in document property; "
                    "pass custom=True to set a custom property"
                ) from exc
            try:
                prop.Value = value
            except Exception as exc:  # read-only stat/date property
                raise OpError(
                    f"built-in property {name!r} is read-only (it's a computed stat/date)"
                ) from exc

    def _set_custom(self, name: str, value: Any) -> None:
        with _com.translate_com_errors():
            bag = self._doc.com.CustomDocumentProperties
            existing = {str(p.Name) for p in bag}
            if name in existing:
                bag(name).Value = value
            else:
                # Positional args (Name, LinkToContent, Type, Value): keyword
                # args are dropped under pywin32 late binding, the same gotcha
                # as Fields.Add / TabStops.Add.
                bag.Add(name, False, _mso_type(value), value)

    def delete(self, name: str) -> None:
        """Delete a custom property by name.

        Only custom properties can be removed — built-in ones are part of the
        format. Raises `AnchorNotFoundError` (kind `"property"`) if no custom
        property of that name exists. Wrap in `doc.edit(...)` for atomic undo.
        """
        with _com.translate_com_errors():
            bag = self._doc.com.CustomDocumentProperties
            existing = {str(p.Name) for p in bag}
            if name not in existing:
                raise AnchorNotFoundError("property", name)
            bag(name).Delete()
