"""Document variables — invisible, named string storage on a document.

`Document.Variables` is a hidden key/value store that survives saves but never
appears in the text. It's the backing store for `{ DOCVARIABLE name }` fields, so
it's the clean way to stash values a template references (a client name, a
revision tag) without parking them in a bookmark or content control.

`doc.variables` is the read/write view. Values are always strings (Word coerces),
so the natural shape is a `{name: value}` mapping rather than the list-of-dicts
the positional collections use. Pair a variable with an `insert_field` of kind
`field` carrying `DOCVARIABLE name` to surface it in the document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._document import Document


class VariableCollection:
    """Read/write view over a document's variables (`doc.variables`).

    `doc.variables.list()` returns `{name: value}`. `set(name, value)` creates or
    updates a variable; `get(name)` reads one; `delete(name)` removes it. Values
    are stored as strings. Wrap writes in `doc.edit(...)` for atomic undo.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.Variables.Count)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        with _com.translate_com_errors():
            return any(str(v.Name) == name for v in self._doc.com.Variables)

    def list(self) -> dict[str, str]:
        """Every variable as a `{name: value}` dict (Word's `Variables` order)."""
        out: dict[str, str] = {}
        with _com.translate_com_errors():
            for var in self._doc.com.Variables:
                out[str(var.Name)] = str(var.Value)
        return out

    def get(self, name: str) -> str:
        """Read one variable's value by name.

        Raises `AnchorNotFoundError` (kind `"variable"`) if no variable of that
        name exists.
        """
        with _com.translate_com_errors():
            for var in self._doc.com.Variables:
                if str(var.Name) == name:
                    return str(var.Value)
        raise AnchorNotFoundError("variable", name)

    def set(self, name: str, value: Any) -> None:
        """Create or update variable `name` with `value` (stored as a string).

        Word's `Variables.Add` errors on a name that already exists, so an
        existing variable is updated in place and a new one is added. Wrap in
        `doc.edit(...)` for atomic undo.
        """
        text = str(value)
        with _com.translate_com_errors():
            variables = self._doc.com.Variables
            existing = {str(v.Name) for v in variables}
            if name in existing:
                variables(name).Value = text
            else:
                # Positional args (Name, Value) — keyword args are unreliable
                # under pywin32 late binding.
                variables.Add(name, text)

    def delete(self, name: str) -> None:
        """Delete variable `name`.

        Raises `AnchorNotFoundError` (kind `"variable"`) if it doesn't exist.
        Wrap in `doc.edit(...)` for atomic undo.
        """
        with _com.translate_com_errors():
            variables = self._doc.com.Variables
            existing = {str(v.Name) for v in variables}
            if name not in existing:
                raise AnchorNotFoundError("variable", name)
            variables(name).Delete()
