"""Fields — read the document's fields (PAGE, REF, TOC, DOCPROPERTY, …) as data.

A field is a self-updating value Word recomputes: a page number, a
cross-reference, a TOC, a document property. wordlive can already *insert* fields
(`anchor.insert_field(...)` / the `insert_field` op) and refresh them
(`update_fields`); `doc.fields` is the read mirror — what fields exist, their raw
code, and their last-rendered result.

Each field reports a `kind` (the leading keyword of its code — ``"PAGE"``,
``"REF"``, ``"TOC"``, …), the full `code`, the rendered `result`, and a
`range:START-END` id over the field. Fields are addressed by 1-based index
(`doc.fields[2]`), matching Word's own `Fields(n)` ordering. Reading is
non-mutating.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from .exceptions import AnchorNotFoundError


def _clean(raw: Any) -> str:
    """Strip Word's trailing paragraph / cell markers from a field's text."""
    return str(raw or "").rstrip("\r\n\x07")


def field_kind(code: str, type_int: int) -> str:
    """The field's kind: the leading keyword of its code (`"PAGE"`, `"REF"`, …).

    Word's field code is a string like ``" PAGE "`` or ``" REF bm \\h "``; the
    first whitespace-delimited token, upper-cased, is the human-meaningful kind.
    Falls back to ``"field:<type>"`` for a field whose code is empty (rare).
    """
    token = code.strip().split(None, 1)[0] if code.strip() else ""
    return token.upper() if token else f"field:{type_int}"


if TYPE_CHECKING:
    from ._document import Document


class Field:
    """A single field, located by its 1-based document index."""

    def __init__(self, doc: Document, com: Any, index: int) -> None:
        self._doc = doc
        self._com = com
        self._index = index

    @property
    def com(self) -> Any:
        """Raw COM Field object — escape hatch (Update, Unlink, ShowCodes, …)."""
        return self._com

    @property
    def index(self) -> int:
        return self._index

    @property
    def code(self) -> str:
        """The raw field code (e.g. ``"PAGE"``, ``"REF bookmark \\h"``)."""
        with _com.translate_com_errors():
            return _clean(self._com.Code.Text).strip()

    @property
    def result(self) -> str:
        """The field's last-rendered value (run `update_fields` to refresh it)."""
        with _com.translate_com_errors():
            return _clean(self._com.Result.Text)

    @property
    def type(self) -> int:
        """Word's numeric `Field.Type` (the `WdFieldType` value)."""
        with _com.translate_com_errors():
            return int(self._com.Type)

    @property
    def kind(self) -> str:
        """The leading keyword of the field code — ``"PAGE"``, ``"REF"``, ``"TOC"``, …."""
        return field_kind(self.code, self.type)

    def to_dict(self) -> dict[str, Any]:
        """`{index, kind, type, code, result, locked, anchor_id, para}` — the `list()` shape.

        `kind` is the leading code keyword; `type` is Word's numeric field type;
        `anchor_id` is a `range:START-END` over the field; `para` is the `para:N`
        it sits in (or ``None``).
        """
        with _com.translate_com_errors():
            code = _clean(self._com.Code.Text).strip()
            type_int = int(self._com.Type)
            result = _clean(self._com.Result.Text)
            try:
                locked = bool(self._com.Locked)
            except Exception:
                locked = False
            rng = self._com.Code
            try:
                start, end = int(rng.Start), int(rng.End)
            except Exception:
                start, end = None, None
        para_id: str | None = None
        if start is not None:
            para = self._doc.paragraphs.at(start)
            para_id = para.anchor_id if para is not None else None
        return {
            "index": self._index,
            "kind": field_kind(code, type_int),
            "type": type_int,
            "code": code,
            "result": result,
            "locked": locked,
            "anchor_id": f"range:{start}-{end}" if start is not None else None,
            "para": para_id,
        }

    def __repr__(self) -> str:
        return f"<Field {self._index} {self.kind!r}>"


class FieldCollection:
    """Indexable, iterable, read-only view over a document's fields (`doc.fields`).

    Scope is the main text story (`doc.Fields`); fields that live only in
    headers/footers are reached through the section's header/footer range on the
    `.com` escape hatch for now.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.Fields.Count)

    def __getitem__(self, index: int) -> Field:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"field index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("field", str(index))
        with _com.translate_com_errors():
            return Field(self._doc, self._doc.com.Fields(index), index)

    def __iter__(self) -> Iterator[Field]:
        with _com.translate_com_errors():
            count = int(self._doc.com.Fields.Count)
        for i in range(1, count + 1):
            with _com.translate_com_errors():
                com = self._doc.com.Fields(i)
            yield Field(self._doc, com, i)

    def list(self) -> list[dict[str, Any]]:
        """Every field as `{index, kind, type, code, result, locked, anchor_id, para}`."""
        return [f.to_dict() for f in self]
