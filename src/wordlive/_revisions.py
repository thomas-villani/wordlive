"""Revisions — read Word's tracked changes as structured data.

When Track Changes is on, every edit becomes a `Revision` the user can accept or
reject. wordlive can already *write* tracked revisions (`doc.tracked_changes()`),
but an agent making tracked edits was otherwise blind to what it had recorded:
snapshots render the *final* text, and plain text reads concatenate the inserted
and deleted runs with no marker. `doc.revisions` closes that gap with a
structured, read-only view that mirrors `doc.comments` — the structured channel,
paired with `doc.snapshot(markup="all")` for the visual one.

Revisions are addressed by 1-based index (`doc.revisions[2]`), matching Word's
own `Revisions(n)` ordering. Reading is non-mutating; accept/reject of individual
revisions stays on the `.com` escape hatch for now.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from .constants import WdRevisionType
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._document import Document


# Revision.Type int -> the human string `list()` reports. Insert / delete are by
# far the common pair (one find/replace records both); the rest cover formatting,
# moves, and table-cell changes. Anything unrecognised reports as "other".
_TYPE_NAMES: dict[int, str] = {
    int(WdRevisionType.INSERT): "insert",
    int(WdRevisionType.DELETE): "delete",
    int(WdRevisionType.PROPERTY): "format",
    int(WdRevisionType.PARAGRAPH_NUMBER): "paragraph-number",
    int(WdRevisionType.DISPLAY_FIELD): "display-field",
    int(WdRevisionType.RECONCILE): "reconcile",
    int(WdRevisionType.CONFLICT): "conflict",
    int(WdRevisionType.STYLE): "style",
    int(WdRevisionType.REPLACE): "replace",
    int(WdRevisionType.PARAGRAPH_PROPERTY): "paragraph-format",
    int(WdRevisionType.TABLE_PROPERTY): "table-format",
    int(WdRevisionType.SECTION_PROPERTY): "section-format",
    int(WdRevisionType.STYLE_DEFINITION): "style-definition",
    int(WdRevisionType.MOVE_SOURCE): "move-source",
    int(WdRevisionType.MOVE_TARGET): "move-target",
    int(WdRevisionType.CELL_INSERTION): "cell-insert",
    int(WdRevisionType.CELL_DELETION): "cell-delete",
    int(WdRevisionType.CELL_MERGE): "cell-merge",
}


def revision_type_name(value: Any) -> str:
    """Map a `Revision.Type` int onto its human string (`"insert"`, `"delete"`, …)."""
    try:
        return _TYPE_NAMES.get(int(value), "other")
    except (TypeError, ValueError):
        return "other"


def _clean(raw: Any) -> str:
    """Strip Word's trailing paragraph / cell markers from a revision's text."""
    return str(raw or "").rstrip("\r\n\x07")


def _iso_date(value: Any) -> str | None:
    """Best-effort ISO-8601 string for a revision's `Date` (None if unavailable)."""
    if value is None:
        return None
    iso = getattr(value, "isoformat", None)
    if callable(iso):
        try:
            return str(iso())
        except Exception:
            return None
    text = str(value).strip()
    return text or None


class Revision:
    """A single tracked change, located by its 1-based document index."""

    def __init__(self, doc: Document, com: Any, index: int) -> None:
        self._doc = doc
        self._com = com
        self._index = index

    @property
    def com(self) -> Any:
        """Raw COM Revision object — escape hatch (Accept/Reject, sub-ranges, …)."""
        return self._com

    @property
    def index(self) -> int:
        return self._index

    @property
    def type(self) -> str:
        """The revision kind: `"insert"`, `"delete"`, `"format"`, … (`"other"` if unknown)."""
        with _com.translate_com_errors():
            return revision_type_name(self._com.Type)

    @property
    def author(self) -> str:
        with _com.translate_com_errors():
            return str(self._com.Author or "")

    @property
    def text(self) -> str:
        """The inserted or deleted text (the run the revision covers)."""
        with _com.translate_com_errors():
            return _clean(self._com.Range.Text)

    @property
    def date(self) -> str | None:
        """When the revision was made, ISO-8601 — `None` if Word doesn't report it."""
        try:
            return _iso_date(self._com.Date)
        except Exception:
            return None

    def to_dict(self) -> dict[str, Any]:
        """`{index, type, author, text, anchor_id, start, end, date}` — the `list()` shape.

        `anchor_id` is a `range:START-END` over the revision's run (so a hit can
        be fed back into `read`/`comments.add`); `text` is the inserted or
        deleted text.
        """
        with _com.translate_com_errors():
            rng = self._com.Range
            start, end = int(rng.Start), int(rng.End)
            return {
                "index": self._index,
                "type": revision_type_name(self._com.Type),
                "author": str(self._com.Author or ""),
                "text": _clean(rng.Text),
                "anchor_id": f"range:{start}-{end}",
                "start": start,
                "end": end,
                "date": self.date,
            }

    def __repr__(self) -> str:
        return f"<Revision {self._index} {self.type!r} by {self.author!r}>"


class RevisionCollection:
    """Indexable, iterable, read-only view over a document's tracked changes (`doc.revisions`)."""

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.Revisions.Count)

    def __getitem__(self, index: int) -> Revision:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"revision index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("revision", str(index))
        with _com.translate_com_errors():
            return Revision(self._doc, self._doc.com.Revisions(index), index)

    def __iter__(self) -> Iterator[Revision]:
        with _com.translate_com_errors():
            count = int(self._doc.com.Revisions.Count)
        for i in range(1, count + 1):
            with _com.translate_com_errors():
                com = self._doc.com.Revisions(i)
            yield Revision(self._doc, com, i)

    def list(self) -> list[dict[str, Any]]:
        """All tracked changes as `{index, type, author, text, anchor_id, start, end, date}` dicts."""
        return [r.to_dict() for r in self]
