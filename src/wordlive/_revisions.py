"""Revisions — read *and* resolve Word's tracked changes as structured data.

When Track Changes is on, every edit becomes a `Revision` the user can accept or
reject. wordlive can already *write* tracked revisions (`doc.tracked_changes()`);
`doc.revisions` is the structured read view that mirrors `doc.comments` — the
structured channel, paired with `doc.snapshot(markup="all")` for the visual one.

Revisions are addressed by 1-based index (`doc.revisions[2]`), matching Word's
own `Revisions(n)` ordering. A single revision can be resolved in place —
`doc.revisions[2].accept()` / `.reject()` — and the whole document (or a single
anchor's range) accepted or rejected at once with
`doc.revisions.accept_all()` / `.reject_all(within=anchor)`.

**Revision-aware reads.** A tracked edit leaves *both* the inserted and the
deleted runs present in the text stream, so a plain `anchor.text` read of a
just-edited paragraph is neither the before nor the after — it concatenates the
two. `segment_runs` reconstructs either view from the revision ranges; the
`Anchor.text_final` / `Anchor.text_original` / `Anchor.revision_segments`
helpers expose it per anchor (see `_anchors.py`).
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from .constants import WdRevisionType
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._anchors import Anchor
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


# Only insert / delete revisions change the text stream; every other kind
# (formatting, paragraph-number, …) leaves characters untouched, so they play no
# part in reconstructing the before/after text.
_TEXT_CHANGE_TYPES: frozenset[str] = frozenset({"insert", "delete"})


def segment_runs(
    final_text: str, base_start: int, runs: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Reconstruct a range's insert / delete / unchanged segments from its revisions.

    The subtlety (confirmed against live Word): a range's `Range.Text` is the
    **final** view — inserted runs are present, *deleted* runs are gone. The
    deleted characters survive only on the delete `Revision` itself (`run["text"]`),
    and Word reports revision offsets in a **markup** coordinate space where the
    deleted text still occupies positions. So we can't just label characters of
    `final_text`; we walk the markup space and splice the deleted text back in.

    `final_text` is the range's `Range.Text`; `base_start` its `Range.Start`;
    `runs` is `[{change, start, end, text}, …]` for the insert / delete revisions
    overlapping the range (document-absolute markup offsets, with the revision's
    own text). Returns `[{text, change}, …]` in document order — `change` is
    ``"insert"``, ``"delete"``, or ``None`` (unchanged) — from which:

    - **final** (accept-all) = the segments whose `change` is not ``"delete"``
      (== `final_text`);
    - **original** (reject-all) = those whose `change` is not ``"insert"``.

    Offsets are UTF-16 code units; outside the Basic Multilingual Plane they
    diverge from Python string indices — the same caveat the rest of the codebase
    carries (see `_utf16_len`).
    """
    ordered = sorted(
        (r for r in runs if r.get("change") in _TEXT_CHANGE_TYPES),
        key=lambda r: int(r["start"]),
    )
    segments: list[dict[str, Any]] = []

    def emit(text: str, change: str | None) -> None:
        if not text:
            return
        if segments and segments[-1]["change"] == change:
            segments[-1]["text"] += text
        else:
            segments.append({"text": text, "change": change})

    markup_cur = base_start  # position in the all-revisions-shown coordinate space
    final_cur = 0  # index into final_text (deletions absent, insertions present)
    for run in ordered:
        start, end = int(run["start"]), int(run["end"])
        if start < markup_cur:  # overlaps already-consumed space; skip defensively
            continue
        gap = start - markup_cur  # unchanged text before this revision
        if gap > 0:
            emit(final_text[final_cur : final_cur + gap], None)
            final_cur += gap
            markup_cur += gap
        width = end - start
        if run.get("change") == "insert":
            # Inserted text lives in final_text at the cursor.
            emit(final_text[final_cur : final_cur + width], "insert")
            final_cur += width
        else:  # delete — text is gone from final_text; restore it from the run
            emit(_clean(run.get("text")), "delete")
        markup_cur += width
    emit(final_text[final_cur:], None)  # trailing unchanged text
    return segments


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

    def accept(self) -> None:
        """Accept this tracked change — make it permanent.

        For an insertion the inserted text stays and loses its revision mark; for
        a deletion the struck-through text is removed. Accepting **renumbers** the
        remaining revisions (this one is consumed), so cached `doc.revisions[N]`
        indices past it shift down by one — re-list between resolves, or use the
        bulk [`accept_all`][wordlive.RevisionCollection.accept_all]. Wrap in
        `doc.edit(...)` for atomic undo.
        """
        with _com.translate_com_errors():
            self._com.Accept()

    def reject(self) -> None:
        """Reject this tracked change — undo it.

        For an insertion the inserted text is removed; for a deletion the
        struck-through text is restored. Like [`accept`][wordlive.Revision.accept]
        this consumes the revision and renumbers the rest. Wrap in `doc.edit(...)`
        for atomic undo.
        """
        with _com.translate_com_errors():
            self._com.Reject()

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

    def _revisions_com(self, within: Anchor | None) -> Any:
        """The COM `Revisions` collection to act on — the doc's, or one anchor's range."""
        if within is None:
            return self._doc.com.Revisions
        return within.com.Revisions

    def accept_all(self, *, within: Anchor | None = None) -> int:
        """Accept every tracked change at once and report how many were resolved.

        With no `within`, accepts the whole document; pass any anchor (heading,
        section range, cell, `range:START-END`, …) as `within` to accept only the
        tracked changes inside that range — "accept all my edits in this section".
        Returns the count accepted (read before the operation, since accepting
        empties the collection). Wrap in `doc.edit(...)` for atomic undo.
        """
        with _com.translate_com_errors():
            revisions = self._revisions_com(within)
            count = int(revisions.Count)
            if count:
                revisions.AcceptAll()
        return count

    def reject_all(self, *, within: Anchor | None = None) -> int:
        """Reject every tracked change at once and report how many were resolved.

        The mirror of [`accept_all`][wordlive.RevisionCollection.accept_all]:
        whole-document by default, or scoped to `within`'s range. Returns the
        count rejected. Wrap in `doc.edit(...)` for atomic undo.
        """
        with _com.translate_com_errors():
            revisions = self._revisions_com(within)
            count = int(revisions.Count)
            if count:
                revisions.RejectAll()
        return count

    def list(self) -> list[dict[str, Any]]:
        """All tracked changes as `{index, type, author, text, anchor_id, start, end, date}` dicts."""
        return [r.to_dict() for r in self]
