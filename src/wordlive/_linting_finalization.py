"""Finalization-hygiene rules — the P3 "is this actually final?" cluster.

`spec-linter.md` §5b·G: leftover review / markup state that shouldn't survive into
a finished document — unresolved comments, unaccepted tracked changes, Track
Changes still on, hidden text, leftover highlighting, and un-refreshed fields.
Detection reuses the shipped read wrappers (`doc.comments`, `doc.revisions`,
`doc.fields`, `doc.track_changes`) plus the two `format_info` character fields
added alongside this batch (`hidden`, `highlight`) — so there's no new COM read
surface beyond those two probes.

Per this batch's decision, the whole cluster ships **off by default**, behind the
`finalization` tag (`default_on=False` on every rule): a mid-authoring document
normally carries comments and revisions, so these are an opt-in pre-send check
(`--rule finalization`), not a default-lint defect. Every rule is **report-only**
except `leftover-highlight`, whose highlight-clear fix is safe and idempotent
(`format_run(highlight="none")` → `WdColorIndex.AUTO`, so a re-run is a no-op).
`stale-fields` is a report-only nudge: Word exposes no "field is stale" flag, so a
presence-based `update_fields` fix would re-flag after fixing and break the
idempotency contract — the fixable version lands with Batch 3's field-code
backbone (§5b·C).

Imported by `_linting` for its side effect of registering the rules (after `Rule`
/ `Finding` / `_register_rule` are defined).
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._linting import Finding, Rule, Span, _overlaps, _register_rule
from .exceptions import ComError

if TYPE_CHECKING:
    from ._document import Document

# Field kinds whose result is computed and can drift as the document changes, so
# their presence is a "refresh before finalizing" signal. Keyed on the string
# `Field.kind` (the leading code keyword) rather than numeric types.
_UPDATABLE_FIELD_KINDS = frozenset(
    {"TOC", "TOA", "SEQ", "REF", "PAGEREF", "PAGE", "NUMPAGES", "INDEX"}
)


def _in_span(span: Span | None, row: dict[str, Any]) -> bool:
    """Whether a `paragraphs.list()` row overlaps the audit span."""
    return _overlaps(span, int(row["start"]), int(row["end"]))


def _iter_paras(doc: Document, span: Span | None) -> Iterator[dict[str, Any]]:
    for row in doc.paragraphs.list():
        if _in_span(span, row):
            yield row


def _range_of(anchor_id: Any) -> tuple[int, int] | None:
    """Parse a `range:START-END` anchor id back to `(start, end)`, else `None`."""
    if isinstance(anchor_id, str) and anchor_id.startswith("range:"):
        try:
            lo, hi = anchor_id[len("range:") :].split("-", 1)
            return int(lo), int(hi)
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------------------
# review / markup state (doc-level; one summarising finding each)
# ---------------------------------------------------------------------------


def _check_comments_present(doc: Document, span: Span | None) -> Iterator[Finding]:
    """Any review comments left in the document. Report-only — resolving or
    removing a comment is a content decision, not a mechanical fix."""
    entries: list[tuple[str, bool]] = []
    try:
        comments = list(doc.comments)
    except ComError:
        return
    for comment in comments:
        try:
            scope = comment.com.Scope
            lo, hi = int(scope.Start), int(scope.End)
            anchor_id = f"range:{lo}-{hi}"
            in_span = _overlaps(span, lo, hi)
        except (ComError, AttributeError):
            anchor_id, in_span = "start", True
        if in_span:
            entries.append((anchor_id, comment.done))
    if not entries:
        return
    unresolved = sum(1 for _, done in entries if not done)
    tail = f" ({unresolved} unresolved)" if unresolved else ""
    yield Finding(
        rule="comments-present",
        kind="structural",
        severity="info",
        anchor_id=entries[0][0],
        message=(
            f"{len(entries)} review comment(s) present{tail}; resolve or remove before finalizing."
        ),
        fixable=False,
        observed=f"{len(entries)} comment(s)",
    )


def _check_unaccepted_revisions(doc: Document, span: Span | None) -> Iterator[Finding]:
    """Tracked changes that were never accepted or rejected. Report-only —
    accepting/rejecting in bulk is loud, so it's left to the user."""
    try:
        rows = doc.revisions.list()
    except ComError:
        return
    in_span: list[dict[str, Any]] = []
    for row in rows:
        start, end = row.get("start"), row.get("end")
        if start is not None and end is not None and not _overlaps(span, int(start), int(end)):
            continue  # a located revision outside the audit span
        in_span.append(row)
    if not in_span:
        return
    anchor_id = in_span[0].get("anchor_id") or "start"
    yield Finding(
        rule="unaccepted-revisions",
        kind="structural",
        severity="warning",
        anchor_id=anchor_id,
        message=(
            f"{len(in_span)} unaccepted tracked change(s) present; "
            "accept or reject them before finalizing."
        ),
        fixable=False,
        observed=f"{len(in_span)} revision(s)",
    )


def _check_track_changes_on(doc: Document, span: Span | None) -> Iterator[Finding]:
    """Track Changes is still enabled — a document-global flag, so `within` doesn't
    scope it. Report-only (turning it off is a one-liner the user should own)."""
    try:
        on = doc.track_changes
    except ComError:
        return
    if on:
        yield Finding(
            rule="track-changes-on",
            kind="structural",
            severity="info",
            anchor_id="start",
            message="Track Changes is on; turn it off before finalizing.",
            fixable=False,
            observed="TrackRevisions=true",
            expected="TrackRevisions=false",
        )


def _check_stale_fields(doc: Document, span: Span | None) -> Iterator[Finding]:
    """Updatable fields (TOC / SEQ / REF / PAGE …) whose rendered result can drift
    as the document changes. Report-only nudge — Word exposes no staleness flag, so
    a presence-based `update_fields` fix couldn't be idempotent (see module note)."""
    try:
        rows = doc.fields.list()
    except ComError:
        return
    matches: list[dict[str, Any]] = []
    for field in rows:
        if field.get("kind") not in _UPDATABLE_FIELD_KINDS:
            continue
        rng = _range_of(field.get("anchor_id"))
        if rng is None or _overlaps(span, *rng):
            matches.append(field)
    if not matches:
        return
    kinds = ", ".join(sorted({str(f["kind"]) for f in matches}))
    yield Finding(
        rule="stale-fields",
        kind="structural",
        severity="info",
        anchor_id=matches[0].get("anchor_id") or "start",
        message=(
            f"{len(matches)} updatable field(s) present ({kinds}); "
            "refresh them (update_fields) before finalizing."
        ),
        fixable=False,
        observed=f"{len(matches)} updatable field(s)",
    )


# ---------------------------------------------------------------------------
# character state (paragraph scan via format_info)
# ---------------------------------------------------------------------------


def _check_hidden_text_present(doc: Document, span: Span | None) -> Iterator[Finding]:
    """Hidden-text runs left in the document (they print/export invisibly). Report-
    only — whether to reveal or delete hidden text is a content decision."""
    for row in _iter_paras(doc, span):
        font = doc.anchor_by_id(row["anchor_id"]).format_info()["font"]
        if font["hidden"]["value"] is True or "hidden" in font["mixed"]:
            yield Finding(
                rule="hidden-text-present",
                kind="structural",
                severity="warning",
                anchor_id=row["anchor_id"],
                message="Hidden text present in this paragraph.",
                fixable=False,
                observed="hidden runs",
            )


def _check_leftover_highlight(doc: Document, span: Span | None) -> Iterator[Finding]:
    """Highlighter colour left on body text. Fix: clear it (idempotent — clearing an
    already-unhighlighted paragraph is a no-op)."""
    for row in _iter_paras(doc, span):
        font = doc.anchor_by_id(row["anchor_id"]).format_info()["font"]
        value = font["highlight"]["value"]
        if (value is not None and value != "none") or "highlight" in font["mixed"]:
            yield Finding(
                rule="leftover-highlight",
                kind="consistency",
                severity="info",
                anchor_id=row["anchor_id"],
                message="Highlighted text present; clear the highlight before finalizing.",
                fixable=True,
                fix={"op": "format_run", "anchor_id": row["anchor_id"], "highlight": "none"},
                observed="highlight present",
                expected="no highlight",
            )


for _rule in (
    Rule(
        id="comments-present",
        kind="structural",
        severity="info",
        tags=("finalization",),
        check=_check_comments_present,
        default_on=False,
    ),
    Rule(
        id="unaccepted-revisions",
        kind="structural",
        severity="warning",
        tags=("finalization",),
        check=_check_unaccepted_revisions,
        default_on=False,
    ),
    Rule(
        id="track-changes-on",
        kind="structural",
        severity="info",
        tags=("finalization",),
        check=_check_track_changes_on,
        default_on=False,
    ),
    Rule(
        id="hidden-text-present",
        kind="structural",
        severity="warning",
        tags=("finalization",),
        check=_check_hidden_text_present,
        default_on=False,
    ),
    Rule(
        id="leftover-highlight",
        kind="consistency",
        severity="info",
        tags=("finalization",),
        check=_check_leftover_highlight,
        default_on=False,
    ),
    Rule(
        id="stale-fields",
        kind="structural",
        severity="info",
        tags=("finalization",),
        check=_check_stale_fields,
        default_on=False,
    ),
):
    _register_rule(_rule)
