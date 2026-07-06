"""Heading & document-structure rules — the §B "is the outline sound?" cluster.

`spec-linter.md` §5b·B (the **P2 + outline walk** primitive — `doc.outline()`
plus a `doc.fields` probe for the TOC). Six rules ship here; two are unambiguous
outline defects and ship **on** by default, the other four are opinionated /
structural nudges behind the `headings` / `structure` tags:

- `heading-level-skip` — the outline jumps a level (an H1 followed by an H3 with no
  H2), so the heading hierarchy is broken for navigation / accessibility. **On.**
- `empty-heading` — a heading paragraph with no text (a stray styled blank line
  that pollutes the outline / TOC). **On.**
- `adjacent-headings` — two headings in a row with no body between them (often a
  title/subtitle pair, sometimes a heading whose body was deleted). **Off** —
  a legitimate pattern often enough that it's a nudge, not a defect.
- `heading-numbering-manual` — the heading text starts with a literal number
  (`3.1 Methods`) that isn't an auto list/heading number, so it won't renumber.
  **Off** (a hand-numbered outline can be deliberate).
- `heading-trailing-period` — the heading text ends in a period (house styles
  usually drop it). **Off**, and the one **fixable** rule here — the fix strips the
  trailing period in place (a `find_replace` regex scoped to the heading), so it
  needs no `adds_content` gate.
- `toc-present-and-current` — the document has top-level headings but no table-of-
  contents field. **Off** — many documents legitimately omit a TOC. Presence-only:
  Word exposes no field-staleness flag (the same limit `stale-fields` hit), so the
  "current" half stays a report, not an auto-`update_fields`.

Detection reuses the shipped `doc.outline()` / `doc.fields` read wrappers plus each
heading anchor's `ListFormat` (to tell a hand-typed number from an auto one) — no
new COM surface. Every rule is document-wide over the outline; `within` clips each
finding to the audit span by the heading's range. Imported by `_linting` for its
side effect of registering the rules.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._linting import Finding, Rule, Span, _anchor_span, _overlaps, _register_rule
from .exceptions import ComError

if TYPE_CHECKING:
    from ._document import Document
    from ._lint_profile import Profile

# The shared tag family: every §B rule joins the `headings` cluster (alongside the
# v1 `heading-keep-with-next` / `manual-heading-formatting`) and the new `structure`
# tag, so `--rules headings` or `--rules structure` lights up the whole outline set.
_TAGS = ("headings", "structure")

# A heading text that opens with a hand-typed section number — `3`, `3.1`, `2.4.1`,
# optionally a trailing `.`/`)` — then a space and real content. Auto list/heading
# numbers never appear in `.Text` (Word renders them), so a match here is literal.
_MANUAL_HEADING_NUMBER = re.compile(r"^\s*(\d+(?:\.\d+){0,6})[.)]?\s+\S")

# A single trailing period to strip — not part of an ellipsis (the lookbehind) and
# allowing trailing whitespace / the paragraph mark after it. Used both to detect
# and, as the fix pattern, to remove it.
_TRAILING_PERIOD = r"(?<!\.)\.(?=\s*\r?$)"


def _para_index(anchor_id: str) -> int | None:
    """The 1-based paragraph index N from a `heading:N` anchor id, or `None`."""
    try:
        return int(anchor_id.split(":", 1)[1])
    except (IndexError, ValueError):
        return None


def _headings(doc: Document) -> list[dict[str, Any]]:
    """The document's headings (`doc.outline()`), or `[]` on a transient COM hiccup."""
    try:
        return doc.outline()
    except ComError:
        return []


def _heading_range(doc: Document, anchor_id: str) -> Span | None:
    """The heading paragraph's character span, or `None` if it can't be resolved."""
    try:
        return _anchor_span(doc.anchor_by_id(anchor_id))
    except ComError:
        return None


def _in_span(doc: Document, anchor_id: str, span: Span | None) -> bool:
    """Whether a heading overlaps the audit span (True when `within` is unset, or
    when the heading's range can't be resolved — never silently drop a finding)."""
    if span is None:
        return True
    rng = _heading_range(doc, anchor_id)
    if rng is None:
        return True
    return _overlaps(span, *rng)


# ---------------------------------------------------------------------------
# outline integrity (on by default)
# ---------------------------------------------------------------------------


def _check_heading_level_skip(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """The outline jumps a level between two adjacent headings (H1 → H3 with no H2).
    Compared against the immediately preceding heading, so a document that simply
    starts deep (all H2s) is left alone; only a genuine downward skip fires.
    Report-only — inserting the missing level or restyling is a human call."""
    prev_level: int | None = None
    for row in _headings(doc):
        level = int(row["level"])
        if prev_level is not None and level - prev_level > 1:
            anchor_id = row["anchor_id"]
            if _in_span(doc, anchor_id, span):
                skipped = ", ".join(str(n) for n in range(prev_level + 1, level))
                yield Finding(
                    rule="heading-level-skip",
                    kind="structural",
                    severity="warning",
                    anchor_id=anchor_id,
                    message=(
                        f"Heading {row['text']!r} is level {level} but follows a level-"
                        f"{prev_level} heading; the outline skips level {skipped}."
                    ),
                    fixable=False,
                    observed=f"level {prev_level} → {level}",
                    expected=f"no gap (level {prev_level + 1} next)",
                )
        prev_level = level


def _check_empty_heading(doc: Document, span: Span | None, profile: Profile) -> Iterator[Finding]:
    """A heading paragraph with no text — a stray styled blank line that shows up as
    an empty entry in the outline / TOC. Report-only: whether to delete it or fill it
    in is a human call (deleting a paragraph is a content edit, deferred anyway)."""
    for row in _headings(doc):
        if str(row["text"]).strip():
            continue
        anchor_id = row["anchor_id"]
        if not _in_span(doc, anchor_id, span):
            continue
        yield Finding(
            rule="empty-heading",
            kind="structural",
            severity="warning",
            anchor_id=anchor_id,
            message=f"Heading paragraph {anchor_id} has no text (an empty outline entry).",
            fixable=False,
            observed="empty heading",
            expected="a non-empty heading",
        )


# ---------------------------------------------------------------------------
# outline hygiene (off by default — `headings` / `structure` tags)
# ---------------------------------------------------------------------------


def _check_adjacent_headings(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """Two headings in a row with no body paragraph between them (their paragraph
    indices are consecutive). Often a legitimate title/subtitle pair, so off by
    default; sometimes a heading whose body was deleted. Report-only."""
    rows = _headings(doc)
    for cur, nxt in zip(rows, rows[1:], strict=False):
        cur_idx, nxt_idx = _para_index(cur["anchor_id"]), _para_index(nxt["anchor_id"])
        if cur_idx is None or nxt_idx is None or nxt_idx != cur_idx + 1:
            continue  # a body paragraph sits between them — fine
        if not _in_span(doc, cur["anchor_id"], span):
            continue
        yield Finding(
            rule="adjacent-headings",
            kind="structural",
            severity="info",
            anchor_id=cur["anchor_id"],
            message=(
                f"Heading {cur['text']!r} is immediately followed by another heading "
                f"({nxt['text']!r}) with no body text between them."
            ),
            fixable=False,
            observed="two adjacent headings",
            expected="body text under each heading",
        )


def _check_heading_numbering_manual(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A heading whose text opens with a hand-typed number (`3.1 Methods`) that isn't
    an auto list/heading number — so it won't renumber if sections are reordered.
    Off by default (deliberate manual numbering is legitimate). Report-only: turning
    it into automatic numbering is a structural change a human should choose."""
    for row in _headings(doc):
        text = str(row["text"])
        match = _MANUAL_HEADING_NUMBER.match(text)
        if match is None:
            continue
        anchor_id = row["anchor_id"]
        if not _in_span(doc, anchor_id, span):
            continue
        # Skip a heading that already carries an auto number (Word renders that
        # outside `.Text`, so digits in the text plus a live ListString would be
        # double-numbering — not this rule's concern).
        try:
            list_string = str(doc.anchor_by_id(anchor_id).com.ListFormat.ListString or "")
        except ComError:
            list_string = ""
        if list_string.strip():
            continue
        yield Finding(
            rule="heading-numbering-manual",
            kind="consistency",
            severity="info",
            anchor_id=anchor_id,
            message=(
                f"Heading {text!r} is numbered by hand ({match.group(1)!r}); use automatic "
                "heading/list numbering so it renumbers when sections move."
            ),
            fixable=False,
            observed=f"literal number {match.group(1)!r}",
            expected="automatic heading numbering",
        )


def _check_heading_trailing_period(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A heading whose text ends in a period — most house styles drop it. Off by
    default (opinion). Fixable: the fix strips the trailing period in place with a
    paragraph-scoped `find_replace` regex, so it never touches surrounding runs and a
    second `regularize` is a clean no-op. An ellipsis is left alone."""
    for row in _headings(doc):
        text = str(row["text"]).rstrip()
        if not text or not text.endswith(".") or text.endswith(".."):
            continue  # empty, no period, or an ellipsis — leave it
        anchor_id = row["anchor_id"]
        if not _in_span(doc, anchor_id, span):
            continue
        # The fix scopes to the heading's *paragraph* (`para:N`), not the heading
        # anchor: `find_replace` expands a heading scope to its body section, so a
        # `heading:N` scope would search under the heading and miss the title text.
        # `heading:N` and `para:N` share the same 1-based paragraph index.
        para_index = _para_index(anchor_id)
        para_anchor = f"para:{para_index}" if para_index is not None else anchor_id
        yield Finding(
            rule="heading-trailing-period",
            kind="consistency",
            severity="info",
            anchor_id=anchor_id,
            message=f"Heading {row['text']!r} ends with a period; headings usually omit it.",
            fixable=True,
            fix={
                "op": "find_replace",
                "find": _TRAILING_PERIOD,
                "text": "",
                "in": para_anchor,
                "all": True,
                "mode": "regex",
                "required": False,
            },
            observed="trailing period",
            expected="no trailing period",
        )


def _check_toc_present_and_current(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """The document has top-level (level-1) headings but no TOC field, so a reader has
    no table of contents. Off by default (many documents omit one deliberately).
    Presence-only and report-only: Word exposes no field-staleness flag, so the
    "current" half can't be auto-checked, and inserting a TOC adds content (an opt-in
    fix, deferred). Document-global, so `within` doesn't scope it."""
    rows = _headings(doc)
    if not any(int(row["level"]) == 1 for row in rows):
        return  # no top-level headings — a TOC isn't expected
    try:
        fields = doc.fields.list()
    except ComError:
        return
    if any(str(field.get("kind") or "").upper() == "TOC" for field in fields):
        return  # a table of contents is already present
    yield Finding(
        rule="toc-present-and-current",
        kind="structural",
        severity="info",
        anchor_id="start",
        message="Document has top-level headings but no table-of-contents (TOC) field.",
        fixable=False,
        observed="no TOC field",
        expected="a TOC field",
    )


for _rule in (
    Rule(
        id="heading-level-skip",
        kind="structural",
        severity="warning",
        tags=_TAGS,
        check=_check_heading_level_skip,
        default_on=True,
    ),
    Rule(
        id="empty-heading",
        kind="structural",
        severity="warning",
        tags=_TAGS,
        check=_check_empty_heading,
        default_on=True,
    ),
    Rule(
        id="adjacent-headings",
        kind="structural",
        severity="info",
        tags=_TAGS,
        check=_check_adjacent_headings,
        default_on=False,
    ),
    Rule(
        id="heading-numbering-manual",
        kind="consistency",
        severity="info",
        tags=_TAGS,
        check=_check_heading_numbering_manual,
        default_on=False,
    ),
    Rule(
        id="heading-trailing-period",
        kind="consistency",
        severity="info",
        tags=_TAGS,
        check=_check_heading_trailing_period,
        default_on=False,
    ),
    Rule(
        id="toc-present-and-current",
        kind="structural",
        severity="info",
        tags=(*_TAGS, "layout"),
        check=_check_toc_present_and_current,
        default_on=False,
    ),
):
    _register_rule(_rule)
