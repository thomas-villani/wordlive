"""Typography-hygiene rules — the P2 "run-walk / text-scan" cluster.

`spec-linter.md` §5b·A: cheap, high-frequency text defects an agent keeps fixing
by hand — trailing / leading / double spaces, space-before-punctuation,
hyphen-as-range, manual "headings" that were never styled, tables on mismatched
styles. Detection is a scan of each paragraph's text (`paragraphs.list()` already
hands back per-paragraph `text` with the trailing mark stripped); the format-aware
ones (`manual-heading-formatting`) reuse `anchor.format_info()`.

Fixes compose the **`find_replace` op in `regex` mode** (added alongside this
batch): scoped to the offending `para:N`, it re-scans live text and rewrites only
the matched span, so it never flattens surrounding inline formatting, has no
precomputed offsets to drift across a batch, and a second `regularize` is a clean
no-op. Every autofix carries `required: false` so an earlier fix in the same pass
that already cleaned an overlapping match (a trailing run is also a double space)
is tolerated rather than failing the batch.

Per `spec-linter.md` §5b "default stance", opinionated rules (`hyphen-as-range`,
`em-dash-usage`, `tabs-for-layout`, `manual-line-break`) ship **off** by default
(`default_on=False`) — named or pulled in by the `typography` tag. The
unambiguous-defect ones (whitespace hygiene, manual-heading-formatting,
table-style-consistent) ship on.

Imported by `_linting` for its side effect of registering the rules (after `Rule`
/ `Finding` / `_register_rule` are defined).
"""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._linting import Finding, Rule, Span, _overlaps, _register_rule
from .exceptions import ComError

if TYPE_CHECKING:
    from ._document import Document

# Paragraph styles where literal whitespace is meaningful (code / preformatted),
# so the whitespace-collapsing rules skip them.
_VERBATIM_STYLES = frozenset({"plain text", "html preformatted", "macro text"})

EN_DASH = "–"
EM_DASH = "—"
LINE_BREAK = "\x0b"  # Chr(11), a Shift+Enter manual line break, as it reads in .Text


def _in_span(span: Span | None, row: dict[str, Any]) -> bool:
    """Whether a `paragraphs.list()` row overlaps the audit span."""
    return _overlaps(span, int(row["start"]), int(row["end"]))


def _is_verbatim(row: dict[str, Any]) -> bool:
    """Skip code / preformatted styles where whitespace carries meaning."""
    style = str(row.get("style") or "").casefold()
    if style in _VERBATIM_STYLES:
        return True
    return any(token in style for token in ("code", "preformat", "verbatim"))


def _regex_fix(anchor_id: str, pattern: str, replace: str) -> dict[str, Any]:
    """A `find_replace` op, regex mode, scoped to one paragraph, tolerant of zero
    matches (an overlapping earlier fix may have cleaned it already)."""
    return {
        "op": "find_replace",
        "find": pattern,
        "text": replace,
        "in": anchor_id,
        "all": True,
        "mode": "regex",
        "required": False,
    }


def _iter_paras(doc: Document, span: Span | None) -> Iterator[dict[str, Any]]:
    for row in doc.paragraphs.list():
        if _in_span(span, row):
            yield row


# ---------------------------------------------------------------------------
# whitespace hygiene (on by default)
# ---------------------------------------------------------------------------

# A trailing whitespace run, asserting (but not consuming) the paragraph end or
# its `\r` mark — so the fix removes the spaces without ever touching the mark
# (deleting it would fuse the paragraph into the next segment).
_TRAILING = r"[ \t]+(?=\r?$)"
_LEADING = r"^[ \t]+"
# Two-plus spaces *between* non-space chars — lookarounds keep it from overlapping
# the leading/trailing rules (which own the edges).
_DOUBLE_SPACE = r"(?<=\S) {2,}(?=\S)"
_SPACE_BEFORE_PUNCT = r"[ \t]+([,.;:\)])"
_NUM_RANGE = r"\b(\d{1,4})-(\d{1,4})\b"


def _check_trailing_whitespace(doc: Document, span: Span | None) -> Iterator[Finding]:
    for row in _iter_paras(doc, span):
        text = row["text"]
        if text != text.rstrip(" \t"):
            yield Finding(
                rule="trailing-whitespace",
                kind="structural",
                severity="warning",
                anchor_id=row["anchor_id"],
                message="Paragraph ends in trailing whitespace.",
                fixable=True,
                fix=_regex_fix(row["anchor_id"], _TRAILING, ""),
                observed="trailing space/tab",
                expected="no trailing whitespace",
            )


def _check_leading_whitespace(doc: Document, span: Span | None) -> Iterator[Finding]:
    for row in _iter_paras(doc, span):
        text = row["text"]
        if text[:1] in (" ", "\t"):
            yield Finding(
                rule="leading-whitespace",
                kind="structural",
                severity="warning",
                anchor_id=row["anchor_id"],
                message="Paragraph starts with literal whitespace (use a paragraph indent).",
                fixable=True,
                fix=_regex_fix(row["anchor_id"], _LEADING, ""),
                observed="leading space/tab",
                expected="no leading whitespace",
            )


def _check_double_space(doc: Document, span: Span | None) -> Iterator[Finding]:
    for row in _iter_paras(doc, span):
        if _is_verbatim(row):
            continue
        if re.search(_DOUBLE_SPACE, row["text"]):
            yield Finding(
                rule="double-space",
                kind="consistency",
                severity="info",
                anchor_id=row["anchor_id"],
                message="Multiple consecutive spaces between words.",
                fixable=True,
                fix=_regex_fix(row["anchor_id"], _DOUBLE_SPACE, " "),
                observed="2+ spaces between words",
                expected="single space",
            )


def _check_space_before_punctuation(doc: Document, span: Span | None) -> Iterator[Finding]:
    for row in _iter_paras(doc, span):
        if _is_verbatim(row):
            continue
        if re.search(_SPACE_BEFORE_PUNCT, row["text"]):
            yield Finding(
                rule="space-before-punctuation",
                kind="consistency",
                severity="info",
                anchor_id=row["anchor_id"],
                message="Whitespace before punctuation.",
                fixable=True,
                fix=_regex_fix(row["anchor_id"], _SPACE_BEFORE_PUNCT, r"\1"),
                observed="space before , . ; : )",
                expected="no space before punctuation",
            )


# ---------------------------------------------------------------------------
# opinionated typography (off by default — `typography` tag / by id)
# ---------------------------------------------------------------------------


def _check_hyphen_as_range(doc: Document, span: Span | None) -> Iterator[Finding]:
    """`1990-1995` / `pp. 10-15` written with a hyphen rather than an en-dash."""
    for row in _iter_paras(doc, span):
        if _is_verbatim(row):
            continue
        if re.search(_NUM_RANGE, row["text"]):
            yield Finding(
                rule="hyphen-as-range",
                kind="consistency",
                severity="info",
                anchor_id=row["anchor_id"],
                message="Numeric range uses a hyphen; an en-dash (–) is conventional.",
                fixable=True,
                fix=_regex_fix(row["anchor_id"], _NUM_RANGE, rf"\1{EN_DASH}\2"),
                observed="digit-hyphen-digit",
                expected="en-dash range",
            )


def _check_em_dash_usage(doc: Document, span: Span | None) -> Iterator[Finding]:
    """An em-dash is present. Report-only — the `--` swap is a loud, opinion-laden
    edit, so it's never applied automatically."""
    for row in _iter_paras(doc, span):
        if EM_DASH in row["text"]:
            yield Finding(
                rule="em-dash-usage",
                kind="policy",
                severity="info",
                anchor_id=row["anchor_id"],
                message="Em-dash (—) present.",
                fixable=False,
                observed="em-dash present",
            )


def _check_tabs_for_layout(doc: Document, span: Span | None) -> Iterator[Finding]:
    """Tabs used mid-paragraph for layout (a tab after content, or a run of them).
    Report-only — the right fix (a table / real indents) needs human judgment."""
    for row in _iter_paras(doc, span):
        if re.search(r"\t\t|\S\t", row["text"]):
            yield Finding(
                rule="tabs-for-layout",
                kind="consistency",
                severity="info",
                anchor_id=row["anchor_id"],
                message="Tabs used for layout; consider a table or paragraph indents.",
                fixable=False,
                observed="layout tabs",
            )


def _check_manual_line_break(doc: Document, span: Span | None) -> Iterator[Finding]:
    """A Shift+Enter manual line break inside a paragraph. Report-only — whether it
    should become a paragraph break depends on intent."""
    for row in _iter_paras(doc, span):
        if LINE_BREAK in row["text"]:
            yield Finding(
                rule="manual-line-break",
                kind="structural",
                severity="info",
                anchor_id=row["anchor_id"],
                message="Manual line break (Shift+Enter) inside a paragraph.",
                fixable=False,
                observed="line break (Chr 11)",
            )


# ---------------------------------------------------------------------------
# structure heuristics
# ---------------------------------------------------------------------------

_SENTENCE_TAIL = ".!?:,;"


def _check_manual_heading_formatting(doc: Document, span: Span | None) -> Iterator[Finding]:
    """A short, fully-bold (or enlarged) body paragraph that reads like a heading
    but was never given a heading style — so it's invisible to the outline / TOC.
    Report-only: the right heading *level* is a judgment call, so we suggest rather
    than auto-apply `apply_style("Heading N")`."""
    for row in _iter_paras(doc, span):
        if row["is_heading"] or _is_verbatim(row):
            continue
        text = row["text"].strip()
        if not text or len(text) > 80 or text[-1] in _SENTENCE_TAIL:
            continue  # empty, long, or sentence-like — not a faux heading
        info = doc.anchor_by_id(row["anchor_id"]).format_info()
        font = info["font"]
        bold = font["bold"]["value"] is True  # True = uniformly bold (None = mixed)
        size = font["size"]
        enlarged = bool(
            size["override"]
            and size["value"]
            and size["style"]
            and size["value"] >= size["style"] * 1.2
        )
        if bold or enlarged:
            yield Finding(
                rule="manual-heading-formatting",
                kind="structural",
                severity="warning",
                anchor_id=row["anchor_id"],
                message=(
                    f"Paragraph {text!r} looks like a heading (short, emphasized, "
                    f"style {info['style']!r}) but isn't a heading style; "
                    'consider apply_style("Heading N").'
                ),
                fixable=False,
                observed=f"emphasized {info['style']!r} paragraph",
                expected="a heading style",
            )


def _check_table_style_consistent(doc: Document, span: Span | None) -> Iterator[Finding]:
    """Tables that don't share the document's dominant table style. Fix: restyle the
    minority tables onto the dominant style (idempotent — re-applying is a no-op)."""
    entries: list[tuple[int, int, str]] = []  # (index, range_start, style_name)
    for table in doc.tables:
        try:
            trng = table.com.Range
            lo, hi = int(trng.Start), int(trng.End)
            style_obj = table.com.Style
            style = str(style_obj.NameLocal) if style_obj is not None else None
        except (ComError, AttributeError):
            continue  # a table whose style can't be read — skip, don't fail the lint
        if style is not None and _overlaps(span, lo, hi):
            entries.append((table.index, lo, style))
    if len(entries) < 2:
        return
    counts = Counter(style for _, _, style in entries)
    (top_style, top_n), *rest = counts.most_common()
    if rest and rest[0][1] == top_n:
        return  # no single dominant style — ambiguous, don't guess
    for index, _start, style in entries:
        if style == top_style:
            continue
        yield Finding(
            rule="table-style-consistent",
            kind="consistency",
            severity="info",
            anchor_id=f"table:{index}:1:1",
            message=(f"Table {index} uses style {style!r}; most tables use {top_style!r}."),
            fixable=True,
            fix={"op": "set_table_style", "table": index, "style": top_style},
            observed=f"style={style!r}",
            expected=f"style={top_style!r}",
        )


for _rule in (
    Rule(
        id="trailing-whitespace",
        kind="structural",
        severity="warning",
        tags=("typography",),
        check=_check_trailing_whitespace,
    ),
    Rule(
        id="leading-whitespace",
        kind="structural",
        severity="warning",
        tags=("typography",),
        check=_check_leading_whitespace,
    ),
    Rule(
        id="double-space",
        kind="consistency",
        severity="info",
        tags=("typography",),
        check=_check_double_space,
    ),
    Rule(
        id="space-before-punctuation",
        kind="consistency",
        severity="info",
        tags=("typography",),
        check=_check_space_before_punctuation,
    ),
    Rule(
        id="hyphen-as-range",
        kind="consistency",
        severity="info",
        tags=("typography", "academia"),
        check=_check_hyphen_as_range,
        default_on=False,
    ),
    Rule(
        id="em-dash-usage",
        kind="policy",
        severity="info",
        tags=("typography",),
        check=_check_em_dash_usage,
        default_on=False,
    ),
    Rule(
        id="tabs-for-layout",
        kind="consistency",
        severity="info",
        tags=("typography",),
        check=_check_tabs_for_layout,
        default_on=False,
    ),
    Rule(
        id="manual-line-break",
        kind="structural",
        severity="info",
        tags=("typography",),
        check=_check_manual_line_break,
        default_on=False,
    ),
    Rule(
        id="manual-heading-formatting",
        kind="structural",
        severity="warning",
        tags=("typography", "headings"),
        check=_check_manual_heading_formatting,
    ),
    Rule(
        id="table-style-consistent",
        kind="consistency",
        severity="info",
        tags=("typography", "tables"),
        check=_check_table_style_consistent,
    ),
):
    _register_rule(_rule)
