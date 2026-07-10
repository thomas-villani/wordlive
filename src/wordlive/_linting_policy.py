"""Policy rules — a value that deviates from a configured house-style target.

`spec-linter.md` §2/§5/§6: unlike consistency/structural rules, a **policy** rule
needs configuration, so it's off in the default set and only runs when a
[`Profile`][wordlive._lint_profile.Profile] opts it in (and supplies its target /
threshold). These three are the first consumers of the profile loader:

- `body-justified` — body paragraphs that aren't justified. Fix: justify them.
- `body-line-spacing` — body paragraphs whose line spacing ≠ the profile's `target`
  (`"single"` / `"1.5"` / `"double"`). No-ops if the profile gives no target.
- `table-numeric-right-align` — a table column whose body cells are mostly numeric
  (≥ `threshold`, default 0.8) but not right-aligned. Fix: right-align those cells.

Every fix goes through the existing `format_paragraph` op, and `format_info` /
`format_paragraph` share the alignment / line-spacing vocabulary, so re-running writes
the same value — a second `regularize` is a no-op (the idempotency contract). Enable
via a profile:

    {"rules": {"body-justified":            {"enabled": true},
               "body-line-spacing":         {"enabled": true, "target": "1.5"},
               "table-numeric-right-align": {"enabled": true, "threshold": 0.8}}}

Imported by `_linting` for its side effect of registering the rules.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._linting import (
    Finding,
    Rule,
    Span,
    _in_table,
    _overlaps,
    _paragraph_rows,
    _register_rule,
    _row_format_info,
    _table_spans,
)
from .exceptions import ComError

if TYPE_CHECKING:
    from ._document import Document
    from ._lint_profile import Profile

# Paragraph styles treated as "body" for the justify / line-spacing rules. Scoping by
# style name keeps headings, captions, and list paragraphs (usually "List Paragraph")
# out of the policy without a per-paragraph list probe.
_BODY_STYLES = frozenset({"Normal", "Body Text"})

_DEFAULT_NUMERIC_THRESHOLD = 0.8

# A cell value that reads as a number: optional sign / currency / open-paren, digits with
# optional thousands separators and a decimal, optional percent / close-paren. Matches
# "$1,234.50", "(1,200)", "45%", "-3.14"; rejects "N/A", "Q3", "".
_NUMERIC_CELL = re.compile(r"^[(\-+]?\s*\$?\s*\d[\d,]*(?:\.\d+)?\s*%?\s*\)?$")


def _is_numeric_cell(text: str) -> bool:
    return bool(_NUMERIC_CELL.match(text.strip()))


def _iter_body_paragraphs(doc: Document, span: Span | None) -> Iterator[dict[str, Any]]:
    """The non-empty body-prose paragraphs (by style, **outside tables**) overlapping
    `span` — the shared walk for the justify and line-spacing rules.

    Table cells are excluded: their alignment/spacing is table policy's domain
    (`table-numeric-right-align`), and justifying a numeric cell would fight it — a
    non-idempotent tug-of-war. So body prose means paragraphs not inside any table."""
    tables = _table_spans(doc)
    for row in _paragraph_rows(doc):
        if row.get("is_heading") or row.get("style") not in _BODY_STYLES:
            continue
        if not str(row.get("text") or "").strip():
            continue  # empty paragraph — no visible alignment / spacing to police
        if _in_table(row, tables):
            continue  # inside a table cell — not body prose
        if not _overlaps(span, int(row["start"]), int(row["end"])):
            continue
        yield row


def _check_body_justified(doc: Document, span: Span | None, profile: Profile) -> Iterator[Finding]:
    """Body paragraphs not set to justified alignment. Off unless a profile enables it
    (the house style wants a justified body). Fix: justify — idempotent."""
    try:
        rows = list(_iter_body_paragraphs(doc, span))
    except ComError:
        return
    for row in rows:
        anchor_id = row["anchor_id"]
        try:
            align = _row_format_info(doc, row)["paragraph"]["alignment"]["value"]
        except ComError:
            continue
        if align == "justify":
            continue
        yield Finding(
            rule="body-justified",
            kind="policy",
            severity="info",
            anchor_id=anchor_id,
            message=f"Body paragraph is {align!r}-aligned; the profile requires justified.",
            fixable=True,
            fix={"op": "format_paragraph", "anchor_id": anchor_id, "alignment": "justify"},
            observed=f"alignment={align!r}",
            expected="alignment='justify'",
        )


def _check_body_line_spacing(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """Body paragraphs whose line spacing differs from the profile's `target`
    (`"single"` / `"1.5"` / `"double"`). Requires a target — no target, no findings.
    Fix: set the target spacing — idempotent (`format_info` reads back the same value)."""
    target = profile.config_for("body-line-spacing").get("target")
    if not target:
        return  # the rule needs an explicit target; without one it polices nothing
    target = str(target)
    try:
        rows = list(_iter_body_paragraphs(doc, span))
    except ComError:
        return
    for row in rows:
        anchor_id = row["anchor_id"]
        try:
            spacing = _row_format_info(doc, row)["paragraph"]["line_spacing"]["value"]
        except ComError:
            continue
        if spacing is None or spacing == target:
            continue
        yield Finding(
            rule="body-line-spacing",
            kind="policy",
            severity="info",
            anchor_id=anchor_id,
            message=f"Body paragraph line spacing is {spacing!r}; the profile requires {target!r}.",
            fixable=True,
            fix={"op": "format_paragraph", "anchor_id": anchor_id, "line_spacing": target},
            observed=f"line_spacing={spacing!r}",
            expected=f"line_spacing={target!r}",
        )


def _check_table_numeric_right_align(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A table body column that's mostly numbers but not right-aligned. The `threshold`
    (default 0.8) is the fraction of non-empty body cells that must parse numeric for
    the column to count as numeric. Fix: right-align the offending cells — idempotent."""
    cfg = profile.config_for("table-numeric-right-align")
    try:
        threshold = float(cfg.get("threshold", _DEFAULT_NUMERIC_THRESHOLD))
    except (TypeError, ValueError):
        threshold = _DEFAULT_NUMERIC_THRESHOLD
    try:
        tables = list(doc.tables)
    except ComError:
        return
    for table in tables:
        try:
            # Physical `table:N:R:C` addressing is only reliable on a uniform grid; a
            # merged/split table would mis-index, so skip it.
            if not table.is_uniform:
                continue
            rows, cols = table.row_count, table.column_count
            trng = table.com.Range
            lo, hi = int(trng.Start), int(trng.End)
        except ComError:
            continue
        if rows < 2 or not _overlaps(span, lo, hi):
            continue  # header-only (or empty) table, or out of the audit span
        for col in range(1, cols + 1):
            try:
                body = [table.cell(r, col) for r in range(2, rows + 1)]
                texts = [(cell, cell.text.strip()) for cell in body]
            except ComError:
                continue
            nonempty = [(cell, txt) for cell, txt in texts if txt]
            if not nonempty:
                continue
            numeric = sum(1 for _, txt in nonempty if _is_numeric_cell(txt))
            if numeric / len(nonempty) < threshold:
                continue
            for cell, _txt in nonempty:
                try:
                    align = cell.format_info()["paragraph"]["alignment"]["value"]
                except ComError:
                    continue
                if align == "right":
                    continue
                yield Finding(
                    rule="table-numeric-right-align",
                    kind="policy",
                    severity="info",
                    anchor_id=cell.anchor_id,
                    message=(
                        f"Numeric table column {col} cell is {align!r}-aligned; the "
                        "profile requires numeric columns right-aligned."
                    ),
                    fixable=True,
                    fix={
                        "op": "format_paragraph",
                        "anchor_id": cell.anchor_id,
                        "alignment": "right",
                    },
                    observed=f"alignment={align!r}",
                    expected="alignment='right'",
                )


for _rule in (
    Rule(
        id="body-justified",
        kind="policy",
        severity="info",
        tags=("alignment", "policy"),
        check=_check_body_justified,
        default_on=False,
    ),
    Rule(
        id="body-line-spacing",
        kind="policy",
        severity="info",
        tags=("spacing", "policy"),
        check=_check_body_line_spacing,
        default_on=False,
    ),
    Rule(
        id="table-numeric-right-align",
        kind="policy",
        severity="info",
        tags=("tables", "policy"),
        check=_check_table_numeric_right_align,
        default_on=False,
    ),
):
    _register_rule(_rule)
