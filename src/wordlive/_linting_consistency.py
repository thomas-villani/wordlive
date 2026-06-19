"""Consistency rules — direct formatting that drifted from the applied style.

The reframing from `spec-linter.md` §2: a "consistency" defect is almost always a
**direct override that deviates from the paragraph's own style** (a `Heading 1`
at 15pt, a body run in a different face). `anchor.format_info()` already reports
this per field (`override: true` with the style baseline alongside), so these
rules are a thin walk over that probe.

The default **targeted, idempotent** fix writes the *style's* value back as a
direct property (`format_run(size=style_size)` / `format_paragraph(...)`): re-running
writes the same value, so a second `regularize` is a no-op (the idempotency
contract), and it never disturbs other intentional formatting in the paragraph.
The aggressive `Font.Reset()` strip-to-style is deferred (a profile flag).

This module is imported by `_linting` for its side effect of registering the
rules; it must not be imported before `_linting` finishes defining `Rule` /
`Finding` / `_register_rule` (it is imported at the bottom of that module).
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._linting import Finding, Rule, Span, _register_rule

if TYPE_CHECKING:
    from ._document import Document

# Which format_info font field maps to which `format_run` op keyword (the read
# mirror and the write verb diverge only on `name` -> `font`).
_FONT_FIX_KEY = {"name": "font", "size": "size", "bold": "bold"}


def _in_span(span: Span | None, row: dict[str, Any]) -> bool:
    if span is None:
        return True
    return int(row["start"]) <= span[1] and int(row["end"]) >= span[0]


def _font_finding(
    rule_id: str, anchor_id: str, field: str, info: dict[str, Any], label: str
) -> Finding | None:
    """A finding for one direct font override, with a targeted style-value fix.

    Returns `None` when the field isn't a genuine, fixable override (not flagged,
    a mixed run, or the style itself has no concrete baseline)."""
    cell = info["font"][field]
    if not cell["override"] or cell["value"] is None:
        return None
    style_val = cell["style"]
    if style_val is None:
        return None
    return Finding(
        rule=rule_id,
        kind="consistency",
        severity="info",
        anchor_id=anchor_id,
        message=(
            f"{label} font {field} is {cell['value']!r} but its style "
            f"({info['style']!r}) specifies {style_val!r}."
        ),
        fixable=True,
        fix={"op": "format_run", "anchor_id": anchor_id, _FONT_FIX_KEY[field]: style_val},
        observed=f"{field}={cell['value']!r}",
        expected=f"{field}={style_val!r}",
    )


def _spacing_finding(
    anchor_id: str, field: str, info: dict[str, Any], label: str
) -> Finding | None:
    cell = info["paragraph"][field]
    if not cell["override"] or cell["value"] is None:
        return None
    style_val = cell["style"]
    if style_val is None:
        return None
    return Finding(
        rule="heading-spacing-consistent",
        kind="consistency",
        severity="info",
        anchor_id=anchor_id,
        message=(
            f"{label} {field} is {cell['value']}pt but its style "
            f"({info['style']!r}) specifies {style_val}pt."
        ),
        fixable=True,
        fix={"op": "format_paragraph", "anchor_id": anchor_id, field: style_val},
        observed=f"{field}={cell['value']}",
        expected=f"{field}={style_val}",
    )


def _check_heading_font_consistent(doc: Document, span: Span | None) -> Iterator[Finding]:
    for row in doc.paragraphs.list():
        if not row["is_heading"] or not _in_span(span, row):
            continue
        info = doc.anchor_by_id(row["anchor_id"]).format_info()
        for field in ("name", "size", "bold"):
            finding = _font_finding(
                "heading-font-consistent", row["anchor_id"], field, info, "Heading"
            )
            if finding is not None:
                yield finding


def _check_heading_spacing_consistent(doc: Document, span: Span | None) -> Iterator[Finding]:
    for row in doc.paragraphs.list():
        if not row["is_heading"] or not _in_span(span, row):
            continue
        info = doc.anchor_by_id(row["anchor_id"]).format_info()
        for field in ("space_before", "space_after"):
            finding = _spacing_finding(row["anchor_id"], field, info, "Heading")
            if finding is not None:
                yield finding


def _check_body_font_consistent(doc: Document, span: Span | None) -> Iterator[Finding]:
    for row in doc.paragraphs.list():
        if row["is_heading"] or not _in_span(span, row):
            continue
        info = doc.anchor_by_id(row["anchor_id"]).format_info()
        finding = _font_finding(
            "body-font-consistent", row["anchor_id"], "name", info, "Body paragraph"
        )
        if finding is not None:
            yield finding


def _check_mixed_run_format(doc: Document, span: Span | None) -> Iterator[Finding]:
    """A heading whose font varies across runs (a `wdUndefined` field). Headings
    are normally a uniform run, so a mixed one is often an accidental stray
    override — but pinpointing the outlier run needs a run-walk, so this is
    **report-only** (no fix)."""
    for row in doc.paragraphs.list():
        if not row["is_heading"] or not _in_span(span, row):
            continue
        mixed = doc.anchor_by_id(row["anchor_id"]).format_info()["font"]["mixed"]
        if mixed:
            yield Finding(
                rule="mixed-run-format",
                kind="consistency",
                severity="info",
                anchor_id=row["anchor_id"],
                message=(
                    f"Heading has mixed character formatting across its runs "
                    f"({', '.join(mixed)}); expected a uniform heading."
                ),
                fixable=False,
                observed=f"mixed={mixed}",
            )


for _rule in (
    Rule(
        id="heading-font-consistent",
        kind="consistency",
        severity="info",
        tags=("headings", "fonts"),
        check=_check_heading_font_consistent,
    ),
    Rule(
        id="heading-spacing-consistent",
        kind="consistency",
        severity="info",
        tags=("headings", "spacing"),
        check=_check_heading_spacing_consistent,
    ),
    Rule(
        id="body-font-consistent",
        kind="consistency",
        severity="info",
        tags=("fonts",),
        check=_check_body_font_consistent,
    ),
    Rule(
        id="mixed-run-format",
        kind="consistency",
        severity="info",
        tags=("headings", "fonts"),
        check=_check_mixed_run_format,
    ),
):
    _register_rule(_rule)
