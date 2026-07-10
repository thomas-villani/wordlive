"""Field-code backbone rules — the P1 "are the computed fields sound?" cluster.

`spec-linter.md` §5b·C (+ §H page-numbers): the academia / cross-reference cluster
unlocked by walking Word's field codes (`Range.Fields`). Three rules ship here:

- `broken-cross-reference` — a `REF` / `PAGEREF` field whose rendered result is Word's
  "reference source not found" (or "bookmark not found") error text. An unambiguous
  defect, so it's **on by default** like the v1 structural set.
- `caption-manual-numbering` — a `Caption`-styled paragraph whose figure/table number is
  **literal text** rather than a `SEQ` field, so it won't renumber. **On (report).**
- `page-numbers-present` — no `PAGE` field in any section header/footer story. A policy
  rule, **off by default** behind the `layout` tag.
- `xref-as-literal-text` — a body paragraph that mentions a figure/table by literal number
  ("see Figure 3") with no `REF`/`PAGEREF` field covering it, so it won't retarget if things
  reorder. Heuristic (a bare "Table 2" in prose is often legitimate), so **off by default**
  behind the `crossref` / `academia` tags — the Batch 3b follow-up promised when Batch 3 shipped.

Detection landed in Batch 3/3b; Batch 6 then wired the one clean fix: `page-numbers-present`
carries an **opt-in fix** (`adds_content=True`) that inserts a `{ PAGE }` field into
`footer:1:primary` via `insert_field` — withheld by the gate unless `allow_content=True`,
and idempotent (`_has_page_field` re-reads the state, so a numbered document stops firing).
The other three stay **report-only**: rebuilding a `caption-manual-numbering` around a `SEQ`
field is a fragile in-place edit, and repairing a `broken-cross-reference` / `xref-as-literal
-text` needs a human-chosen target. The write verbs already exist (`insert_field`,
`insert_caption`, `insert_cross_reference`), so no new COM write surface.

Detection reuses the shipped `doc.fields` / `doc.paragraphs` / `doc.sections` read wrappers.
Imported by `_linting` for its side effect of registering the rules.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._linting import Finding, Rule, Span, _overlaps, _paragraph_rows, _register_rule
from .constants import WdFieldType
from .exceptions import ComError

if TYPE_CHECKING:
    from ._document import Document
    from ._lint_profile import Profile

# Field kinds (leading code token) that resolve to a cross-reference into a bookmark.
_CROSS_REF_KINDS = frozenset({"REF", "PAGEREF", "NOTEREF"})

# The rendered-result sentinels Word writes when a cross-reference target is gone.
# Locale-sensitive: this covers Word's English strings; a non-English install renders
# the equivalent message, which this substring match would miss (documented limitation).
_BROKEN_REF_MARKERS = ("reference source not found", "bookmark not found")

# Header/footer stories to probe for a PAGE field (the three per-section variants).
_HF_WHICH = ("primary", "first", "even")

# A caption line that carries a literal figure/table number ("Figure 3", "Table 2",
# "Eq. 1") — the signal that the number was typed, not emitted by a SEQ field.
_CAPTION_NUMBER = re.compile(
    r"\b(?:figure|fig\.?|table|tbl\.?|equation|eq\.?)\s+\d+",
    re.IGNORECASE,
)


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
# cross-references (main-story field walk)
# ---------------------------------------------------------------------------


def _check_broken_cross_reference(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A REF / PAGEREF field rendering Word's "source not found" error — its target
    bookmark was deleted or renamed. Report-only: repairing it needs a human to pick
    the intended target."""
    try:
        rows = doc.fields.list()
    except ComError:
        return
    for field in rows:
        if field.get("kind") not in _CROSS_REF_KINDS:
            continue
        result = str(field.get("result") or "").casefold()
        if not any(marker in result for marker in _BROKEN_REF_MARKERS):
            continue
        rng = _range_of(field.get("anchor_id"))
        if rng is not None and not _overlaps(span, *rng):
            continue
        yield Finding(
            rule="broken-cross-reference",
            kind="structural",
            severity="warning",
            anchor_id=field.get("anchor_id") or "start",
            message=(
                f"{field.get('kind')} field renders a broken-reference error "
                "(target bookmark missing); repair the cross-reference."
            ),
            fixable=False,
            observed=str(field.get("result") or "").strip(),
            expected="a resolved reference",
        )


# ---------------------------------------------------------------------------
# captions (paragraph scan cross-checked against SEQ fields)
# ---------------------------------------------------------------------------


def _check_caption_manual_numbering(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A Caption-styled paragraph whose number is literal text with no overlapping SEQ
    field — it won't renumber when figures are added/reordered. Report-only: rebuilding
    the number as a SEQ field adds content (an opt-in fix, deferred)."""
    try:
        paras = _paragraph_rows(doc)
        fields = doc.fields.list()
    except ComError:
        return
    seq_spans = [
        rng
        for f in fields
        if f.get("kind") == "SEQ" and (rng := _range_of(f.get("anchor_id"))) is not None
    ]
    for row in paras:
        if row.get("style") != "Caption":
            continue
        start, end = int(row["start"]), int(row["end"])
        if not _overlaps(span, start, end):
            continue
        if not _CAPTION_NUMBER.search(str(row.get("text") or "")):
            continue
        if any(_overlaps((start, end), lo, hi) for lo, hi in seq_spans):
            continue  # already backed by a SEQ field
        yield Finding(
            rule="caption-manual-numbering",
            kind="structural",
            severity="info",
            anchor_id=row["anchor_id"],
            message=(
                "Caption number appears to be literal text, not a SEQ field; "
                "it won't renumber automatically."
            ),
            fixable=False,
            observed="literal caption number",
            expected="a SEQ field",
        )


# ---------------------------------------------------------------------------
# unlinked cross-references (body text scan cross-checked against REF fields)
# ---------------------------------------------------------------------------


def _check_xref_as_literal_text(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A body paragraph that refers to a figure/table by literal number ("see Figure 3")
    with no REF/PAGEREF field covering it — the reference is typed text, so it won't
    retarget when figures are added or reordered. Heuristic and false-positive-prone (a
    plain "Table 2" in prose can be legitimate), so off by default. Report-only: an
    auto-fix would have to guess which caption the mention points at."""
    try:
        paras = _paragraph_rows(doc)
        fields = doc.fields.list()
    except ComError:
        return
    ref_spans = [
        rng
        for f in fields
        if f.get("kind") in _CROSS_REF_KINDS and (rng := _range_of(f.get("anchor_id"))) is not None
    ]
    for row in paras:
        # A caption legitimately reads "Figure 3" (that's `caption-manual-numbering`'s
        # job); a heading numbered "Table 2" isn't a cross-reference either.
        if row.get("style") == "Caption" or row.get("is_heading"):
            continue
        start, end = int(row["start"]), int(row["end"])
        if not _overlaps(span, start, end):
            continue
        if not _CAPTION_NUMBER.search(str(row.get("text") or "")):
            continue
        if any(_overlaps((start, end), lo, hi) for lo, hi in ref_spans):
            continue  # a real cross-reference field already covers this paragraph
        yield Finding(
            rule="xref-as-literal-text",
            kind="structural",
            severity="info",
            anchor_id=row["anchor_id"],
            message=(
                "Paragraph refers to a figure/table by literal number, not a cross-reference "
                "(REF) field; it won't update if the target is renumbered or moved."
            ),
            fixable=False,
            observed="literal figure/table reference",
            expected="a REF cross-reference field",
        )


# ---------------------------------------------------------------------------
# page numbers (header/footer field walk)
# ---------------------------------------------------------------------------


def _has_page_field(hf: Any) -> bool:
    """Whether a header/footer story carries a PAGE field."""
    try:
        if not hf.exists:
            return False
        for field in hf.com.Fields:
            if int(field.Type) == int(WdFieldType.PAGE):
                return True
    except (ComError, AttributeError):
        return False
    return False


def _check_page_numbers_present(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """No PAGE field in any header or footer — the document has no automatic page
    numbers. Policy + off by default (some documents deliberately omit them). Report-only
    (inserting a page number adds content). Document-global, so `within` doesn't scope it."""
    try:
        for section in doc.sections:
            for which in _HF_WHICH:
                if _has_page_field(section.header(which)) or _has_page_field(section.footer(which)):
                    return  # a page number exists somewhere — nothing to report
    except ComError:
        return
    yield Finding(
        rule="page-numbers-present",
        kind="policy",
        severity="info",
        anchor_id="start",
        message="No page-number (PAGE) field found in any header or footer.",
        fixable=True,
        # Adds content — withheld by the gate unless the caller opts in. Idempotent:
        # once a PAGE field lives in the first section's primary footer, `_has_page_field`
        # finds it on the next pass and the rule falls silent. `footer:1:primary` is the
        # conventional home for a page number.
        adds_content=True,
        fix={"op": "insert_field", "anchor_id": "footer:1:primary", "kind": "page"},
        observed="no PAGE field",
        expected="a PAGE field in a header/footer",
    )


for _rule in (
    Rule(
        id="broken-cross-reference",
        kind="structural",
        severity="warning",
        tags=("crossref", "academia"),
        check=_check_broken_cross_reference,
        default_on=True,
    ),
    Rule(
        id="caption-manual-numbering",
        kind="structural",
        severity="info",
        tags=("captions", "academia"),
        check=_check_caption_manual_numbering,
        default_on=True,
    ),
    Rule(
        id="page-numbers-present",
        kind="policy",
        severity="info",
        tags=("layout",),
        check=_check_page_numbers_present,
        default_on=False,
    ),
    Rule(
        id="xref-as-literal-text",
        kind="structural",
        severity="info",
        tags=("crossref", "academia"),
        check=_check_xref_as_literal_text,
        default_on=False,
    ),
):
    _register_rule(_rule)
