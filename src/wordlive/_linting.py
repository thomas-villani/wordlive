"""Linter + formatting regularizer — `doc.lint()` / `doc.regularize()`.

Audit a document for publishing-quality defects (`lint`, a pure read), then
autofix the mechanical ones in one atomic-undo step (`regularize`, a write). Pure
composition over shipped primitives — `format_info` (the format probe),
`format_paragraph`, `apply_list`/`remove_list`, `set_heading_row` — so there is
**no new COM write surface** here; the new work is the rule engine.

Each rule is `consistency` (a direct override fighting the applied style),
`structural` (an objective layout/structure defect), or `policy` (a value that
deviates from a configured house-style target — deferred to a later pass). A
rule emits [`Finding`][wordlive._linting.Finding]s; a *fixable* finding carries
an op-shaped `fix` (literally an `exec` op, or a list of them), so `regularize`
is "lint, then run each finding's `fix` through the existing `run_batch` loop" —
the fix path reuses the audited op vocabulary rather than a parallel writer.

Design: `spec-linter.md`. Mirrors `_proofing.py`'s module shape (rule functions +
a `run_lint` entry that `Document.lint` delegates to).
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import asdict, dataclass, replace
from typing import TYPE_CHECKING, Any, cast

from . import _com
from ._anchors._anchor_format import style_baseline_cache
from ._lint_profile import Profile
from .exceptions import ComError

if TYPE_CHECKING:
    from ._anchors import Anchor
    from ._document import Document

# A finding's fix is one exec op, or a list of ops applied in order (the
# list-continuity repair is remove-then-reapply).
FixOps = dict[str, Any] | list[dict[str, Any]]

_SEVERITY_RANK = {"error": 0, "warning": 1, "info": 2}


@dataclass(frozen=True)
class Finding:
    """One linter result. `fix` is present iff `fixable` — an op-shaped dict (or
    list of them) `regularize` runs verbatim through the batch op loop."""

    rule: str
    kind: str  # consistency | structural | policy
    severity: str  # error | warning | info
    anchor_id: str
    message: str
    fixable: bool = False
    fix: FixOps | None = None
    # A fixable fix that **adds or destroys content** (inserts a caption/notice,
    # deletes a stray paragraph, strips a watermark) rather than just re-formatting
    # existing content. Such fixes are withheld by `regularize` unless the caller
    # opts in with `allow_content=True` — see `spec-linter.md` §8. Pure in-place
    # formatting/text fixes leave this `False` and apply by default.
    adds_content: bool = False
    observed: str | None = None
    expected: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


# A span is an inclusive-ish (start, end) character range used by `within=` to
# scope the audit; `None` means the whole document.
Span = tuple[int, int]


@dataclass(frozen=True)
class Rule:
    """A registered rule: identity/metadata plus a `check` that yields findings.

    `check(doc, span, profile)` walks the document (optionally clipped to `span`)
    and yields `Finding`s. Every check receives the resolved
    [`Profile`][wordlive._lint_profile.Profile]; most ignore it (consistency /
    structural rules need no config), while **policy** rules read their target /
    threshold from it. `tags` lets a caller select a family (`["headings"]`) as
    well as by id.

    `default_on` controls whether the rule runs in the **default** set (`rules=
    None`). Unambiguous-defect structural/consistency rules ship on; opinionated
    or stylistic rules (`spec-linter.md` §5b "default stance") ship **off** —
    they only run when named explicitly or pulled in by a tag (`--rules
    typography`). Policy rules are off in the default set regardless (by `kind`).
    """

    id: str
    kind: str
    severity: str
    tags: tuple[str, ...]
    check: Callable[[Document, Span | None, Profile], Iterator[Finding]]
    default_on: bool = True


# ---------------------------------------------------------------------------
# small COM/range helpers
# ---------------------------------------------------------------------------


def _anchor_span(anchor: Anchor) -> Span:
    rng = anchor.com
    return int(rng.Start), int(rng.End)


def _overlaps(span: Span | None, lo: int, hi: int) -> bool:
    """Does `[lo, hi]` intersect the audit `span` (None = whole doc)?"""
    if span is None:
        return True
    return lo <= span[1] and hi >= span[0]


_NO_PASS = object()  # distinguishes "no lint pass in flight" from "not yet filled"
_ROW_CACHE: ContextVar[Any] = ContextVar("wordlive_lint_rows", default=_NO_PASS)
_OUTLINE_CACHE: ContextVar[Any] = ContextVar("wordlive_lint_outline", default=_NO_PASS)


@contextmanager
def _document_walk_cache() -> Iterator[None]:
    """Memoise the whole-document walks for the duration of one lint pass.

    `paragraphs.list()` and `outline()` each enumerate every paragraph over COM,
    and a default run has ~18 rules that between them asked for those walks ten
    times over. A lint pass is a pure read, so one walk apiece is enough — as is
    one read of each *style's* baseline, which `format_info` otherwise re-reads
    (~25 COM properties) for every same-styled paragraph. Scoped to the pass (not
    the `Document`) so any edit — a `regularize` fix, a user keystroke —
    invalidates it by construction."""
    rows = _ROW_CACHE.set(None)
    outline = _OUTLINE_CACHE.set(None)
    try:
        with style_baseline_cache():
            yield
    finally:
        _OUTLINE_CACHE.reset(outline)
        _ROW_CACHE.reset(rows)


def _paragraph_rows(doc: Document) -> list[dict[str, Any]]:
    """`doc.paragraphs.list()`, walked once per lint pass. Reads straight through
    when no pass is in flight, so a rule stays callable on its own."""
    cached = _ROW_CACHE.get()
    if cached is _NO_PASS:
        return doc.paragraphs.list()
    if cached is None:
        cached = doc.paragraphs.list()
        _ROW_CACHE.set(cached)
    return cast("list[dict[str, Any]]", cached)


def _outline_rows(doc: Document) -> list[dict[str, Any]]:
    """`doc.outline()`, walked once per lint pass. See `_paragraph_rows`."""
    cached = _OUTLINE_CACHE.get()
    if cached is _NO_PASS:
        return doc.outline()
    if cached is None:
        cached = doc.outline()
        _OUTLINE_CACHE.set(cached)
    return cast("list[dict[str, Any]]", cached)


def _row_format_info(doc: Document, row: dict[str, Any]) -> dict[str, Any]:
    """`format_info()` for a `paragraphs.list()` row, resolved by character offset.

    Never route this through `anchor_by_id(row["anchor_id"])`: resolving a `para:N`
    re-walks the whole `Paragraphs` collection (one cross-process COM step each), so
    calling it once per row makes a rule quadratic in paragraph count. Every table
    cell is a paragraph, so a big table used to push a default lint past four
    minutes. The row already carries `para.Range.Start/End`, and `doc.range(...)` is
    a single `Range(start, end)` call, so this is O(1) and yields the identical
    dict — bar its `anchor_id` key, which no caller reads (they report `row`'s)."""
    return doc.range(int(row["start"]), int(row["end"])).format_info()


def _table_spans(doc: Document) -> list[tuple[int, int]]:
    """The character ranges of every table, so a body-prose walk can skip cells."""
    spans: list[tuple[int, int]] = []
    try:
        for table in doc.tables:
            rng = table.com.Range
            spans.append((int(rng.Start), int(rng.End)))
    except ComError:
        return spans
    return spans


def _in_table(row: dict[str, Any], tables: list[tuple[int, int]]) -> bool:
    """Whether a `paragraphs.list()` row sits inside one of `tables`' spans."""
    start = int(row["start"])
    return any(lo <= start < hi for lo, hi in tables)


# A trailing sentence mark means prose, not a heading.
_SENTENCE_TAIL = ".!?:,;"
_FAUX_HEADING_MAX_CHARS = 80
_ENLARGED_RATIO = 1.2


def _heading_shaped(row: dict[str, Any]) -> bool:
    """Cheap text-only half of the faux-heading test: short, and not a sentence.

    Kept separate so a caller can gate on it *before* paying for `_row_format_info`
    (one COM round-trip per row — the cost that a default lint is dominated by)."""
    text = str(row["text"]).strip()
    return bool(text) and len(text) <= _FAUX_HEADING_MAX_CHARS and text[-1] not in _SENTENCE_TAIL


def _emphasized(info: dict[str, Any]) -> bool:
    """Uniformly bold, or set ≥20% larger than its style — the emphasis half."""
    font = info["font"]
    if font["bold"]["value"] is True:  # True = uniformly bold (None = mixed runs)
        return True
    size = font["size"]
    return bool(
        size["override"]
        and size["value"]
        and size["style"]
        and size["value"] >= size["style"] * _ENLARGED_RATIO
    )


def _emphasized_like_heading(row: dict[str, Any], info: dict[str, Any]) -> bool:
    """Whether a body paragraph is short + emphasized enough to read as a heading.

    Shared by two rules that must agree about the same paragraph.
    `manual-heading-formatting` *reports* one of these (suggesting a real heading
    style); `body-font-consistent` must not *fix* the very emphasis that makes it
    detectable — regularize would strip the bold, and the paragraph would then
    stop matching this predicate, silently erasing the finding. So the emphasis
    fields (`size` / `bold`) are exempt from drift-fixing there, while the font
    `name` — never a heading signal — stays audited.

    Callers apply their own pre-filters (table cells, verbatim styles); this is
    only the shape-and-emphasis test. Where `info` isn't already in hand, gate on
    `_heading_shaped(row)` first — it's free, and this isn't."""
    return _heading_shaped(row) and _emphasized(info)


# ---------------------------------------------------------------------------
# structural rules (no config; objective defects)
# ---------------------------------------------------------------------------


def _check_heading_keep_with_next(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A heading paragraph with keep-with-next off can be stranded at a page foot,
    its body starting on the next page. Fix: turn keep-with-next on."""
    for row in _outline_rows(doc):
        anchor_id = row["anchor_id"]
        anchor = doc.anchor_by_id(anchor_id)
        lo, hi = _anchor_span(anchor)
        if not _overlaps(span, lo, hi):
            continue
        kwn = anchor.format_info()["paragraph"]["keep_with_next"]["value"]
        if kwn is False:
            yield Finding(
                rule="heading-keep-with-next",
                kind="structural",
                severity="warning",
                anchor_id=anchor_id,
                message=(
                    f"Heading {row['text']!r} has keep-with-next off; "
                    "it may dangle at the foot of a page."
                ),
                fixable=True,
                fix={"op": "format_paragraph", "anchor_id": anchor_id, "keep_with_next": True},
                observed="keep_with_next=false",
                expected="keep_with_next=true",
            )


def _check_table_repeat_header(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A table that breaks across a page boundary should repeat its header row so
    the column labels carry over. Fix: mark row 1 as a heading row."""
    for table in doc.tables:
        # One location() over the table's whole range gives both its first and
        # last page (page = start, end_page = end) in a single repaginate —
        # cheaper than two per-cell calls, and it never resolves a cell so a
        # merged/odd grid doesn't trip it.
        try:
            trng = table.com.Range
            lo, hi = int(trng.Start), int(trng.End)
            loc = doc.range(lo, hi).location()
            first_page, last_page = int(loc["page"]), int(loc["end_page"])
        except ComError:
            continue  # transient COM hiccup — skip rather than fail the whole lint
        if first_page == last_page:
            continue  # single-page table; no repeating header needed
        if not _overlaps(span, lo, lo):
            continue
        try:
            already = bool(table.com.Rows(1).HeadingFormat)
        except ComError:
            already = False
        if already:
            continue
        anchor_id = f"table:{table.index}:1:1"
        yield Finding(
            rule="table-repeat-header",
            kind="structural",
            severity="warning",
            anchor_id=anchor_id,
            message=(
                f"Table {table.index} spans pages {first_page}–{last_page} "
                "but row 1 doesn't repeat as a header."
            ),
            fixable=True,
            fix={"op": "set_heading_row", "table": table.index, "row": 1},
            observed="heading_row=false",
            expected="heading_row=true",
        )


# Every list type whose items carry an auto-number — all can suffer the "split
# into independent 1. lists" footgun. Word can report an applied numbered list as
# any of these depending on template/version, so the continuity rule must accept
# them all, not just simple/outline numbering.
_NUMBERED_LIST_TYPES = frozenset({"numbered", "outline", "number-only", "mixed"})


def _numbered_list_spans(doc: Document) -> list[Span]:
    """The character spans of every *numbered* (incl. outline / number-only /
    mixed) Word list, in document order. Word models each distinct list as one
    `Document.Lists` entry, so a numbered list that got split into independent
    "1." lists shows up as several entries whose ranges abut."""
    spans: list[Span] = []
    with _com.translate_com_errors():
        count = int(doc.com.Lists.Count)
        for i in range(1, count + 1):
            rng = doc.com.Lists(i).Range
            start, end = int(rng.Start), int(rng.End)
            from ._lists import read_list_info

            if read_list_info(rng)["type"] in _NUMBERED_LIST_TYPES:
                spans.append((start, end))
    return sorted(spans)


def _check_list_numbering_continuity(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """The "N independent 1. lists" footgun: a run of numbered paragraphs Word
    split into separate lists, each restarting at 1. Detected as numbered lists
    whose character ranges **abut** (no body content between). Fix: drop the list
    formatting over the whole run and reapply one numbered list."""
    spans = _numbered_list_spans(doc)
    if len(spans) < 2:
        return
    # Chain runs of abutting numbered lists (next starts at/before prev's end).
    chains: list[list[Span]] = []
    current: list[Span] = [spans[0]]
    for prev, nxt in zip(spans, spans[1:], strict=False):
        if nxt[0] <= prev[1]:
            current.append(nxt)
        else:
            chains.append(current)
            current = [nxt]
    chains.append(current)

    for chain in chains:
        if len(chain) < 2:
            continue
        lo = chain[0][0]
        hi = chain[-1][1]
        if not _overlaps(span, lo, hi):
            continue
        anchor_id = f"range:{lo}-{hi}"
        yield Finding(
            rule="list-numbering-continuity",
            kind="structural",
            severity="warning",
            anchor_id=anchor_id,
            message=(
                f"{len(chain)} adjacent numbered lists look like one list Word split "
                "into independent runs (each restarts at 1)."
            ),
            fixable=True,
            fix=[
                {"op": "remove_list", "anchor_id": anchor_id},
                {"op": "apply_list", "anchor_id": anchor_id, "type": "numbered"},
            ],
            observed=f"{len(chain)} abutting numbered lists",
            expected="one continuous numbered list",
        )


# ---------------------------------------------------------------------------
# registry
# ---------------------------------------------------------------------------

# Consistency rules are appended in `_register_consistency_rules` (Step 3) to keep
# this list readable; structural rules live here.
_RULES: list[Rule] = [
    Rule(
        id="heading-keep-with-next",
        kind="structural",
        severity="warning",
        tags=("headings", "pagination"),
        check=_check_heading_keep_with_next,
    ),
    Rule(
        id="table-repeat-header",
        kind="structural",
        severity="warning",
        tags=("tables", "pagination"),
        check=_check_table_repeat_header,
    ),
    Rule(
        id="list-numbering-continuity",
        kind="structural",
        severity="warning",
        tags=("lists",),
        check=_check_list_numbering_continuity,
    ),
]


def _registry() -> dict[str, Rule]:
    return {r.id: r for r in _RULES}


def _select_rules(rules: Any, profile: Profile) -> list[Rule]:
    """Resolve the `rules=` selector (composed with `profile`) into the rule set to run.

    - `None` — the default set: every on-by-default consistency + structural rule
      (policy rules stay off until a profile enables them; opinionated rules with
      `default_on=False` stay off until named/tagged — `spec-linter.md` §5b).
    - a list of ids/tags — only rules matching an id or carrying a tag. This path
      **ignores** `default_on`, so an off-by-default rule runs when asked for by
      id or via its tag (`--rules typography` lights up the whole cluster).
    - `{"exclude": [...]}` — the default set minus the listed ids/tags.

    The `profile` then composes with that base set: any rule the profile enables is
    **unioned in** (this is how a policy rule — off in the default set by `kind` —
    opts in), and any rule the profile disables is **removed** (turn a default rule
    off). An explicit id/tag selection is left untouched by a profile-*enable* (the
    caller already chose the set) but a profile-*disable* still applies.
    """
    if rules is None:
        selected = [r for r in _RULES if r.kind in ("consistency", "structural") and r.default_on]
    elif isinstance(rules, dict):
        excluded = set(rules.get("exclude", []))
        default = [r for r in _RULES if r.kind in ("consistency", "structural") and r.default_on]
        selected = [r for r in default if not (r.id in excluded or excluded & set(r.tags))]
    else:
        wanted = set(rules)
        selected = [r for r in _RULES if r.id in wanted or wanted & set(r.tags)]

    chosen = {r.id: r for r in selected}
    # Profile opts policy (and any off-by-default) rules in by id, and disables rules
    # it turns off. Enabling only augments the *default-based* paths (None / exclude);
    # an explicit id/tag list already picked its set, so there we apply profile
    # *disables* only.
    default_based = rules is None or isinstance(rules, dict)
    for r in _RULES:
        state = profile.is_enabled(r.id)
        if state is True and default_based:
            chosen[r.id] = r
        elif state is False:
            chosen.pop(r.id, None)
    return [r for r in _RULES if r.id in chosen]


# ---------------------------------------------------------------------------
# public entry points (Document.lint / Document.regularize delegate here)
# ---------------------------------------------------------------------------


def _resolve_span(doc: Document, within: Any) -> Span | None:
    if within is None:
        return None
    anchor = doc.anchor_by_id(within) if isinstance(within, str) else within
    return _anchor_span(anchor)


def run_lint(
    doc: Document, *, rules: Any = None, within: Any = None, profile: Any = None
) -> list[Finding]:
    """Run the selected rules and return findings, ranked by severity.

    Pure read. The layout rules repaginate (content-neutrally, via
    `anchor.location()`), but nothing is mutated. `profile` (a path / dict / `Profile`
    / `None`) opts policy rules in, supplies their targets, and can override a rule's
    severity — resolved once here and threaded to every rule. See
    [`Document.lint`][wordlive.Document.lint].
    """
    resolved = Profile.load(profile)
    span = _resolve_span(doc, within)
    findings: list[Finding] = []
    with _document_walk_cache():
        for rule in _select_rules(rules, resolved):
            for finding in rule.check(doc, span, resolved):
                sev = resolved.severity_for(finding.rule)
                if sev is not None and sev != finding.severity:
                    finding = replace(finding, severity=sev)
                findings.append(finding)
    findings.sort(key=lambda f: _SEVERITY_RANK.get(f.severity, 9))
    return findings


def _fix_ops(fix: FixOps) -> list[dict[str, Any]]:
    return list(fix) if isinstance(fix, list) else [fix]


def _delete_order_key(op: dict[str, Any]) -> int:
    """The document position a `delete_paragraph` op targets — its `para:N` index or a
    `range:START-END` start. `regularize` applies deletes in *descending* order of this
    key so that removing a later paragraph never renumbers/shifts an earlier fix's
    anchor within the same atomic pass (the multi-blank `stray-empty-paragraph` case)."""
    anchor = str(op.get("anchor_id") or "")
    kind, _, value = anchor.partition(":")
    try:
        if kind == "para":
            return int(value)
        if kind == "range":
            return int(value.split("-", 1)[0])
    except (ValueError, IndexError):
        return 0
    return 0


def regularize(
    doc: Document,
    *,
    rules: Any = None,
    within: Any = None,
    profile: Any = None,
    dry_run: bool = False,
    allow_content: bool = False,
    own_undo: bool = True,
) -> dict[str, Any]:
    """Apply the fixable findings; return the
    `{applied, skipped, deferred, findings}` report.

    Runs `run_lint`, then applies every fixable finding's `fix` op(s). With
    `own_undo=True` (the `Document.regularize` path) the fixes run through
    `run_batch` — one `doc.edit("Regularize formatting")`, one Ctrl-Z. With
    `own_undo=False` (the `regularize` *exec op*, already inside a batch's
    `doc.edit`) they apply via `apply_op` directly, so the surrounding batch stays
    a single undo record rather than nesting one. `dry_run=True` plans without
    writing.

    Fixes flagged `adds_content` (they insert or destroy content, not just
    re-format) are **withheld** unless `allow_content=True`; withheld fixes land in
    the `deferred` bucket so the caller sees what an opt-in would apply. See
    `spec-linter.md` §8 and [`Document.regularize`][wordlive.Document.regularize].
    """
    from ._ops import apply_op, run_batch

    findings = run_lint(doc, rules=rules, within=within, profile=profile)
    fixable = [f for f in findings if f.fixable and f.fix is not None]
    # Split the fixable set on the content gate: `deferred` fixes only apply when
    # the caller opts in with `allow_content`. `skipped` stays "not fixable at all".
    to_apply = [f for f in fixable if allow_content or not f.adds_content]
    deferred = [f for f in fixable if not allow_content and f.adds_content]
    skipped = [f.to_dict() for f in findings if not (f.fixable and f.fix is not None)]
    report: dict[str, Any] = {
        "applied": [],
        "skipped": skipped,
        "deferred": [f.to_dict() for f in deferred],
        "findings": [f.to_dict() for f in findings],
    }
    if dry_run or not to_apply:
        report["dry_run"] = dry_run
        return report

    ops: list[dict[str, Any]] = []
    deletes: list[dict[str, Any]] = []
    for f in to_apply:
        for op in _fix_ops(f.fix):  # type: ignore[arg-type]
            # Paragraph deletions shift every later anchor, so they can't run inline
            # with the other fixes: collect them and apply last, descending by position
            # (see `_delete_order_key`). Every other fix is anchor-stable in one pass.
            (deletes if op.get("op") == "delete_paragraph" else ops).append(op)
    deletes.sort(key=_delete_order_key, reverse=True)
    ops.extend(deletes)
    if own_undo:
        result, exc = run_batch(doc, ops, label="Regularize formatting")
        if exc is not None:
            # `run_batch` stops at the first failing op and records a structured
            # `failure` dict (op index, the op, error message/type, partial
            # `ops_run`). Surface it on the raised error rather than dropping it,
            # so the caller can see which fix failed and how far the batch got.
            failure = result.get("failure")
            if failure is not None:
                exc.failure = failure  # type: ignore[attr-defined]
                exc.ops_run = result.get("ops_run", 0)  # type: ignore[attr-defined]
            raise exc
        report["ops_run"] = result.get("ops_run", 0)
        if result.get("warnings"):
            report["warnings"] = result["warnings"]
    else:
        for op in ops:
            apply_op(doc, op)
        report["ops_run"] = len(ops)
    report["applied"] = [f.to_dict() for f in to_apply]
    return report


def _register_rule(rule: Rule) -> None:
    """Append a rule to the registry (used by the consistency-rule module split)."""
    _RULES.append(rule)


# Defer rule registration to keep import order simple; each module is imported
# for its side effect of extending `_RULES` (it must run after `Rule` / `Finding`
# / `_register_rule` are defined, hence the bottom-of-module import).
from . import (  # noqa: E402,F401
    _linting_consistency,
    _linting_fields,
    _linting_finalization,
    _linting_headings,
    _linting_hyperlinks,
    _linting_layout,
    _linting_policy,
    _linting_typography,
)
