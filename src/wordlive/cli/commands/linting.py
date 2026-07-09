"""Lint / regularize / proofing."""

from __future__ import annotations

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_lint,
    _fmt_proofing,
    _fmt_regularize,
    _rules_selector,
)


@click.command(name="proofing")
@click.pass_context
def proofing_cmd(ctx: click.Context) -> None:
    """Spelling/grammar errors and readability statistics for the document.

    Runs Word's proofing tools: `spelling`/`grammar` report a count plus a
    (capped) list of flagged runs with `range:START-END` ids, and `readability`
    reports Flesch Reading Ease, Flesch-Kincaid Grade Level, passive-sentence %,
    and averages. Heavier than `stats` (it (re)checks the document) but still a
    pure read.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data = doc.proofing()
            emit(data, as_text=not ctx.obj["as_json"], text=_fmt_proofing(data))

    _run(ctx, go)


@click.command(name="lint")
@click.option("--rule", "rule", multiple=True, help="Only run these rule ids/tags (repeatable).")
@click.option("--exclude", "exclude", multiple=True, help="Skip these rule ids/tags (repeatable).")
@click.option(
    "--within",
    "within",
    default=None,
    help="Scope the audit to an anchor id (heading:N, range:S-E, table:N:R:C).",
)
@click.option(
    "--profile",
    "profile",
    default=None,
    help="Path to a JSON house-style profile (enables policy rules + their targets).",
)
@click.pass_context
def lint_cmd(
    ctx: click.Context,
    rule: tuple[str, ...],
    exclude: tuple[str, ...],
    within: str | None,
    profile: str | None,
) -> None:
    """Audit the document for publishing-quality defects (pure read).

    Emits a severity-ranked list of findings — dangling headings, multi-page
    tables with no repeating header, numbered lists Word split into independent
    runs, direct formatting that drifted from the applied style. Each `fixable`
    finding can be applied by `regularize`. `--rule`/`--exclude` select rules by
    id or tag; `--within` scopes to an anchor. `--profile PATH` loads a JSON
    house-style config that enables **policy** rules and supplies their targets.
    """
    selector = _rules_selector(rule, exclude)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            findings = doc.lint(rules=selector, within=within, profile=profile)
            emit(findings, as_text=not ctx.obj["as_json"], text=_fmt_lint(findings))

    _run(ctx, go)


@click.command(name="regularize")
@click.option("--rule", "rule", multiple=True, help="Only run these rule ids/tags (repeatable).")
@click.option("--exclude", "exclude", multiple=True, help="Skip these rule ids/tags (repeatable).")
@click.option("--within", "within", default=None, help="Scope to an anchor id.")
@click.option(
    "--profile",
    "profile",
    default=None,
    help="Path to a JSON house-style profile (enables policy rules + their targets).",
)
@click.option(
    "--dry-run",
    "dry_run",
    is_flag=True,
    default=False,
    help="Plan the fixes (in the findings) without writing anything.",
)
@click.option(
    "--allow-content",
    "allow_content",
    is_flag=True,
    default=False,
    help="Also apply content-changing fixes (insert/delete content), not just formatting.",
)
@click.pass_context
def regularize_cmd(
    ctx: click.Context,
    rule: tuple[str, ...],
    exclude: tuple[str, ...],
    within: str | None,
    profile: str | None,
    dry_run: bool,
    allow_content: bool,
) -> None:
    """Apply the fixable lint findings in one atomic-undo step.

    Runs `lint`, then applies every fixable finding's fix inside a single
    edit (one Ctrl-Z reverts the whole pass; selection/scroll preserved). The
    default fixes are targeted and idempotent — a second `regularize` is a no-op.
    Returns `{applied, skipped, deferred, findings}`. `--dry-run` plans without
    writing. `--profile PATH` enables policy rules (justify, line-spacing,
    numeric-column alignment) and their fixes.

    Formatting fixes apply by default; content-changing fixes (insert a caption,
    delete a stray paragraph, strip a watermark) are withheld into `deferred`
    unless you pass `--allow-content`.
    """
    selector = _rules_selector(rule, exclude)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            report = doc.regularize(
                rules=selector,
                within=within,
                profile=profile,
                dry_run=dry_run,
                allow_content=allow_content,
            )
            emit(report, as_text=not ctx.obj["as_json"], text=_fmt_regularize(report))

    _run(ctx, go)
