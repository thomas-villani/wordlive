"""Equation commands."""

from __future__ import annotations

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_equations,
)


@click.command(name="equations")
@click.pass_context
def equations_cmd(ctx: click.Context) -> None:
    """List the document's equations (equation:N id, type, linear preview, para:N).

    The discovery half of equation editing: see what math is in the document and
    address it by `equation:N`. Reading is non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.equations.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_equations(rows))

    _run(ctx, go)


@click.command(name="insert-equation")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor to insert the equation relative to."
)
@click.option(
    "--unicodemath",
    "unicodemath",
    default=None,
    help="UnicodeMath linear string (native, no extra).",
)
@click.option("--latex", "latex", default=None, help="LaTeX math string (needs the 'latex' extra).")
@click.option(
    "--mathml", "mathml", default=None, help="MathML string, or '-' to read it from stdin."
)
@click.option(
    "--display/--inline",
    "display",
    default=True,
    show_default="--display",
    help="Display equation (own centred line) vs inline (left-aligned).",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def insert_equation_cmd(
    ctx: click.Context,
    anchor_id: str,
    unicodemath: str | None,
    latex: str | None,
    mathml: str | None,
    display: bool,
    before: bool,
) -> None:
    """Insert an equation at any anchor (atomic-undo), from UnicodeMath, LaTeX, or MathML.

    Exactly one of --unicodemath / --latex / --mathml is required. UnicodeMath is
    native (no extra); LaTeX needs the 'latex' extra; MathML (or '--mathml -' from
    stdin) goes through Office's own transform. The equation lands on its own
    paragraph — --display centres it, --inline left-aligns it.
    """
    given = [
        n
        for n, v in (("--unicodemath", unicodemath), ("--latex", latex), ("--mathml", mathml))
        if v is not None
    ]
    if len(given) != 1:
        raise click.UsageError("pass exactly one of --unicodemath / --latex / --mathml")
    if mathml == "-":
        mathml = click.get_text_stream("stdin").read()
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert equation {where} {anchor_id}"):
                equation = anchor.insert_equation(
                    unicodemath=unicodemath,
                    latex=latex,
                    mathml=mathml,
                    where=where,
                    display=display,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "equation": equation.index,
                    "equation_anchor_id": equation.anchor_id,
                    "display": display,
                    "where": where,
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {equation.anchor_id} {where} {anchor_id}",
            )

    _run(ctx, go)
