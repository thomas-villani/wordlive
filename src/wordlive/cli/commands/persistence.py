"""Save / save-as / export-pdf."""

from __future__ import annotations

from pathlib import Path

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _parse_pages_range,
)


@click.command(name="save")
@click.pass_context
def save_cmd(ctx: click.Context) -> None:
    """Save the document to its existing file (gated).

    Fails if the document has never been saved — use `save-as PATH` first. The
    existing path must itself sit inside a whitelisted `--save-dir`
    (or `WORDLIVE_SAVE_DIRS`); with no whitelist, saving is off.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            # The save target is the doc's own path; it must also be whitelisted.
            ctx.obj["policy"].resolve_save_target(doc.path)
            path = doc.save()
            emit(
                {"ok": True, "path": path, "saved": True},
                as_text=not ctx.obj["as_json"],
                text=f"saved {path}",
            )

    _run(ctx, go)


@click.command(name="save-as")
@click.argument("path", type=click.Path(path_type=Path))
@click.option(
    "--format",
    "fmt",
    type=click.Choice(["docx"]),
    default="docx",
    show_default=True,
    help="Output format (PDF is `export-pdf`).",
)
@click.option(
    "--overwrite",
    "overwrite",
    is_flag=True,
    default=False,
    help="Allow overwriting an existing file (default: refuse).",
)
@click.pass_context
def save_as_cmd(ctx: click.Context, path: Path, fmt: str, overwrite: bool) -> None:
    """Save the document to PATH (gated).

    PATH must resolve inside a whitelisted `--save-dir` (or `WORDLIVE_SAVE_DIRS`);
    with no whitelist, saving is off. Refuses to clobber an existing file unless
    `--overwrite` is given. For PDF, use `export-pdf`.
    """

    def go() -> None:
        target = ctx.obj["policy"].resolve_save_target(path)
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            written = doc.save_as(target, fmt=fmt, overwrite=overwrite)
            emit(
                {"ok": True, "path": written, "format": fmt},
                as_text=not ctx.obj["as_json"],
                text=f"saved {written}",
            )

    _run(ctx, go)


@click.command(name="export-pdf")
@click.argument("path", type=click.Path(path_type=Path))
@click.option(
    "--pages",
    "pages_range",
    default=None,
    help="Export an inclusive page span, e.g. '2-4' (or a single page '3'). "
    "Default: the whole document.",
)
@click.pass_context
def export_pdf_cmd(ctx: click.Context, path: Path, pages_range: str | None) -> None:
    """Export the document (or a page span) to a PDF at PATH (gated).

    PATH must resolve inside a whitelisted `--save-dir` (or `WORDLIVE_SAVE_DIRS`).
    The recommended "hand back a deliverable" path — a pixel-faithful render via
    Word's PDF engine (the same one `snapshot` uses). Overwrites an existing PDF.
    """
    from_page: int | None = None
    to_page: int | None = None
    if pages_range is not None:
        if "-" in pages_range:
            from_page, to_page = _parse_pages_range(pages_range)
        else:
            try:
                from_page = int(pages_range)
            except ValueError as e:
                raise click.UsageError("--pages must be 'N' or 'A-B' (inclusive)") from e
            if from_page < 1:
                raise click.UsageError("--pages must be >= 1")

    def go() -> None:
        target = ctx.obj["policy"].resolve_save_target(path)
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            written = doc.export_pdf(target, from_page=from_page, to_page=to_page)
            emit(
                {"ok": True, "path": written},
                as_text=not ctx.obj["as_json"],
                text=f"exported {written}",
            )

    _run(ctx, go)
