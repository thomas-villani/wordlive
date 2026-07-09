"""Sections, headers/footers, page setup, watermark."""

from __future__ import annotations

from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _SECTION_OPTION,
    _WHICH_OPTION,
    _emit_section_list,
)


@click.command(name="watermark")
@click.option("--text", "text", default=None, help="Watermark text (e.g. DRAFT). Required to set.")
@click.option(
    "--remove", "remove", is_flag=True, default=False, help="Remove any existing text watermark."
)
@click.option("--font", "font", default="Calibri", show_default=True, help="Watermark font.")
@click.option("--color", "color", default="#C0C0C0", show_default=True, help="Fill colour.")
@click.option(
    "--layout",
    "layout",
    type=click.Choice(["diagonal", "horizontal"]),
    default="diagonal",
    show_default=True,
    help="Diagonal (45°) or horizontal.",
)
@click.option(
    "--transparent/--solid",
    "transparent",
    default=True,
    show_default="--transparent",
    help="Wash the watermark out (50%) so body text stays readable.",
)
@click.pass_context
def watermark_cmd(
    ctx: click.Context,
    text: str | None,
    remove: bool,
    font: str,
    color: str,
    layout: str,
    transparent: bool,
) -> None:
    """Stamp (or --remove) a text watermark behind every page (atomic-undo)."""
    if remove == (text is not None):
        raise click.UsageError("provide either --text (to set) or --remove (not both)")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            if remove:
                with doc.edit("CLI: remove watermark"):
                    n = doc.remove_watermark()
                emit(
                    {"ok": True, "removed": n},
                    as_text=not ctx.obj["as_json"],
                    text=f"removed {n} watermark shape(s)",
                )
                return
            assert text is not None  # guaranteed by the validation above
            with doc.edit("CLI: set watermark"):
                n = doc.set_watermark(
                    text, font=font, color=color, layout=layout, semitransparent=transparent
                )
            emit(
                {"ok": True, "text": text, "sections": n},
                as_text=not ctx.obj["as_json"],
                text=f"watermarked {n} section(s) with {text!r}",
            )

    _run(ctx, go)


@click.command(name="sections")
@click.pass_context
def sections_cmd(ctx: click.Context) -> None:
    """List sections with their page setup (orientation, margins, page size).

    Per-section page geometry is written with `page-setup --section N`;
    headers/footers live in `header` / `footer`.
    """
    _run(ctx, lambda: _emit_section_list(ctx))


@click.group(name="section", hidden=True)
def section() -> None:
    """Deprecated: use the top-level `sections` command. Kept one release."""


@section.command(name="list")
@click.pass_context
def section_list(ctx: click.Context) -> None:
    """Deprecated alias for the top-level `sections` command."""
    _run(ctx, lambda: _emit_section_list(ctx))


@click.group(name="header")
def header() -> None:
    """Read or write section headers (anchor id: header:S:WHICH)."""


@header.command(name="read")
@_SECTION_OPTION
@_WHICH_OPTION
@click.pass_context
def header_read(ctx: click.Context, section_index: int, which: str) -> None:
    """Read the text of a section header."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            hf = doc.sections[section_index].header(which)
            text = hf.text
            emit(
                {"anchor_id": hf.anchor_id, "section": section_index, "which": which, "text": text},
                as_text=not ctx.obj["as_json"],
                text=text,
            )

    _run(ctx, go)


@header.command(name="write")
@_SECTION_OPTION
@_WHICH_OPTION
@click.option("--text", "text", required=True, help="New header text.")
@click.pass_context
def header_write(ctx: click.Context, section_index: int, which: str, text: str) -> None:
    """Set the text of a section header (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            hf = doc.sections[section_index].header(which)
            with doc.edit(f"CLI: write {hf.anchor_id}"):
                hf.set_text(text)
            emit(
                {"ok": True, "anchor_id": hf.anchor_id, "section": section_index, "which": which},
                as_text=not ctx.obj["as_json"],
                text=f"wrote {hf.anchor_id}",
            )

    _run(ctx, go)


@click.group(name="footer")
def footer() -> None:
    """Read or write section footers (anchor id: footer:S:WHICH)."""


@footer.command(name="read")
@_SECTION_OPTION
@_WHICH_OPTION
@click.pass_context
def footer_read(ctx: click.Context, section_index: int, which: str) -> None:
    """Read the text of a section footer."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            hf = doc.sections[section_index].footer(which)
            text = hf.text
            emit(
                {"anchor_id": hf.anchor_id, "section": section_index, "which": which, "text": text},
                as_text=not ctx.obj["as_json"],
                text=text,
            )

    _run(ctx, go)


@footer.command(name="write")
@_SECTION_OPTION
@_WHICH_OPTION
@click.option("--text", "text", required=True, help="New footer text.")
@click.pass_context
def footer_write(ctx: click.Context, section_index: int, which: str, text: str) -> None:
    """Set the text of a section footer (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            hf = doc.sections[section_index].footer(which)
            with doc.edit(f"CLI: write {hf.anchor_id}"):
                hf.set_text(text)
            emit(
                {"ok": True, "anchor_id": hf.anchor_id, "section": section_index, "which": which},
                as_text=not ctx.obj["as_json"],
                text=f"wrote {hf.anchor_id}",
            )

    _run(ctx, go)


@click.command(name="page-setup")
@_SECTION_OPTION
@click.option(
    "--margins", "margins", default=None, help="All four margins at once (points or '1in')."
)
@click.option("--top-margin", "top_margin", default=None, help="Top margin (points or '1in').")
@click.option(
    "--bottom-margin", "bottom_margin", default=None, help="Bottom margin (points or '1in')."
)
@click.option("--left-margin", "left_margin", default=None, help="Left margin (points or '1in').")
@click.option(
    "--right-margin", "right_margin", default=None, help="Right margin (points or '1in')."
)
@click.option("--gutter", "gutter", default=None, help="Binding gutter (points or a unit string).")
@click.option(
    "--orientation",
    "orientation",
    type=click.Choice(["portrait", "landscape"]),
    default=None,
    help="Page orientation.",
)
@click.option(
    "--paper-size",
    "paper_size",
    type=click.Choice(["letter", "legal", "tabloid", "a3", "a4", "a5"]),
    default=None,
    help="Paper size (resizes the page).",
)
@click.option(
    "--columns", "columns", type=int, default=None, help="Number of equal newspaper columns."
)
@click.option(
    "--column-spacing", "column_spacing", default=None, help="Gap between columns (points/unit)."
)
@click.pass_context
def page_setup_cmd(
    ctx: click.Context,
    section_index: int,
    margins: str | None,
    top_margin: str | None,
    bottom_margin: str | None,
    left_margin: str | None,
    right_margin: str | None,
    gutter: str | None,
    orientation: str | None,
    paper_size: str | None,
    columns: int | None,
    column_spacing: str | None,
) -> None:
    """Set a section's page geometry: margins, orientation, paper size, columns (atomic-undo)."""
    applied: dict[str, Any] = {
        k: v
        for k, v in {
            "margins": margins,
            "top_margin": top_margin,
            "bottom_margin": bottom_margin,
            "left_margin": left_margin,
            "right_margin": right_margin,
            "gutter": gutter,
            "orientation": orientation,
            "paper_size": paper_size,
            "columns": columns,
            "column_spacing": column_spacing,
        }.items()
        if v is not None
    }
    if not applied:
        raise click.UsageError("give at least one page-setup option to change.")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: page setup section {section_index}"):
                doc.sections[section_index].set_page_setup(**applied)
            emit(
                {"ok": True, "section": section_index, "applied": applied},
                as_text=not ctx.obj["as_json"],
                text=f"set page setup on section {section_index}: {applied}",
            )

    _run(ctx, go)
