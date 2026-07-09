"""Styles plus character/paragraph formatting."""

from __future__ import annotations

from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_style_list,
    _parse_color,
)


@click.group(name="style")
def style() -> None:
    """Read or apply paragraph and character styles."""


@style.command(name="list")
@click.pass_context
def style_list(ctx: click.Context) -> None:
    """List every style defined in the document."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.styles.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_style_list(rows))

    _run(ctx, go)


@style.command(name="apply")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to apply the style to.")
@click.option(
    "--name", "name", required=True, help="Style name (must already exist in the document)."
)
@click.pass_context
def style_apply(ctx: click.Context, anchor_id: str, name: str) -> None:
    """Apply STYLE NAME to the anchor identified by ANCHOR-ID (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: apply style {name!r} to {anchor_id}"):
                anchor.apply_style(name)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "style": name,
                },
                as_text=not ctx.obj["as_json"],
                text=f"applied style {name!r} to {anchor_id}",
            )

    _run(ctx, go)


@style.command(name="add")
@click.argument("name")
@click.option(
    "--type",
    "style_type",
    default="paragraph",
    type=click.Choice(["paragraph", "character", "table", "list"]),
    help="Kind of style to create.",
)
@click.option("--based-on", "based_on", default=None, help="Existing style to inherit from.")
@click.option(
    "--next-style", "next_style", default=None, help="Style applied to the following paragraph."
)
@click.pass_context
def style_add(
    ctx: click.Context, name: str, style_type: str, based_on: str | None, next_style: str | None
) -> None:
    """Define a new style NAME (atomic-undo). Style its defaults with `style set`."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: add style {name!r}"):
                new = doc.styles.add(
                    name, type=style_type, based_on=based_on, next_style=next_style
                )
            emit(
                {"ok": True, "style": new.name, "type": style_type},
                as_text=not ctx.obj["as_json"],
                text=f"added style {name!r} ({style_type})",
            )

    _run(ctx, go)


@style.command(name="set")
@click.argument("name")
@click.option("--bold/--no-bold", "bold", default=None, help="Bold.")
@click.option("--italic/--no-italic", "italic", default=None, help="Italic.")
@click.option("--underline/--no-underline", "underline", default=None, help="Single underline.")
@click.option("--font", "font", default=None, help="Font family name.")
@click.option("--size", "size", default=None, help="Font size in points or a unit string.")
@click.option("--color", "color", default=None, help="Font colour: name, hex, or r,g,b.")
@click.option(
    "--alignment",
    "alignment",
    default=None,
    type=click.Choice(["left", "center", "centre", "right", "justify"], case_sensitive=False),
    help="Paragraph alignment.",
)
@click.option("--space-before", "space_before", default=None, help="Space before in points/units.")
@click.option("--space-after", "space_after", default=None, help="Space after in points/units.")
@click.option(
    "--line-spacing",
    "line_spacing",
    default=None,
    help="Leading within the paragraph: a multiple (1, 1.5, 2), single/1.5/double, "
    "or an exact length (e.g. 14pt).",
)
@click.option("--based-on", "based_on", default=None, help="Existing style to inherit from.")
@click.option(
    "--next-style", "next_style", default=None, help="Style applied to the following paragraph."
)
@click.pass_context
def style_set(
    ctx: click.Context,
    name: str,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    font: str | None,
    size: str | None,
    color: str | None,
    alignment: str | None,
    space_before: str | None,
    space_after: str | None,
    line_spacing: str | None,
    based_on: str | None,
    next_style: str | None,
) -> None:
    """Set the font / paragraph defaults of an existing style NAME (atomic-undo)."""
    run_raw: dict[str, Any] = {
        "bold": bold,
        "italic": italic,
        "underline": underline,
        "font": font,
        "size": size,
        "color": _parse_color(color),
    }
    para_raw: dict[str, Any] = {
        "alignment": alignment,
        "space_before": space_before,
        "space_after": space_after,
        "line_spacing": line_spacing,
    }
    run_kwargs = {k: v for k, v in run_raw.items() if v is not None}
    para_kwargs = {k: v for k, v in para_raw.items() if v is not None}
    if not run_kwargs and not para_kwargs and based_on is None and next_style is None:
        raise click.UsageError("pass at least one style property to set")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: set style {name!r}"):
                style_obj = doc.styles[name]
                if run_kwargs:
                    style_obj.format_run(**run_kwargs)
                if para_kwargs:
                    style_obj.format_paragraph(**para_kwargs)
                if based_on is not None:
                    style_obj.base_style = based_on
                if next_style is not None:
                    style_obj.next_paragraph_style = next_style
            emit(
                {
                    "ok": True,
                    "style": name,
                    "applied": {**run_kwargs, **para_kwargs},
                    "based_on": based_on,
                    "next_style": next_style,
                },
                as_text=not ctx.obj["as_json"],
                text=f"set style {name!r}",
            )

    _run(ctx, go)


@click.command(name="format-paragraph")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor whose paragraph(s) to format."
)
@click.option(
    "--alignment",
    "alignment",
    default=None,
    type=click.Choice(["left", "center", "centre", "right", "justify"], case_sensitive=False),
    help="Paragraph alignment.",
)
@click.option(
    "--left-indent", "left_indent", type=float, default=None, help="Left indent in points."
)
@click.option(
    "--right-indent", "right_indent", type=float, default=None, help="Right indent in points."
)
@click.option(
    "--first-line-indent",
    "first_line_indent",
    type=float,
    default=None,
    help="First-line indent in points.",
)
@click.option(
    "--space-before",
    "space_before",
    type=float,
    default=None,
    help="Space before paragraph in points.",
)
@click.option(
    "--space-after",
    "space_after",
    type=float,
    default=None,
    help="Space after paragraph in points.",
)
@click.option(
    "--line-spacing",
    "line_spacing",
    default=None,
    help="Leading within the paragraph: a multiple (1, 1.5, 2), single/1.5/double, "
    "or an exact length (e.g. 14pt).",
)
@click.option(
    "--page-break-before/--no-page-break-before",
    "page_break_before",
    default=None,
    help="Force (or clear) a page break before the paragraph — the clean, "
    "reflow-safe way to page-break (e.g. on every Heading 1).",
)
@click.option(
    "--keep-together/--no-keep-together",
    "keep_together",
    default=None,
    help="Keep all lines of the paragraph on one page.",
)
@click.option(
    "--keep-with-next/--no-keep-with-next",
    "keep_with_next",
    default=None,
    help="Keep the paragraph on the same page as the next one (e.g. a heading "
    "with its first body line).",
)
@click.option(
    "--widow-control/--no-widow-control",
    "widow_control",
    default=None,
    help="Prevent a lone first/last line stranded at a page boundary.",
)
@click.pass_context
def format_paragraph_cmd(
    ctx: click.Context,
    anchor_id: str,
    alignment: str | None,
    left_indent: float | None,
    right_indent: float | None,
    first_line_indent: float | None,
    space_before: float | None,
    space_after: float | None,
    line_spacing: str | None,
    page_break_before: bool | None,
    keep_together: bool | None,
    keep_with_next: bool | None,
    widow_control: bool | None,
) -> None:
    """Set paragraph-formatting properties on the anchor's range (atomic-undo)."""
    kwargs: dict[str, Any] = {}
    if alignment is not None:
        kwargs["alignment"] = alignment
    if left_indent is not None:
        kwargs["left_indent"] = left_indent
    if right_indent is not None:
        kwargs["right_indent"] = right_indent
    if first_line_indent is not None:
        kwargs["first_line_indent"] = first_line_indent
    if space_before is not None:
        kwargs["space_before"] = space_before
    if space_after is not None:
        kwargs["space_after"] = space_after
    if line_spacing is not None:
        kwargs["line_spacing"] = line_spacing
    if page_break_before is not None:
        kwargs["page_break_before"] = page_break_before
    if keep_together is not None:
        kwargs["keep_together"] = keep_together
    if keep_with_next is not None:
        kwargs["keep_with_next"] = keep_with_next
    if widow_control is not None:
        kwargs["widow_control"] = widow_control
    if not kwargs:
        raise click.UsageError("pass at least one formatting option")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: format paragraph {anchor_id}"):
                anchor.format_paragraph(**kwargs)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "applied": kwargs,
                },
                as_text=not ctx.obj["as_json"],
                text=f"formatted {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="format-run")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor whose text run(s) to format.")
@click.option("--bold/--no-bold", "bold", default=None, help="Bold.")
@click.option("--italic/--no-italic", "italic", default=None, help="Italic.")
@click.option("--underline/--no-underline", "underline", default=None, help="Single underline.")
@click.option(
    "--strikethrough/--no-strikethrough", "strikethrough", default=None, help="Strikethrough."
)
@click.option("--font", "font", default=None, help="Font family name.")
@click.option("--size", "size", default=None, help="Font size in points or a unit string (12pt).")
@click.option("--color", "color", default=None, help="Font colour: name, hex (#FF0000), or r,g,b.")
@click.option(
    "--highlight",
    "highlight",
    default=None,
    help="Text-highlight colour name (yellow, green, …) or 'none' to clear.",
)
@click.option("--subscript/--no-subscript", "subscript", default=None, help="Subscript.")
@click.option("--superscript/--no-superscript", "superscript", default=None, help="Superscript.")
@click.option("--small-caps/--no-small-caps", "small_caps", default=None, help="Small caps.")
@click.option("--all-caps/--no-all-caps", "all_caps", default=None, help="All caps.")
@click.option(
    "--spacing", "spacing", default=None, help="Character spacing in points or a unit string."
)
@click.pass_context
def format_run_cmd(
    ctx: click.Context,
    anchor_id: str,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    strikethrough: bool | None,
    font: str | None,
    size: str | None,
    color: str | None,
    highlight: str | None,
    subscript: bool | None,
    superscript: bool | None,
    small_caps: bool | None,
    all_caps: bool | None,
    spacing: str | None,
) -> None:
    """Set character-formatting (run-level) properties on the anchor (atomic-undo).

    A colour may be a name, hex (#FF0000), or comma-separated r,g,b. Sizes and
    spacing accept a number (points) or a unit string like 12pt / 1.5mm.
    """
    raw: dict[str, Any] = {
        "bold": bold,
        "italic": italic,
        "underline": underline,
        "strikethrough": strikethrough,
        "font": font,
        "size": size,
        "color": _parse_color(color),
        "highlight": highlight,
        "subscript": subscript,
        "superscript": superscript,
        "small_caps": small_caps,
        "all_caps": all_caps,
        "spacing": spacing,
    }
    kwargs: dict[str, Any] = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one formatting option")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: format run {anchor_id}"):
                anchor.format_run(**kwargs)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "applied": kwargs,
                },
                as_text=not ctx.obj["as_json"],
                text=f"formatted run {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="shading")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor (or cell) to shade.")
@click.option("--fill", "fill", required=True, help="Fill colour: name, hex (#FFFF00), or r,g,b.")
@click.pass_context
def shading_cmd(ctx: click.Context, anchor_id: str, fill: str) -> None:
    """Set the background-fill shading of the anchor's range (atomic-undo)."""
    fill_value = _parse_color(fill)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: shading {anchor_id}"):
                anchor.set_shading(fill=fill_value)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "applied": {"fill": fill_value},
                },
                as_text=not ctx.obj["as_json"],
                text=f"shaded {anchor_id}: {fill_value}",
            )

    _run(ctx, go)


@click.command(name="borders")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor (or cell) to border.")
@click.option(
    "--sides",
    "sides",
    default="all",
    help="Edges: all/box, top, bottom, left, right, horizontal, vertical "
    "(comma-separated for several).",
)
@click.option(
    "--style",
    "style",
    default="single",
    help="Line style: single, double, dot, dash, … or none. "
    "(In exec/MCP this field is named `line_style` to avoid colliding with a "
    "paragraph/table `style` name.)",
)
@click.option("--weight", "weight", type=float, default=0.5, help="Line width in points (snapped).")
@click.option("--color", "color", default=None, help="Border colour: name, hex, or r,g,b.")
@click.pass_context
def borders_cmd(
    ctx: click.Context, anchor_id: str, sides: str, style: str, weight: float, color: str | None
) -> None:
    """Draw borders on the anchor's range or cell (atomic-undo)."""
    side_list = [s.strip() for s in sides.split(",") if s.strip()]
    color_value = _parse_color(color)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: borders {anchor_id}"):
                anchor.set_borders(sides=side_list, style=style, weight=weight, color=color_value)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "applied": {
                        "sides": side_list,
                        "style": style,
                        "weight": weight,
                        "color": color_value,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"bordered {anchor_id}: {side_list} {style}",
            )

    _run(ctx, go)


@click.command(name="drop-cap")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor whose paragraph's first letter to drop."
)
@click.option(
    "--position",
    "position",
    default="dropped",
    type=click.Choice(["dropped", "normal", "margin", "none"], case_sensitive=False),
    help="dropped (into the text), margin (in the left margin), or none (remove).",
)
@click.option("--lines", "lines", type=int, default=3, help="How many lines tall the letter is.")
@click.option(
    "--distance", "distance", default="0", help="Gap from the body text in points or a unit string."
)
@click.option("--font", "font", default=None, help="Font family for the dropped letter.")
@click.pass_context
def drop_cap_cmd(
    ctx: click.Context,
    anchor_id: str,
    position: str,
    lines: int,
    distance: str,
    font: str | None,
) -> None:
    """Turn the first letter of the anchor's paragraph into a drop cap (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: drop cap {anchor_id}"):
                anchor.drop_cap(lines, position=position, distance=distance, font=font)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "applied": {"position": position, "lines": lines},
                },
                as_text=not ctx.obj["as_json"],
                text=f"drop cap on {anchor_id}: {position}",
            )

    _run(ctx, go)


@click.command(name="tab-stop")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor whose paragraph(s) to tab.")
@click.option(
    "--position",
    "position",
    required=True,
    help="Distance from the left margin in points or a unit string (3in).",
)
@click.option(
    "--align",
    "align",
    default="left",
    type=click.Choice(["left", "center", "centre", "right", "decimal", "bar"]),
    help="Tab alignment.",
)
@click.option(
    "--leader",
    "leader",
    default=None,
    type=click.Choice(["none", "dots", "dashes", "lines", "heavy", "middle-dot"]),
    help="Leader fill drawn up to the stop.",
)
@click.pass_context
def tab_stop_cmd(
    ctx: click.Context, anchor_id: str, position: str, align: str, leader: str | None
) -> None:
    """Add a tab stop to the anchor's paragraph(s) (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: tab stop {anchor_id}"):
                anchor.add_tab_stop(position, align=align, leader=leader)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "applied": {"position": position, "align": align, "leader": leader},
                },
                as_text=not ctx.obj["as_json"],
                text=f"tab stop on {anchor_id}: {position} {align}",
            )

    _run(ctx, go)
