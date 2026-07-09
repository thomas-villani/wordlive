"""Floating-shape commands."""

from __future__ import annotations

from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_shapes,
    _shape_anchor,
)


@click.command(name="shapes")
@click.pass_context
def shapes_cmd(ctx: click.Context) -> None:
    """List the document's floating shapes (shape:N id, kind, size, wrap, para:N).

    Text boxes, floating images, and WordArt — the things `shape:N` addresses.
    Restyle one with set-shape-wrap / -position / -size / format-shape /
    replace-shape-image. Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.shapes.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_shapes(rows))

    _run(ctx, go)


@click.command(name="set-shape-wrap")
@click.option("--anchor-id", "anchor_id", required=True, help="Shape anchor (shape:N).")
@click.option(
    "--wrap",
    "wrap",
    default=None,
    type=click.Choice(["square", "tight", "through", "top-bottom", "front", "behind"]),
    help="How body text flows around the shape.",
)
@click.option(
    "--side",
    "side",
    default=None,
    type=click.Choice(["both", "left", "right", "largest"]),
    help="Which sides text flows past (square/tight/through only).",
)
@click.option("--distance-top", "distance_top", default=None, help="Top standoff (pt / '0.1in').")
@click.option(
    "--distance-bottom", "distance_bottom", default=None, help="Bottom standoff (pt / '0.1in')."
)
@click.option(
    "--distance-left", "distance_left", default=None, help="Left standoff (pt / '0.1in')."
)
@click.option(
    "--distance-right", "distance_right", default=None, help="Right standoff (pt / '0.1in')."
)
@click.pass_context
def set_shape_wrap_cmd(
    ctx: click.Context,
    anchor_id: str,
    wrap: str | None,
    side: str | None,
    distance_top: str | None,
    distance_bottom: str | None,
    distance_left: str | None,
    distance_right: str | None,
) -> None:
    """Set how body text wraps around a floating shape (atomic-undo). Pass at least one option."""
    raw: dict[str, Any] = {
        "wrap": wrap,
        "side": side,
        "distance_top": distance_top,
        "distance_bottom": distance_bottom,
        "distance_left": distance_left,
        "distance_right": distance_right,
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --wrap / --side / --distance-*")

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set shape wrap {anchor_id}"):
                anchor.set_wrap(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"set wrap of {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="set-shape-crop")
@click.option("--anchor-id", "anchor_id", required=True, help="Picture shape anchor (shape:N).")
@click.option("--left", "left", default=None, help="Trim off the left edge (pt / '0.2in').")
@click.option("--top", "top", default=None, help="Trim off the top edge (pt / '0.2in').")
@click.option("--right", "right", default=None, help="Trim off the right edge (pt / '0.2in').")
@click.option("--bottom", "bottom", default=None, help="Trim off the bottom edge (pt / '0.2in').")
@click.pass_context
def set_shape_crop_cmd(
    ctx: click.Context,
    anchor_id: str,
    left: str | None,
    top: str | None,
    right: str | None,
    bottom: str | None,
) -> None:
    """Crop a floating picture shape in from its edges (atomic-undo). Pass at least one edge."""
    raw: dict[str, Any] = {"left": left, "top": top, "right": right, "bottom": bottom}
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --left / --top / --right / --bottom")

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: crop shape {anchor_id}"):
                anchor.set_crop(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"cropped {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="set-shape-position")
@click.option("--anchor-id", "anchor_id", required=True, help="Shape anchor (shape:N).")
@click.option("--left", "left", default=None, help="Horizontal offset (pt / '2in' / 'center').")
@click.option("--top", "top", default=None, help="Vertical offset (pt / '2in' / 'center').")
@click.option(
    "--relative-to",
    "relative_to",
    type=click.Choice(["margin", "page"]),
    default=None,
    help="Frame the offsets are measured from.",
)
@click.pass_context
def set_shape_position_cmd(
    ctx: click.Context, anchor_id: str, left: str | None, top: str | None, relative_to: str | None
) -> None:
    """Reposition a floating shape (atomic-undo). Pass at least one option."""
    raw: dict[str, Any] = {"left": left, "top": top, "relative_to": relative_to}
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --left / --top / --relative-to")

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set shape position {anchor_id}"):
                anchor.set_position(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"repositioned {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="set-shape-size")
@click.option("--anchor-id", "anchor_id", required=True, help="Shape anchor (shape:N).")
@click.option("--width", "width", default=None, help="Width (pt or '3in').")
@click.option("--height", "height", default=None, help="Height (pt or '2cm').")
@click.option(
    "--lock-aspect/--no-lock-aspect",
    "lock_aspect",
    default=None,
    help="Lock the aspect ratio for proportional scaling.",
)
@click.pass_context
def set_shape_size_cmd(
    ctx: click.Context,
    anchor_id: str,
    width: str | None,
    height: str | None,
    lock_aspect: bool | None,
) -> None:
    """Resize a floating shape (atomic-undo). Pass at least one option."""
    raw: dict[str, Any] = {"width": width, "height": height, "lock_aspect": lock_aspect}
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --width / --height / --lock-aspect")

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set shape size {anchor_id}"):
                anchor.set_size(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"resized {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="format-shape")
@click.option("--anchor-id", "anchor_id", required=True, help="Shape anchor (shape:N).")
@click.option("--fill", "fill", default=None, help="Fill colour (e.g. '#eeeeff' / 'navy').")
@click.option("--border-color", "border_color", default=None, help="Outline colour.")
@click.option("--no-border", "no_border", is_flag=True, default=False, help="No outline.")
@click.option(
    "--default-border", "default_border", is_flag=True, default=False, help="Default outline."
)
@click.option(
    "--border-weight", "border_weight", default=None, help="Outline thickness (pt / '1.5pt')."
)
@click.pass_context
def format_shape_cmd(
    ctx: click.Context,
    anchor_id: str,
    fill: str | None,
    border_color: str | None,
    no_border: bool,
    default_border: bool,
    border_weight: str | None,
) -> None:
    """Set a floating shape's fill and outline (atomic-undo). Pass at least one option."""
    if sum(bool(x) for x in (no_border, default_border, border_color is not None)) > 1:
        raise click.UsageError(
            "pass at most one of --no-border / --default-border / --border-color"
        )
    border: str | bool | None = None
    if no_border:
        border = False
    elif default_border:
        border = True
    elif border_color is not None:
        border = border_color
    raw: dict[str, Any] = {"fill": fill, "border": border, "border_weight": border_weight}
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one formatting option")

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: format shape {anchor_id}"):
                anchor.format(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"formatted {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="set-shape-alt-text")
@click.option("--anchor-id", "anchor_id", required=True, help="Shape anchor (shape:N).")
@click.option("--text", "text", required=True, help="Accessibility (alt) text.")
@click.pass_context
def set_shape_alt_text_cmd(ctx: click.Context, anchor_id: str, text: str) -> None:
    """Set a floating shape's accessibility (alt) text (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set shape alt text {anchor_id}"):
                anchor.set_alt_text(text)
            emit(
                {"ok": True, "anchor_id": anchor_id, "alt_text": text},
                as_text=not ctx.obj["as_json"],
                text=f"set alt text of {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="set-shape-text")
@click.option("--anchor-id", "anchor_id", required=True, help="Text-box shape anchor (shape:N).")
@click.option("--text", "text", required=True, help="New text-box contents.")
@click.pass_context
def set_shape_text_cmd(ctx: click.Context, anchor_id: str, text: str) -> None:
    """Replace a text box's contents (atomic-undo). Needs a text-box shape."""

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set shape text {anchor_id}"):
                anchor.set_text(text)
            emit(
                {"ok": True, "anchor_id": anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"set text of {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="replace-shape-image")
@click.option("--anchor-id", "anchor_id", required=True, help="Picture shape anchor (shape:N).")
@click.option("--path", "path", default=None, help="Replacement image file (local only).")
@click.option("--base64", "base64_value", default=None, help="Replacement image as base64.")
@click.pass_context
def replace_shape_image_cmd(
    ctx: click.Context, anchor_id: str, path: str | None, base64_value: str | None
) -> None:
    """Swap a floating picture's image in place (atomic-undo). Needs a picture shape.

    Pass exactly one of --path / --base64. Preserves wrap / position / size / alt
    text (delete + reinsert at the same anchor).
    """
    if (path is None) == (base64_value is None):
        raise click.UsageError("pass exactly one of --path / --base64")
    image: str = path if path is not None else base64_value  # type: ignore[assignment]

    def go() -> None:
        if path is not None:
            ctx.obj["policy"].screen_image_path(path)
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: replace shape image {anchor_id}"):
                anchor.replace_image(image)
            emit(
                {"ok": True, "anchor_id": anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"replaced image of {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="delete-shape")
@click.option("--anchor-id", "anchor_id", required=True, help="Shape anchor (shape:N).")
@click.pass_context
def delete_shape_cmd(ctx: click.Context, anchor_id: str) -> None:
    """Delete a floating shape (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: delete shape {anchor_id}"):
                anchor.delete()
            emit(
                {"ok": True, "anchor_id": anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"deleted {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="set-shape-rotation")
@click.option("--anchor-id", "anchor_id", required=True, help="Shape anchor (shape:N).")
@click.option("--degrees", "degrees", required=True, help="Clockwise rotation in degrees.")
@click.pass_context
def set_shape_rotation_cmd(ctx: click.Context, anchor_id: str, degrees: str) -> None:
    """Rotate a floating shape (atomic-undo). Absolute angle in degrees."""

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set shape rotation {anchor_id}"):
                anchor.set_rotation(degrees)
            emit(
                {"ok": True, "anchor_id": anchor_id, "rotation": float(degrees)},
                as_text=not ctx.obj["as_json"],
                text=f"rotated {anchor_id} -> {degrees}deg",
            )

    _run(ctx, go)


@click.command(name="set-shape-z-order")
@click.option("--anchor-id", "anchor_id", required=True, help="Shape anchor (shape:N).")
@click.option(
    "--order",
    "order",
    required=True,
    type=click.Choice(["front", "back", "forward", "backward"]),
    help="Restack the shape in the floating layer.",
)
@click.pass_context
def set_shape_z_order_cmd(ctx: click.Context, anchor_id: str, order: str) -> None:
    """Restack a floating shape (atomic-undo) — front / back / forward / backward."""

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set shape z-order {anchor_id}"):
                anchor.set_z_order(order)
            emit(
                {"ok": True, "anchor_id": anchor_id, "order": order},
                as_text=not ctx.obj["as_json"],
                text=f"restacked {anchor_id} -> {order}",
            )

    _run(ctx, go)


@click.command(name="set-shape-text-frame")
@click.option("--anchor-id", "anchor_id", required=True, help="Text-box shape anchor (shape:N).")
@click.option("--margin-left", "margin_left", default=None, help="Left inset (pt / '0.1in').")
@click.option("--margin-right", "margin_right", default=None, help="Right inset (pt / '0.1in').")
@click.option("--margin-top", "margin_top", default=None, help="Top inset (pt / '0.1in').")
@click.option("--margin-bottom", "margin_bottom", default=None, help="Bottom inset (pt / '0.1in').")
@click.option(
    "--word-wrap/--no-word-wrap",
    "word_wrap",
    default=None,
    help="Wrap text to the box width.",
)
@click.pass_context
def set_shape_text_frame_cmd(
    ctx: click.Context,
    anchor_id: str,
    margin_left: str | None,
    margin_right: str | None,
    margin_top: str | None,
    margin_bottom: str | None,
    word_wrap: bool | None,
) -> None:
    """Set a text box's internal margins and word-wrap (atomic-undo). Pass at least one option."""
    raw: dict[str, Any] = {
        "margin_left": margin_left,
        "margin_right": margin_right,
        "margin_top": margin_top,
        "margin_bottom": margin_bottom,
        "word_wrap": word_wrap,
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --margin-* / --word-wrap")

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set shape text frame {anchor_id}"):
                anchor.set_text_frame(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"set text frame of {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="group-shapes")
@click.option(
    "--anchor-id",
    "anchor_ids",
    required=True,
    multiple=True,
    help="A shape to group (pass two or more times).",
)
@click.pass_context
def group_shapes_cmd(ctx: click.Context, anchor_ids: tuple[str, ...]) -> None:
    """Group two or more floating shapes into one (atomic-undo). Returns the group's shape:N."""
    if len(anchor_ids) < 2:
        raise click.UsageError("pass --anchor-id at least twice (group needs two or more shapes)")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: group {len(anchor_ids)} shapes"):
                group = doc.group_shapes(*anchor_ids)
            emit(
                {"ok": True, "anchor_id": group.anchor_id, "members": list(anchor_ids)},
                as_text=not ctx.obj["as_json"],
                text=f"grouped {list(anchor_ids)} -> {group.anchor_id}",
            )

    _run(ctx, go)


@click.command(name="ungroup-shape")
@click.option("--anchor-id", "anchor_id", required=True, help="Group shape anchor (shape:N).")
@click.pass_context
def ungroup_shape_cmd(ctx: click.Context, anchor_id: str) -> None:
    """Dissolve a group shape into its members (atomic-undo). Returns the members' shape:N ids."""

    def go() -> None:
        with attach() as word:
            doc, anchor = _shape_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: ungroup {anchor_id}"):
                members = anchor.ungroup()
            ids = [m.anchor_id for m in members]
            emit(
                {"ok": True, "anchor_id": anchor_id, "members": ids},
                as_text=not ctx.obj["as_json"],
                text=f"ungrouped {anchor_id} -> {ids}",
            )

    _run(ctx, go)
