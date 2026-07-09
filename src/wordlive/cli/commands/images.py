"""Inline-image reads and restyle."""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_images,
    _image_anchor,
)


@click.command(name="images")
@click.pass_context
def images_cmd(ctx: click.Context) -> None:
    """List the document's embedded images (image:N id, MIME, size, alt text, para:N).

    The discovery half of image extraction: see what pictures are in the document
    before pulling any bytes with `read-image`. Reading is non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.images.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_images(rows))

    _run(ctx, go)


@click.command(name="read-image")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="The image to read: an image:N id (or any anchor whose range holds one picture).",
)
@click.option(
    "--out",
    "out",
    type=click.Path(path_type=Path),
    default=None,
    help="Write the raw image bytes here. Without --out, base64 data is returned in the JSON.",
)
@click.pass_context
def read_image_cmd(ctx: click.Context, anchor_id: str, out: Path | None) -> None:
    """Extract an embedded image's bytes + MIME type (the read side for vision models).

    Resolve the picture by `--anchor-id image:N` (discover them with `images`) or
    any anchor whose range contains exactly one image. With `--out` the raw bytes
    are written to that file and the JSON reports `{path, mime, bytes}`; without
    it, base64 data is returned inline (`{mime, bytes, base64}`). A range with no
    image — or more than one — is a bad-input error (exit 1). Read-only.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data, mime = doc.anchor_by_id(anchor_id).read_image()
            result: dict[str, Any] = {
                "ok": True,
                "anchor_id": anchor_id,
                "mime": mime,
                "bytes": len(data),
            }
            if out is not None:
                out.write_bytes(data)
                result["path"] = str(out)
            else:
                result["base64"] = base64.b64encode(data).decode("ascii")
            where = result.get("path") or "base64"
            emit(
                result,
                as_text=not ctx.obj["as_json"],
                text=f"{anchor_id}: {mime}, {len(data)} bytes → {where}",
            )

    _run(ctx, go)


@click.command(name="set-image-alt-text")
@click.option("--anchor-id", "anchor_id", required=True, help="Inline image anchor (image:N).")
@click.option("--text", "text", required=True, help="Accessibility (alt) text.")
@click.pass_context
def set_image_alt_text_cmd(ctx: click.Context, anchor_id: str, text: str) -> None:
    """Set an inline picture's accessibility (alt) text (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc, anchor = _image_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set image alt text {anchor_id}"):
                anchor.set_alt_text(text)
            emit(
                {"ok": True, "anchor_id": anchor_id, "alt_text": text},
                as_text=not ctx.obj["as_json"],
                text=f"set alt text of {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="set-image-size")
@click.option("--anchor-id", "anchor_id", required=True, help="Inline image anchor (image:N).")
@click.option("--width", "width", default=None, help="Width (pt or '3in').")
@click.option("--height", "height", default=None, help="Height (pt or '2cm').")
@click.option(
    "--lock-aspect/--no-lock-aspect",
    "lock_aspect",
    default=None,
    help="Lock the aspect ratio for proportional scaling.",
)
@click.pass_context
def set_image_size_cmd(
    ctx: click.Context,
    anchor_id: str,
    width: str | None,
    height: str | None,
    lock_aspect: bool | None,
) -> None:
    """Resize an inline picture (atomic-undo). Pass at least one option.

    Re-wrapping (floating) an image isn't here — that crosses it into shape:N
    via insert-image --wrap.
    """
    raw: dict[str, Any] = {"width": width, "height": height, "lock_aspect": lock_aspect}
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --width / --height / --lock-aspect")

    def go() -> None:
        with attach() as word:
            doc, anchor = _image_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set image size {anchor_id}"):
                anchor.set_size(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"resized {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="set-image-crop")
@click.option("--anchor-id", "anchor_id", required=True, help="Inline image anchor (image:N).")
@click.option("--left", "left", default=None, help="Trim off the left edge (pt / '0.2in').")
@click.option("--top", "top", default=None, help="Trim off the top edge (pt / '0.2in').")
@click.option("--right", "right", default=None, help="Trim off the right edge (pt / '0.2in').")
@click.option("--bottom", "bottom", default=None, help="Trim off the bottom edge (pt / '0.2in').")
@click.pass_context
def set_image_crop_cmd(
    ctx: click.Context,
    anchor_id: str,
    left: str | None,
    top: str | None,
    right: str | None,
    bottom: str | None,
) -> None:
    """Crop an inline picture in from its edges (atomic-undo). Pass at least one edge."""
    raw: dict[str, Any] = {"left": left, "top": top, "right": right, "bottom": bottom}
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --left / --top / --right / --bottom")

    def go() -> None:
        with attach() as word:
            doc, anchor = _image_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: crop image {anchor_id}"):
                anchor.set_crop(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"cropped {anchor_id}: {kwargs}",
            )

    _run(ctx, go)
