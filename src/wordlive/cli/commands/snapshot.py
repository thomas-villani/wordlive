"""Page/section snapshot rendering."""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_snapshot,
    _parse_pages_range,
)


@click.command(name="snapshot")
@click.option(
    "--anchor-id",
    "anchor_id",
    default=None,
    help="Render the page(s) this anchor sits on (a heading expands to its whole section).",
)
@click.option("--page", "page", type=int, default=None, help="Render a single 1-based page.")
@click.option(
    "--pages",
    "pages_range",
    default=None,
    help="Render an inclusive page span, e.g. '2-4'.",
)
@click.option(
    "--out",
    "out",
    type=click.Path(path_type=Path),
    default=None,
    help="Write the PNG here. Multiple pages are written as <stem>-p<N><suffix>. "
    "Without --out, base64 PNG data is returned inline in the JSON.",
)
@click.option(
    "--dpi", "dpi", type=int, default=150, show_default=True, help="Render resolution (dots/inch)."
)
@click.option(
    "--max-dim",
    "max_dim",
    type=int,
    default=None,
    help="Cap each page's long edge to this many pixels (only ever lowers resolution). "
    "The lever for a cheap whole-document layout check — ~1000 stays legible at a "
    "fraction of the tokens; predictable per-page cost regardless of paper size.",
)
@click.option(
    "--markup",
    "markup",
    type=click.Choice(["none", "all"]),
    default="none",
    show_default=True,
    help="'all' renders tracked changes and comments as visible revision marks.",
)
@click.pass_context
def snapshot_cmd(
    ctx: click.Context,
    anchor_id: str | None,
    page: int | None,
    pages_range: str | None,
    out: Path | None,
    dpi: int,
    max_dim: int | None,
    markup: str,
) -> None:
    """Render document page(s) to PNG so a vision model can see the layout.

    Word exports a pixel-faithful PDF of the document it has open and wordlive
    rasterises the requested pages — a true WYSIWYG image (real fonts, spacing,
    page geometry) for iterating on style and formatting. Read-only.

    Choose at most one target: `--anchor-id` (the page(s) an anchor occupies; a
    `heading:` expands to its whole section), `--page N`, or `--pages A-B`. With
    none, the whole document is rendered. With `--out` the image is written to
    disk (one file per page); otherwise base64 PNG data is returned inline.

    `--max-dim N` caps each page's long edge to N pixels — pair it with no page
    target to eyeball the whole document's layout cheaply (a vision model is
    billed on pixel area, so the cap gives a predictable per-page token budget;
    ~1000 stays legible). `--dpi 72` is a coarser alternative.

    `--markup all` shows tracked changes and comments as visible revision marks
    and balloons (the structured list is the `revisions` command).

    Requires the `snapshot` extra: `pip install "wordlive[snapshot]"`.
    """
    targets = [t is not None for t in (anchor_id, page, pages_range)]
    if sum(targets) > 1:
        raise click.UsageError("provide at most one of --anchor-id, --page, or --pages")
    if dpi < 1:
        raise click.UsageError("--dpi must be >= 1")
    if max_dim is not None and max_dim < 1:
        raise click.UsageError("--max-dim must be >= 1")
    if page is not None and page < 1:
        raise click.UsageError("--page must be >= 1")
    pages_arg: int | tuple[int, int] | None = None
    if page is not None:
        pages_arg = page
    elif pages_range is not None:
        pages_arg = _parse_pages_range(pages_range)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            if anchor_id is not None:
                anchor = doc.anchor_by_id(anchor_id)
                shots = doc.snapshot_anchor(anchor, out, dpi=dpi, max_dim=max_dim, markup=markup)
                selector: Any = anchor_id
            else:
                shots = doc.snapshot(out, pages=pages_arg, dpi=dpi, max_dim=max_dim, markup=markup)
                selector = pages_range or page or "all"
            images: list[dict[str, Any]] = []
            for s in shots:
                entry: dict[str, Any] = {"page": s.page, "bytes": len(s.png)}
                if s.path is not None:
                    entry["path"] = str(s.path)
                else:
                    entry["base64"] = base64.b64encode(s.png).decode("ascii")
                images.append(entry)
            payload: dict[str, Any] = {
                "ok": True,
                "selector": selector,
                "dpi": dpi,
                "count": len(images),
                "images": images,
            }
            if max_dim is not None:
                payload["max_dim"] = max_dim
            emit(
                payload,
                as_text=not ctx.obj["as_json"],
                text=_fmt_snapshot(images, dpi),
            )

    _run(ctx, go)
