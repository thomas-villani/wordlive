"""The `read` command group (structured reads)."""

from __future__ import annotations

import click

from ... import attach
from ..._anchors import Heading
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _WITHIN_HELP,
    _fmt_format_info,
    _fmt_nearest_heading,
)


@click.group(name="read")
def read() -> None:
    """Read structured values from the target document."""


@read.command(name="bookmark")
@click.argument("name", required=False)
@click.option(
    "--list",
    "list_all",
    is_flag=True,
    default=False,
    help="List every bookmark name instead of reading one.",
)
@click.option(
    "--include-hidden",
    "include_hidden",
    is_flag=True,
    default=False,
    help="With --list, also include Word's internal bookmarks (_Toc…, _Ref…).",
)
@click.pass_context
def read_bookmark(
    ctx: click.Context, name: str | None, list_all: bool, include_hidden: bool
) -> None:
    """Read the text of bookmark NAME, or list all bookmarks with --list."""
    if list_all == (name is not None):
        raise click.UsageError("provide either NAME or --list (not both)")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            if list_all:
                names = doc.bookmarks.list(include_hidden=include_hidden)
                emit(
                    names,
                    as_text=not ctx.obj["as_json"],
                    text="\n".join(names) if names else "(no bookmarks)",
                )
            else:
                assert name is not None  # guaranteed by the validation above
                text = doc.bookmarks[name].text
                emit({"text": text}, as_text=not ctx.obj["as_json"], text=text)

    _run(ctx, go)


@read.command(name="cc")
@click.argument("name")
@click.pass_context
def read_cc(ctx: click.Context, name: str) -> None:
    """Read the text of content control NAME."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            text = doc.content_controls[name].text
            emit({"text": text}, as_text=not ctx.obj["as_json"], text=text)

    _run(ctx, go)


@read.command(name="section")
@click.argument("heading", required=False)
@click.option(
    "--anchor-id",
    "anchor_id",
    default=None,
    help="Resolve heading by anchor id (e.g. 'heading:3') instead of by visible text.",
)
@click.pass_context
def read_section(ctx: click.Context, heading: str | None, anchor_id: str | None) -> None:
    """Read the body text under HEADING (up to the next same-or-higher heading)."""
    if (heading is None) == (anchor_id is None):
        raise click.UsageError("provide either HEADING or --anchor-id (not both)")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            if anchor_id is not None:
                h = doc.anchor_by_id(anchor_id)
                if not isinstance(h, Heading):
                    raise click.UsageError(f"--anchor-id must reference a heading, got {h.kind!r}")
            else:
                assert heading is not None  # guaranteed by the validation above
                h = doc.heading(heading)
            body = h.section_text()
            emit(
                {
                    "heading": h.text,
                    "anchor_id": h.anchor_id,
                    "level": h.level,
                    "text": body,
                },
                as_text=not ctx.obj["as_json"],
                text=body,
            )

    _run(ctx, go)


@read.command(name="format")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor to read effective formatting from."
)
@click.pass_context
def read_format(ctx: click.Context, anchor_id: str) -> None:
    """Effective paragraph + character formatting at an anchor.

    The read mirror of `format-paragraph` / `format-run`: each field carries its
    effective `value`, the applied style's `style` baseline, and whether a direct
    `override` sits on top — the substrate the linter's consistency rules use.
    `font.mixed` lists fields that vary across the range's runs. Pure read.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            info = doc.anchor_by_id(anchor_id).format_info()
            emit(info, as_text=not ctx.obj["as_json"], text=_fmt_format_info(info))

    _run(ctx, go)


@read.command(name="watermark")
@click.pass_context
def read_watermark(ctx: click.Context) -> None:
    """The text watermark stamped behind the pages, or none.

    The read mirror of `watermark set` / `watermark remove`: emits
    `{text, sections}` for the watermark Word draws in the header story, or
    `null` when there is none. Pure read.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            info = doc.watermark()
            payload = info.to_dict() if info is not None else None
            emit(
                payload,
                as_text=not ctx.obj["as_json"],
                text=(info.text if info is not None else "(no watermark)"),
            )

    _run(ctx, go)


@read.command(name="markdown")
@click.option("--within", "within", default=None, help=_WITHIN_HELP)
@click.pass_context
def read_markdown(ctx: click.Context, within: str | None) -> None:
    """Serialise the document (or an anchor's range) to clean Markdown.

    The read mirror of `insert-markdown`: headings, lists, **bold**/*italic*,
    GFM tables, `![alt](image:N)`, and `[text](url)`. Lossy by design.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            md = doc.to_markdown(within=within)
            emit({"markdown": md}, as_text=not ctx.obj["as_json"], text=md)

    _run(ctx, go)


@read.command(name="html")
@click.option("--within", "within", default=None, help=_WITHIN_HELP)
@click.pass_context
def read_html(ctx: click.Context, within: str | None) -> None:
    """Serialise the document (or an anchor's range) to an HTML fragment."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            html = doc.to_html(within=within)
            emit({"html": html}, as_text=not ctx.obj["as_json"], text=html)

    _run(ctx, go)


@read.command(name="digest")
@click.option(
    "--budget",
    "budget",
    type=int,
    default=6000,
    show_default=True,
    help="Approximate token budget (~4 chars/token) for the whole-document digest.",
)
@click.option(
    "--depth",
    "depth",
    type=int,
    default=None,
    help="Cap how deep a section keeps body (deeper sections collapse to a marker).",
)
@click.pass_context
def read_digest(ctx: click.Context, budget: int, depth: int | None) -> None:
    """A token-budgeted, anchor-addressable digest of the **whole** document.

    Headings verbatim (the navigation spine), tables as one-line shape stubs,
    body sampled to fit `--budget` — so a large document loads into context
    cheaply while every anchor stays addressable. Drill into any elided region
    with `read markdown --within …`.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            digest = doc.read(budget=budget, depth=depth)
            emit({"digest": digest}, as_text=not ctx.obj["as_json"], text=digest)

    _run(ctx, go)


@read.command(name="between")
@click.option("--start", "start", required=True, help="Start anchor id (e.g. 'heading:1').")
@click.option("--end", "end", required=True, help="End anchor id (e.g. 'heading:3').")
@click.option(
    "--inclusive",
    "inclusive",
    is_flag=True,
    default=False,
    help="Include both bounding paragraphs (default: only the content strictly between).",
)
@click.pass_context
def read_between(ctx: click.Context, start: str, end: str, inclusive: bool) -> None:
    """Read the content between two anchors (read-only).

    Default spans the gap strictly between START and END (e.g. the body between
    two headings, excluding the heading lines); --inclusive covers both.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            span = doc.between(start, end, inclusive=inclusive)
            body = span.text
            emit(
                {
                    "start": start,
                    "end": end,
                    "inclusive": inclusive,
                    "anchor_id": span.anchor_id,
                    "text": body,
                },
                as_text=not ctx.obj["as_json"],
                text=body,
            )

    _run(ctx, go)


@read.command(name="nearest-heading")
@click.option("--anchor-id", "anchor_id", required=True, help="Position to scan from (any anchor).")
@click.option(
    "--direction",
    "direction",
    type=click.Choice(["before", "after"]),
    default="before",
    help="before = enclosing/preceding heading; after = next heading.",
)
@click.pass_context
def read_nearest_heading(ctx: click.Context, anchor_id: str, direction: str) -> None:
    """Find the heading nearest to ANCHOR (read-only)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            row = doc.nearest_heading(anchor_id, direction=direction)
            emit(
                {"anchor_id": anchor_id, "direction": direction, "heading": row},
                as_text=not ctx.obj["as_json"],
                text=_fmt_nearest_heading(anchor_id, direction, row),
            )

    _run(ctx, go)


@read.command(name="text")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor whose text to read.")
@click.option(
    "--view",
    "view",
    type=click.Choice(["raw", "final", "original", "segments"]),
    default="raw",
    help=(
        "raw = text as-is (deleted+inserted runs both present); final = as if tracked "
        "changes accepted; original = as if rejected; segments = per-run breakdown."
    ),
)
@click.pass_context
def read_text(ctx: click.Context, anchor_id: str, view: str) -> None:
    """Read an anchor's text, optionally resolving tracked changes (--view)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            if view == "segments":
                segs = anchor.revision_segments()
                emit(
                    {"anchor_id": anchor_id, "segments": segs},
                    as_text=not ctx.obj["as_json"],
                    text="".join(
                        s["text"] if s["change"] is None else f"[{s['change']}:{s['text']}]"
                        for s in segs
                    ),
                )
                return
            text = {
                "raw": lambda: anchor.text,
                "final": lambda: anchor.text_final,
                "original": lambda: anchor.text_original,
            }[view]()
            emit(
                {"anchor_id": anchor_id, "view": view, "text": text},
                as_text=not ctx.obj["as_json"],
                text=text,
            )

    _run(ctx, go)
