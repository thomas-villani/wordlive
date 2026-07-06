"""CLI subcommands wired against the wordlive library."""

from __future__ import annotations

import base64
import json
import os
import sys
from pathlib import Path
from typing import Any

import click

from .. import attach
from .._anchors import Heading
from .._guide import bundled_skill as _bundled_skill
from .._guide import skill_body as _skill_body
from .._guide import skill_name as _skill_name
from .._ops import OP_REQUIRED_FIELDS as _OP_REQUIRED_FIELDS  # noqa: F401  (back-compat re-export)
from .._ops import apply_op as _apply_op  # noqa: F401  (back-compat re-export)
from .._ops import op_before as _op_before  # noqa: F401  (back-compat re-export)
from .._ops import pick_doc as _pick_doc
from .._ops import run_batch as _run_batch
from .._ops import validate_op as _validate_op  # noqa: F401  (back-compat re-export)
from ..exceptions import AmbiguousMatchError, OpError, WordNotRunningError
from .main import _run, emit

# ---------------------------------------------------------------------------
# Text formatters
# ---------------------------------------------------------------------------


def _fmt_status(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no documents open)"
    lines: list[str] = []
    width = max(len(str(r.get("name", ""))) for r in rows)
    for r in rows:
        marker = "*" if r.get("is_active") else " "
        lines.append(f"{marker} {str(r.get('name', '')):<{width}}  {r.get('path', '')}")
    return "\n".join(lines)


def _fmt_outline(items: list[dict[str, Any]]) -> str:
    if not items:
        return "(no headings)"
    lines: list[str] = []
    for it in items:
        level = int(it.get("level", 1))
        indent = "  " * max(level - 1, 0)
        lines.append(f"{indent}{it.get('text', '')}  [{it.get('anchor_id', '')}]")
    return "\n".join(lines)


def _fmt_paragraphs(items: list[dict[str, Any]]) -> str:
    if not items:
        return "(no paragraphs)"
    lines: list[str] = []
    for it in items:
        marker = f"H{it.get('level', 1)}" if it.get("is_heading") else "  "
        text = it.get("text", "")
        snippet = text if len(text) <= 60 else text[:57] + "…"
        lines.append(
            f"{marker} [{it.get('anchor_id', '')}] "
            f"{it.get('start', 0)}-{it.get('end', 0)}  {snippet}"
        )
    return "\n".join(lines)


def _fmt_find(matches: list[dict[str, Any]]) -> str:
    if not matches:
        return "(no matches)"
    return "\n".join(f"{m['start']:>6}–{m['end']:<6}  {m['text']!r}" for m in matches)


def _fmt_find_paragraphs(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no matches)"
    return "\n".join(f"{r['score']:.2f}  {r['anchor_id']:<10}  {r['text']!r}" for r in rows)


def _fmt_nearest_heading(anchor_id: str, direction: str, row: dict[str, Any] | None) -> str:
    if row is None:
        return f"(no heading {direction} {anchor_id})"
    return f"{row['anchor_id']}  (L{row['level']})  {row['text']!r}"


def _fmt_replace_summary(replacements: list[dict[str, Any]]) -> str:
    n = len(replacements)
    return f"replaced {n} occurrence{'s' if n != 1 else ''}"


def register(group: click.Group) -> None:
    group.add_command(status)
    group.add_command(outline)
    group.add_command(paragraphs_cmd)
    group.add_command(read)
    group.add_command(write)
    group.add_command(insert)
    group.add_command(insert_block_cmd)
    group.add_command(insert_section_cmd)
    group.add_command(insert_markdown_cmd)
    group.add_command(replace_section_cmd)
    group.add_command(delete_paragraph_cmd)
    group.add_command(insert_break_cmd)
    group.add_command(insert_field_cmd)
    group.add_command(update_fields_cmd)
    group.add_command(insert_footnote_cmd)
    group.add_command(insert_endnote_cmd)
    group.add_command(insert_toc_cmd)
    group.add_command(footnotes_cmd)
    group.add_command(endnotes_cmd)
    group.add_command(revisions_cmd)
    group.add_command(locate_cmd)
    group.add_command(stats_cmd)
    group.add_command(proofing_cmd)
    group.add_command(lint_cmd)
    group.add_command(regularize_cmd)
    group.add_command(checkpoint_cmd)
    group.add_command(diff_cmd)
    group.add_command(hyperlinks_cmd)
    group.add_command(set_hyperlink_cmd)
    group.add_command(fields_cmd)
    group.add_command(properties)
    group.add_command(variables)
    group.add_command(images_cmd)
    group.add_command(read_image_cmd)
    group.add_command(equations_cmd)
    group.add_command(insert_equation_cmd)
    group.add_command(charts_cmd)
    group.add_command(insert_chart_cmd)
    group.add_command(format_chart_cmd)
    group.add_command(format_axis_cmd)
    group.add_command(add_trendline_cmd)
    group.add_command(set_series_color_cmd)
    group.add_command(format_series_cmd)
    group.add_command(add_error_bars_cmd)
    group.add_command(shapes_cmd)
    group.add_command(set_shape_wrap_cmd)
    group.add_command(set_shape_crop_cmd)
    group.add_command(set_shape_position_cmd)
    group.add_command(set_shape_size_cmd)
    group.add_command(format_shape_cmd)
    group.add_command(set_shape_alt_text_cmd)
    group.add_command(set_shape_text_cmd)
    group.add_command(set_shape_rotation_cmd)
    group.add_command(set_shape_z_order_cmd)
    group.add_command(set_shape_text_frame_cmd)
    group.add_command(replace_shape_image_cmd)
    group.add_command(delete_shape_cmd)
    group.add_command(group_shapes_cmd)
    group.add_command(ungroup_shape_cmd)
    group.add_command(set_image_alt_text_cmd)
    group.add_command(set_image_size_cmd)
    group.add_command(set_image_crop_cmd)
    group.add_command(bookmark)
    group.add_command(pin_cmd)
    group.add_command(pin_outline_cmd)
    group.add_command(link_cmd)
    group.add_command(cross_ref_cmd)
    group.add_command(caption_cmd)
    group.add_command(create_content_control_cmd)
    group.add_command(set_cc_properties_cmd)
    group.add_command(set_cc_items_cmd)
    group.add_command(mark_index_entry_cmd)
    group.add_command(insert_index_cmd)
    group.add_command(table_of_figures_cmd)
    group.add_command(bibliography_style_cmd)
    group.add_command(add_source_cmd)
    group.add_command(insert_citation_cmd)
    group.add_command(insert_bibliography_cmd)
    group.add_command(mark_citation_cmd)
    group.add_command(table_of_authorities_cmd)
    group.add_command(theme_cmd)
    group.add_command(list_themes_cmd)
    group.add_command(apply_theme_cmd)
    group.add_command(set_theme_colors_cmd)
    group.add_command(set_theme_fonts_cmd)
    group.add_command(page_setup_cmd)
    group.add_command(prepend_cmd)
    group.add_command(append_cmd)
    group.add_command(insert_image_cmd)
    group.add_command(snapshot_cmd)
    group.add_command(save_cmd)
    group.add_command(save_as_cmd)
    group.add_command(export_pdf_cmd)
    group.add_command(cursor)
    group.add_command(find_cmd)
    group.add_command(find_paragraph_cmd)
    group.add_command(replace)
    group.add_command(go_to)
    group.add_command(style)
    group.add_command(format_paragraph_cmd)
    group.add_command(format_run_cmd)
    group.add_command(shading_cmd)
    group.add_command(borders_cmd)
    group.add_command(cell_valign_cmd)
    group.add_command(drop_cap_cmd)
    group.add_command(tab_stop_cmd)
    group.add_command(table)
    group.add_command(comment)
    group.add_command(track)
    group.add_command(revision)
    group.add_command(watermark_cmd)
    group.add_command(insert_text_box_cmd)
    group.add_command(list_cmd)
    group.add_command(sections_cmd)
    group.add_command(section)
    group.add_command(header)
    group.add_command(footer)
    group.add_command(exec_)
    group.add_command(llm_help_cmd)
    group.add_command(install_skill_cmd)
    group.add_command(install_mcp_cmd)


# ---------------------------------------------------------------------------
# status
# ---------------------------------------------------------------------------


@click.command(name="status")
@click.pass_context
def status(ctx: click.Context) -> None:
    """List open documents and which one is active."""

    def go() -> None:
        try:
            with attach() as word:
                rows = word.documents.list()
                emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_status(rows))
        except WordNotRunningError:
            emit([], as_text=not ctx.obj["as_json"], text=_fmt_status([]))
            raise

    _run(ctx, go)


# ---------------------------------------------------------------------------
# outline
# ---------------------------------------------------------------------------


@click.command(name="outline")
@click.option(
    "--all",
    "show_all",
    is_flag=True,
    default=False,
    help="List every paragraph (para:N), not just headings — same as `paragraphs`.",
)
@click.pass_context
def outline(ctx: click.Context, show_all: bool) -> None:
    """Print the heading outline (or every paragraph with --all)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            if show_all:
                items = doc.paragraphs.list()
                emit(items, as_text=not ctx.obj["as_json"], text=_fmt_paragraphs(items))
            else:
                items = doc.outline()
                emit(items, as_text=not ctx.obj["as_json"], text=_fmt_outline(items))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# paragraphs
# ---------------------------------------------------------------------------


@click.command(name="paragraphs")
@click.pass_context
def paragraphs_cmd(ctx: click.Context) -> None:
    """List every paragraph with its para:N anchor, level, offsets, and text.

    Includes headings, body paragraphs, and list items in document order — the
    everything view (`outline --all` is an alias). Use the emitted offsets to
    build a `range:START-END` target for a mid-paragraph insertion.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            items = doc.paragraphs.list()
            emit(items, as_text=not ctx.obj["as_json"], text=_fmt_paragraphs(items))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# pin / pin-outline — durable handles
# ---------------------------------------------------------------------------


@click.command(name="pin")
@click.argument("anchor_id")
@click.option(
    "--name",
    "name",
    default=None,
    help="A readable slug for the pin (lowercase words joined by hyphens, "
    "e.g. budget-intro); omit for a random code.",
)
@click.pass_context
def pin_cmd(ctx: click.Context, anchor_id: str, name: str | None) -> None:
    """Plant a durable handle on ANCHOR_ID and print its pin: id.

    A pin survives the inserts/deletes that renumber positional para:N / heading:N
    ids — resolve it later with `--anchor-id pin:CODE`. If the pinned content is
    deleted the handle vanishes (exit 2 on the next resolve).
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: pin {anchor_id}"):
                result = doc.pin(anchor_id, name=name)
            emit(
                result,
                as_text=not ctx.obj["as_json"],
                text=f"{result['pin']} -> {result['target']}",
            )

    _run(ctx, go)


@click.command(name="pin-outline")
@click.option(
    "--levels",
    "levels",
    nargs=2,
    type=int,
    default=None,
    help="Pin only headings in this inclusive level band, e.g. --levels 1 2.",
)
@click.pass_context
def pin_outline_cmd(ctx: click.Context, levels: tuple[int, int] | None) -> None:
    """Pin every heading at once; print the {heading:N -> pin:CODE} map.

    Idempotent — a heading already pinned reuses its handle. A durable navigation
    scaffold to set up before a batch of structural edits.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: pin-outline"):
                pins = doc.pin_outline(levels=levels)
            text = "\n".join(f"{k} -> {v}" for k, v in pins.items()) or "(no headings)"
            emit(pins, as_text=not ctx.obj["as_json"], text=text)

    _run(ctx, go)


# ---------------------------------------------------------------------------
# read bookmark|cc|section NAME
# ---------------------------------------------------------------------------


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


def _fmt_format_info(info: dict[str, Any]) -> str:
    lines = [f"{info['anchor_id']}  style={info['style']!r}"]
    for group in ("paragraph", "font"):
        for field, cell in info[group].items():
            if field == "mixed":
                continue
            mark = " *override*" if cell.get("override") else ""
            lines.append(f"  {field}: {cell['value']} (style {cell['style']}){mark}")
    if info["font"].get("mixed"):
        lines.append(f"  mixed runs: {', '.join(info['font']['mixed'])}")
    return "\n".join(lines)


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


_WITHIN_HELP = (
    "Scope to an anchor's range (e.g. 'heading:3', 'range:120-540'); default is the whole document."
)


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


# ---------------------------------------------------------------------------
# write bookmark|cc NAME --text "..."
# ---------------------------------------------------------------------------


@click.group(name="write")
def write() -> None:
    """Write structured values into the target document."""


@write.command(name="bookmark")
@click.argument("name")
@click.option("--text", "text", default=None, help="New text for an existing bookmark.")
@click.option(
    "--create",
    "create",
    is_flag=True,
    default=False,
    help="Create the bookmark over --anchor-id (instead of writing text).",
)
@click.option(
    "--anchor-id",
    "anchor_id",
    default=None,
    help="With --create: the anchor whose range the new bookmark covers "
    "(e.g. heading:2, range:120-140).",
)
@click.pass_context
def write_bookmark(
    ctx: click.Context, name: str, text: str | None, create: bool, anchor_id: str | None
) -> None:
    """Create a bookmark, or set an existing one's text (atomic-undo).

    Two modes:

    \b
      write bookmark NAME --text "…"                set an existing bookmark's text
      write bookmark NAME --create --anchor-id ID   create NAME over an anchor's range

    Creating a bookmark is the prerequisite for internal links
    (`link --bookmark NAME`) and cross-references (`cross-ref --target
    bookmark:NAME`). NAME must start with a letter and contain only letters,
    digits, and underscores.
    """
    if create:
        if text is not None:
            raise click.UsageError("--create and --text are mutually exclusive")
        if anchor_id is None:
            raise click.UsageError("--create requires --anchor-id")
    else:
        if anchor_id is not None:
            raise click.UsageError("--anchor-id is only valid with --create")
        if text is None:
            raise click.UsageError(
                "provide --text (write an existing bookmark) or "
                "--create --anchor-id ID (create a new one)"
            )

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            if create:
                assert anchor_id is not None  # guaranteed by the validation above
                with doc.edit(f"CLI: add bookmark {name}"):
                    doc.bookmarks.add(name, anchor_id)
                emit(
                    {"ok": True, "bookmark": name, "anchor_id": anchor_id, "created": True},
                    as_text=not ctx.obj["as_json"],
                    text=f"added bookmark:{name} over {anchor_id}",
                )
            else:
                assert text is not None  # guaranteed by the validation above
                bm = doc.bookmarks[name]
                with doc.edit(f"CLI: write bookmark {name}"):
                    bm.set_text(text)
                emit(
                    {"ok": True, "anchor": {"kind": bm.kind, "name": name}},
                    as_text=not ctx.obj["as_json"],
                    text=f"wrote bookmark:{name}",
                )

    _run(ctx, go)


@write.command(name="cc")
@click.argument("name")
@click.option("--text", "text", required=True, help="New text for the content control.")
@click.pass_context
def write_cc(ctx: click.Context, name: str, text: str) -> None:
    """Set the text of content control NAME (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            cc = doc.content_controls[name]
            with doc.edit(f"CLI: write cc {name}"):
                cc.set_text(text)
            emit(
                {"ok": True, "anchor": {"kind": cc.kind, "name": name}},
                as_text=not ctx.obj["as_json"],
                text=f"wrote cc:{name}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# insert --anchor-id ID --text "..." [--before|--after] [--style "..."]
# ---------------------------------------------------------------------------


@click.command(name="insert")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert a new paragraph relative to (e.g. heading:1, para:3).",
)
@click.option(
    "--text",
    "text",
    default=None,
    help="Paragraph text to insert (literal — no markup). For inline formatting "
    "use --runs, or `insert-block` for a styled multi-paragraph run.",
)
@click.option(
    "--runs",
    "runs",
    default=None,
    help='JSON array of inline runs (e.g. \'[{"text":"Fast","bold":true},'
    '{"text":" — quick"}]\'), or \'-\' to read it from stdin. Each run is '
    "{text, bold?, italic?, underline?, style?}. Mutually exclusive with --text.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the new paragraph before the anchor instead of after it.",
)
@click.option(
    "--style", "style", default=None, help="Optional Word style name for the new paragraph."
)
@click.pass_context
def insert(
    ctx: click.Context,
    anchor_id: str,
    text: str | None,
    runs: str | None,
    before: bool,
    style: str | None,
) -> None:
    """Insert a new paragraph before/after any anchor (atomic-undo).

    Addresses anchors the same way every other command does — `--anchor-id`
    (headings, paragraphs, bookmarks, cells, ranges). Pass either `--text`
    (literal) or `--runs` (inline-formatted spans); for a contiguous run of
    several styled paragraphs in one shot, use `insert-block` instead. To insert
    text *inside* a paragraph at an offset, target a collapsed range:
    `replace --anchor-id range:120-120 --text "…"` (offsets come from
    `paragraphs` / `find`).
    """
    if (text is None) == (runs is None):
        raise click.UsageError("provide exactly one of --text or --runs")
    parsed_runs: list[Any] | None = None
    if runs is not None:
        raw = click.get_text_stream("stdin").read() if runs == "-" else runs
        try:
            parsed_runs = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--runs must be a JSON array: {e}") from e
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert {where} {anchor_id}"):
                if parsed_runs is not None:
                    anchor.insert_block([{"runs": parsed_runs, "style": style}], where=where)
                else:
                    assert text is not None  # xor-validated above; narrows for mypy
                    if before:
                        anchor.insert_paragraph_before(text, style=style)
                    else:
                        anchor.insert_paragraph_after(text, style=style)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "where": where,
                    "style": style,
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {where} {anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# insert-block --anchor-id ID --items JSON
# ---------------------------------------------------------------------------


@click.command(name="insert-block")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the block of paragraphs relative to (heading:/para:/end/…).",
)
@click.option(
    "--items",
    "items",
    required=True,
    help="JSON array of paragraphs, or '-' to read it from stdin. Each item is a "
    'string ("plain text") or an object {text|runs, style?}. `text` carries '
    "tiny inline markdown (**bold**, *italic*, ***both***; escape with \\*); "
    "`runs` is [{text, bold?, italic?, underline?, style?}].",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the block before the anchor instead of after it.",
)
@click.pass_context
def insert_block_cmd(ctx: click.Context, anchor_id: str, items: str, before: bool) -> None:
    """Insert a contiguous run of styled paragraphs at an anchor (atomic-undo).

    The multi-paragraph insert — drop a whole styled section (a feature list,
    a heading plus its body) in ONE op, in natural reading order, instead of a
    reverse-order storm of `insert` calls. Each item is one paragraph; `text`
    supports inline markdown and `runs` the structured form, so the "**Bold
    lead** — rest" bullet is a single op with no second formatting pass.

    Reports the spanning `range:START-END` of the inserted block, so a follow-up
    op can target the whole run — e.g. `list apply --anchor-id range:… --type
    bulleted` to bullet the section you just inserted.
    """
    raw = click.get_text_stream("stdin").read() if items == "-" else items
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--items must be a JSON array: {e}") from e
    if not isinstance(parsed, list):
        raise click.UsageError("--items must be a JSON array of paragraphs")
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert block {where} {anchor_id}"):
                rng = anchor.insert_block(parsed, where=where)
            emit(
                {
                    "ok": True,
                    "anchor_id": rng.anchor_id,
                    "paragraphs": len(parsed),
                    "where": where,
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {len(parsed)} paragraph(s) {where} {anchor_id} → {rng.anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# insert-section --anchor-id ID --heading TEXT --body JSON [--level N]
# ---------------------------------------------------------------------------


@click.command(name="insert-section")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the section relative to (heading:/para:/end/…).",
)
@click.option(
    "--heading", "heading", required=True, help="Heading text (inline **bold**/*italic* ok)."
)
@click.option(
    "--body",
    "body",
    required=True,
    help="JSON array of body paragraphs (insert-block items shape), or '-' for stdin.",
)
@click.option("--level", "level", type=int, default=1, show_default=True, help="Heading level 1–9.")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the section before the anchor instead of after it.",
)
@click.pass_context
def insert_section_cmd(
    ctx: click.Context, anchor_id: str, heading: str, body: str, level: int, before: bool
) -> None:
    """Insert a heading plus its body in one atomic op.

    The opinionated common case over `insert-block`: a `Heading {level}`
    paragraph followed by the body paragraphs, in reading order. `--body` is the
    same items shape `insert-block` takes. Reports the section's spanning
    `range:START-END`.
    """
    raw = click.get_text_stream("stdin").read() if body == "-" else body
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--body must be a JSON array: {e}") from e
    if not isinstance(parsed, list):
        raise click.UsageError("--body must be a JSON array of paragraphs")
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert section {where} {anchor_id}"):
                rng = anchor.insert_section(heading, parsed, level=level, where=where)
            emit(
                {"ok": True, "anchor_id": rng.anchor_id, "where": where},
                as_text=not ctx.obj["as_json"],
                text=f"inserted section {where} {anchor_id} → {rng.anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# insert-markdown --anchor-id ID --markdown TEXT
# ---------------------------------------------------------------------------


@click.command(name="insert-markdown")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the markdown relative to (heading:/para:/end/…).",
)
@click.option(
    "--markdown",
    "markdown",
    required=True,
    help="Constrained-Markdown text, or '-' to read it from stdin. Subset: "
    "#/##/### headings, -/* bullets, 1. numbers, blank-line paragraphs, "
    "inline **bold**/*italic*. No code fences/nested lists/tables.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the markdown before the anchor instead of after it.",
)
@click.pass_context
def insert_markdown_cmd(ctx: click.Context, anchor_id: str, markdown: str, before: bool) -> None:
    """Insert a constrained-Markdown block as real Word structure (atomic-undo).

    Maps a tiny block dialect to paragraphs/headings/lists — a documented subset,
    not CommonMark. Path-bearing or multi-line input is easiest via `--markdown -`
    (stdin). Reports the spanning `range:START-END` of everything inserted.
    """
    md = click.get_text_stream("stdin").read() if markdown == "-" else markdown
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert markdown {where} {anchor_id}"):
                rng = anchor.insert_markdown(md, where=where)
            emit(
                {"ok": True, "anchor_id": rng.anchor_id, "where": where},
                as_text=not ctx.obj["as_json"],
                text=f"inserted markdown {where} {anchor_id} → {rng.anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# replace-section --anchor-id heading:N (--body JSON | --markdown TEXT)
# ---------------------------------------------------------------------------


@click.command(name="replace-section")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="The heading whose body to replace (heading:N). The heading itself is kept.",
)
@click.option(
    "--body", "body", default=None, help="JSON array of new body paragraphs, or '-' for stdin."
)
@click.option(
    "--markdown",
    "markdown",
    default=None,
    help="New body as constrained Markdown, or '-' for stdin.",
)
@click.pass_context
def replace_section_cmd(
    ctx: click.Context, anchor_id: str, body: str | None, markdown: str | None
) -> None:
    """Rewrite a heading's body, preserving the heading paragraph.

    Clears the span under `--anchor-id` (up to the next same-or-higher heading)
    and inserts the new body after the heading. Give exactly one of `--body`
    (insert-block items) or `--markdown` (constrained Markdown).
    """
    if (body is None) == (markdown is None):
        raise click.UsageError("give exactly one of --body or --markdown")
    if markdown is not None:
        new_body = click.get_text_stream("stdin").read() if markdown == "-" else markdown
    else:
        assert body is not None
        raw = click.get_text_stream("stdin").read() if body == "-" else body
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--body must be a JSON array: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--body must be a JSON array of paragraphs")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            if not hasattr(anchor, "replace_section_body"):
                raise click.UsageError(
                    f"replace-section needs a heading anchor; {anchor_id} is a {anchor.kind}"
                )
            with doc.edit(f"CLI: replace section {anchor_id}"):
                if markdown is not None:
                    rng = anchor.replace_section_body(new_body, markdown=True)
                else:
                    rng = anchor.replace_section_body(parsed)
            emit(
                {"ok": True, "anchor_id": rng.anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"replaced section body of {anchor_id} → {rng.anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# delete-paragraph --anchor-id ID
# ---------------------------------------------------------------------------


@click.command(name="delete-paragraph")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Paragraph anchor to delete (e.g. para:1, heading:2).",
)
@click.pass_context
def delete_paragraph_cmd(ctx: click.Context, anchor_id: str) -> None:
    """Delete the paragraph(s) at an anchor — text and the trailing mark (atomic-undo).

    Removes the whole paragraph so the surrounding text closes up (no empty line
    left, unlike `replace --text ""`). Useful for a stray leading empty paragraph.
    Deleting the document's last paragraph clears it but keeps Word's mandatory
    final mark.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: delete paragraph {anchor_id}"):
                doc.delete_paragraph(anchor_id)
            emit(
                {"ok": True, "anchor_id": anchor_id, "deleted": True},
                as_text=not ctx.obj["as_json"],
                text=f"deleted paragraph {anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# insert-break --anchor-id ID [--kind page|column|...] [--before|--after]
# ---------------------------------------------------------------------------


@click.command(name="insert-break")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the break relative to (e.g. heading:1, para:3, end).",
)
@click.option(
    "--kind",
    "kind",
    type=click.Choice(["page", "column", "section_next", "section_continuous"]),
    default="page",
    show_default=True,
    help="Break kind. Page is the common case; section breaks pair with `section`.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the break before the anchor instead of after it.",
)
@click.pass_context
def insert_break_cmd(ctx: click.Context, anchor_id: str, kind: str, before: bool) -> None:
    """Insert a page / column / section break at an anchor (atomic-undo).

    The explicit one-off break, the clean alternative to a literal form-feed
    paragraph. To make a *style* (e.g. every Heading 1) open a new page without
    a stray break character, prefer
    `format-paragraph --anchor-id ID --page-break-before` instead.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert {kind} break {where} {anchor_id}"):
                anchor.insert_break(kind, where=where)
            emit(
                {"ok": True, "anchor_id": anchor_id, "kind": kind, "where": where},
                as_text=not ctx.obj["as_json"],
                text=f"inserted {kind} break {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-field")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the field at (e.g. footer:1:primary, end).",
)
@click.option(
    "--kind",
    "kind",
    type=click.Choice(["page", "numpages", "date", "time", "filename", "author", "title", "field"]),
    required=True,
    help="Field kind. Use 'field' with --text for a raw field code.",
)
@click.option(
    "--text",
    "text",
    default=None,
    help="Raw field code, required when --kind field (e.g. 'REF myBookmark').",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the field before the anchor instead of after it.",
)
@click.pass_context
def insert_field_cmd(
    ctx: click.Context, anchor_id: str, kind: str, text: str | None, before: bool
) -> None:
    """Insert a self-updating field (page number, date, …) at an anchor (atomic-undo).

    Page numbers belong in a footer/header: `insert-field --anchor-id
    footer:1:primary --kind page`. Refresh stale fields with `update-fields`.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert {kind} field {where} {anchor_id}"):
                anchor.insert_field(kind, text=text, where=where)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "applied": {"kind": kind, "text": text, "where": where},
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {kind} field {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="update-fields")
@click.pass_context
def update_fields_cmd(ctx: click.Context) -> None:
    """Refresh the document's fields — recompute page numbers, refs, dates (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: update fields"):
                doc.update_fields()
            emit(
                {"ok": True, "updated": True},
                as_text=not ctx.obj["as_json"],
                text="updated fields",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# insert-footnote / insert-endnote / insert-toc + footnotes/endnotes lists
# ---------------------------------------------------------------------------


def _fmt_notes(rows: list[dict[str, Any]], scheme: str) -> str:
    if not rows:
        return f"(no {scheme}s)"
    lines = []
    for r in rows:
        where = f" @ {r['para']}" if r.get("para") else ""
        lines.append(f"{scheme}:{r['index']}{where}  {r.get('text', '')}")
    return "\n".join(lines)


@click.command(name="insert-footnote")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor the footnote's reference mark attaches to (e.g. range:120-140, para:3).",
)
@click.option("--text", "text", required=True, help="The footnote body text.")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Place the reference mark before the anchor instead of after it.",
)
@click.pass_context
def insert_footnote_cmd(ctx: click.Context, anchor_id: str, text: str, before: bool) -> None:
    """Insert a footnote at an anchor (atomic-undo). Reports the new footnote:N."""
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert footnote {where} {anchor_id}"):
                note = anchor.insert_footnote(text, where=where)
                note_id = note.anchor_id
                index = note.index
            emit(
                {"ok": True, "anchor_id": anchor_id, "footnote": index, "note_id": note_id},
                as_text=not ctx.obj["as_json"],
                text=f"inserted {note_id} {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-endnote")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor the endnote's reference mark attaches to (e.g. range:120-140, para:3).",
)
@click.option("--text", "text", required=True, help="The endnote body text.")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Place the reference mark before the anchor instead of after it.",
)
@click.pass_context
def insert_endnote_cmd(ctx: click.Context, anchor_id: str, text: str, before: bool) -> None:
    """Insert an endnote at an anchor (atomic-undo). Reports the new endnote:N."""
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert endnote {where} {anchor_id}"):
                note = anchor.insert_endnote(text, where=where)
                note_id = note.anchor_id
                index = note.index
            emit(
                {"ok": True, "anchor_id": anchor_id, "endnote": index, "note_id": note_id},
                as_text=not ctx.obj["as_json"],
                text=f"inserted {note_id} {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-toc")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="start",
    show_default=True,
    help="Where to insert the TOC (default: the document start).",
)
@click.option(
    "--levels",
    "levels",
    default="1-3",
    show_default=True,
    help="Heading levels to include, as 'upper-lower' (e.g. 1-3).",
)
@click.option(
    "--heading-styles/--no-heading-styles",
    "heading_styles",
    default=True,
    show_default=True,
    help="Source entries from the built-in Heading styles.",
)
@click.option(
    "--hyperlinks/--no-hyperlinks",
    "hyperlinks",
    default=True,
    show_default=True,
    help="Make each entry a clickable link (and a real link in exported PDFs).",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the TOC before the anchor instead of after it.",
)
@click.pass_context
def insert_toc_cmd(
    ctx: click.Context,
    anchor_id: str,
    levels: str,
    heading_styles: bool,
    hyperlinks: bool,
    before: bool,
) -> None:
    """Insert a table of contents (atomic-undo).

    Page numbers populate after repagination — run `update-fields` (or take a
    `snapshot`) before reading them.
    """
    upper_str, sep, lower_str = levels.partition("-")
    if not sep:
        raise click.UsageError("--levels must be 'upper-lower', e.g. 1-3")
    try:
        level_pair = (int(upper_str), int(lower_str))
    except ValueError as e:
        raise click.UsageError("--levels must be two integers, e.g. 1-3") from e
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert TOC {where} {anchor_id}"):
                anchor.insert_toc(
                    levels=level_pair,
                    use_heading_styles=heading_styles,
                    hyperlinks=hyperlinks,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "levels": list(level_pair),
                        "use_heading_styles": heading_styles,
                        "hyperlinks": hyperlinks,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted TOC {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="footnotes")
@click.pass_context
def footnotes_cmd(ctx: click.Context) -> None:
    """List the document's footnotes with their footnote:N id, text, and para:N."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.footnotes.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_notes(rows, "footnote"))

    _run(ctx, go)


@click.command(name="endnotes")
@click.pass_context
def endnotes_cmd(ctx: click.Context) -> None:
    """List the document's endnotes with their endnote:N id, text, and para:N."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.endnotes.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_notes(rows, "endnote"))

    _run(ctx, go)


def _fmt_revisions(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "no tracked changes"
    lines = []
    for r in rows:
        who = r.get("author") or "?"
        text = (r.get("text") or "").replace("\r", " ").replace("\n", " ")
        if len(text) > 60:
            text = text[:57] + "…"
        lines.append(f"{r['index']}. [{r.get('type')}] {who}: {text!r} ({r.get('anchor_id')})")
    return "\n".join(lines)


@click.command(name="revisions")
@click.pass_context
def revisions_cmd(ctx: click.Context) -> None:
    """List the document's tracked changes (type, author, text, and range).

    The structured counterpart to `snapshot --markup all`: each revision is an
    insert / delete / format change with its author, the affected text, and a
    `range:START-END` id. Reading is non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.revisions.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_revisions(rows))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# locate / stats  (non-visual layout introspection — reads, no snapshot)
# ---------------------------------------------------------------------------


def _fmt_location(anchor_id: str, loc: dict[str, Any]) -> str:
    span = (
        f"page {loc['page']}"
        if loc["page"] == loc["end_page"]
        else f"pages {loc['page']}–{loc['end_page']}"
    )
    where = f"{anchor_id}: {span}, line {loc['line']}, col {loc['column']}"
    return where + (" (in table)" if loc["in_table"] else "")


@click.command(name="locate")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to locate (e.g. 'para:5', 'heading:3', 'table:1', 'image:2').",
)
@click.pass_context
def locate_cmd(ctx: click.Context, anchor_id: str) -> None:
    """Report where an anchor sits in the laid-out document.

    Returns the anchor's page span (`page`/`end_page`), `line`, `column`, and
    `in_table` — a non-visual layout read that answers "what page is this on"
    without a `snapshot` vision pass. Page numbers are print-layout truth, so
    the document is repaginated first; the user's selection, scroll, and view
    are left untouched. Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            loc = doc.anchor_by_id(anchor_id).location()
            emit(
                {"anchor_id": anchor_id, **loc},
                as_text=not ctx.obj["as_json"],
                text=_fmt_location(anchor_id, loc),
            )

    _run(ctx, go)


def _fmt_stats(s: dict[str, Any]) -> str:
    order = [
        "pages",
        "words",
        "characters",
        "paragraphs",
        "lines",
        "sections",
        "headings",
        "tables",
        "images",
        "comments",
        "revisions",
    ]
    parts = [f"{k}: {s[k]}" for k in order if k in s]
    parts.append("saved" if s.get("saved") else "unsaved")
    return "  ".join(parts)


@click.command(name="stats")
@click.pass_context
def stats_cmd(ctx: click.Context) -> None:
    """Summarise the document in one read — counts plus structure.

    Returns `{pages, words, characters, paragraphs, lines, sections, headings,
    tables, images, comments, revisions, saved}`: the "what am I looking at
    before I act" read. The page/line counts are print-layout truth, so the
    document is repaginated first (selection/scroll/view untouched). Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            s = doc.stats()
            emit(s, as_text=not ctx.obj["as_json"], text=_fmt_stats(s))

    _run(ctx, go)


def _fmt_proofing(data: dict[str, Any]) -> str:
    sp, gr = data.get("spelling", {}), data.get("grammar", {})
    read = data.get("readability", {})
    lines = [
        f"spelling errors: {sp.get('count')}",
        f"grammar errors: {gr.get('count')}",
    ]
    for key in (
        "flesch_reading_ease",
        "flesch_kincaid_grade_level",
        "passive_sentences",
        "words_per_sentence",
    ):
        if key in read:
            lines.append(f"{key}: {read[key]}")
    return "\n".join(lines)


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


# ---------------------------------------------------------------------------
# lint / regularize  (audit + autofix publishing-quality defects)
# ---------------------------------------------------------------------------


def _rules_selector(rule: tuple[str, ...], exclude: tuple[str, ...]) -> Any:
    """Build the `rules=` selector from the repeatable --rule / --exclude flags."""
    if rule and exclude:
        raise click.UsageError("pass either --rule or --exclude, not both")
    if rule:
        return list(rule)
    if exclude:
        return {"exclude": list(exclude)}
    return None


def _fmt_lint(findings: list[dict[str, Any]]) -> str:
    if not findings:
        return "(no findings)"
    lines = []
    for f in findings:
        fix = " [fixable]" if f.get("fixable") else ""
        lines.append(f"[{f['severity']}] {f['rule']} ({f['anchor_id']}): {f['message']}{fix}")
    return "\n".join(lines)


def _fmt_regularize(report: dict[str, Any]) -> str:
    applied, skipped = report.get("applied", []), report.get("skipped", [])
    deferred = report.get("deferred", [])
    verb = "would fix" if report.get("dry_run") else "fixed"
    summary = f"{verb} {len(applied)}; skipped {len(skipped)} (report-only / not fixable)"
    if deferred:
        summary += f"; deferred {len(deferred)} content fix(es) (pass --allow-content)"
    lines = [summary]
    for f in applied:
        lines.append(f"  {verb}: {f['rule']} ({f['anchor_id']})")
    for f in deferred:
        lines.append(f"  deferred: {f['rule']} ({f['anchor_id']})")
    return "\n".join(lines)


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


# ---------------------------------------------------------------------------
# hyperlinks / fields  (read-only discovery mirrors of link_to / insert_field)
# ---------------------------------------------------------------------------


def _fmt_hyperlinks(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no hyperlinks)"
    lines: list[str] = []
    for r in rows:
        dest = r.get("address") or (f"#{r['sub_address']}" if r.get("sub_address") else "?")
        text = r.get("text") or ""
        para = r.get("para") or ""
        lines.append(f"[{r['anchor_id']}] {text!r} -> {dest}  {para}".rstrip())
    return "\n".join(lines)


@click.command(name="hyperlinks")
@click.pass_context
def hyperlinks_cmd(ctx: click.Context) -> None:
    """List the document's hyperlinks (text, destination, range:START-END id).

    The read mirror of `link`: each link's visible text, external `address` or
    internal `sub_address` bookmark, screen tip, and the range/para it sits in.
    Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.hyperlinks.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_hyperlinks(rows))

    _run(ctx, go)


@click.command(name="set-hyperlink")
@click.option(
    "--index", "index", type=int, required=True, help="1-based hyperlink index (see `hyperlinks`)."
)
@click.option("--address", "address", default=None, help="External URL to retarget to.")
@click.option(
    "--sub-address", "sub_address", default=None, help='In-document bookmark ("" clears it).'
)
@click.option("--text", "text", default=None, help="Visible link text.")
@click.option("--screen-tip", "screen_tip", default=None, help='Hover tooltip ("" clears it).')
@click.pass_context
def set_hyperlink_cmd(
    ctx: click.Context,
    index: int,
    address: str | None,
    sub_address: str | None,
    text: str | None,
    screen_tip: str | None,
) -> None:
    """Retarget or relabel an existing hyperlink in place (atomic-undo).

    Address the link by its 1-based --index (from `hyperlinks`). Pass at least
    one field; omitting one leaves it untouched. These retarget a link, they
    don't unlink it: --sub-address / --screen-tip clear with "", but --address /
    --text cannot be emptied (Word keeps a link pointing somewhere).
    """
    raw: dict[str, Any] = {
        "address": address,
        "sub_address": sub_address,
        "text": text,
        "screen_tip": screen_tip,
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --address/--sub-address/--text/--screen-tip")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: set hyperlink {index}"):
                doc.hyperlinks[index].update(**kwargs)
            emit(
                {"ok": True, "index": index, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"updated hyperlink {index}: {kwargs}",
            )

    _run(ctx, go)


def _fmt_fields(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no fields)"
    lines: list[str] = []
    for r in rows:
        result = r.get("result") or ""
        suffix = f" = {result!r}" if result else ""
        lines.append(f"[{r['anchor_id']}] {r['kind']}: {r['code']}{suffix}")
    return "\n".join(lines)


@click.command(name="fields")
@click.pass_context
def fields_cmd(ctx: click.Context) -> None:
    """List the document's fields (kind, code, rendered result, range:START-END id).

    The read mirror of `insert-field`: each field's `kind` (the code's leading
    keyword — PAGE, REF, TOC, …), raw `code`, and last-rendered `result`. Run
    `update-fields` first to refresh stale results. Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.fields.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_fields(rows))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# properties / variables  (document metadata + invisible named storage)
# ---------------------------------------------------------------------------


def _fmt_properties(data: dict[str, Any]) -> str:
    lines: list[str] = []
    for bag in ("builtin", "custom"):
        items = data.get(bag, {})
        if items:
            lines.append(f"[{bag}]")
            lines.extend(f"  {k}: {v}" for k, v in items.items())
    return "\n".join(lines) if lines else "(no properties)"


@click.group(name="properties")
def properties() -> None:
    """Read and edit document properties (metadata): built-in + custom."""


@properties.command(name="list")
@click.pass_context
def properties_list(ctx: click.Context) -> None:
    """List the document's built-in and custom properties."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data = doc.properties.read()
            emit(data, as_text=not ctx.obj["as_json"], text=_fmt_properties(data))

    _run(ctx, go)


@properties.command(name="set")
@click.option("--name", "name", required=True, help="Property name (e.g. 'Title', 'Author').")
@click.option("--value", "value", required=True, help="New value.")
@click.option(
    "--custom/--builtin",
    "custom",
    default=False,
    show_default=True,
    help="Set a custom property (created if absent) instead of a built-in one.",
)
@click.pass_context
def properties_set(ctx: click.Context, name: str, value: str, custom: bool) -> None:
    """Set a built-in (default) or custom document property (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: set property {name}"):
                doc.properties.set(name, value, custom=custom)
            label = "custom property" if custom else "property"
            emit(
                {"ok": True, "name": name, "value": value, "custom": custom},
                as_text=not ctx.obj["as_json"],
                text=f"set {label} {name!r} = {value!r}",
            )

    _run(ctx, go)


@properties.command(name="delete")
@click.option("--name", "name", required=True, help="Custom property name to delete.")
@click.pass_context
def properties_delete(ctx: click.Context, name: str) -> None:
    """Delete a custom document property (atomic-undo). Built-ins can't be removed."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: delete property {name}"):
                doc.properties.delete(name)
            emit(
                {"ok": True, "name": name},
                as_text=not ctx.obj["as_json"],
                text=f"deleted custom property {name!r}",
            )

    _run(ctx, go)


def _fmt_variables(data: dict[str, str]) -> str:
    if not data:
        return "(no variables)"
    return "\n".join(f"{k}: {v}" for k, v in data.items())


@click.group(name="variables")
def variables() -> None:
    """Read and edit document variables (invisible named string storage)."""


@variables.command(name="list")
@click.pass_context
def variables_list(ctx: click.Context) -> None:
    """List the document's variables as name: value pairs."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data = doc.variables.list()
            emit(data, as_text=not ctx.obj["as_json"], text=_fmt_variables(data))

    _run(ctx, go)


@variables.command(name="set")
@click.option("--name", "name", required=True, help="Variable name.")
@click.option("--value", "value", required=True, help="Value (stored as a string).")
@click.pass_context
def variables_set(ctx: click.Context, name: str, value: str) -> None:
    """Create or update a document variable (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: set variable {name}"):
                doc.variables.set(name, value)
            emit(
                {"ok": True, "name": name, "value": value},
                as_text=not ctx.obj["as_json"],
                text=f"set variable {name!r} = {value!r}",
            )

    _run(ctx, go)


@variables.command(name="delete")
@click.option("--name", "name", required=True, help="Variable name to delete.")
@click.pass_context
def variables_delete(ctx: click.Context, name: str) -> None:
    """Delete a document variable (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: delete variable {name}"):
                doc.variables.delete(name)
            emit(
                {"ok": True, "name": name},
                as_text=not ctx.obj["as_json"],
                text=f"deleted variable {name!r}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# images / read-image  (image extraction — read embedded pictures out)
# ---------------------------------------------------------------------------


def _fmt_images(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no images)"
    lines: list[str] = []
    for r in rows:
        w, h = r.get("width"), r.get("height")
        size = f"  {w:.0f}×{h:.0f}pt" if w and h else ""
        mime = r.get("mime") or "?"
        para = r.get("para") or ""
        alt = r.get("alt_text") or ""
        crop = "  cropped" if r.get("crop") else ""
        suffix = f"  {alt!r}" if alt else ""
        lines.append(f"[{r['anchor_id']}] {mime}{size}{crop}  {para}{suffix}")
    return "\n".join(lines)


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


# ---------------------------------------------------------------------------
# bookmark add / link / cross-ref / caption  (anchoring & linking)
# ---------------------------------------------------------------------------


@click.group(name="bookmark", hidden=True)
def bookmark() -> None:
    """Deprecated: use `write bookmark NAME --create --anchor-id ID`.

    Kept as a hidden alias for one release.
    """


@bookmark.command(name="add")
@click.argument("name")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor whose range the bookmark covers (e.g. heading:2, range:120-140).",
)
@click.pass_context
def bookmark_add(ctx: click.Context, name: str, anchor_id: str) -> None:
    """Deprecated alias for `write bookmark NAME --create --anchor-id ID`.

    Create a bookmark NAME over an anchor's range (atomic-undo). The prerequisite
    for internal links (`link --bookmark NAME`) and cross-references
    (`cross-ref --target bookmark:NAME`). NAME must start with a letter and
    contain only letters, digits, and underscores.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: add bookmark {name}"):
                doc.bookmarks.add(name, anchor_id)
            emit(
                {"ok": True, "bookmark": name, "anchor_id": anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"added bookmark:{name} over {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="link")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to turn into a hyperlink.")
@click.option("--url", "url", default=None, help="External link target (URL, mailto:, file path).")
@click.option(
    "--bookmark", "bookmark_target", default=None, help="Internal target: a bookmark name."
)
@click.option(
    "--text", "text", default=None, help="Visible link text (replaces the range content)."
)
@click.option("--screen-tip", "screen_tip", default=None, help="Hover tooltip.")
@click.pass_context
def link_cmd(
    ctx: click.Context,
    anchor_id: str,
    url: str | None,
    bookmark_target: str | None,
    text: str | None,
    screen_tip: str | None,
) -> None:
    """Turn an anchor into a hyperlink — external `--url` or internal `--bookmark` (atomic-undo)."""
    if (url is None) == (bookmark_target is None):
        raise click.UsageError("pass exactly one of --url or --bookmark")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: link {anchor_id}"):
                anchor.link_to(
                    address=url, bookmark=bookmark_target, text=text, screen_tip=screen_tip
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {"url": url, "bookmark": bookmark_target, "text": text},
                },
                as_text=not ctx.obj["as_json"],
                text=f"linked {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="cross-ref")
@click.option("--anchor-id", "anchor_id", required=True, help="Where to insert the reference.")
@click.option(
    "--target",
    "target",
    required=True,
    help="Anchor id to reference: bookmark:NAME | heading:N | footnote:N | endnote:N.",
)
@click.option(
    "--kind",
    "kind",
    type=click.Choice(["text", "page", "number", "above_below"]),
    default="text",
    show_default=True,
    help="What the reference shows.",
)
@click.option(
    "--hyperlink/--no-hyperlink",
    "hyperlink",
    default=True,
    show_default=True,
    help="Make the inserted reference a clickable jump.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def cross_ref_cmd(
    ctx: click.Context, anchor_id: str, target: str, kind: str, hyperlink: bool, before: bool
) -> None:
    """Insert a cross-reference to another anchor (atomic-undo).

    `--target` resolves a bookmark by name, a heading/footnote/endnote by its id.
    Refresh stale references (page numbers move) with `update-fields`.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: cross-ref {target} {where} {anchor_id}"):
                anchor.insert_cross_reference(target, kind=kind, hyperlink=hyperlink, where=where)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "target": target,
                        "kind": kind,
                        "hyperlink": hyperlink,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted cross-reference to {target} {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="caption")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to caption (e.g. a figure).")
@click.option(
    "--label", "label", default="Figure", show_default=True, help="Caption label (Figure/Table/…)."
)
@click.option("--text", "text", default=None, help="Caption title after the label and number.")
@click.option(
    "--position",
    "position",
    type=click.Choice(["above", "below"], case_sensitive=False),
    default=None,
    help="Place the caption above or below the anchor "
    "(default: above for a Table, below otherwise).",
)
@click.pass_context
def caption_cmd(
    ctx: click.Context, anchor_id: str, label: str, text: str | None, position: str | None
) -> None:
    """Insert a numbered caption (Figure 1, Table 2, …) as its own paragraph (atomic-undo).

    The caption always becomes its own Caption-styled paragraph; on a table cell
    it is placed above/below the whole table. Table captions default to above,
    figures to below — pass --position to override.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: caption {label} {anchor_id}"):
                anchor.insert_caption(label, text=text, position=position)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {"label": label, "text": text, "position": position},
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {label} caption at {anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# create-content-control / mark-index-entry / insert-index / table-of-figures
# ---------------------------------------------------------------------------


@click.command(name="create-content-control")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor to wrap (or insert the control at)."
)
@click.option(
    "--kind",
    "kind",
    type=click.Choice(
        [
            "rich_text",
            "text",
            "picture",
            "combo_box",
            "dropdown",
            "date",
            "checkbox",
            "building_block",
            "group",
            "repeating_section",
        ],
        case_sensitive=False,
    ),
    default="rich_text",
    show_default=True,
    help="Content control type.",
)
@click.option(
    "--title", "title", default=None, help="Control title (addressable later as cc:TITLE)."
)
@click.option(
    "--tag", "tag", default=None, help="Control tag (a hidden name; cc: falls back to it)."
)
@click.option(
    "--item",
    "items",
    multiple=True,
    help="A combo_box/dropdown choice (repeatable). 'Text' or 'Text=Value'.",
)
@click.option(
    "--where",
    "where",
    type=click.Choice(["wrap", "before", "after"], case_sensitive=False),
    default="wrap",
    show_default=True,
    help="Wrap the anchor's range, or insert an empty control before/after it.",
)
@click.option("--lock-contents", "lock_contents", is_flag=True, help="Stop edits to the value.")
@click.option("--lock-control", "lock_control", is_flag=True, help="Stop deletion of the control.")
@click.pass_context
def create_content_control_cmd(
    ctx: click.Context,
    anchor_id: str,
    kind: str,
    title: str | None,
    tag: str | None,
    items: tuple[str, ...],
    where: str,
    lock_contents: bool,
    lock_control: bool,
) -> None:
    """Create a content control over an anchor (atomic-undo).

    The form-building primitive: wrap a range (or insert an empty control) of the
    given --kind. Give it a --title to address it later as cc:TITLE. For a
    combo_box/dropdown, pass --item once per choice ('Text' or 'Text=Value').
    """
    parsed_items: list[Any] | None = None
    if items:
        parsed_items = []
        for raw in items:
            label, sep, value = raw.partition("=")
            parsed_items.append({"text": label, "value": value} if sep else label)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: content control {kind} {where} {anchor_id}"):
                cc = anchor.insert_content_control(
                    kind,
                    title=title,
                    tag=tag,
                    items=parsed_items,
                    where=where,
                    lock_contents=lock_contents,
                    lock_control=lock_control,
                )
            name = cc.name or None
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "content_control": name,
                    "cc_anchor_id": cc.anchor_id if name else None,
                    "applied": {
                        "kind": kind,
                        "title": title,
                        "tag": tag,
                        "items": parsed_items,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"created {kind} content control at {anchor_id}"
                + (f" (cc:{name})" if name else ""),
            )

    _run(ctx, go)


def _cc_anchor(word: Any, doc_name: str | None, anchor_id: str) -> Any:
    """Resolve `anchor_id` to a content control, raising a clean usage error otherwise."""
    from .._anchors import ContentControl

    doc = _pick_doc(word, doc_name)
    anchor = doc.anchor_by_id(anchor_id)
    if not isinstance(anchor, ContentControl):
        raise click.UsageError(f"{anchor_id!r} is not a content control; pass a cc:NAME anchor")
    return doc, anchor


@click.command(name="set-cc-properties")
@click.option("--anchor-id", "anchor_id", required=True, help="Content control anchor (cc:NAME).")
@click.option("--title", "title", default=None, help='Control title (pass "" to clear it).')
@click.option("--tag", "tag", default=None, help='Control tag (pass "" to clear it).')
@click.option(
    "--lock-contents/--no-lock-contents",
    "lock_contents",
    default=None,
    help="Stop / allow edits to the value.",
)
@click.option(
    "--lock-control/--no-lock-control",
    "lock_control",
    default=None,
    help="Stop / allow deletion of the control.",
)
@click.pass_context
def set_cc_properties_cmd(
    ctx: click.Context,
    anchor_id: str,
    title: str | None,
    tag: str | None,
    lock_contents: bool | None,
    lock_control: bool | None,
) -> None:
    """Re-set a content control's title/tag/locks in place (atomic-undo).

    Pass at least one option; "" clears --title/--tag, omitting leaves it. A
    title (or tag) rename changes the control's cc:NAME anchor id.
    """
    raw: dict[str, Any] = {
        "title": title,
        "tag": tag,
        "lock_contents": lock_contents,
        "lock_control": lock_control,
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --title/--tag/--lock-contents/--lock-control")

    def go() -> None:
        with attach() as word:
            doc, anchor = _cc_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set content control properties {anchor_id}"):
                anchor.set_properties(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"updated {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="set-cc-items")
@click.option("--anchor-id", "anchor_id", required=True, help="Content control anchor (cc:NAME).")
@click.option(
    "--item",
    "items",
    multiple=True,
    required=True,
    help="A choice (repeatable). 'Text' or 'Text=Value'. Replaces the existing list.",
)
@click.pass_context
def set_cc_items_cmd(ctx: click.Context, anchor_id: str, items: tuple[str, ...]) -> None:
    """Replace a combo_box/dropdown's choice list in place (atomic-undo).

    Pass --item once per choice ('Text' or 'Text=Value'); the new list replaces
    the existing entries. Only valid on a combo_box/dropdown control.
    """
    parsed_items: list[Any] = []
    for raw in items:
        label, sep, value = raw.partition("=")
        parsed_items.append({"text": label, "value": value} if sep else label)

    def go() -> None:
        with attach() as word:
            doc, anchor = _cc_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set content control items {anchor_id}"):
                anchor.set_items(parsed_items)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": {"items": parsed_items}},
                as_text=not ctx.obj["as_json"],
                text=f"set {len(parsed_items)} item(s) on {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="mark-index-entry")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor whose range to index.")
@click.option(
    "--entry",
    "entry",
    required=True,
    help="Index text; use 'main:sub' for a subentry.",
)
@click.option(
    "--cross-reference",
    "cross_reference",
    default=None,
    help="Replace the page number with a 'see …' pointer.",
)
@click.option("--bold", "bold", is_flag=True, help="Bold the entry's page number.")
@click.option("--italic", "italic", is_flag=True, help="Italicise the entry's page number.")
@click.pass_context
def mark_index_entry_cmd(
    ctx: click.Context,
    anchor_id: str,
    entry: str,
    cross_reference: str | None,
    bold: bool,
    italic: bool,
) -> None:
    """Mark an anchor's range as a back-of-book index entry (atomic-undo).

    The per-term step; build the index itself with `insert-index`.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: mark index entry {anchor_id}"):
                anchor.mark_index_entry(
                    entry, cross_reference=cross_reference, bold=bold, italic=italic
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "entry": entry,
                        "cross_reference": cross_reference,
                        "bold": bold,
                        "italic": italic,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"marked index entry {entry!r} at {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-index")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="end",
    show_default=True,
    help="Where to insert the index (default: the document end).",
)
@click.option("--columns", "columns", type=int, default=2, show_default=True, help="Column count.")
@click.option("--run-in", "run_in", is_flag=True, help="Pack subentries into one paragraph.")
@click.option(
    "--right-align-page-numbers",
    "right_align",
    is_flag=True,
    help="Flush page numbers to the right margin.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def insert_index_cmd(
    ctx: click.Context,
    anchor_id: str,
    columns: int,
    run_in: bool,
    right_align: bool,
    before: bool,
) -> None:
    """Insert a back-of-book index from the marked entries (atomic-undo).

    Mark entries first with `mark-index-entry`. Page numbers populate after
    repagination — run `update-fields` (or take a `snapshot`) before reading them.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert index {where} {anchor_id}"):
                anchor.insert_index(
                    columns=columns,
                    run_in=run_in,
                    right_align_page_numbers=right_align,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "columns": columns,
                        "run_in": run_in,
                        "right_align_page_numbers": right_align,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted index {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="table-of-figures")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="start",
    show_default=True,
    help="Where to insert the table of figures (default: the document start).",
)
@click.option(
    "--label",
    "label",
    default="Figure",
    show_default=True,
    help="Caption label to gather (Figure/Table/Equation/…).",
)
@click.option(
    "--no-label",
    "no_label",
    is_flag=True,
    help="Drop the 'Figure 1' label prefix from each entry.",
)
@click.option(
    "--hyperlinks/--no-hyperlinks",
    "hyperlinks",
    default=True,
    show_default=True,
    help="Make each entry a clickable link.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def table_of_figures_cmd(
    ctx: click.Context,
    anchor_id: str,
    label: str,
    no_label: bool,
    hyperlinks: bool,
    before: bool,
) -> None:
    """Insert a table of figures built from captions of one label (atomic-undo).

    Page numbers populate after repagination — run `update-fields` (or take a
    `snapshot`) before reading them.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: table of figures {label} {where} {anchor_id}"):
                anchor.insert_table_of_figures(
                    label=label,
                    include_label=not no_label,
                    hyperlinks=hyperlinks,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "label": label,
                        "include_label": not no_label,
                        "hyperlinks": hyperlinks,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted table of figures ({label}) {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="bibliography-style")
@click.option(
    "--style",
    "style",
    required=True,
    help="Citation style, e.g. APA/MLA/Chicago/IEEE (build-dependent).",
)
@click.pass_context
def bibliography_style_cmd(ctx: click.Context, style: str) -> None:
    """Set the document's citation/bibliography style (atomic-undo).

    Refresh existing citations/bibliography with `update-fields` afterwards.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: bibliography style {style}"):
                doc.bibliography_style = style
            emit(
                {"ok": True, "applied": {"style": style}},
                as_text=not ctx.obj["as_json"],
                text=f"set bibliography style to {style!r}",
            )

    _run(ctx, go)


@click.command(name="add-source")
@click.option(
    "--type",
    "source_type",
    default="book",
    show_default=True,
    help="Source type (book/journal_article/conference_proceedings/case/…).",
)
@click.option(
    "--tag",
    "tag",
    default=None,
    help="Citation tag; auto-derived from first author + year if omitted.",
)
@click.option("--author", "authors", multiple=True, help="Author 'Last, First' (repeatable).")
@click.option("--title", "title", default=None, help="Source title.")
@click.option("--year", "year", default=None, help="Publication year.")
@click.option("--publisher", "publisher", default=None, help="Publisher.")
@click.option("--city", "city", default=None, help="City of publication.")
@click.option("--journal-name", "journal_name", default=None, help="Journal/periodical name.")
@click.option("--volume", "volume", default=None, help="Volume.")
@click.option("--issue", "issue", default=None, help="Issue.")
@click.option("--pages", "pages", default=None, help="Page range.")
@click.option("--url", "url", default=None, help="URL.")
@click.option("--edition", "edition", default=None, help="Edition.")
@click.option("--doi", "doi", default=None, help="DOI.")
@click.option(
    "--xml",
    "xml",
    default=None,
    help="Raw <b:Source> XML (escape hatch; supersedes the typed fields).",
)
@click.pass_context
def add_source_cmd(
    ctx: click.Context,
    source_type: str,
    tag: str | None,
    authors: tuple[str, ...],
    title: str | None,
    year: str | None,
    publisher: str | None,
    city: str | None,
    journal_name: str | None,
    volume: str | None,
    issue: str | None,
    pages: str | None,
    url: str | None,
    edition: str | None,
    doi: str | None,
    xml: str | None,
) -> None:
    """Add a bibliography source to the document's store (atomic-undo).

    Cite it with `insert-citation --tag TAG`, then list cited sources with
    `insert-bibliography`.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: add source"):
                if xml:
                    src = doc.sources.add_xml(xml)
                else:
                    src = doc.sources.add(
                        source_type,
                        tag=tag,
                        author=list(authors) or None,
                        title=title,
                        year=year,
                        publisher=publisher,
                        city=city,
                        journal_name=journal_name,
                        volume=volume,
                        issue=issue,
                        pages=pages,
                        url=url,
                        edition=edition,
                        doi=doi,
                    )
            emit(
                {
                    "ok": True,
                    "source": src.tag,
                    "applied": {"source_type": source_type, "tag": src.tag},
                },
                as_text=not ctx.obj["as_json"],
                text=f"added source {src.tag!r}",
            )

    _run(ctx, go)


@click.command(name="insert-citation")
@click.option("--anchor-id", "anchor_id", required=True, help="Where to insert the citation.")
@click.option("--tag", "tag", required=True, help="Source tag to cite.")
@click.option("--pages", "pages", default=None, help="Page locator, e.g. '15'.")
@click.option("--prefix", "prefix", default=None, help="Text before the citation, e.g. 'see '.")
@click.option("--suffix", "suffix", default=None, help="Text after the citation, e.g. ', at 12'.")
@click.option("--volume", "volume", default=None, help="Volume.")
@click.option("--suppress-author", "suppress_author", is_flag=True, help="Drop the author.")
@click.option("--suppress-year", "suppress_year", is_flag=True, help="Drop the year.")
@click.option("--suppress-title", "suppress_title", is_flag=True, help="Drop the title.")
@click.option("--locale", "locale", type=int, default=1033, show_default=True, help="Format LCID.")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def insert_citation_cmd(
    ctx: click.Context,
    anchor_id: str,
    tag: str,
    pages: str | None,
    prefix: str | None,
    suffix: str | None,
    volume: str | None,
    suppress_author: bool,
    suppress_year: bool,
    suppress_title: bool,
    locale: int,
    before: bool,
) -> None:
    """Insert an in-text citation referencing a source by tag (atomic-undo)."""
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: citation {tag} {where} {anchor_id}"):
                anchor.insert_citation(
                    tag,
                    pages=pages,
                    prefix=prefix,
                    suffix=suffix,
                    volume=volume,
                    suppress_author=suppress_author,
                    suppress_year=suppress_year,
                    suppress_title=suppress_title,
                    locale=locale,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "citation": tag,
                    "applied": {
                        "tag": tag,
                        "pages": pages,
                        "prefix": prefix,
                        "suffix": suffix,
                        "volume": volume,
                        "suppress_author": suppress_author,
                        "suppress_year": suppress_year,
                        "suppress_title": suppress_title,
                        "locale": locale,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted citation {tag!r} {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-bibliography")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="end",
    show_default=True,
    help="Where to insert the bibliography (default: the document end).",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def insert_bibliography_cmd(ctx: click.Context, anchor_id: str, before: bool) -> None:
    """Insert a bibliography of the cited sources (atomic-undo).

    Page numbers/entries populate after repagination — run `update-fields` (or take
    a `snapshot`) before reading them.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert bibliography {where} {anchor_id}"):
                anchor.insert_bibliography(where=where)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": {"where": where}},
                as_text=not ctx.obj["as_json"],
                text=f"inserted bibliography {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="mark-citation")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor whose range to mark.")
@click.option(
    "--long",
    "long_citation",
    required=True,
    help="Full citation as it appears in the table.",
)
@click.option(
    "--short",
    "short_citation",
    default=None,
    help="Abbreviated form Word matches elsewhere (defaults to --long).",
)
@click.option(
    "--category",
    "category",
    default="cases",
    show_default=True,
    help="cases/statutes/other/rules/treatises/regulations/constitutional, or 1-16.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def mark_citation_cmd(
    ctx: click.Context,
    anchor_id: str,
    long_citation: str,
    short_citation: str | None,
    category: str,
    before: bool,
) -> None:
    """Mark an anchor's range as a table-of-authorities citation (atomic-undo).

    The per-authority step; build the table with `table-of-authorities`.
    """
    where = "before" if before else "after"
    cat: str | int = int(category) if category.isdigit() else category

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: mark citation {anchor_id}"):
                anchor.mark_citation(
                    long_citation, short_citation=short_citation, category=cat, where=where
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "long_citation": long_citation,
                        "short_citation": short_citation,
                        "category": category,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"marked citation {long_citation!r} at {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="table-of-authorities")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="end",
    show_default=True,
    help="Where to insert the table (default: the document end).",
)
@click.option(
    "--category",
    "category",
    default="all",
    show_default=True,
    help="all/cases/statutes/other/rules/treatises/regulations/constitutional, or 1-16.",
)
@click.option("--no-passim", "no_passim", is_flag=True, help="Don't collapse 5+ refs to 'passim'.")
@click.option(
    "--no-keep-formatting",
    "no_keep_formatting",
    is_flag=True,
    help="Don't preserve each entry's character formatting.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def table_of_authorities_cmd(
    ctx: click.Context,
    anchor_id: str,
    category: str,
    no_passim: bool,
    no_keep_formatting: bool,
    before: bool,
) -> None:
    """Insert a table of authorities from the marked citations (atomic-undo).

    Mark citations first with `mark-citation`. Page numbers populate after
    repagination — run `update-fields` (or take a `snapshot`) before reading them.
    """
    where = "before" if before else "after"
    cat: str | int = int(category) if category.isdigit() else category

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: table of authorities {where} {anchor_id}"):
                anchor.insert_table_of_authorities(
                    category=cat,
                    passim=not no_passim,
                    keep_entry_formatting=not no_keep_formatting,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "category": category,
                        "passim": not no_passim,
                        "keep_entry_formatting": not no_keep_formatting,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted table of authorities {where} {anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# document themes — theme / list-themes / apply-theme / set-theme-colors / -fonts
# ---------------------------------------------------------------------------


@click.command(name="theme")
@click.pass_context
def theme_cmd(ctx: click.Context) -> None:
    """Show the document's current theme (colours + major/minor fonts). Non-mutating."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data = doc.theme.to_dict()
            text = "\n".join(
                [
                    f"major font: {data['major_font']}",
                    f"minor font: {data['minor_font']}",
                    *(f"{k}: {v}" for k, v in data["colors"].items()),
                ]
            )
            emit(data, as_text=not ctx.obj["as_json"], text=text)

    _run(ctx, go)


@click.command(name="list-themes")
@click.pass_context
def list_themes_cmd(ctx: click.Context) -> None:
    """List the built-in themes, colour schemes, and font schemes Office ships.

    These names feed `apply-theme --theme`, `set-theme-colors --scheme`, and
    `set-theme-fonts --scheme`. Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data = doc.theme.list_available()
            text = "\n".join(
                [
                    f"themes: {', '.join(data['themes'])}",
                    f"color schemes: {', '.join(data['color_schemes'])}",
                    f"font schemes: {', '.join(data['font_schemes'])}",
                ]
            )
            emit(data, as_text=not ctx.obj["as_json"], text=text)

    _run(ctx, go)


@click.command(name="apply-theme")
@click.option(
    "--theme",
    "theme",
    required=True,
    help="Built-in theme name (e.g. Facet, Ion) or a .thmx file path.",
)
@click.pass_context
def apply_theme_cmd(ctx: click.Context, theme: str) -> None:
    """Apply a whole document theme — colours, fonts, and effects (atomic-undo).

    See `list-themes` for the built-in names. Brand colours/fonts can then be
    overridden with `set-theme-colors` / `set-theme-fonts`.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: apply theme {theme}"):
                applied = doc.theme.apply(theme)
            emit(
                {"ok": True, "applied": {"theme": applied}},
                as_text=not ctx.obj["as_json"],
                text=f"applied theme {applied!r}",
            )

    _run(ctx, go)


@click.command(name="set-theme-colors")
@click.option("--scheme", "scheme", default=None, help="Built-in colour scheme name or .xml path.")
@click.option("--text1", "text1", default=None, help="Text 1 (dark 1) colour.")
@click.option("--background1", "background1", default=None, help="Background 1 (light 1) colour.")
@click.option("--text2", "text2", default=None, help="Text 2 (dark 2) colour.")
@click.option("--background2", "background2", default=None, help="Background 2 (light 2) colour.")
@click.option("--accent1", "accent1", default=None, help="Accent 1 colour (name/hex).")
@click.option("--accent2", "accent2", default=None, help="Accent 2 colour.")
@click.option("--accent3", "accent3", default=None, help="Accent 3 colour.")
@click.option("--accent4", "accent4", default=None, help="Accent 4 colour.")
@click.option("--accent5", "accent5", default=None, help="Accent 5 colour.")
@click.option("--accent6", "accent6", default=None, help="Accent 6 colour.")
@click.option("--hyperlink", "hyperlink", default=None, help="Hyperlink colour.")
@click.option(
    "--followed-hyperlink", "followed_hyperlink", default=None, help="Followed-hyperlink colour."
)
@click.pass_context
def set_theme_colors_cmd(
    ctx: click.Context,
    scheme: str | None,
    text1: str | None,
    background1: str | None,
    text2: str | None,
    background2: str | None,
    accent1: str | None,
    accent2: str | None,
    accent3: str | None,
    accent4: str | None,
    accent5: str | None,
    accent6: str | None,
    hyperlink: str | None,
    followed_hyperlink: str | None,
) -> None:
    """Set the theme's colour scheme and/or individual brand colours (atomic-undo).

    Pass `--scheme` for a named built-in scheme, and/or any `--accentN`/`--text*`
    flag to override a single colour (a name like `navy` or hex like `#1A73E8`).
    """
    overrides = {
        k: v
        for k, v in {
            "text1": text1,
            "background1": background1,
            "text2": text2,
            "background2": background2,
            "accent1": accent1,
            "accent2": accent2,
            "accent3": accent3,
            "accent4": accent4,
            "accent5": accent5,
            "accent6": accent6,
            "hyperlink": hyperlink,
            "followed_hyperlink": followed_hyperlink,
        }.items()
        if v is not None
    }

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: set theme colors"):
                colors = doc.theme.set_colors(scheme=scheme, **overrides)
            emit(
                {"ok": True, "colors": colors, "applied": {"scheme": scheme, **overrides}},
                as_text=not ctx.obj["as_json"],
                text="set theme colours",
            )

    _run(ctx, go)


@click.command(name="set-theme-fonts")
@click.option("--scheme", "scheme", default=None, help="Built-in font scheme name or .xml path.")
@click.option("--major", "major", default=None, help="Major (heading) font name.")
@click.option("--minor", "minor", default=None, help="Minor (body) font name.")
@click.pass_context
def set_theme_fonts_cmd(
    ctx: click.Context, scheme: str | None, major: str | None, minor: str | None
) -> None:
    """Set the theme's fonts via a named scheme and/or explicit names (atomic-undo).

    `--scheme` loads a named built-in font scheme; `--major`/`--minor` override
    the heading/body font names.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: set theme fonts"):
                fonts = doc.theme.set_fonts(scheme=scheme, major=major, minor=minor)
            emit(
                {
                    "ok": True,
                    **fonts,
                    "applied": {"scheme": scheme, "major": major, "minor": minor},
                },
                as_text=not ctx.obj["as_json"],
                text=f"set theme fonts (major={fonts['major_font']}, minor={fonts['minor_font']})",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# prepend / append --text "..." [--paragraph|--inline] [--style "..."]
# ---------------------------------------------------------------------------


@click.command(name="prepend")
@click.option("--text", "text", required=True, help="Text to prepend at the start of the document.")
@click.option(
    "--inline/--paragraph",
    "inline",
    default=False,
    show_default="--paragraph",
    help="Prepend inline (join the first paragraph) instead of as a new paragraph.",
)
@click.option(
    "--style",
    "style",
    default=None,
    help="Optional Word style for the prepended paragraph (paragraph mode only).",
)
@click.pass_context
def prepend_cmd(ctx: click.Context, text: str, inline: bool, style: str | None) -> None:
    """Prepend text to the start of the document (atomic-undo).

    The mirror of `append` — no anchor needed. By default `text` becomes a new
    first paragraph (`--style` optional, validated first); pass `--inline` to
    join the document's first paragraph instead. Equivalent to
    `insert --anchor-id start --text "…"`.
    """
    if inline and style is not None:
        raise click.UsageError("--style is only valid in --paragraph mode")
    mode = "inline" if inline else "paragraph"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: prepend to start of document"):
                if inline:
                    doc.prepend(text)
                else:
                    doc.prepend_paragraph(text, style=style)
            emit(
                {"ok": True, "mode": mode, "style": None if inline else style},
                as_text=not ctx.obj["as_json"],
                text=f"prepended ({mode}) to start of document",
            )

    _run(ctx, go)


@click.command(name="append")
@click.option("--text", "text", required=True, help="Text to append at the end of the document.")
@click.option(
    "--inline/--paragraph",
    "inline",
    default=False,
    show_default="--paragraph",
    help="Append inline (continue the last paragraph) instead of as a new paragraph.",
)
@click.option(
    "--style",
    "style",
    default=None,
    help="Optional Word style for the appended paragraph (paragraph mode only).",
)
@click.pass_context
def append_cmd(ctx: click.Context, text: str, inline: bool, style: str | None) -> None:
    """Append text to the end of the document (atomic-undo).

    The high-level "end of doc" helper — no anchor needed. By default `text`
    becomes a new final paragraph (`--style` optional, validated first); pass
    `--inline` to continue the document's last paragraph instead. Equivalent to
    `insert --anchor-id end --text "…"`.
    """
    if inline and style is not None:
        raise click.UsageError("--style is only valid in --paragraph mode")
    mode = "inline" if inline else "paragraph"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: append to end of document"):
                if inline:
                    doc.append(text)
                else:
                    doc.append_paragraph(text, style=style)
            emit(
                {"ok": True, "mode": mode, "style": None if inline else style},
                as_text=not ctx.obj["as_json"],
                text=f"appended ({mode}) to end of document",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# insert-image --anchor-id ID (--path FILE | --base64 VALUE) --wrap WRAP ...
# ---------------------------------------------------------------------------


_WRAP_CHOICES = ["inline", "auto", "square", "tight", "through", "top-bottom", "behind", "front"]


@click.command(name="insert-image")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor to insert the image relative to."
)
@click.option(
    "--path", "path", default=None, type=click.Path(path_type=Path), help="Path to the image file."
)
@click.option(
    "--base64", "b64", default=None, help="Base64 image data, or '-' to read base64 from stdin."
)
@click.option(
    "--wrap",
    "wrap",
    required=True,
    type=click.Choice(_WRAP_CHOICES),
    help="Layout / text-wrap (required). 'inline' stays in the text flow; "
    "'auto' floats Square when small, else top-bottom.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.option(
    "--block/--no-block",
    "block",
    default=False,
    show_default="--no-block",
    help="Place the image on its own new (Normal) line instead of in the anchor's text run.",
)
@click.option("--width", "width", type=float, default=None, help="Width in points (optional).")
@click.option("--height", "height", type=float, default=None, help="Height in points (optional).")
@click.option("--alt-text", "alt_text", default=None, help="Alternative (accessibility) text.")
@click.option(
    "--lock-aspect/--no-lock-aspect",
    "lock_aspect",
    default=True,
    show_default=True,
    help="Keep the image's aspect ratio when resizing.",
)
@click.pass_context
def insert_image_cmd(
    ctx: click.Context,
    anchor_id: str,
    path: Path | None,
    b64: str | None,
    wrap: str,
    before: bool,
    block: bool,
    width: float | None,
    height: float | None,
    alt_text: str | None,
    lock_aspect: bool,
) -> None:
    """Insert an image at any anchor, from a file or base64 (atomic-undo).

    Exactly one of --path / --base64 is required. --path is best for large
    images; base64 (or '--base64 -' from stdin) suits an LLM holding image
    data in memory. --wrap is required so layout is always explicit.
    """
    if (path is None) == (b64 is None):
        raise click.UsageError("pass exactly one of --path or --base64")
    if b64 == "-":
        b64 = click.get_text_stream("stdin").read()
    image: str | Path = path if path is not None else (b64 or "")
    where = "before" if before else "after"

    def go() -> None:
        # Screen a --path source against the policy *before* the COM/filesystem
        # probe: a UNC path's own existence check would authenticate to a remote
        # SMB server. Inside go() so _run() maps a denial to the right exit code.
        if path is not None:
            ctx.obj["policy"].screen_image_path(path)
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert image {where} {anchor_id}"):
                shape = anchor.insert_image(
                    image,
                    wrap=wrap,
                    where=where,
                    block=block,
                    width=width,
                    height=height,
                    alt_text=alt_text,
                    lock_aspect=lock_aspect,
                )
            # A floating image returns its shape:N handle (inline stays None).
            shape_id = shape.anchor_id if shape is not None else None
            text = f"inserted image {where} {anchor_id} (wrap={wrap})"
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "shape": shape_id,
                    "wrap": wrap,
                    "where": where,
                    "block": block,
                },
                as_text=not ctx.obj["as_json"],
                text=(f"{text} -> {shape_id}" if shape_id else text),
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# equations / insert-equation
# ---------------------------------------------------------------------------


def _fmt_equations(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no equations)"
    lines: list[str] = []
    for r in rows:
        para = r.get("para") or ""
        linear = r.get("linear") or ""
        preview = f"  {linear}" if linear else ""
        lines.append(f"[{r['anchor_id']}] {r.get('type', '?')}  {para}{preview}")
    return "\n".join(lines)


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


def _fmt_charts(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no charts)"
    lines: list[str] = []
    for r in rows:
        para = r.get("para") or ""
        title = r.get("title")
        suffix = f"  {title!r}" if title else ""
        lines.append(f"[{r['anchor_id']}] {r.get('kind', '?')}  {para}{suffix}")
    return "\n".join(lines)


@click.command(name="charts")
@click.pass_context
def charts_cmd(ctx: click.Context) -> None:
    """List the document's charts (chart:N id, kind, title, para:N).

    The discovery half of charting: see what charts are in the document and
    address them by `chart:N`. Metadata only (the series data is static) and
    non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.charts.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_charts(rows))

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


@click.command(name="insert-chart")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor to insert the chart relative to."
)
@click.option(
    "--kind",
    "kind",
    required=True,
    type=click.Choice(["bar", "pie", "line", "scatter"]),
    help="Chart kind: bar (clustered columns), pie, line, or scatter.",
)
@click.option(
    "--data",
    "data",
    required=True,
    help="Chart data as JSON, or '-' to read it from stdin: an object "
    '{"label": value} (bar/pie/line) or an array of [x, y] pairs (scatter/line).',
)
@click.option("--title", "title", default=None, help="Chart title (and series name).")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def insert_chart_cmd(
    ctx: click.Context,
    anchor_id: str,
    kind: str,
    data: str,
    title: str | None,
    before: bool,
) -> None:
    """Insert an Excel-backed chart at any anchor (atomic-undo).

    --data is JSON (or '-' for stdin): an object {"Q1": 10, "Q2": 25} for
    bar/pie/line, or an array of [x, y] pairs [[1.2, 3.4], [2.5, 6.1]] for
    scatter (both axes numeric; duplicate x preserved). Charts embed a hidden
    Excel workbook, so Excel must be installed (exit 6 if not); the data link is
    then broken, so the chart's data is static.
    """
    raw = click.get_text_stream("stdin").read() if data == "-" else data
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--data is not valid JSON: {e}") from e
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert chart {where} {anchor_id}"):
                chart = anchor.insert_chart(kind, parsed, title=title, where=where)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "chart": chart.index,
                    "chart_anchor_id": chart.anchor_id,
                    "kind": kind,
                    "where": where,
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {chart.anchor_id} ({kind}) {where} {anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# format-chart / format-axis / add-trendline / set-series-color (chart:N)
# ---------------------------------------------------------------------------


def _chart_anchor(word: Any, doc_name: str | None, anchor_id: str) -> Any:
    """Resolve `anchor_id` to a chart, raising a clean usage error otherwise."""
    from .._anchors import ChartAnchor

    doc = _pick_doc(word, doc_name)
    anchor = doc.anchor_by_id(anchor_id)
    if not isinstance(anchor, ChartAnchor):
        raise click.UsageError(f"{anchor_id!r} is not a chart; pass a chart:N anchor")
    return doc, anchor


@click.command(name="format-chart")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N) to format.")
@click.option("--title", "title", default=None, help='Chart title (pass "" to clear it).')
@click.option("--legend/--no-legend", "legend", default=None, help="Show or hide the legend.")
@click.option(
    "--legend-position",
    "legend_position",
    type=click.Choice(["right", "left", "top", "bottom", "corner"]),
    default=None,
    help="Where the legend sits (implies it is shown).",
)
@click.option(
    "--chart-style", "chart_style", type=int, default=None, help="Design-gallery style id."
)
@click.option("--background", "background", default=None, help="Chart-area fill colour.")
@click.option("--plot-background", "plot_background", default=None, help="Plot-area fill colour.")
@click.option("--font", "font", default=None, help="Whole-chart font family.")
@click.option("--font-size", "font_size", default=None, help="Whole-chart font size (pt or unit).")
@click.option("--font-color", "font_color", default=None, help="Whole-chart font colour.")
@click.option(
    "--data-labels/--no-data-labels", "data_labels", default=None, help="Show point data labels."
)
@click.option(
    "--data-label-format", "data_label_format", default=None, help="Data-label number format."
)
@click.option(
    "--chart-type",
    "chart_type",
    type=click.Choice(["bar", "pie", "line", "scatter"]),
    default=None,
    help="Re-type the chart in place.",
)
@click.option(
    "--gap-width", "gap_width", type=int, default=None, help="Bar gap width (bar/column charts)."
)
@click.option(
    "--overlap", "overlap", type=int, default=None, help="Bar overlap (bar/column charts)."
)
@click.option(
    "--data-table/--no-data-table",
    "data_table",
    default=None,
    help="Show the data-table grid beneath the plot.",
)
@click.pass_context
def format_chart_cmd(
    ctx: click.Context,
    anchor_id: str,
    title: str | None,
    legend: bool | None,
    legend_position: str | None,
    chart_style: int | None,
    background: str | None,
    plot_background: str | None,
    font: str | None,
    font_size: str | None,
    font_color: str | None,
    data_labels: bool | None,
    data_label_format: str | None,
    chart_type: str | None,
    gap_width: int | None,
    overlap: int | None,
    data_table: bool | None,
) -> None:
    """Apply whole-chart / design formatting to a chart (atomic-undo).

    Operates on the static chart — no Excel needed. Colours are a name, hex
    (#2E86C1), or comma-separated r,g,b. Pass at least one option.
    """
    raw: dict[str, Any] = {
        "title": title,
        "legend": legend,
        "legend_position": legend_position,
        "chart_style": chart_style,
        "background": _parse_color(background),
        "plot_background": _parse_color(plot_background),
        "font": font,
        "font_size": font_size,
        "font_color": _parse_color(font_color),
        "data_labels": data_labels,
        "data_label_format": data_label_format,
        "chart_type": chart_type,
        "gap_width": gap_width,
        "overlap": overlap,
        "data_table": data_table,
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one formatting option")

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: format chart {anchor_id}"):
                anchor.format(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"formatted {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="format-axis")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N) to format.")
@click.option(
    "--which",
    "which",
    required=True,
    type=click.Choice(["value", "y", "category", "x"]),
    help="Which axis: value/y or category/x.",
)
@click.option("--title", "title", default=None, help='Axis title (pass "" to clear it).')
@click.option("--minimum", "minimum", type=float, default=None, help="Axis minimum.")
@click.option("--maximum", "maximum", type=float, default=None, help="Axis maximum.")
@click.option(
    "--scale",
    "scale",
    type=click.Choice(["linear", "log"]),
    default=None,
    help="Axis scale type.",
)
@click.option("--number-format", "number_format", default=None, help="Tick-label number format.")
@click.option("--gridlines/--no-gridlines", "gridlines", default=None, help="Show major gridlines.")
@click.pass_context
def format_axis_cmd(
    ctx: click.Context,
    anchor_id: str,
    which: str,
    title: str | None,
    minimum: float | None,
    maximum: float | None,
    scale: str | None,
    number_format: str | None,
    gridlines: bool | None,
) -> None:
    """Format one axis of a chart (atomic-undo). `log` scale suits order-of-magnitude data."""
    raw: dict[str, Any] = {
        "title": title,
        "minimum": minimum,
        "maximum": maximum,
        "scale": scale,
        "number_format": number_format,
        "gridlines": gridlines,
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one axis option")

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: format axis {anchor_id}"):
                anchor.set_axis(which, **kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "which": which, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"formatted {which} axis of {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="add-trendline")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N).")
@click.option("--series", "series", type=int, default=1, show_default=True, help="1-based series.")
@click.option(
    "--kind",
    "kind",
    type=click.Choice(
        ["linear", "exponential", "logarithmic", "moving_average", "polynomial", "power"]
    ),
    default="linear",
    show_default=True,
    help="Trendline curve type.",
)
@click.option(
    "--display-equation", "display_equation", is_flag=True, default=False, help="Show the equation."
)
@click.option(
    "--display-r-squared", "display_r_squared", is_flag=True, default=False, help="Show R²."
)
@click.option("--forward", "forward", type=float, default=None, help="Forecast forward N units.")
@click.option("--backward", "backward", type=float, default=None, help="Forecast backward N units.")
@click.option(
    "--order", "order", type=int, default=None, help="Polynomial degree 2–6 (kind=polynomial)."
)
@click.option(
    "--period",
    "period",
    type=int,
    default=None,
    help="Moving-average window (kind=moving_average).",
)
@click.pass_context
def add_trendline_cmd(
    ctx: click.Context,
    anchor_id: str,
    series: int,
    kind: str,
    display_equation: bool,
    display_r_squared: bool,
    forward: float | None,
    backward: float | None,
    order: int | None,
    period: int | None,
) -> None:
    """Fit a trendline to a chart series (atomic-undo).

    A power/exponential fit with --display-equation draws the law of best fit.
    """
    kwargs: dict[str, Any] = {
        "series": series,
        "kind": kind,
        "display_equation": display_equation,
        "display_r_squared": display_r_squared,
    }
    if forward is not None:
        kwargs["forward"] = forward
    if backward is not None:
        kwargs["backward"] = backward
    if order is not None:
        kwargs["order"] = order
    if period is not None:
        kwargs["period"] = period

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: add trendline {anchor_id}"):
                anchor.add_trendline(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"added {kind} trendline to {anchor_id} series {series}",
            )

    _run(ctx, go)


@click.command(name="set-series-color")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N).")
@click.option(
    "--color", "color", required=True, help="Colour: name, hex (#2E86C1), or comma-separated r,g,b."
)
@click.option("--series", "series", type=int, default=1, show_default=True, help="1-based series.")
@click.option(
    "--point",
    "point",
    type=int,
    default=None,
    help="1-based point/slice to recolour (omit to colour the whole series).",
)
@click.pass_context
def set_series_color_cmd(
    ctx: click.Context, anchor_id: str, color: str, series: int, point: int | None
) -> None:
    """Recolour a chart series, or a single point / pie slice (atomic-undo)."""
    parsed = _parse_color(color)

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set series color {anchor_id}"):
                anchor.set_series_color(parsed, series=series, point=point)
            target = f"point {point}" if point is not None else f"series {series}"
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "series": series,
                    "point": point,
                    "color": color,
                },
                as_text=not ctx.obj["as_json"],
                text=f"recoloured {target} of {anchor_id} -> {color}",
            )

    _run(ctx, go)


@click.command(name="format-series")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N).")
@click.option("--series", "series", type=int, default=1, show_default=True, help="1-based series.")
@click.option(
    "--point",
    "point",
    type=int,
    default=None,
    help="1-based point/slice to target (omit for the whole series).",
)
@click.option(
    "--marker",
    "marker",
    default=None,
    help="Marker glyph: circle/square/diamond/triangle/x/star/dot/dash/plus/none/auto.",
)
@click.option("--marker-size", "marker_size", type=int, default=None, help="Marker size 2–72.")
@click.option(
    "--smooth/--no-smooth", "smooth", default=None, help="Curve a line/scatter through its points."
)
@click.option(
    "--explosion", "explosion", type=int, default=None, help="Pull a pie slice out 0–400."
)
@click.option(
    "--data-labels/--no-data-labels",
    "data_labels",
    default=None,
    help="Show this series' point labels.",
)
@click.option(
    "--data-label-size", "data_label_size", type=float, default=None, help="Data-label font size."
)
@click.option(
    "--data-label-color", "data_label_color", default=None, help="Data-label font colour."
)
@click.pass_context
def format_series_cmd(
    ctx: click.Context,
    anchor_id: str,
    series: int,
    point: int | None,
    marker: str | None,
    marker_size: int | None,
    smooth: bool | None,
    explosion: int | None,
    data_labels: bool | None,
    data_label_size: float | None,
    data_label_color: str | None,
) -> None:
    """Format a chart series, or a single point / slice (atomic-undo).

    Markers and --smooth suit line/scatter; --explosion a pie slice. Colours are
    a name, hex (#2E86C1), or comma-separated r,g,b. Pass at least one option.
    """
    raw: dict[str, Any] = {
        "marker": marker,
        "marker_size": marker_size,
        "smooth": smooth,
        "explosion": explosion,
        "data_labels": data_labels,
        "data_label_size": data_label_size,
        "data_label_color": _parse_color(data_label_color),
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one formatting option")
    kwargs["series"] = series
    kwargs["point"] = point

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: format series {anchor_id}"):
                anchor.format_series(**kwargs)
            target = f"point {point}" if point is not None else f"series {series}"
            emit(
                {"ok": True, "anchor_id": anchor_id, "series": series, "point": point},
                as_text=not ctx.obj["as_json"],
                text=f"formatted {target} of {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="add-error-bars")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N).")
@click.option("--series", "series", type=int, default=1, show_default=True, help="1-based series.")
@click.option(
    "--kind",
    "kind",
    type=click.Choice(["fixed", "percent", "stdev", "sterror"]),
    default="fixed",
    show_default=True,
    help="How the error amount is computed.",
)
@click.option(
    "--amount",
    "amount",
    type=float,
    default=None,
    help="Error magnitude (required unless kind=sterror).",
)
@click.option(
    "--include",
    "include",
    type=click.Choice(["both", "plus", "minus"]),
    default="both",
    show_default=True,
    help="Which side(s) to draw.",
)
@click.option(
    "--axis",
    "axis",
    type=click.Choice(["y", "value", "x", "category"]),
    default="y",
    show_default=True,
    help="Which axis the bars run along.",
)
@click.pass_context
def add_error_bars_cmd(
    ctx: click.Context,
    anchor_id: str,
    series: int,
    kind: str,
    amount: float | None,
    include: str,
    axis: str,
) -> None:
    """Draw error bars on a chart series (atomic-undo)."""
    kwargs: dict[str, Any] = {"series": series, "kind": kind, "include": include, "axis": axis}
    if amount is not None:
        kwargs["amount"] = amount

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: add error bars {anchor_id}"):
                anchor.add_error_bars(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "series": series, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"added {kind} error bars to {anchor_id} series {series}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# shapes | set-shape-wrap | set-shape-position | set-shape-size | format-shape
#        | set-shape-alt-text | set-shape-text | replace-shape-image | delete-shape
# ---------------------------------------------------------------------------


def _fmt_shapes(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no shapes)"
    lines: list[str] = []
    for r in rows:
        para = r.get("para") or ""
        w, h = r.get("width"), r.get("height")
        dims = f"{round(w)}x{round(h)}pt" if w is not None and h is not None else ""
        wrap = r.get("wrap")
        side = r.get("wrap_side")
        wrap_txt = f"wrap={wrap}" + (f"/{side}" if side and side != "both" else "")
        crop_txt = "  cropped" if r.get("crop") else ""
        lines.append(
            f"[{r['anchor_id']}] {r.get('shape_type', '?')}  {dims}  "
            f"{wrap_txt}{crop_txt}  {para}".rstrip()
        )
    return "\n".join(lines)


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


def _shape_anchor(word: Any, doc_name: str | None, anchor_id: str) -> Any:
    """Resolve `anchor_id` to a floating shape, raising a clean usage error otherwise."""
    from .._anchors import ShapeAnchor

    doc = _pick_doc(word, doc_name)
    anchor = doc.anchor_by_id(anchor_id)
    if not isinstance(anchor, ShapeAnchor):
        raise click.UsageError(f"{anchor_id!r} is not a shape; pass a shape:N anchor")
    return doc, anchor


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


def _image_anchor(word: Any, doc_name: str | None, anchor_id: str) -> Any:
    """Resolve `anchor_id` to an inline image, raising a clean usage error otherwise."""
    from .._anchors import ImageAnchor

    doc = _pick_doc(word, doc_name)
    anchor = doc.anchor_by_id(anchor_id)
    if not isinstance(anchor, ImageAnchor):
        raise click.UsageError(f"{anchor_id!r} is not an inline image; pass an image:N anchor")
    return doc, anchor


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


# ---------------------------------------------------------------------------
# snapshot [--anchor-id ID | --page N | --pages A-B] [--out FILE] [--dpi N]
# ---------------------------------------------------------------------------


def _parse_pages_range(value: str) -> tuple[int, int]:
    """Parse a `--pages` value like `2-4` into an inclusive `(start, end)` span."""
    start_str, sep, end_str = value.partition("-")
    if not sep:
        raise click.UsageError("--pages must look like 'A-B' (inclusive), e.g. '2-4'")
    try:
        start, end = int(start_str), int(end_str)
    except ValueError as e:
        raise click.UsageError("--pages must look like 'A-B' (inclusive), e.g. '2-4'") from e
    if start < 1 or end < start:
        raise click.UsageError(f"invalid page span {value!r}: need 1 <= start <= end")
    return start, end


def _parse_rc(value: str) -> tuple[int, int]:
    """Parse a 1-based cell coordinate like `2:3` into `(row, col)`."""
    row_str, sep, col_str = value.partition(":")
    if not sep:
        raise click.UsageError("cell must look like 'R:C' (1-based), e.g. '2:3'")
    try:
        row, col = int(row_str), int(col_str)
    except ValueError as e:
        raise click.UsageError("cell must look like 'R:C' (1-based), e.g. '2:3'") from e
    if row < 1 or col < 1:
        raise click.UsageError(f"invalid cell {value!r}: row and column are 1-based")
    return row, col


def _parse_color(value: str | None) -> str | tuple[int, int, int] | None:
    """Turn a `--color` value into something `to_bgr` understands.

    A comma-separated `r,g,b` becomes an `(r, g, b)` tuple; anything else
    (a colour name or hex string) passes through unchanged for the helper to
    resolve. Returns `None` for `None` (option not given).
    """
    if value is None:
        return None
    if "," in value:
        try:
            r, g, b = (int(p.strip()) for p in value.split(","))
        except ValueError as e:
            raise click.UsageError(f"--color as r,g,b needs three integers; got {value!r}") from e
        return (r, g, b)
    return value


def _fmt_snapshot(images: list[dict[str, Any]], dpi: int) -> str:
    if not images:
        return "(no pages rendered)"
    lines: list[str] = []
    for im in images:
        size = f"{im['bytes']} bytes"
        where = im.get("path") or "base64"
        lines.append(f"page {im['page']}: {size} → {where}")
    head = f"rendered {len(images)} page(s) at {dpi} dpi"
    return head + "\n" + "\n".join(lines)


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


# ---------------------------------------------------------------------------
# save | save-as PATH [--format] [--overwrite] | export-pdf PATH [--pages]
#   (gated: writes only under a --save-dir / WORDLIVE_SAVE_DIRS directory)
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# cursor read | cursor write --text "..." [--no-replace]
# ---------------------------------------------------------------------------


def _fmt_cursor(info: dict[str, Any]) -> str:
    para = info.get("paragraph")
    where = f"  in {para['anchor_id']}" if para else ""
    if info.get("collapsed"):
        return f"cursor at {info.get('start', 0)}{where}"
    sel = info.get("text") or ""
    return f"selection {info.get('start', 0)}-{info.get('end', 0)}: {sel!r}{where}"


@click.group(name="cursor")
def cursor() -> None:
    """Read or write at the user's live cursor — the explicit, non-anchored surface.

    Unlike every other command, `cursor write` deliberately moves the user's
    cursor and is *not* addressable by `--anchor-id` — that's the structural
    signal that it's the non-preferred mode. Prefer anchor-addressed edits
    (`replace`/`insert --anchor-id …`); reach for `cursor` only when the user
    genuinely wants something at their current position.
    """


@cursor.command(name="read")
@click.pass_context
def cursor_read(ctx: click.Context) -> None:
    """Report the cursor position, any selected text, and the containing paragraph."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            info = doc.selection.info()
            para = doc.paragraphs.at(info["start"])
            info["paragraph"] = {"anchor_id": para.anchor_id} if para is not None else None
            emit(info, as_text=not ctx.obj["as_json"], text=_fmt_cursor(info))

    _run(ctx, go)


@cursor.command(name="write")
@click.option("--text", "text", required=True, help="Text to insert at the cursor.")
@click.option(
    "--replace/--no-replace",
    "replace",
    default=True,
    show_default=True,
    help="Overwrite the selected text (if any). --no-replace inserts at the selection start.",
)
@click.pass_context
def cursor_write(ctx: click.Context, text: str, replace: bool) -> None:
    """Insert text at the cursor (deliberately moves the cursor; atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: write at cursor") as scope:
                scope.allow_cursor_move()
                doc.selection.write(text, replace=replace)
            emit(
                {"ok": True, "replace": replace},
                as_text=not ctx.obj["as_json"],
                text="wrote at cursor",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# find --text "..." [--in ANCHOR_ID]
# ---------------------------------------------------------------------------


@click.command(name="find")
@click.option(
    "--text", "text", required=True, help="Text to locate (whitespace + smart-quote fuzzy match)."
)
@click.option(
    "--in", "in_", default=None, help="Optional anchor id to scope the search (e.g. 'heading:3')."
)
@click.option(
    "--mode",
    "mode",
    type=click.Choice(["fuzzy", "literal", "regex"]),
    default="fuzzy",
    show_default=True,
    help="Matcher: fuzzy (Unicode/whitespace-tolerant), literal (exact), or regex (Python).",
)
@click.pass_context
def find_cmd(ctx: click.Context, text: str, in_: str | None, mode: str) -> None:
    """Locate every occurrence of TEXT (read-only)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            scope = doc.anchor_by_id(in_) if in_ else None
            matches = doc.find(text, scope=scope, mode=mode)
            emit(matches, as_text=not ctx.obj["as_json"], text=_fmt_find(matches))

    _run(ctx, go)


@click.command(name="find-paragraph")
@click.option("--text", "text", required=True, help="Approximate paragraph text to locate.")
@click.option("--limit", "limit", type=int, default=5, show_default=True, help="Max candidates.")
@click.option(
    "--min-score",
    "min_score",
    type=float,
    default=0.6,
    show_default=True,
    help="Minimum similarity score (0–1) to include.",
)
@click.pass_context
def find_paragraph_cmd(ctx: click.Context, text: str, limit: int, min_score: float) -> None:
    """Fuzzy-rank paragraphs by similarity to TEXT (typo/paraphrase tolerant, read-only)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.find_paragraphs(text, limit=limit, min_score=min_score)
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_find_paragraphs(rows))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# checkpoint / diff — structural fingerprint + content-aligned change list
# ---------------------------------------------------------------------------


def _load_checkpoint(path: Path) -> Any:
    """Read a checkpoint token from a file, mapping IO/parse errors to OpError
    (clean exit 1) rather than a traceback."""
    from .._checkpoint import Checkpoint

    try:
        return Checkpoint.from_json(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise OpError(f"cannot read checkpoint file {str(path)!r}: {exc}") from exc
    except (ValueError, KeyError) as exc:
        raise OpError(f"invalid checkpoint file {str(path)!r}: {exc}") from exc


@click.command(name="checkpoint")
@click.option(
    "--include",
    "include",
    type=click.Choice(["text", "text+style", "text+format"]),
    default="text+style",
    show_default=True,
    help="Fingerprint depth: text < text+style (restyle visible) < text+format (reformat visible).",
)
@click.option(
    "--within",
    "within",
    default=None,
    help="Fingerprint just one anchor's range (heading:N, range:S-E, table:N:R:C).",
)
@click.option(
    "--out",
    "out",
    default=None,
    type=click.Path(path_type=Path),
    help="Write the checkpoint token here. Without --out the JSON is emitted on stdout.",
)
@click.pass_context
def checkpoint_cmd(ctx: click.Context, include: str, within: str | None, out: Path | None) -> None:
    """Fingerprint the document's structure now → a checkpoint token (pure read).

    Store the token, edit the document, then `diff --since FILE` (or
    `diff --from A --to B`) for a structured change list — the way an agent
    verifies its edits landed, or sees what the user changed (Word emits no
    content-change event). Without `--out` the token is the JSON object on stdout.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            cp = doc.checkpoint(include=include, within=within)
        if out is not None:
            try:
                out.write_text(cp.to_json() + "\n", encoding="utf-8")
            except OSError as exc:
                raise OpError(f"cannot write checkpoint to {str(out)!r}: {exc}") from exc
            summary = {
                "out": str(out),
                "include": cp.include,
                "scope": cp.scope,
                "paragraphs": len(cp.paragraphs),
                "tables": len(cp.tables),
            }
            emit(summary, as_text=not ctx.obj["as_json"])
        else:
            emit(json.loads(cp.to_json()), as_text=not ctx.obj["as_json"])

    _run(ctx, go)


@click.command(name="diff")
@click.option(
    "--since",
    "since",
    default=None,
    type=click.Path(path_type=Path),
    help="Diff this stored checkpoint against the document now (needs live Word).",
)
@click.option(
    "--from",
    "frm",
    default=None,
    type=click.Path(path_type=Path),
    help="With --to: diff two stored checkpoints (no live Word needed).",
)
@click.option(
    "--to",
    "to",
    default=None,
    type=click.Path(path_type=Path),
    help="With --from: the second stored checkpoint.",
)
@click.pass_context
def diff_cmd(ctx: click.Context, since: Path | None, frm: Path | None, to: Path | None) -> None:
    """Diff a checkpoint against the document now (--since) or two checkpoints
    (--from/--to) → a content-aligned change list (pure read).

    Each change is `replace` / `insert` / `delete` / `restyle` / `reformat`,
    carrying the current `para:N` so the caller can act on it immediately.
    """
    from .._checkpoint import diff_checkpoints

    def go() -> None:
        if since is not None:
            if frm is not None or to is not None:
                raise OpError("use either --since, or --from/--to — not both")
        elif frm is None or to is None:
            raise OpError("provide --since FILE, or both --from FILE and --to FILE")
        if since is not None:
            cp = _load_checkpoint(since)
            with attach() as word:
                doc = _pick_doc(word, ctx.obj["doc_name"])
                changes = doc.changes_since(cp)
        else:
            assert frm is not None and to is not None  # narrowed above
            changes = diff_checkpoints(_load_checkpoint(frm), _load_checkpoint(to))
        emit(changes, as_text=not ctx.obj["as_json"])

    _run(ctx, go)


# ---------------------------------------------------------------------------
# replace
#   --anchor-id ID --text "..."                          (anchor mode)
#   --find OLD --text NEW [--in ID] [--all|--occurrence N]   (fuzzy mode)
# ---------------------------------------------------------------------------


@click.command(name="replace")
@click.option(
    "--anchor-id", "anchor_id", default=None, help="Replace the entire range at this anchor."
)
@click.option(
    "--find", "find", default=None, help="Fuzzy text to locate (alternative to --anchor-id)."
)
@click.option("--text", "text", required=True, help="Replacement text.")
@click.option(
    "--in", "in_", default=None, help="In fuzzy mode, scope the search to this anchor id."
)
@click.option(
    "--all", "replace_all", is_flag=True, default=False, help="In fuzzy mode, replace every match."
)
@click.option(
    "--occurrence",
    "occurrence",
    type=int,
    default=None,
    help="In find mode, replace only the Nth match (1-based).",
)
@click.option(
    "--mode",
    "mode",
    type=click.Choice(["fuzzy", "literal", "regex"]),
    default="fuzzy",
    show_default=True,
    help="In find mode: fuzzy (tolerant), literal (exact), or regex (Python; --text may use \\1).",
)
@click.pass_context
def replace(
    ctx: click.Context,
    anchor_id: str | None,
    find: str | None,
    text: str,
    in_: str | None,
    replace_all: bool,
    occurrence: int | None,
    mode: str,
) -> None:
    """Replace text. Either at an anchor (entire range) or via find."""
    if (anchor_id is None) == (find is None):
        raise click.UsageError("provide exactly one of --anchor-id or --find")
    if anchor_id is not None and (in_ or replace_all or occurrence is not None):
        raise click.UsageError("--in / --all / --occurrence are only valid with --find")
    if replace_all and occurrence is not None:
        raise click.UsageError("--all and --occurrence are mutually exclusive")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            if anchor_id is not None:
                anchor = doc.anchor_by_id(anchor_id)
                with doc.edit(f"CLI: replace {anchor_id}"):
                    anchor.set_text(text)
                emit(
                    {
                        "ok": True,
                        "anchor_id": anchor_id,
                        "anchor": {"kind": anchor.kind, "name": anchor.name},
                    },
                    as_text=not ctx.obj["as_json"],
                    text=f"replaced {anchor_id}",
                )
                return

            assert find is not None  # guaranteed by the validation above
            scope = doc.anchor_by_id(in_) if in_ else None
            try:
                with doc.edit(f"CLI: find/replace {find!r}"):
                    applied = doc.find_replace(
                        find,
                        text,
                        scope=scope,
                        all=replace_all,
                        occurrence=occurrence,
                        mode=mode,
                    )
            except AmbiguousMatchError as exc:
                emit(
                    {
                        "ok": False,
                        "error": "ambiguous_match",
                        "find": exc.find,
                        "matches": exc.matches,
                    },
                    as_text=not ctx.obj["as_json"],
                    text=f"{len(exc.matches)} matches — use --all or --occurrence N",
                )
                raise
            emit(
                {"ok": True, "replacements": applied},
                as_text=not ctx.obj["as_json"],
                text=_fmt_replace_summary(applied),
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# go-to --anchor-id ID
# ---------------------------------------------------------------------------


@click.command(name="go-to")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor ID to move the cursor to.")
@click.option(
    "--scroll/--no-scroll", default=True, help="Scroll the view to the anchor (default: yes)."
)
@click.pass_context
def go_to(ctx: click.Context, anchor_id: str, scroll: bool) -> None:
    """Move the user's cursor to the anchor (deliberate, opt-in cursor move)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            doc.go_to(anchor, scroll=scroll)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                },
                as_text=not ctx.obj["as_json"],
                text=f"moved to {anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# style list | style apply --anchor-id ID --name NAME
# ---------------------------------------------------------------------------


def _fmt_style_list(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no styles)"
    name_w = max(len(r["name"]) for r in rows)
    return "\n".join(
        f"{r['name']:<{name_w}}  {r['type']:<10}  "
        f"{'builtin' if r['builtin'] else 'custom':<8}  "
        f"{'in-use' if r['in_use'] else ''}"
        for r in rows
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


# ---------------------------------------------------------------------------
# format-paragraph --anchor-id ID [--alignment ...] [--left-indent N] ...
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# table list | read N | add-row | delete-row
# ---------------------------------------------------------------------------


def _fmt_table_list(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no tables)"
    return "\n".join(
        f"table:{r['index']}  {r['rows']}x{r['columns']}"
        + (f"  {r['title']!r}" if r.get("title") else "")
        for r in rows
    )


def _fmt_table_read(grid: dict[str, Any]) -> str:
    cells = grid.get("cells") or []
    if not cells:
        return f"table:{grid.get('index')} (empty)"
    # Rows can be ragged on a merged / split table, so size columns off the
    # widest row and guard each per-column scan against shorter rows.
    ncols = max((len(row) for row in cells), default=0)
    widths = [
        max((len(row[c]["text"]) for row in cells if c < len(row)), default=0) for c in range(ncols)
    ]
    lines = []
    for row in cells:
        lines.append("  ".join(cell["text"].ljust(widths[i]) for i, cell in enumerate(row)))
    return "\n".join(lines)


@click.group(name="table")
def table() -> None:
    """Read and edit tables (cells are anchors: table:N:R:C)."""


@table.command(name="list")
@click.pass_context
def table_list(ctx: click.Context) -> None:
    """List every table with its position, size, and title."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.tables.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_table_list(rows))

    _run(ctx, go)


@table.command(name="read")
@click.argument("index", type=int)
@click.pass_context
def table_read(ctx: click.Context, index: int) -> None:
    """Read table INDEX (1-based) as a grid of cells with anchor IDs."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            grid = doc.tables[index].read()
            emit(grid, as_text=not ctx.obj["as_json"], text=_fmt_table_read(grid))

    _run(ctx, go)


@table.command(name="add-row")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--values", "values", default=None, help="Optional JSON array of cell values for the new row."
)
@click.pass_context
def table_add_row(ctx: click.Context, table_index: int, values: str | None) -> None:
    """Append a row to the table (atomic-undo)."""
    parsed: list[Any] | None = None
    if values is not None:
        try:
            parsed = json.loads(values)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--values must be a JSON array: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--values must be a JSON array")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: add row to table {table_index}"):
                t.add_row(parsed)
            emit(
                {"ok": True, "table": table_index, "rows": t.row_count},
                as_text=not ctx.obj["as_json"],
                text=f"added row to table:{table_index} (now {t.row_count} rows)",
            )

    _run(ctx, go)


@table.command(name="delete-row")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--row", "row", type=int, required=True, help="1-based row to delete.")
@click.pass_context
def table_delete_row(ctx: click.Context, table_index: int, row: int) -> None:
    """Delete a row from the table (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: delete row {row} from table {table_index}"):
                t.delete_row(row)
            emit(
                {"ok": True, "table": table_index, "rows": t.row_count},
                as_text=not ctx.obj["as_json"],
                text=f"deleted row {row} from table:{table_index} (now {t.row_count} rows)",
            )

    _run(ctx, go)


@table.command(name="add-column")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--values",
    "values",
    default=None,
    help="Optional JSON array of cell values for the new column (top-to-bottom).",
)
@click.pass_context
def table_add_column(ctx: click.Context, table_index: int, values: str | None) -> None:
    """Append a column to the table (atomic-undo)."""
    parsed: list[Any] | None = None
    if values is not None:
        try:
            parsed = json.loads(values)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--values must be a JSON array: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--values must be a JSON array")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: add column to table {table_index}"):
                t.add_column(parsed)
            emit(
                {"ok": True, "table": table_index, "columns": t.column_count},
                as_text=not ctx.obj["as_json"],
                text=f"added column to table:{table_index} (now {t.column_count} columns)",
            )

    _run(ctx, go)


@table.command(name="delete-column")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--column", "column", type=int, required=True, help="1-based column to delete.")
@click.pass_context
def table_delete_column(ctx: click.Context, table_index: int, column: int) -> None:
    """Delete a column from the table (atomic-undo).

    Fails with an OpError on a table with merged / mixed-width cells — Word can't
    address an individual column there; delete its cells via table:N:R:C instead.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: delete column {column} from table {table_index}"):
                t.delete_column(column)
            emit(
                {"ok": True, "table": table_index, "columns": t.column_count},
                as_text=not ctx.obj["as_json"],
                text=f"deleted column {column} from table:{table_index} (now {t.column_count} columns)",
            )

    _run(ctx, go)


@table.command(name="merge-cells")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--from", "from_cell", required=True, help='Anchor cell "R:C" (1-based).')
@click.option("--to", "to_cell", required=True, help='Opposite cell "R:C" (1-based).')
@click.pass_context
def table_merge_cells(ctx: click.Context, table_index: int, from_cell: str, to_cell: str) -> None:
    """Merge two cells (and the rectangle they span) into one (atomic-undo)."""
    fr, fc = _parse_rc(from_cell)
    tr, tc = _parse_rc(to_cell)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: merge cells in table {table_index}"):
                t.cell(fr, fc).merge(t.cell(tr, tc))
            emit(
                {"ok": True, "table": table_index, "anchor_id": f"table:{table_index}:{fr}:{fc}"},
                as_text=not ctx.obj["as_json"],
                text=f"merged into table:{table_index}:{fr}:{fc} (table is now non-uniform)",
            )

    _run(ctx, go)


@table.command(name="split-cell")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--cell", "cell", required=True, help='Cell to split, "R:C" (1-based).')
@click.option("--rows", "rows", type=int, default=1, show_default=True, help="Rows to split into.")
@click.option(
    "--columns", "columns", type=int, default=2, show_default=True, help="Columns to split into."
)
@click.pass_context
def table_split_cell(
    ctx: click.Context, table_index: int, cell: str, rows: int, columns: int
) -> None:
    """Split one cell into a rows x columns grid (atomic-undo)."""
    cr, cc = _parse_rc(cell)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: split cell in table {table_index}"):
                t.cell(cr, cc).split(rows, columns)
            emit(
                {"ok": True, "table": table_index, "anchor_id": f"table:{table_index}:{cr}:{cc}"},
                as_text=not ctx.obj["as_json"],
                text=f"split table:{table_index}:{cr}:{cc} into {rows}x{columns} (table is now non-uniform)",
            )

    _run(ctx, go)


@table.command(name="set-heading-row")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--row", "row", type=int, default=1, show_default=True, help="1-based row.")
@click.option(
    "--heading/--no-heading",
    "heading",
    default=True,
    show_default=True,
    help="Make the row a repeating table heading (repeats on every page).",
)
@click.option(
    "--allow-break/--no-allow-break",
    "allow_break",
    default=None,
    help="Allow the row to split across a page (default: off for a heading row).",
)
@click.pass_context
def table_set_heading_row(
    ctx: click.Context, table_index: int, row: int, heading: bool, allow_break: bool | None
) -> None:
    """Mark a row as a repeating table heading row (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set heading row {row} on table {table_index}"):
                t.set_heading_row(row, heading=heading, allow_break=allow_break)
            emit(
                {"ok": True, "table": table_index, "row": row, "heading": heading},
                as_text=not ctx.obj["as_json"],
                text=(
                    f"{'set' if heading else 'cleared'} heading row {row} on table:{table_index}"
                ),
            )

    _run(ctx, go)


@table.command(name="records")
@click.argument("index", type=int)
@click.pass_context
def table_records(ctx: click.Context, index: int) -> None:
    """Read table INDEX as records — body rows as dicts keyed by the header row.

    The read mirror of `table create --data` from records: row 1 is the header,
    each row below becomes a `{header: value}` object. Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            recs = doc.tables[index].records()
            emit(recs, as_text=not ctx.obj["as_json"])

    _run(ctx, go)


@table.command(name="append-record")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--record", "record", required=True, help="JSON object mapping header names to cell values."
)
@click.pass_context
def table_append_record(ctx: click.Context, table_index: int, record: str) -> None:
    """Append a row from a JSON record, mapping keys to header columns (atomic-undo)."""
    try:
        parsed = json.loads(record)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--record must be a JSON object: {e}") from e
    if not isinstance(parsed, dict):
        raise click.UsageError("--record must be a JSON object")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: append record to table {table_index}"):
                t.append_record(parsed)
            emit(
                {"ok": True, "table": table_index, "rows": t.row_count},
                as_text=not ctx.obj["as_json"],
                text=f"appended record to table:{table_index} (now {t.row_count} rows)",
            )

    _run(ctx, go)


@table.command(name="update-row")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--key", "key", required=True, help="Value to match in the key column.")
@click.option(
    "--values", "values", required=True, help="JSON object of {header: new_value} cells to set."
)
@click.option(
    "--column",
    "column",
    default=None,
    help="Header name of the key column to match (default: first column).",
)
@click.pass_context
def table_update_row(
    ctx: click.Context, table_index: int, key: str, values: str, column: str | None
) -> None:
    """Update the first row whose key-column cell equals --key, by header (atomic-undo)."""
    try:
        parsed = json.loads(values)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--values must be a JSON object: {e}") from e
    if not isinstance(parsed, dict):
        raise click.UsageError("--values must be a JSON object")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: update row in table {table_index}"):
                t.update_row(key, parsed, column=column)
            emit(
                {"ok": True, "table": table_index, "key": key},
                as_text=not ctx.obj["as_json"],
                text=f"updated row {key!r} in table:{table_index}",
            )

    _run(ctx, go)


@table.command(name="create")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Position anchor for the new table (heading:/para:/start/end/range:…).",
)
@click.option(
    "--rows",
    "rows",
    type=int,
    default=None,
    help="Number of rows (>= 1). Optional when --data is given (inferred from it).",
)
@click.option(
    "--cols",
    "cols",
    type=int,
    default=None,
    help="Number of columns (>= 1). Optional when --data is given (inferred from it).",
)
@click.option(
    "--style",
    "style",
    default=None,
    help="Table style name (default: the built-in 'Table Grid', so borders show).",
)
@click.option(
    "--header/--no-header",
    "header",
    default=False,
    show_default=True,
    help="Bold the first row as a header.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.option(
    "--data",
    "data",
    default=None,
    help="JSON to populate cells, or '-' to read it from stdin: a row-major 2-D "
    'array (\'[["Name","Qty"],["Widget","3"]]\') OR records — a list of objects '
    '(\'[{"Name":"Widget","Qty":"3"}]\'), whose keys become a header row. '
    "Reading from stdin avoids quoting/backslash fights on Windows.",
)
@click.pass_context
def table_create(
    ctx: click.Context,
    anchor_id: str,
    rows: int | None,
    cols: int | None,
    style: str | None,
    header: bool,
    before: bool,
    data: str | None,
) -> None:
    """Create a table at an anchor (atomic-undo).

    Builds new table structure where wordlive's other verbs only edit existing
    structure. Fill cells at creation with --data — a row-major JSON array, or
    records (a list of objects whose keys become a header row), or '--data -' to
    read it from stdin; a short array leaves trailing cells empty. --rows/--cols
    are optional when --data is given (inferred from its shape), required
    otherwise. --style defaults to 'Table Grid' (visible borders); a style name
    not defined in the document raises (exit 2). Reports the new table's 1-based
    index for a follow-up `table set-cell` / `add-row`.
    """
    parsed: list[Any] | None = None
    if data is not None:
        raw = click.get_text_stream("stdin").read() if data == "-" else data
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--data must be JSON (a 2-D array or records): {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--data must be a JSON array of rows or records")
    if data is None and (rows is None or cols is None):
        raise click.UsageError("--rows and --cols are required when --data is not given")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: create table at {anchor_id}"):
                t = anchor.insert_table(
                    rows,
                    cols,
                    where=("before" if before else "after"),
                    style=style,
                    data=parsed,
                    header=header,
                )
            emit(
                {"ok": True, "table": t.index, "rows": t.row_count, "columns": t.column_count},
                as_text=not ctx.obj["as_json"],
                text=f"created table:{t.index} ({t.row_count}x{t.column_count}) at {anchor_id}",
            )

    _run(ctx, go)


@table.command(name="delete")
@click.argument("index", type=int)
@click.pass_context
def table_delete(ctx: click.Context, index: int) -> None:
    """Delete table INDEX (1-based) and all its cells (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[index]  # AnchorNotFoundError (exit 2) if missing
            with doc.edit(f"CLI: delete table {index}"):
                t.delete()
            emit(
                {"ok": True, "deleted": index},
                as_text=not ctx.obj["as_json"],
                text=f"deleted table:{index}",
            )

    _run(ctx, go)


@table.command(name="autofit")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--mode",
    type=click.Choice(["content", "window", "fixed"]),
    default="content",
    show_default=True,
    help="content: fit columns to cells · window: stretch to page · fixed: pin widths.",
)
@click.pass_context
def table_autofit(ctx: click.Context, table_index: int, mode: str) -> None:
    """Resize a table's columns — fit to content/window, or pin them (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: autofit table {table_index}"):
                t.autofit(mode)
            emit(
                {"ok": True, "table": table_index, "mode": mode},
                as_text=not ctx.obj["as_json"],
                text=f"autofit table:{table_index} ({mode})",
            )

    _run(ctx, go)


@table.command(name="set-style")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--style", "style", required=True, help="Table style name (e.g. 'Grid Table 4 - Accent 1')."
)
@click.pass_context
def table_set_style(ctx: click.Context, table_index: int, style: str) -> None:
    """Restyle an existing table (restyle first, then layer cell overrides; atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set-style table {table_index}"):
                t.set_style(style)
            emit(
                {"ok": True, "table": table_index, "style": style},
                as_text=not ctx.obj["as_json"],
                text=f"styled table:{table_index} ({style})",
            )

    _run(ctx, go)


@table.command(name="set-alignment")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--alignment",
    "alignment",
    type=click.Choice(["left", "center", "right"]),
    required=True,
    help="Align the whole table across the page width.",
)
@click.pass_context
def table_set_alignment(ctx: click.Context, table_index: int, alignment: str) -> None:
    """Align a whole table left/center/right across the page (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set-alignment table {table_index}"):
                t.set_alignment(alignment)
            emit(
                {"ok": True, "table": table_index, "alignment": alignment},
                as_text=not ctx.obj["as_json"],
                text=f"aligned table:{table_index} ({alignment})",
            )

    _run(ctx, go)


@table.command(name="set-borders")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--sides",
    "sides",
    default="all",
    help="Edges: all/box, top, bottom, left, right, horizontal, vertical "
    "(comma-separated for several; interior gridlines need horizontal/vertical).",
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
def table_set_borders(
    ctx: click.Context,
    table_index: int,
    sides: str,
    style: str,
    weight: float,
    color: str | None,
) -> None:
    """Draw borders across a whole table grid in one call (atomic-undo)."""
    side_list = [s.strip() for s in sides.split(",") if s.strip()]
    color_value = _parse_color(color)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set-borders table {table_index}"):
                t.set_borders(sides=side_list, style=style, weight=weight, color=color_value)
            emit(
                {
                    "ok": True,
                    "table": table_index,
                    "applied": {
                        "sides": side_list,
                        "style": style,
                        "weight": weight,
                        "color": color_value,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"bordered table:{table_index}: {side_list} {style}",
            )

    _run(ctx, go)


@table.command(name="set-banding")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--first-row/--no-first-row", "first_row", default=None, help="Header-row banding.")
@click.option("--last-row/--no-last-row", "last_row", default=None, help="Total-row banding.")
@click.option(
    "--first-column/--no-first-column", "first_column", default=None, help="First-column banding."
)
@click.option(
    "--last-column/--no-last-column", "last_column", default=None, help="Last-column banding."
)
@click.option(
    "--banded-rows/--no-banded-rows", "banded_rows", default=None, help="Alternating row stripes."
)
@click.option(
    "--banded-columns/--no-banded-columns",
    "banded_columns",
    default=None,
    help="Alternating column stripes.",
)
@click.pass_context
def table_set_banding(
    ctx: click.Context,
    table_index: int,
    first_row: bool | None,
    last_row: bool | None,
    first_column: bool | None,
    last_column: bool | None,
    banded_rows: bool | None,
    banded_columns: bool | None,
) -> None:
    """Toggle a table's style options / banding (needs a real table style applied; atomic-undo)."""
    flags = {
        "first_row": first_row,
        "last_row": last_row,
        "first_column": first_column,
        "last_column": last_column,
        "banded_rows": banded_rows,
        "banded_columns": banded_columns,
    }
    applied = {k: v for k, v in flags.items() if v is not None}

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set-banding table {table_index}"):
                t.set_banding(**applied)
            emit(
                {"ok": True, "table": table_index, "applied": applied},
                as_text=not ctx.obj["as_json"],
                text=f"banding table:{table_index}: {applied or '(no flags)'}",
            )

    _run(ctx, go)


@click.command(name="cell-valign")
@click.option("--anchor-id", "anchor_id", required=True, help="Cell anchor (table:N:R:C) to align.")
@click.option(
    "--align",
    "align",
    type=click.Choice(["top", "center", "bottom"]),
    required=True,
    help="Where the cell's content sits vertically.",
)
@click.pass_context
def cell_valign_cmd(ctx: click.Context, anchor_id: str, align: str) -> None:
    """Set a table cell's vertical alignment — top, center, or bottom (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            from .._tables import Cell

            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            if not isinstance(anchor, Cell):
                raise OpError(
                    f"cell-valign needs a cell anchor (table:N:R:C); "
                    f"{anchor_id!r} resolved to {anchor.kind}"
                )
            with doc.edit(f"CLI: cell-valign {anchor_id}"):
                anchor.set_vertical_alignment(align)
            emit(
                {"ok": True, "anchor_id": anchor_id, "align": align},
                as_text=not ctx.obj["as_json"],
                text=f"valigned {anchor_id}: {align}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# comment list | add | resolve | delete
# ---------------------------------------------------------------------------


def _fmt_comment_list(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no comments)"
    lines: list[str] = []
    for r in rows:
        author = r.get("author") or "?"
        state = "resolved" if r.get("done") else "open"
        scope = r.get("scope") or ""
        on = f"  on {scope!r}" if scope else ""
        lines.append(f"[{r['index']}] {author} ({state}): {r.get('text', '')}{on}")
    return "\n".join(lines)


@click.group(name="comment")
def comment() -> None:
    """Add, list, resolve, and delete review comments."""


@comment.command(name="list")
@click.pass_context
def comment_list(ctx: click.Context) -> None:
    """List every comment with its index, author, body, and scope."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.comments.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_comment_list(rows))

    _run(ctx, go)


@comment.command(name="add")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to attach the comment to.")
@click.option("--text", "text", required=True, help="Comment body.")
@click.option("--author", "author", default=None, help="Optional comment author name.")
@click.pass_context
def comment_add(ctx: click.Context, anchor_id: str, text: str, author: str | None) -> None:
    """Attach a comment to the anchor's range — the document text is untouched."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: comment on {anchor_id}"):
                c = doc.comments.add(anchor, text, author=author)
                index, c_author = c.index, c.author
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "comment": {"index": index, "author": c_author},
                },
                as_text=not ctx.obj["as_json"],
                text=f"added comment {index} on {anchor_id}",
            )

    _run(ctx, go)


@comment.command(name="resolve")
@click.option(
    "--index", "index", type=int, required=True, help="1-based comment index (see `comment list`)."
)
@click.pass_context
def comment_resolve(ctx: click.Context, index: int) -> None:
    """Mark comment INDEX as resolved/done (Word 2013+)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            c = doc.comments[index]
            with doc.edit(f"CLI: resolve comment {index}"):
                c.resolve()
            emit(
                {"ok": True, "index": index, "done": True},
                as_text=not ctx.obj["as_json"],
                text=f"resolved comment {index}",
            )

    _run(ctx, go)


@comment.command(name="delete")
@click.option(
    "--index", "index", type=int, required=True, help="1-based comment index (see `comment list`)."
)
@click.pass_context
def comment_delete(ctx: click.Context, index: int) -> None:
    """Delete comment INDEX."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            c = doc.comments[index]
            with doc.edit(f"CLI: delete comment {index}"):
                c.delete()
            emit(
                {"ok": True, "index": index, "deleted": True},
                as_text=not ctx.obj["as_json"],
                text=f"deleted comment {index}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# track status | on | off
# ---------------------------------------------------------------------------


@click.group(name="track")
def track() -> None:
    """Inspect or toggle the document's Track Changes setting.

    The toggle is *persistent* — `track on` leaves Word recording revisions
    until `track off`. For a self-restoring scope, prefer the library's
    `doc.tracked_changes()` or `exec`'s `"tracked": true` payload key.
    """


@track.command(name="status")
@click.pass_context
def track_status(ctx: click.Context) -> None:
    """Report whether Track Changes is currently on."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            state = doc.track_changes
            emit(
                {"tracked": state},
                as_text=not ctx.obj["as_json"],
                text=f"track changes: {'on' if state else 'off'}",
            )

    _run(ctx, go)


@track.command(name="on")
@click.pass_context
def track_on(ctx: click.Context) -> None:
    """Turn Track Changes on (persists until `track off`)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            doc.track_changes = True
            emit(
                {"ok": True, "tracked": True},
                as_text=not ctx.obj["as_json"],
                text="track changes: on",
            )

    _run(ctx, go)


@track.command(name="off")
@click.pass_context
def track_off(ctx: click.Context) -> None:
    """Turn Track Changes off."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            doc.track_changes = False
            emit(
                {"ok": True, "tracked": False},
                as_text=not ctx.obj["as_json"],
                text="track changes: off",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# revision list | accept | reject | accept-all | reject-all
# (the top-level `revisions` command above is an alias for `revision list`)
# ---------------------------------------------------------------------------


@click.group(name="revision")
def revision() -> None:
    """List tracked changes and accept / reject them (the write side of `revisions`)."""


@revision.command(name="list")
@click.pass_context
def revision_list(ctx: click.Context) -> None:
    """List the document's tracked changes (alias of the top-level `revisions`)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.revisions.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_revisions(rows))

    _run(ctx, go)


@revision.command(name="accept")
@click.option(
    "--index", "index", type=int, required=True, help="1-based revision index (see `revisions`)."
)
@click.pass_context
def revision_accept(ctx: click.Context, index: int) -> None:
    """Accept revision INDEX — make that tracked change permanent."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rev = doc.revisions[index]
            with doc.edit(f"CLI: accept revision {index}"):
                rev.accept()
            emit(
                {"ok": True, "index": index, "accepted": True},
                as_text=not ctx.obj["as_json"],
                text=f"accepted revision {index}",
            )

    _run(ctx, go)


@revision.command(name="reject")
@click.option(
    "--index", "index", type=int, required=True, help="1-based revision index (see `revisions`)."
)
@click.pass_context
def revision_reject(ctx: click.Context, index: int) -> None:
    """Reject revision INDEX — undo that tracked change."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rev = doc.revisions[index]
            with doc.edit(f"CLI: reject revision {index}"):
                rev.reject()
            emit(
                {"ok": True, "index": index, "rejected": True},
                as_text=not ctx.obj["as_json"],
                text=f"rejected revision {index}",
            )

    _run(ctx, go)


@revision.command(name="accept-all")
@click.option(
    "--anchor-id",
    "anchor_id",
    default=None,
    help="Scope to one anchor's range (e.g. 'heading:3'); whole document if omitted.",
)
@click.pass_context
def revision_accept_all(ctx: click.Context, anchor_id: str | None) -> None:
    """Accept every tracked change (optionally only inside --anchor-id)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            within = doc.anchor_by_id(anchor_id) if anchor_id else None
            with doc.edit("CLI: accept all revisions"):
                n = doc.revisions.accept_all(within=within)
            emit(
                {"ok": True, "accepted": n, "anchor_id": anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"accepted {n} revision(s)",
            )

    _run(ctx, go)


@revision.command(name="reject-all")
@click.option(
    "--anchor-id",
    "anchor_id",
    default=None,
    help="Scope to one anchor's range (e.g. 'heading:3'); whole document if omitted.",
)
@click.pass_context
def revision_reject_all(ctx: click.Context, anchor_id: str | None) -> None:
    """Reject every tracked change (optionally only inside --anchor-id)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            within = doc.anchor_by_id(anchor_id) if anchor_id else None
            with doc.edit("CLI: reject all revisions"):
                n = doc.revisions.reject_all(within=within)
            emit(
                {"ok": True, "rejected": n, "anchor_id": anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"rejected {n} revision(s)",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# watermark / insert-text-box  (floating-shape publishing flourishes)
# ---------------------------------------------------------------------------


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


@click.command(name="insert-text-box")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to attach the text box to.")
@click.option("--text", "text", required=True, help="Text box body.")
@click.option("--width", "width", default="200", show_default=True, help="Width (pt or '3in').")
@click.option("--height", "height", default="100", show_default=True, help="Height (pt or '2cm').")
@click.option(
    "--wrap",
    "wrap",
    type=click.Choice(["square", "tight", "through", "top-bottom", "front", "behind"]),
    default="square",
    show_default=True,
    help="How body text flows around the box.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Anchor before the anchor instead of after it.",
)
@click.option("--font", "font", default=None, help="Text font.")
@click.option("--size", "size", default=None, help="Font size (pt or unit string).")
@click.option("--bold/--no-bold", "bold", default=None, help="Bold the text.")
@click.option("--italic/--no-italic", "italic", default=None, help="Italicise the text.")
@click.option(
    "--align",
    "alignment",
    type=click.Choice(["left", "center", "right", "justify"]),
    default=None,
    help="Paragraph alignment.",
)
@click.option("--fill", "fill", default=None, help="Background colour (e.g. '#eeeeff' / 'navy').")
@click.option("--border-color", "border_color", default=None, help="Outline colour.")
@click.option("--no-border", "no_border", is_flag=True, default=False, help="No outline.")
@click.pass_context
def insert_text_box_cmd(
    ctx: click.Context,
    anchor_id: str,
    text: str,
    width: str,
    height: str,
    wrap: str,
    before: bool,
    font: str | None,
    size: str | None,
    bold: bool | None,
    italic: bool | None,
    alignment: str | None,
    fill: str | None,
    border_color: str | None,
    no_border: bool,
) -> None:
    """Insert a floating text box / pull quote at an anchor (atomic-undo)."""
    if no_border and border_color is not None:
        raise click.UsageError("pass either --no-border or --border-color (not both)")
    border: str | bool | None = False if no_border else border_color
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert text box {where} {anchor_id}"):
                shape = anchor.insert_text_box(
                    text,
                    width=width,
                    height=height,
                    wrap=wrap,
                    where=where,
                    font=font,
                    size=size,
                    bold=bold,
                    italic=italic,
                    alignment=alignment,
                    fill=fill,
                    border=border,
                )
            emit(
                {"ok": True, "anchor_id": shape.anchor_id, "wrap": wrap},
                as_text=not ctx.obj["as_json"],
                text=f"inserted text box {where} {anchor_id} -> {shape.anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# list show | apply | remove | info | restart | indent | outdent
# ---------------------------------------------------------------------------


def _fmt_list_show(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no lists)"
    return "\n".join(
        f"list:{r['index']}  {r['type']}  "
        f"{r['count']} item{'s' if r['count'] != 1 else ''}  [{r['anchor_id']}]"
        for r in rows
    )


def _fmt_list_info(info: dict[str, Any]) -> str:
    if info.get("type") == "none":
        return "not in a list"
    return (
        f"{info['type']} (level {info['level']}, number {info['number']}, "
        f"marker {info['string']!r})"
    )


def _fmt_list_levels(levels: list[dict[str, Any]]) -> str:
    if not levels:
        return "not in a list"
    return "\n".join(
        f"L{lv['level']}: {lv['kind']} {lv['format']!r}"
        + (f" ({lv['style']})" if lv["kind"] == "number" else f" font {lv['font']!r}")
        + f"  trailing={lv['trailing']}  num@{lv['number_position']:g}pt text@{lv['text_position']:g}pt"
        for lv in levels
    )


@click.group(name="list")
def list_cmd() -> None:
    """Apply, inspect, and manage bullet / numbered lists."""


@list_cmd.command(name="show")
@click.pass_context
def list_show(ctx: click.Context) -> None:
    """List every bullet/numbered list in the document, with its range anchor id."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.lists.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_list_show(rows))

    _run(ctx, go)


@list_cmd.command(name="apply")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor whose paragraphs to format as a list."
)
@click.option(
    "--type",
    "list_type",
    type=click.Choice(["bulleted", "numbered", "outline"], case_sensitive=False),
    default="bulleted",
    show_default=True,
    help="List style to apply.",
)
@click.option(
    "--continue",
    "continue_previous",
    is_flag=True,
    default=False,
    help="Continue numbering from the previous list instead of starting at 1.",
)
@click.pass_context
def list_apply(ctx: click.Context, anchor_id: str, list_type: str, continue_previous: bool) -> None:
    """Turn the anchor's paragraphs into a bulleted/numbered/outline list (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: apply {list_type} list to {anchor_id}"):
                anchor.apply_list(list_type, continue_previous=continue_previous)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "type": list_type,
                    "continue_previous": continue_previous,
                },
                as_text=not ctx.obj["as_json"],
                text=f"applied {list_type} list to {anchor_id}",
            )

    _run(ctx, go)


@list_cmd.command(name="remove")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor whose list formatting to strip."
)
@click.pass_context
def list_remove(ctx: click.Context, anchor_id: str) -> None:
    """Strip list formatting (bullets / numbers) from the anchor's paragraphs (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: remove list from {anchor_id}"):
                anchor.remove_list()
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                },
                as_text=not ctx.obj["as_json"],
                text=f"removed list from {anchor_id}",
            )

    _run(ctx, go)


@list_cmd.command(name="info")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to inspect.")
@click.pass_context
def list_info_cmd(ctx: click.Context, anchor_id: str) -> None:
    """Report the list type / level / number for the anchor (read-only)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            info = doc.anchor_by_id(anchor_id).list_info()
            emit(info, as_text=not ctx.obj["as_json"], text=_fmt_list_info(info))

    _run(ctx, go)


@list_cmd.command(name="format")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor whose paragraphs to format as a list."
)
@click.option(
    "--levels",
    "levels",
    required=True,
    help='JSON array of per-level specs, e.g. \'[{"kind":"number","format":"%1)","style":"lower-letter"}]\'.',
)
@click.option(
    "--continue",
    "continue_previous",
    is_flag=True,
    default=False,
    help="Continue numbering from the previous list instead of starting at 1.",
)
@click.pass_context
def list_format(ctx: click.Context, anchor_id: str, levels: str, continue_previous: bool) -> None:
    """Author a custom multi-level list (per-level number/bullet format) and apply it (atomic-undo)."""
    try:
        parsed = json.loads(levels)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--levels must be a JSON array: {e}") from e
    if not isinstance(parsed, list) or not parsed:
        raise click.UsageError("--levels must be a non-empty JSON array of level objects")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: apply custom list format to {anchor_id}"):
                anchor.apply_list_format(parsed, continue_previous=continue_previous)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "levels": len(parsed),
                },
                as_text=not ctx.obj["as_json"],
                text=f"applied custom {len(parsed)}-level list format to {anchor_id}",
            )

    _run(ctx, go)


@list_cmd.command(name="levels")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor inside the list to inspect.")
@click.pass_context
def list_levels_cmd(ctx: click.Context, anchor_id: str) -> None:
    """Report the per-level format of the list at the anchor (read-only)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            levels = doc.anchor_by_id(anchor_id).read_list_levels()
            emit(
                {"anchor_id": anchor_id, "levels": levels},
                as_text=not ctx.obj["as_json"],
                text=_fmt_list_levels(levels),
            )

    _run(ctx, go)


@list_cmd.command(name="restart")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor inside the list to restart.")
@click.pass_context
def list_restart(ctx: click.Context, anchor_id: str) -> None:
    """Restart the list's numbering at 1 (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: restart numbering at {anchor_id}"):
                anchor.restart_numbering()
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                },
                as_text=not ctx.obj["as_json"],
                text=f"restarted numbering at {anchor_id}",
            )

    _run(ctx, go)


@list_cmd.command(name="indent")
@click.option("--anchor-id", "anchor_id", required=True, help="List item to demote one level.")
@click.pass_context
def list_indent(ctx: click.Context, anchor_id: str) -> None:
    """Demote the list item(s) one level (e.g. level 1 -> 2; atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: indent list {anchor_id}"):
                anchor.indent_list()
            emit(
                {"ok": True, "anchor_id": anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"indented {anchor_id}",
            )

    _run(ctx, go)


@list_cmd.command(name="outdent")
@click.option("--anchor-id", "anchor_id", required=True, help="List item to promote one level.")
@click.pass_context
def list_outdent(ctx: click.Context, anchor_id: str) -> None:
    """Promote the list item(s) one level (e.g. level 2 -> 1; atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: outdent list {anchor_id}"):
                anchor.outdent_list()
            emit(
                {"ok": True, "anchor_id": anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"outdented {anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# section list  |  header read|write  |  footer read|write
# ---------------------------------------------------------------------------


def _fmt_section_list(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no sections)"
    lines: list[str] = []
    for r in rows:
        ps = r.get("page_setup", {})
        lines.append(
            f"section:{r['index']}  {ps.get('orientation', '?')}  "
            f"{ps.get('page_width', 0):.0f}x{ps.get('page_height', 0):.0f}pt"
        )
    return "\n".join(lines)


def _emit_section_list(ctx: click.Context) -> None:
    with attach() as word:
        doc = _pick_doc(word, ctx.obj["doc_name"])
        rows = doc.sections.list()
        emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_section_list(rows))


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


_WHICH_OPTION = click.option(
    "--which",
    "which",
    type=click.Choice(["primary", "first", "even"], case_sensitive=False),
    default="primary",
    show_default=True,
    help="Which header/footer: primary, first-page, or even-pages.",
)
_SECTION_OPTION = click.option(
    "--section",
    "section_index",
    type=int,
    default=1,
    show_default=True,
    help="1-based section index.",
)


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


# ---------------------------------------------------------------------------
# exec --script ops.json
# ---------------------------------------------------------------------------


# The batch-op core (`_OP_REQUIRED_FIELDS`, `_validate_op`, `_apply_op`,
# `_op_before`, `_run_batch`) now lives in `wordlive._ops` so the MCP server can
# reuse it without importing Click. The names are re-imported above for any
# existing callers; `exec` below drives the batch through `_run_batch`.


@click.command(name="exec")
@click.option(
    "--script",
    "script",
    default=None,
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help="Path to an ops JSON file.",
)
@click.option(
    "--ops",
    "ops_inline",
    default=None,
    help="Inline JSON batch — the same content as a --script file, passed "
    'directly. Accepts the full {"label", "ops", …} object or a bare '
    "[…] ops array, or '-' to read JSON from stdin. Alternative to --script.",
)
@click.pass_context
def exec_(ctx: click.Context, script: Path | None, ops_inline: str | None) -> None:
    """Apply a batch of ops in a single atomic-undo scope.

    Provide the batch either as a file (`--script ops.json`) or inline
    (`--ops '{"ops": [...]}'`, or `--ops -` to read JSON from stdin — best for
    large payloads such as base64 images, which can exceed the command-line
    length limit). The script shape is
    `{"label": "…", "ops": [{"op": "...", ...}, ...]}`; a bare `[...]` array is
    accepted as shorthand for `{"ops": [...]}`. Set `"tracked": true` at the top
    level to record the whole batch as Word revisions (Track Changes is restored
    to its prior state afterwards).
    Supported ops: write_bookmark, write_cc, insert_paragraph, insert_block,
    insert_section, insert_markdown, replace_section,
    delete_paragraph, append, append_inline, prepend, prepend_inline,
    insert_image, insert_equation, insert_chart, format_chart, format_axis, add_trendline,
    set_series_color, format_series, add_error_bars,
    set_shape_wrap, set_shape_crop, set_shape_position, set_shape_size,
    format_shape, set_shape_alt_text, set_shape_text, set_shape_rotation, set_shape_z_order,
    set_shape_text_frame, replace_shape_image, delete_shape, group_shapes, ungroup_shape,
    set_image_alt_text, set_image_size, set_image_crop,
    replace, find_replace, apply_style, format_paragraph,
    format_run, set_shading, set_borders, drop_cap, add_tab_stop, add_style, set_style,
    insert_field, set_page_setup, update_fields, regularize, insert_footnote, insert_endnote,
    insert_toc, add_bookmark, pin, pin_outline, add_hyperlink, set_hyperlink,
    insert_cross_reference,
    insert_caption, create_content_control, set_cc_properties, set_cc_items,
    mark_index_entry, insert_index,
    insert_table_of_figures, set_bibliography_style, add_source, insert_citation,
    insert_bibliography, mark_citation, insert_table_of_authorities,
    apply_theme, set_theme_colors, set_theme_fonts,
    set_cell, add_row, append_record, update_row, delete_row,
    add_column, delete_column, merge_cells, split_cell,
    set_heading_row, autofit_table, create_table, delete_table,
    set_table_style, set_table_alignment, set_table_borders, set_table_banding,
    set_cell_vertical_alignment,
    set_property, delete_property, set_variable, delete_variable,
    insert_break, add_comment,
    resolve_comment, delete_comment,
    accept_revision, reject_revision, accept_all_revisions, reject_all_revisions,
    set_watermark, remove_watermark, insert_text_box,
    apply_list, apply_list_format, remove_list, restart_numbering, indent_list, outdent_list,
    write_header, write_footer. (append/prepend add a new paragraph + optional
    style; append_inline/prepend_inline continue the adjacent paragraph, text
    only. append_paragraph/prepend_paragraph remain as synonyms.) A field an op
    doesn't use is reported in the result's `warnings`, not silently dropped.

    Durable handles: `bind: "slug"` (or `true`) on insert/insert_block/
    insert_section/insert_markdown/create_table mints a `pin:` handle on the new
    content and returns it in that op's `outputs` entry. An op field of the exact
    form `$ops[N].field` is replaced with an earlier op's output before the op
    runs — e.g. create a table at op 0, then `{"op": "set_cell", "table":
    "$ops[0].table", ...}`.
    See docs/cli.md for each op's required and optional fields.
    """
    if (script is None) == (ops_inline is None):
        raise click.UsageError("provide exactly one of --script or --ops")

    def go() -> None:
        if ops_inline is not None:
            raw = click.get_text_stream("stdin").read() if ops_inline == "-" else ops_inline
            source = "inline"
        else:
            assert script is not None  # guaranteed by the validation above
            raw = script.read_text(encoding="utf-8")
            source = script.name
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.ClickException(f"ops JSON is malformed: {e}") from e
        if isinstance(payload, list):
            payload = {"ops": payload}
        if not isinstance(payload, dict):
            raise click.ClickException(
                'ops JSON must be an object {"ops": [...]} or an array of ops'
            )
        label = str(payload.get("label") or f"CLI: exec {source}")
        tracked = bool(payload.get("tracked", False))
        ops = payload.get("ops") or []
        if not isinstance(ops, list):
            raise click.ClickException("'ops' must be a list")
        # Vet image-source paths before any COM/filesystem access (a UNC path's
        # own existence probe would authenticate to a remote SMB server).
        ctx.obj["policy"].screen_op_image_paths(ops)

        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            result, failure_exc = _run_batch(doc, ops, label=label, tracked=tracked)
            if failure_exc is None:
                emit(
                    result,
                    as_text=not ctx.obj["as_json"],
                    text=f"applied {result['ops_run']} op(s): {label!r}",
                )
            else:
                failure = result["failure"]
                emit(
                    result,
                    as_text=not ctx.obj["as_json"],
                    text=f"failed at op {failure['index']}: {failure['error']}",
                )
                # Re-raise the original so _run() maps it to the right exit code
                # (e.g. anchor-not-found → 2, busy → 3, ambiguous → 5).
                raise failure_exc

    _run(ctx, go)


# ---------------------------------------------------------------------------
# llm-help / install-skill / install-mcp  (offline — no Word needed)
# ---------------------------------------------------------------------------


@click.command(name="llm-help")
@click.option(
    "--python",
    "python",
    is_flag=True,
    default=False,
    help="Print the Python-API guide instead of the CLI guide.",
)
def llm_help_cmd(python: bool) -> None:
    """Print the full wordlive agent guide (the bundled skill) to stdout.

    One-shot orientation for an LLM: the anchor model, every verb, image
    insertion, the `exec` batch format, and the exit-code taxonomy. `wordlive
    --help` points here. Defaults to the CLI guide; `--python` prints the
    Python-API (`import wordlive as wl`) guide instead. Output is raw Markdown —
    not JSON, and unaffected by `--json/--text` — so it reads cleanly straight
    into a model's context, exactly like `--help`. Offline: never touches Word.
    """
    kind = "python" if python else "cli"
    try:
        click.echo(_skill_body(kind))
    except (FileNotFoundError, ModuleNotFoundError, OSError, ValueError) as e:
        raise click.ClickException(f"could not read the bundled skill: {e}") from e


@click.command(name="install-skill")
@click.option(
    "--cli", "cli", is_flag=True, default=False, help="Install the CLI skill (the default)."
)
@click.option(
    "--python", "python", is_flag=True, default=False, help="Install only the Python-API skill."
)
@click.option(
    "--both",
    "both",
    is_flag=True,
    default=False,
    help="Install both the CLI and Python-API skills.",
)
@click.option(
    "--system",
    "system",
    is_flag=True,
    default=False,
    help="Install to ~/.agents/skills/ instead of the current project's ./.agents/skills/.",
)
@click.option(
    "--force", "force", is_flag=True, default=False, help="Overwrite an existing SKILL.md."
)
@click.pass_context
def install_skill_cmd(
    ctx: click.Context, cli: bool, python: bool, both: bool, system: bool, force: bool
) -> None:
    """Install wordlive's agent skill(s) (SKILL.md) for LLM coding tools.

    wordlive ships two skills — `wordlive-cli` (the command-line workflow) and
    `wordlive-python` (the `import wordlive as wl` API). By default only the
    **CLI** skill is installed; pass `--python` for just the Python one, or
    `--both` for both. They land under `.agents/skills/<name>/SKILL.md` in the
    current directory (default) or your home directory (`--system`). Offline —
    this doesn't touch Word.
    """
    if both or (cli and python):
        kinds = ["cli", "python"]
    elif python:
        kinds = ["python"]
    else:
        kinds = ["cli"]

    base = Path.home() if system else Path.cwd()
    scope = "system" if system else "local"
    dests = [(kind, base / ".agents" / "skills" / _skill_name(kind) / "SKILL.md") for kind in kinds]

    # Check every target up front so we never half-write when --force is absent.
    clashes = [str(dest) for _, dest in dests if dest.exists()]
    if clashes and not force:
        raise click.ClickException(
            "already exists (pass --force to overwrite): " + ", ".join(clashes)
        )

    installed = []
    try:
        for kind, dest in dests:
            content = _bundled_skill(kind)
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_text(content, encoding="utf-8")
            installed.append(
                {
                    "kind": kind,
                    "name": _skill_name(kind),
                    "path": str(dest),
                    "bytes": len(content.encode("utf-8")),
                }
            )
    except (FileNotFoundError, ModuleNotFoundError, OSError, ValueError) as e:
        raise click.ClickException(f"could not install the skill: {e}") from e

    emit(
        {"ok": True, "scope": scope, "installed": installed},
        as_text=not ctx.obj["as_json"],
        text="installed:\n" + "\n".join(f"  {r['name']} → {r['path']}" for r in installed),
    )


# ---------------------------------------------------------------------------
# install-mcp — register the MCP server in an agent's config
# ---------------------------------------------------------------------------


def _claude_desktop_config_path() -> Path:
    """Where Claude Desktop keeps `claude_desktop_config.json` on this OS."""
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA")
        base = Path(appdata) if appdata else Path.home() / "AppData" / "Roaming"
        return base / "Claude" / "claude_desktop_config.json"
    if sys.platform == "darwin":
        return (
            Path.home()
            / "Library"
            / "Application Support"
            / "Claude"
            / "claude_desktop_config.json"
        )
    return Path.home() / ".config" / "Claude" / "claude_desktop_config.json"


def _mcp_server_entry(directory: str | None) -> dict[str, Any]:
    """The `mcpServers` entry that launches the wordlive stdio server.

    Default (repo-less) form runs the published package straight from PyPI with
    `uvx` — `wordlive-mcp` is a console script *inside* `wordlive`, so it needs
    `--from "wordlive[mcp,snapshot]"` to tell uv which package provides it (and
    the `snapshot` extra enables the vision tool). With `--directory` (a local
    checkout) wordlive *is* the project, so a plain `uv run wordlive-mcp`
    resolves it without `--from`.
    """
    if directory:
        return {"command": "uv", "args": ["run", "--directory", directory, "wordlive-mcp"]}
    return {"command": "uvx", "args": ["--from", "wordlive[mcp,snapshot]", "wordlive-mcp"]}


@click.command(name="install-mcp")
@click.option(
    "--client",
    type=click.Choice(["claude-desktop", "claude-code"]),
    default="claude-desktop",
    help="Which MCP client's config to write (default: claude-desktop).",
)
@click.option(
    "--name", "server_name", default="wordlive", help="Server key to register (default: wordlive)."
)
@click.option(
    "--directory",
    "directory",
    default=None,
    help="Register a local checkout via `uv run --directory DIR` (dev), instead of the default `uvx --from wordlive[mcp,snapshot]`.",
)
@click.option(
    "--config",
    "config_path",
    default=None,
    type=click.Path(),
    help="Write to this config file instead of the client's default location.",
)
@click.option(
    "--print",
    "print_only",
    is_flag=True,
    default=False,
    help="Print the JSON server snippet to stdout instead of writing any file.",
)
@click.option(
    "--force", "force", is_flag=True, default=False, help="Overwrite an existing server entry."
)
@click.pass_context
def install_mcp_cmd(
    ctx: click.Context,
    client: str,
    server_name: str,
    directory: str | None,
    config_path: str | None,
    print_only: bool,
    force: bool,
) -> None:
    """Register the wordlive MCP server in an agent's config.

    Merges an `mcpServers.<name>` entry into Claude Desktop's
    `claude_desktop_config.json` (default) or a Claude Code `.mcp.json`
    (`--client claude-code`, project-local). The entry launches the stdio server
    with `uvx --from "wordlive[mcp,snapshot]" wordlive-mcp` (no separate install
    needed), or `uv run --directory DIR wordlive-mcp` for a local checkout. Use
    `--print` to just emit the snippet for any client. Offline — never touches
    Word; restart the client to pick up the change.
    """
    entry = _mcp_server_entry(directory)

    if print_only:
        emit(
            {"ok": True, "server": server_name, "entry": entry, "mcpServers": {server_name: entry}},
            as_text=not ctx.obj["as_json"],
            text=json.dumps({"mcpServers": {server_name: entry}}, indent=2),
        )
        return

    if config_path is not None:
        target = Path(config_path)
    elif client == "claude-desktop":
        target = _claude_desktop_config_path()
    else:  # claude-code: portable, project-local server file
        target = Path.cwd() / ".mcp.json"

    cfg: dict[str, Any] = {}
    if target.exists():
        try:
            raw = target.read_text(encoding="utf-8").strip()
            cfg = json.loads(raw) if raw else {}
        except (OSError, json.JSONDecodeError) as e:
            raise click.ClickException(f"could not read existing config {target}: {e}") from e
        if not isinstance(cfg, dict):
            raise click.ClickException(f"existing config {target} is not a JSON object")

    servers = cfg.setdefault("mcpServers", {})
    if not isinstance(servers, dict):
        raise click.ClickException(f"'mcpServers' in {target} is not a JSON object")
    action = "updated" if server_name in servers else "created"
    if server_name in servers and not force:
        raise click.ClickException(
            f"server '{server_name}' is already in {target}; pass --force to overwrite"
        )
    servers[server_name] = entry

    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(cfg, indent=2) + "\n", encoding="utf-8")
    except OSError as e:
        raise click.ClickException(f"could not write {target}: {e}") from e

    emit(
        {
            "ok": True,
            "client": client,
            "path": str(target),
            "server": server_name,
            "action": action,
            "entry": entry,
        },
        as_text=not ctx.obj["as_json"],
        text=f"{action} server '{server_name}' → {target}\n(restart {client} to load it)",
    )
