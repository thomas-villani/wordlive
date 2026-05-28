"""CLI subcommands wired against the wordlive library."""

from __future__ import annotations

import base64
import json
from contextlib import nullcontext
from importlib.resources import files
from pathlib import Path
from typing import Any

import click

from .. import attach
from .._anchors import Heading
from .._document import Document
from ..exceptions import AmbiguousMatchError, WordliveError, WordNotRunningError
from .main import _run, emit


def _pick_doc(word: Any, doc_name: str | None) -> Document:
    if doc_name is None:
        return word.documents.active
    return word.documents[doc_name]


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
    group.add_command(prepend_cmd)
    group.add_command(append_cmd)
    group.add_command(insert_image_cmd)
    group.add_command(snapshot_cmd)
    group.add_command(cursor)
    group.add_command(find_cmd)
    group.add_command(replace)
    group.add_command(go_to)
    group.add_command(style)
    group.add_command(format_paragraph_cmd)
    group.add_command(table)
    group.add_command(comment)
    group.add_command(track)
    group.add_command(list_cmd)
    group.add_command(section)
    group.add_command(header)
    group.add_command(footer)
    group.add_command(exec_)
    group.add_command(llm_help_cmd)
    group.add_command(install_skill_cmd)


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
# read bookmark|cc|section NAME
# ---------------------------------------------------------------------------


@click.group(name="read")
def read() -> None:
    """Read structured values from the target document."""


@read.command(name="bookmark")
@click.argument("name")
@click.pass_context
def read_bookmark(ctx: click.Context, name: str) -> None:
    """Read the text of bookmark NAME."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
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


# ---------------------------------------------------------------------------
# write bookmark|cc NAME --text "..."
# ---------------------------------------------------------------------------


@click.group(name="write")
def write() -> None:
    """Write structured values into the target document."""


@write.command(name="bookmark")
@click.argument("name")
@click.option("--text", "text", required=True, help="New text for the bookmark.")
@click.pass_context
def write_bookmark(ctx: click.Context, name: str, text: str) -> None:
    """Set the text of bookmark NAME (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
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
@click.option("--text", "text", required=True, help="Paragraph text to insert.")
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
def insert(ctx: click.Context, anchor_id: str, text: str, before: bool, style: str | None) -> None:
    """Insert a new paragraph before/after any anchor (atomic-undo).

    Addresses anchors the same way every other command does — `--anchor-id`
    (headings, paragraphs, bookmarks, cells, ranges). To insert text *inside* a
    paragraph at an offset, target a collapsed range instead:
    `replace --anchor-id range:120-120 --text "…"` (offsets come from
    `paragraphs` / `find`).
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert {where} {anchor_id}"):
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
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert image {where} {anchor_id}"):
                anchor.insert_image(
                    image,
                    wrap=wrap,
                    where=where,
                    width=width,
                    height=height,
                    alt_text=alt_text,
                    lock_aspect=lock_aspect,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "wrap": wrap,
                    "where": where,
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted image {where} {anchor_id} (wrap={wrap})",
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
@click.pass_context
def snapshot_cmd(
    ctx: click.Context,
    anchor_id: str | None,
    page: int | None,
    pages_range: str | None,
    out: Path | None,
    dpi: int,
) -> None:
    """Render document page(s) to PNG so a vision model can see the layout.

    Word exports a pixel-faithful PDF of the document it has open and wordlive
    rasterises the requested pages — a true WYSIWYG image (real fonts, spacing,
    page geometry) for iterating on style and formatting. Read-only.

    Choose at most one target: `--anchor-id` (the page(s) an anchor occupies; a
    `heading:` expands to its whole section), `--page N`, or `--pages A-B`. With
    none, the whole document is rendered. With `--out` the image is written to
    disk (one file per page); otherwise base64 PNG data is returned inline.

    Requires the `snapshot` extra: `pip install "wordlive[snapshot]"`.
    """
    targets = [t is not None for t in (anchor_id, page, pages_range)]
    if sum(targets) > 1:
        raise click.UsageError("provide at most one of --anchor-id, --page, or --pages")
    if dpi < 1:
        raise click.UsageError("--dpi must be >= 1")
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
                shots = doc.snapshot_anchor(anchor, out, dpi=dpi)
                selector: Any = anchor_id
            else:
                shots = doc.snapshot(out, pages=pages_arg, dpi=dpi)
                selector = pages_range or page or "all"
            images: list[dict[str, Any]] = []
            for s in shots:
                entry: dict[str, Any] = {"page": s.page, "bytes": len(s.png)}
                if s.path is not None:
                    entry["path"] = str(s.path)
                else:
                    entry["base64"] = base64.b64encode(s.png).decode("ascii")
                images.append(entry)
            emit(
                {
                    "ok": True,
                    "selector": selector,
                    "dpi": dpi,
                    "count": len(images),
                    "images": images,
                },
                as_text=not ctx.obj["as_json"],
                text=_fmt_snapshot(images, dpi),
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
@click.pass_context
def find_cmd(ctx: click.Context, text: str, in_: str | None) -> None:
    """Locate every fuzzy occurrence of TEXT (read-only)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            scope = doc.anchor_by_id(in_) if in_ else None
            matches = doc.find(text, scope=scope)
            emit(matches, as_text=not ctx.obj["as_json"], text=_fmt_find(matches))

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
    help="In fuzzy mode, replace only the Nth match (1-based).",
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
) -> None:
    """Replace text. Either at an anchor (entire range) or via fuzzy find."""
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
    widths = [
        max((len(row[c]["text"]) for row in cells), default=0)
        for c in range(grid.get("columns", 0))
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


@click.group(name="section")
def section() -> None:
    """Inspect document sections (headers/footers live in `header` / `footer`)."""


@section.command(name="list")
@click.pass_context
def section_list(ctx: click.Context) -> None:
    """List sections with their page setup (orientation, margins, page size)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.sections.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_section_list(rows))

    _run(ctx, go)


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


# ---------------------------------------------------------------------------
# exec --script ops.json
# ---------------------------------------------------------------------------


# Required fields per op kind. Validated up-front so a malformed payload
# raises a clean click.ClickException ("exec op 'write_bookmark' requires
# field 'name'") instead of a Python KeyError traceback that would land an
# LLM tool-use loop on exit code 1 with no actionable signal.
_OP_REQUIRED_FIELDS: dict[str, tuple[str, ...]] = {
    "write_bookmark": ("name", "text"),
    "write_cc": ("name", "text"),
    "insert_paragraph": ("anchor_id", "text"),
    "append_paragraph": ("text",),
    "append": ("text",),
    "prepend_paragraph": ("text",),
    "prepend": ("text",),
    "insert_image": ("anchor_id", "wrap"),
    "replace": ("anchor_id", "text"),
    "find_replace": ("find", "text"),
    "apply_style": ("anchor_id", "name"),
    "format_paragraph": ("anchor_id",),
    "set_cell": ("table", "row", "col", "text"),
    "add_row": ("table",),
    "delete_row": ("table", "row"),
    "add_comment": ("anchor_id", "text"),
    "resolve_comment": ("index",),
    "delete_comment": ("index",),
    "apply_list": ("anchor_id",),
    "remove_list": ("anchor_id",),
    "restart_numbering": ("anchor_id",),
    "indent_list": ("anchor_id",),
    "outdent_list": ("anchor_id",),
    "write_header": ("section", "text"),
    "write_footer": ("section", "text"),
}


def _op_before(op: dict[str, Any]) -> bool:
    """Whether an insert op targets *before* its anchor (default: after).

    Accepts either the verbose `"where": "before"|"after"` or the boolean
    `"before": true` / `"after": true` — the latter mirrors the CLI's
    `--before/--after` flags, so the natural JSON encoding works regardless of
    which form an LLM reaches for. An explicit `"before"` wins if both appear.
    """
    if "before" in op:
        return bool(op["before"])
    if "after" in op:
        return not bool(op["after"])
    return op.get("where") == "before"


def _validate_op(op: dict[str, Any]) -> str:
    """Return the op kind after asserting it's known and required keys exist."""
    if not isinstance(op, dict):
        raise click.ClickException(f"each op must be an object; got {type(op).__name__}")
    kind = op.get("op")
    if kind is None:
        raise click.ClickException("op is missing the 'op' field")
    if kind not in _OP_REQUIRED_FIELDS:
        raise click.ClickException(f"unknown op: {kind!r}")
    missing = [k for k in _OP_REQUIRED_FIELDS[kind] if k not in op]
    if missing:
        raise click.ClickException(
            f"op {kind!r} is missing required field(s): {', '.join(repr(m) for m in missing)}"
        )
    return kind


def _apply_op(doc: Document, op: dict[str, Any]) -> None:
    """Apply a single op from an exec script. Raises WordliveError on bad input."""
    kind = _validate_op(op)
    if kind == "write_bookmark":
        doc.bookmarks[op["name"]].set_text(op["text"])
    elif kind == "write_cc":
        doc.content_controls[op["name"]].set_text(op["text"])
    elif kind == "insert_paragraph":
        anchor = doc.anchor_by_id(op["anchor_id"])
        if _op_before(op):
            anchor.insert_paragraph_before(op["text"], style=op.get("style"))
        else:
            anchor.insert_paragraph_after(op["text"], style=op.get("style"))
    elif kind == "append_paragraph":
        doc.append_paragraph(op["text"], style=op.get("style"))
    elif kind == "append":
        doc.append(op["text"])
    elif kind == "prepend_paragraph":
        doc.prepend_paragraph(op["text"], style=op.get("style"))
    elif kind == "prepend":
        doc.prepend(op["text"])
    elif kind == "insert_image":
        if ("path" in op) == ("base64" in op):
            raise click.ClickException(
                "op 'insert_image' requires exactly one of 'path' or 'base64'"
            )
        image: str | Path = Path(op["path"]) if "path" in op else op["base64"]
        kwargs = {k: op[k] for k in ("width", "height", "alt_text", "lock_aspect") if k in op}
        doc.anchor_by_id(op["anchor_id"]).insert_image(
            image, wrap=op["wrap"], where=("before" if _op_before(op) else "after"), **kwargs
        )
    elif kind == "replace":
        doc.anchor_by_id(op["anchor_id"]).set_text(op["text"])
    elif kind == "find_replace":
        scope = doc.anchor_by_id(op["in"]) if op.get("in") else None
        doc.find_replace(
            op["find"],
            op["text"],
            scope=scope,
            all=bool(op.get("all", False)),
            occurrence=op.get("occurrence"),
        )
    elif kind == "apply_style":
        doc.anchor_by_id(op["anchor_id"]).apply_style(op["name"])
    elif kind == "format_paragraph":
        kwargs = {
            k: op[k]
            for k in (
                "alignment",
                "left_indent",
                "right_indent",
                "first_line_indent",
                "space_before",
                "space_after",
            )
            if k in op
        }
        doc.anchor_by_id(op["anchor_id"]).format_paragraph(**kwargs)
    elif kind == "set_cell":
        doc.tables[op["table"]].cell(op["row"], op["col"]).set_text(op["text"])
    elif kind == "add_row":
        doc.tables[op["table"]].add_row(op.get("values"))
    elif kind == "delete_row":
        doc.tables[op["table"]].delete_row(op["row"])
    elif kind == "add_comment":
        anchor = doc.anchor_by_id(op["anchor_id"])
        doc.comments.add(anchor, op["text"], author=op.get("author"))
    elif kind == "resolve_comment":
        doc.comments[op["index"]].resolve()
    elif kind == "delete_comment":
        doc.comments[op["index"]].delete()
    elif kind == "apply_list":
        continue_previous = bool(op.get("continue_previous", op.get("continue", False)))
        doc.anchor_by_id(op["anchor_id"]).apply_list(
            op.get("type", "bulleted"), continue_previous=continue_previous
        )
    elif kind == "remove_list":
        doc.anchor_by_id(op["anchor_id"]).remove_list()
    elif kind == "restart_numbering":
        doc.anchor_by_id(op["anchor_id"]).restart_numbering()
    elif kind == "indent_list":
        doc.anchor_by_id(op["anchor_id"]).indent_list()
    elif kind == "outdent_list":
        doc.anchor_by_id(op["anchor_id"]).outdent_list()
    elif kind == "write_header":
        doc.sections[op["section"]].header(op.get("which", "primary")).set_text(op["text"])
    elif kind == "write_footer":
        doc.sections[op["section"]].footer(op.get("which", "primary")).set_text(op["text"])


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
    Supported ops: write_bookmark, write_cc, insert_paragraph, append_paragraph,
    append, prepend_paragraph, prepend, insert_image, replace, find_replace,
    apply_style, format_paragraph, set_cell, add_row, delete_row, add_comment,
    resolve_comment, delete_comment, apply_list, remove_list, restart_numbering,
    indent_list, outdent_list, write_header, write_footer.
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

        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            ops_run = 0
            failure_exc: WordliveError | None = None
            failure_meta: dict[str, Any] | None = None
            tracking = doc.tracked_changes() if tracked else nullcontext()
            with tracking, doc.edit(label):
                for i, op in enumerate(ops):
                    try:
                        _apply_op(doc, op)
                    except WordliveError as exc:
                        failure_exc = exc
                        failure_meta = {
                            "index": i,
                            "op": op,
                            "error": str(exc),
                            "type": type(exc).__name__,
                        }
                        if isinstance(exc, AmbiguousMatchError):
                            failure_meta["matches"] = exc.matches
                        break
                    ops_run += 1
            if failure_exc is None:
                emit(
                    {"ok": True, "ops_run": ops_run, "label": label},
                    as_text=not ctx.obj["as_json"],
                    text=f"applied {ops_run} op(s): {label!r}",
                )
            else:
                assert failure_meta is not None  # set together with failure_exc
                emit(
                    {"ok": False, "ops_run": ops_run, "label": label, "failure": failure_meta},
                    as_text=not ctx.obj["as_json"],
                    text=f"failed at op {failure_meta['index']}: {failure_meta['error']}",
                )
                # Re-raise the original so _run() maps it to the right exit code
                # (e.g. anchor-not-found → 2, busy → 3, ambiguous → 5).
                raise failure_exc

    _run(ctx, go)


# ---------------------------------------------------------------------------
# install-skill [--system] [--force]
# ---------------------------------------------------------------------------


def _bundled_skill() -> str:
    """The packaged agent skill (SKILL.md) text."""
    return (files("wordlive") / "_skill" / "SKILL.md").read_text(encoding="utf-8")


def _strip_frontmatter(md: str) -> str:
    """Drop a leading YAML frontmatter block (--- … ---), if present.

    The bundled SKILL.md opens with `name:` / `description:` frontmatter for the
    agent-skill loader. That metadata is noise when the doc is read straight off
    stdout, so `llm-help` emits just the Markdown body.
    """
    lines = md.splitlines()
    if lines and lines[0].strip() == "---":
        for i in range(1, len(lines)):
            if lines[i].strip() == "---":
                return "\n".join(lines[i + 1 :]).lstrip("\n")
    return md


@click.command(name="llm-help")
def llm_help_cmd() -> None:
    """Print the full wordlive agent guide (the bundled skill) to stdout.

    One-shot orientation for an LLM: the anchor model, every read/write verb,
    image insertion, the `exec` batch format, and the exit-code taxonomy.
    `wordlive --help` points here. Output is raw Markdown — not JSON, and
    unaffected by `--json/--text` — so it reads cleanly straight into a model's
    context, exactly like `--help` itself. Offline: never touches Word.
    """
    try:
        content = _bundled_skill()
    except (FileNotFoundError, ModuleNotFoundError, OSError) as e:
        raise click.ClickException(f"could not read the bundled skill: {e}") from e
    click.echo(_strip_frontmatter(content))


@click.command(name="install-skill")
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
def install_skill_cmd(ctx: click.Context, system: bool, force: bool) -> None:
    """Install the wordlive agent skill (SKILL.md) for LLM coding tools.

    Writes `.agents/skills/wordlive/SKILL.md` under the current directory
    (default) or your home directory (`--system`). This is offline — it doesn't
    touch Word.
    """
    base = Path.home() if system else Path.cwd()
    dest = base / ".agents" / "skills" / "wordlive" / "SKILL.md"
    scope = "system" if system else "local"
    if dest.exists() and not force:
        raise click.ClickException(f"{dest} already exists; pass --force to overwrite")
    try:
        content = _bundled_skill()
    except (FileNotFoundError, ModuleNotFoundError, OSError) as e:
        raise click.ClickException(f"could not read the bundled skill: {e}") from e
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(content, encoding="utf-8")
    except OSError as e:
        raise click.ClickException(f"could not write {dest}: {e}") from e
    emit(
        {"ok": True, "scope": scope, "path": str(dest), "bytes": len(content.encode("utf-8"))},
        as_text=not ctx.obj["as_json"],
        text=f"installed wordlive skill → {dest}",
    )
