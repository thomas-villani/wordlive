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
from ..exceptions import AmbiguousMatchError, WordNotRunningError
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
    group.add_command(insert_break_cmd)
    group.add_command(insert_field_cmd)
    group.add_command(update_fields_cmd)
    group.add_command(insert_footnote_cmd)
    group.add_command(insert_endnote_cmd)
    group.add_command(insert_toc_cmd)
    group.add_command(footnotes_cmd)
    group.add_command(endnotes_cmd)
    group.add_command(page_setup_cmd)
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
    group.add_command(format_run_cmd)
    group.add_command(shading_cmd)
    group.add_command(borders_cmd)
    group.add_command(tab_stop_cmd)
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
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert image {where} {anchor_id}"):
                anchor.insert_image(
                    image,
                    wrap=wrap,
                    where=where,
                    block=block,
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
                    "block": block,
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
    "--page-break-before/--no-page-break-before",
    "page_break_before",
    default=None,
    help="Force (or clear) a page break before the paragraph — the clean, "
    "reflow-safe way to page-break (e.g. on every Heading 1).",
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
    page_break_before: bool | None,
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
    if page_break_before is not None:
        kwargs["page_break_before"] = page_break_before
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
    "--style", "style", default="single", help="Line style: single, double, dot, dash, … or none."
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


@table.command(name="create")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Position anchor for the new table (heading:/para:/start/end/range:…).",
)
@click.option("--rows", "rows", type=int, required=True, help="Number of rows (>= 1).")
@click.option("--cols", "cols", type=int, required=True, help="Number of columns (>= 1).")
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
    help="Row-major JSON 2-D array to populate cells "
    '(e.g. \'[["Name","Qty"],["Widget","3"]]\'), or \'-\' to read it from '
    "stdin. Reading from stdin avoids quoting/backslash fights on Windows.",
)
@click.pass_context
def table_create(
    ctx: click.Context,
    anchor_id: str,
    rows: int,
    cols: int,
    style: str | None,
    header: bool,
    before: bool,
    data: str | None,
) -> None:
    """Create a ROWS x COLS table at an anchor (atomic-undo).

    Builds new table structure where wordlive's other verbs only edit existing
    structure. Fill cells at creation with --data (a row-major JSON array, or
    '--data -' to read it from stdin); a short array leaves trailing cells
    empty. --style defaults to 'Table Grid' (visible borders); a style name not
    defined in the document raises (exit 2). Reports the new table's 1-based
    index for a follow-up `table set-cell` / `add-row`.
    """
    parsed: list[Any] | None = None
    if data is not None:
        raw = click.get_text_stream("stdin").read() if data == "-" else data
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--data must be a JSON 2-D array: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--data must be a JSON array of rows")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: create {rows}x{cols} table at {anchor_id}"):
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
    Supported ops: write_bookmark, write_cc, insert_paragraph, append,
    append_inline, prepend, prepend_inline, insert_image, replace, find_replace,
    apply_style, format_paragraph, set_cell, add_row, delete_row, create_table,
    delete_table, insert_break, add_comment, resolve_comment, delete_comment, apply_list,
    remove_list, restart_numbering, indent_list, outdent_list, write_header,
    write_footer. (append/prepend add a new paragraph + optional style;
    append_inline/prepend_inline continue the adjacent paragraph, text only.
    append_paragraph/prepend_paragraph remain as synonyms.) A field an op doesn't
    use is reported in the result's `warnings`, not silently dropped.
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
