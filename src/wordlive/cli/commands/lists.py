"""The `list` command group."""

from __future__ import annotations

import json

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_list_info,
    _fmt_list_levels,
    _fmt_list_show,
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
