"""Content-control creation and configuration."""

from __future__ import annotations

from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _cc_anchor,
)


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
