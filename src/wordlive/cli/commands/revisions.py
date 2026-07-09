"""Track-changes and revision commands."""

from __future__ import annotations

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_revisions,
)


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
