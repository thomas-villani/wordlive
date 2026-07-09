"""The `comment` command group."""

from __future__ import annotations

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_comment_list,
)


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
