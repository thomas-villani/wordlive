"""CLI subcommands wired against the wordlive library."""

from __future__ import annotations

from typing import Any

import click

from .. import attach
from .._document import Document
from ..exceptions import WordNotRunningError
from .main import emit, _run


def _pick_doc(word: Any, doc_name: str | None) -> Document:
    if doc_name is None:
        return word.documents.active
    return word.documents[doc_name]


def register(group: click.Group) -> None:
    group.add_command(status)
    group.add_command(outline)
    group.add_command(read)
    group.add_command(write)
    group.add_command(insert)


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
                emit(word.documents.list(), as_text=not ctx.obj["as_json"])
        except WordNotRunningError:
            emit([], as_text=not ctx.obj["as_json"])
            raise
    _run(ctx, go)


# ---------------------------------------------------------------------------
# outline
# ---------------------------------------------------------------------------


@click.command(name="outline")
@click.pass_context
def outline(ctx: click.Context) -> None:
    """Print the heading outline of the target document as JSON."""
    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            emit(doc.outline(), as_text=not ctx.obj["as_json"])
    _run(ctx, go)


# ---------------------------------------------------------------------------
# read bookmark|cc NAME
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
            emit({"text": doc.bookmarks[name].text}, as_text=not ctx.obj["as_json"])
    _run(ctx, go)


@read.command(name="cc")
@click.argument("name")
@click.pass_context
def read_cc(ctx: click.Context, name: str) -> None:
    """Read the text of content control NAME."""
    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            emit({"text": doc.content_controls[name].text}, as_text=not ctx.obj["as_json"])
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
            with doc.edit(f"CLI: write bookmark {name}"):
                doc.bookmarks[name].set_text(text)
            emit({"ok": True, "anchor": {"kind": "bookmark", "name": name}}, as_text=not ctx.obj["as_json"])
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
            with doc.edit(f"CLI: write cc {name}"):
                doc.content_controls[name].set_text(text)
            emit({"ok": True, "anchor": {"kind": "content_control", "name": name}}, as_text=not ctx.obj["as_json"])
    _run(ctx, go)


# ---------------------------------------------------------------------------
# insert --after-heading "Intro" --text "..." [--style "..."]
# ---------------------------------------------------------------------------


@click.command(name="insert")
@click.option("--after-heading", "after_heading", required=True, help="Heading text to anchor against.")
@click.option("--text", "text", required=True, help="Paragraph text to insert.")
@click.option("--style", "style", default=None, help="Optional Word style name.")
@click.pass_context
def insert(ctx: click.Context, after_heading: str, text: str, style: str | None) -> None:
    """Insert a paragraph after the named heading (atomic-undo)."""
    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: insert after {after_heading!r}"):
                doc.heading(after_heading).insert_paragraph_after(text, style=style)
            emit(
                {
                    "ok": True,
                    "after_heading": after_heading,
                    "style": style,
                },
                as_text=not ctx.obj["as_json"],
            )
    _run(ctx, go)
