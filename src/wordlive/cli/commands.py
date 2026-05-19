"""CLI subcommands wired against the wordlive library."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import click

from .. import attach
from .._document import Document
from ..exceptions import WordNotRunningError, WordliveError
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
    group.add_command(replace)
    group.add_command(go_to)
    group.add_command(exec_)


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


# ---------------------------------------------------------------------------
# replace --anchor-id ID --text "..."
# ---------------------------------------------------------------------------


@click.command(name="replace")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor ID (e.g. 'heading:3', 'bookmark:Address', 'cc:Signatory').")
@click.option("--text", "text", required=True, help="Replacement text.")
@click.pass_context
def replace(ctx: click.Context, anchor_id: str, text: str) -> None:
    """Replace the text at an anchor identified by anchor-id (atomic-undo)."""
    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
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
            )
    _run(ctx, go)


# ---------------------------------------------------------------------------
# go-to --anchor-id ID
# ---------------------------------------------------------------------------


@click.command(name="go-to")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor ID to move the cursor to.")
@click.option("--scroll/--no-scroll", default=True, help="Scroll the view to the anchor (default: yes).")
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
            )
    _run(ctx, go)


# ---------------------------------------------------------------------------
# exec --script ops.json
# ---------------------------------------------------------------------------


def _apply_op(doc: Document, op: dict[str, Any]) -> None:
    """Apply a single op from an exec script. Raises WordliveError on bad input."""
    kind = op.get("op")
    if kind == "write_bookmark":
        doc.bookmarks[op["name"]].set_text(op["text"])
    elif kind == "write_cc":
        doc.content_controls[op["name"]].set_text(op["text"])
    elif kind == "insert_after_heading":
        doc.heading(op["heading"]).insert_paragraph_after(op["text"], style=op.get("style"))
    elif kind == "replace":
        doc.anchor_by_id(op["anchor_id"]).set_text(op["text"])
    else:
        raise click.ClickException(f"unknown op: {kind!r}")


@click.command(name="exec")
@click.option("--script", "script", required=True, type=click.Path(exists=True, dir_okay=False, path_type=Path), help="Path to ops JSON file.")
@click.pass_context
def exec_(ctx: click.Context, script: Path) -> None:
    """Apply a batch of ops in a single atomic-undo scope.

    Script shape: `{"label": "…", "ops": [{"op": "...", ...}, ...]}`.
    Supported ops: write_bookmark, write_cc, insert_after_heading, replace.
    """
    def go() -> None:
        payload = json.loads(script.read_text(encoding="utf-8"))
        label = str(payload.get("label") or f"CLI: exec {script.name}")
        ops = payload.get("ops") or []
        if not isinstance(ops, list):
            raise click.ClickException("'ops' must be a list")

        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            ops_run = 0
            failure_exc: WordliveError | None = None
            failure_meta: dict[str, Any] | None = None
            with doc.edit(label):
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
                        break
                    ops_run += 1
            if failure_exc is None:
                emit({"ok": True, "ops_run": ops_run, "label": label}, as_text=not ctx.obj["as_json"])
            else:
                emit(
                    {"ok": False, "ops_run": ops_run, "label": label, "failure": failure_meta},
                    as_text=not ctx.obj["as_json"],
                )
                # Re-raise the original so _run() maps it to the right exit code
                # (e.g. anchor-not-found → 2, busy → 3).
                raise failure_exc
    _run(ctx, go)
