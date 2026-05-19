"""CLI subcommands wired against the wordlive library."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import click

from .. import attach
from .._anchors import Heading
from .._document import Document
from ..exceptions import AmbiguousMatchError, WordNotRunningError, WordliveError
from .main import emit, _run


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


def _fmt_find(matches: list[dict[str, Any]]) -> str:
    if not matches:
        return "(no matches)"
    return "\n".join(
        f"{m['start']:>6}–{m['end']:<6}  {m['text']!r}" for m in matches
    )


def _fmt_replace_summary(replacements: list[dict[str, Any]]) -> str:
    n = len(replacements)
    return f"replaced {n} occurrence{'s' if n != 1 else ''}"


def register(group: click.Group) -> None:
    group.add_command(status)
    group.add_command(outline)
    group.add_command(read)
    group.add_command(write)
    group.add_command(insert)
    group.add_command(find_cmd)
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
@click.pass_context
def outline(ctx: click.Context) -> None:
    """Print the heading outline of the target document."""
    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            items = doc.outline()
            emit(items, as_text=not ctx.obj["as_json"], text=_fmt_outline(items))
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
@click.option("--anchor-id", "anchor_id", default=None, help="Resolve heading by anchor id (e.g. 'heading:3') instead of by visible text.")
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
                    raise click.UsageError(
                        f"--anchor-id must reference a heading, got {h.kind!r}"
                    )
            else:
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
            with doc.edit(f"CLI: write bookmark {name}"):
                doc.bookmarks[name].set_text(text)
            emit(
                {"ok": True, "anchor": {"kind": "bookmark", "name": name}},
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
            with doc.edit(f"CLI: write cc {name}"):
                doc.content_controls[name].set_text(text)
            emit(
                {"ok": True, "anchor": {"kind": "content_control", "name": name}},
                as_text=not ctx.obj["as_json"],
                text=f"wrote cc:{name}",
            )
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
                text=f"inserted after {after_heading!r}",
            )
    _run(ctx, go)


# ---------------------------------------------------------------------------
# find --text "..." [--in ANCHOR_ID]
# ---------------------------------------------------------------------------


@click.command(name="find")
@click.option("--text", "text", required=True, help="Text to locate (whitespace + smart-quote fuzzy match).")
@click.option("--in", "in_", default=None, help="Optional anchor id to scope the search (e.g. 'heading:3').")
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
@click.option("--anchor-id", "anchor_id", default=None, help="Replace the entire range at this anchor.")
@click.option("--find", "find", default=None, help="Fuzzy text to locate (alternative to --anchor-id).")
@click.option("--text", "text", required=True, help="Replacement text.")
@click.option("--in", "in_", default=None, help="In fuzzy mode, scope the search to this anchor id.")
@click.option("--all", "replace_all", is_flag=True, default=False, help="In fuzzy mode, replace every match.")
@click.option("--occurrence", "occurrence", type=int, default=None, help="In fuzzy mode, replace only the Nth match (1-based).")
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
                text=f"moved to {anchor_id}",
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
    elif kind == "find_replace":
        scope = doc.anchor_by_id(op["in"]) if op.get("in") else None
        doc.find_replace(
            op["find"],
            op["text"],
            scope=scope,
            all=bool(op.get("all", False)),
            occurrence=op.get("occurrence"),
        )
    else:
        raise click.ClickException(f"unknown op: {kind!r}")


@click.command(name="exec")
@click.option("--script", "script", required=True, type=click.Path(exists=True, dir_okay=False, path_type=Path), help="Path to ops JSON file.")
@click.pass_context
def exec_(ctx: click.Context, script: Path) -> None:
    """Apply a batch of ops in a single atomic-undo scope.

    Script shape: `{"label": "…", "ops": [{"op": "...", ...}, ...]}`.
    Supported ops: write_bookmark, write_cc, insert_after_heading, replace,
    find_replace.
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
                emit(
                    {"ok": False, "ops_run": ops_run, "label": label, "failure": failure_meta},
                    as_text=not ctx.obj["as_json"],
                    text=f"failed at op {failure_meta['index']}: {failure_meta['error']}",
                )
                # Re-raise the original so _run() maps it to the right exit code
                # (e.g. anchor-not-found → 2, busy → 3, ambiguous → 5).
                raise failure_exc
    _run(ctx, go)
