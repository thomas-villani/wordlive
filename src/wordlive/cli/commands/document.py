"""Document-level and navigation commands."""

from __future__ import annotations

import json
from pathlib import Path

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ...exceptions import OpError, WordNotRunningError
from ..main import _run, emit
from ._common import (
    _fmt_cursor,
    _fmt_find,
    _fmt_find_paragraphs,
    _fmt_location,
    _fmt_outline,
    _fmt_paragraphs,
    _fmt_stats,
    _fmt_status,
    _load_checkpoint,
)


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
    from ..._checkpoint import diff_checkpoints

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
