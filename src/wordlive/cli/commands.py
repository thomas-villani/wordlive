"""CLI subcommands wired against the wordlive library."""

from __future__ import annotations

import json
from contextlib import nullcontext
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
    group.add_command(style)
    group.add_command(format_paragraph_cmd)
    group.add_command(table)
    group.add_command(comment)
    group.add_command(track)
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
@click.option("--name", "name", required=True, help="Style name (must already exist in the document).")
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
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor whose paragraph(s) to format.")
@click.option("--alignment", "alignment", default=None,
              type=click.Choice(["left", "center", "centre", "right", "justify"], case_sensitive=False),
              help="Paragraph alignment.")
@click.option("--left-indent", "left_indent", type=float, default=None, help="Left indent in points.")
@click.option("--right-indent", "right_indent", type=float, default=None, help="Right indent in points.")
@click.option("--first-line-indent", "first_line_indent", type=float, default=None, help="First-line indent in points.")
@click.option("--space-before", "space_before", type=float, default=None, help="Space before paragraph in points.")
@click.option("--space-after", "space_after", type=float, default=None, help="Space after paragraph in points.")
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
        lines.append(
            "  ".join(cell["text"].ljust(widths[i]) for i, cell in enumerate(row))
        )
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
@click.option("--values", "values", default=None, help="Optional JSON array of cell values for the new row.")
@click.pass_context
def table_add_row(ctx: click.Context, table_index: int, values: str | None) -> None:
    """Append a row to the table (atomic-undo)."""
    parsed: list[Any] | None = None
    if values is not None:
        try:
            parsed = json.loads(values)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--values must be a JSON array: {e}")
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
@click.option("--index", "index", type=int, required=True, help="1-based comment index (see `comment list`).")
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
@click.option("--index", "index", type=int, required=True, help="1-based comment index (see `comment list`).")
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
# exec --script ops.json
# ---------------------------------------------------------------------------


# Required fields per op kind. Validated up-front so a malformed payload
# raises a clean click.ClickException ("exec op 'write_bookmark' requires
# field 'name'") instead of a Python KeyError traceback that would land an
# LLM tool-use loop on exit code 1 with no actionable signal.
_OP_REQUIRED_FIELDS: dict[str, tuple[str, ...]] = {
    "write_bookmark": ("name", "text"),
    "write_cc": ("name", "text"),
    "insert_after_heading": ("heading", "text"),
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
}


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


@click.command(name="exec")
@click.option("--script", "script", required=True, type=click.Path(exists=True, dir_okay=False, path_type=Path), help="Path to ops JSON file.")
@click.pass_context
def exec_(ctx: click.Context, script: Path) -> None:
    """Apply a batch of ops in a single atomic-undo scope.

    Script shape: `{"label": "…", "ops": [{"op": "...", ...}, ...]}`. Set
    `"tracked": true` at the top level to record the whole batch as Word
    revisions (Track Changes is restored to its prior state afterwards).
    Supported ops: write_bookmark, write_cc, insert_after_heading, replace,
    find_replace, apply_style, format_paragraph, set_cell, add_row, delete_row,
    add_comment, resolve_comment, delete_comment. See docs/cli.md for each op's
    required and optional fields.
    """
    def go() -> None:
        payload = json.loads(script.read_text(encoding="utf-8"))
        label = str(payload.get("label") or f"CLI: exec {script.name}")
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
                emit(
                    {"ok": False, "ops_run": ops_run, "label": label, "failure": failure_meta},
                    as_text=not ctx.obj["as_json"],
                    text=f"failed at op {failure_meta['index']}: {failure_meta['error']}",
                )
                # Re-raise the original so _run() maps it to the right exit code
                # (e.g. anchor-not-found → 2, busy → 3, ambiguous → 5).
                raise failure_exc
    _run(ctx, go)
