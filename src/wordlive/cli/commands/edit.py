"""Text-editing commands (write / replace / prepend / append / delete)."""

from __future__ import annotations

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ...exceptions import AmbiguousMatchError
from ..main import _run, emit
from ._common import (
    _fmt_replace_summary,
)


@click.group(name="write")
def write() -> None:
    """Write structured values into the target document."""


@write.command(name="bookmark")
@click.argument("name")
@click.option("--text", "text", default=None, help="New text for an existing bookmark.")
@click.option(
    "--create",
    "create",
    is_flag=True,
    default=False,
    help="Create the bookmark over --anchor-id (instead of writing text).",
)
@click.option(
    "--anchor-id",
    "anchor_id",
    default=None,
    help="With --create: the anchor whose range the new bookmark covers "
    "(e.g. heading:2, range:120-140).",
)
@click.pass_context
def write_bookmark(
    ctx: click.Context, name: str, text: str | None, create: bool, anchor_id: str | None
) -> None:
    """Create a bookmark, or set an existing one's text (atomic-undo).

    Two modes:

    \b
      write bookmark NAME --text "…"                set an existing bookmark's text
      write bookmark NAME --create --anchor-id ID   create NAME over an anchor's range

    Creating a bookmark is the prerequisite for internal links
    (`link --bookmark NAME`) and cross-references (`cross-ref --target
    bookmark:NAME`). NAME must start with a letter and contain only letters,
    digits, and underscores.
    """
    if create:
        if text is not None:
            raise click.UsageError("--create and --text are mutually exclusive")
        if anchor_id is None:
            raise click.UsageError("--create requires --anchor-id")
    else:
        if anchor_id is not None:
            raise click.UsageError("--anchor-id is only valid with --create")
        if text is None:
            raise click.UsageError(
                "provide --text (write an existing bookmark) or "
                "--create --anchor-id ID (create a new one)"
            )

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            if create:
                assert anchor_id is not None  # guaranteed by the validation above
                with doc.edit(f"CLI: add bookmark {name}"):
                    doc.bookmarks.add(name, anchor_id)
                emit(
                    {"ok": True, "bookmark": name, "anchor_id": anchor_id, "created": True},
                    as_text=not ctx.obj["as_json"],
                    text=f"added bookmark:{name} over {anchor_id}",
                )
            else:
                assert text is not None  # guaranteed by the validation above
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


@click.command(name="delete-paragraph")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Paragraph anchor to delete (e.g. para:1, heading:2).",
)
@click.pass_context
def delete_paragraph_cmd(ctx: click.Context, anchor_id: str) -> None:
    """Delete the paragraph(s) at an anchor — text and the trailing mark (atomic-undo).

    Removes the whole paragraph so the surrounding text closes up (no empty line
    left, unlike `replace --text ""`). Useful for a stray leading empty paragraph.
    Deleting the document's last paragraph clears it but keeps Word's mandatory
    final mark.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: delete paragraph {anchor_id}"):
                doc.delete_paragraph(anchor_id)
            emit(
                {"ok": True, "anchor_id": anchor_id, "deleted": True},
                as_text=not ctx.obj["as_json"],
                text=f"deleted paragraph {anchor_id}",
            )

    _run(ctx, go)


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
    help="In find mode, replace only the Nth match (1-based).",
)
@click.option(
    "--mode",
    "mode",
    type=click.Choice(["fuzzy", "literal", "regex"]),
    default="fuzzy",
    show_default=True,
    help="In find mode: fuzzy (tolerant), literal (exact), or regex (Python; --text may use \\1).",
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
    mode: str,
) -> None:
    """Replace text. Either at an anchor (entire range) or via find."""
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
                        mode=mode,
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
