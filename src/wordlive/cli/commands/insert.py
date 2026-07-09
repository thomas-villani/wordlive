"""Content-insertion commands."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _WRAP_CHOICES,
)


@click.command(name="insert")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert a new paragraph relative to (e.g. heading:1, para:3).",
)
@click.option(
    "--text",
    "text",
    default=None,
    help="Paragraph text to insert (literal — no markup). For inline formatting "
    "use --runs, or `insert-block` for a styled multi-paragraph run.",
)
@click.option(
    "--runs",
    "runs",
    default=None,
    help='JSON array of inline runs (e.g. \'[{"text":"Fast","bold":true},'
    '{"text":" — quick"}]\'), or \'-\' to read it from stdin. Each run is '
    "{text, bold?, italic?, underline?, style?}. Mutually exclusive with --text.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the new paragraph before the anchor instead of after it.",
)
@click.option(
    "--style", "style", default=None, help="Optional Word style name for the new paragraph."
)
@click.pass_context
def insert(
    ctx: click.Context,
    anchor_id: str,
    text: str | None,
    runs: str | None,
    before: bool,
    style: str | None,
) -> None:
    """Insert a new paragraph before/after any anchor (atomic-undo).

    Addresses anchors the same way every other command does — `--anchor-id`
    (headings, paragraphs, bookmarks, cells, ranges). Pass either `--text`
    (literal) or `--runs` (inline-formatted spans); for a contiguous run of
    several styled paragraphs in one shot, use `insert-block` instead. To insert
    text *inside* a paragraph at an offset, target a collapsed range:
    `replace --anchor-id range:120-120 --text "…"` (offsets come from
    `paragraphs` / `find`).
    """
    if (text is None) == (runs is None):
        raise click.UsageError("provide exactly one of --text or --runs")
    parsed_runs: list[Any] | None = None
    if runs is not None:
        raw = click.get_text_stream("stdin").read() if runs == "-" else runs
        try:
            parsed_runs = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--runs must be a JSON array: {e}") from e
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert {where} {anchor_id}"):
                if parsed_runs is not None:
                    anchor.insert_block([{"runs": parsed_runs, "style": style}], where=where)
                else:
                    assert text is not None  # xor-validated above; narrows for mypy
                    if before:
                        anchor.insert_paragraph_before(text, style=style)
                    else:
                        anchor.insert_paragraph_after(text, style=style)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "where": where,
                    "style": style,
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-block")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the block of paragraphs relative to (heading:/para:/end/…).",
)
@click.option(
    "--items",
    "items",
    required=True,
    help="JSON array of paragraphs, or '-' to read it from stdin. Each item is a "
    'string ("plain text") or an object {text|runs, style?}. `text` carries '
    "tiny inline markdown (**bold**, *italic*, ***both***; escape with \\*); "
    "`runs` is [{text, bold?, italic?, underline?, style?}].",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the block before the anchor instead of after it.",
)
@click.pass_context
def insert_block_cmd(ctx: click.Context, anchor_id: str, items: str, before: bool) -> None:
    """Insert a contiguous run of styled paragraphs at an anchor (atomic-undo).

    The multi-paragraph insert — drop a whole styled section (a feature list,
    a heading plus its body) in ONE op, in natural reading order, instead of a
    reverse-order storm of `insert` calls. Each item is one paragraph; `text`
    supports inline markdown and `runs` the structured form, so the "**Bold
    lead** — rest" bullet is a single op with no second formatting pass.

    Reports the spanning `range:START-END` of the inserted block, so a follow-up
    op can target the whole run — e.g. `list apply --anchor-id range:… --type
    bulleted` to bullet the section you just inserted.
    """
    raw = click.get_text_stream("stdin").read() if items == "-" else items
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--items must be a JSON array: {e}") from e
    if not isinstance(parsed, list):
        raise click.UsageError("--items must be a JSON array of paragraphs")
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert block {where} {anchor_id}"):
                rng = anchor.insert_block(parsed, where=where)
            emit(
                {
                    "ok": True,
                    "anchor_id": rng.anchor_id,
                    "paragraphs": len(parsed),
                    "where": where,
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {len(parsed)} paragraph(s) {where} {anchor_id} → {rng.anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-section")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the section relative to (heading:/para:/end/…).",
)
@click.option(
    "--heading", "heading", required=True, help="Heading text (inline **bold**/*italic* ok)."
)
@click.option(
    "--body",
    "body",
    required=True,
    help="JSON array of body paragraphs (insert-block items shape), or '-' for stdin.",
)
@click.option("--level", "level", type=int, default=1, show_default=True, help="Heading level 1–9.")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the section before the anchor instead of after it.",
)
@click.pass_context
def insert_section_cmd(
    ctx: click.Context, anchor_id: str, heading: str, body: str, level: int, before: bool
) -> None:
    """Insert a heading plus its body in one atomic op.

    The opinionated common case over `insert-block`: a `Heading {level}`
    paragraph followed by the body paragraphs, in reading order. `--body` is the
    same items shape `insert-block` takes. Reports the section's spanning
    `range:START-END`.
    """
    raw = click.get_text_stream("stdin").read() if body == "-" else body
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--body must be a JSON array: {e}") from e
    if not isinstance(parsed, list):
        raise click.UsageError("--body must be a JSON array of paragraphs")
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert section {where} {anchor_id}"):
                rng = anchor.insert_section(heading, parsed, level=level, where=where)
            emit(
                {"ok": True, "anchor_id": rng.anchor_id, "where": where},
                as_text=not ctx.obj["as_json"],
                text=f"inserted section {where} {anchor_id} → {rng.anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-markdown")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the markdown relative to (heading:/para:/end/…).",
)
@click.option(
    "--markdown",
    "markdown",
    required=True,
    help="Constrained-Markdown text, or '-' to read it from stdin. Subset: "
    "#/##/### headings, -/* bullets, 1. numbers, blank-line paragraphs, "
    "inline **bold**/*italic*. No code fences/nested lists/tables.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the markdown before the anchor instead of after it.",
)
@click.pass_context
def insert_markdown_cmd(ctx: click.Context, anchor_id: str, markdown: str, before: bool) -> None:
    """Insert a constrained-Markdown block as real Word structure (atomic-undo).

    Maps a tiny block dialect to paragraphs/headings/lists — a documented subset,
    not CommonMark. Path-bearing or multi-line input is easiest via `--markdown -`
    (stdin). Reports the spanning `range:START-END` of everything inserted.
    """
    md = click.get_text_stream("stdin").read() if markdown == "-" else markdown
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert markdown {where} {anchor_id}"):
                rng = anchor.insert_markdown(md, where=where)
            emit(
                {"ok": True, "anchor_id": rng.anchor_id, "where": where},
                as_text=not ctx.obj["as_json"],
                text=f"inserted markdown {where} {anchor_id} → {rng.anchor_id}",
            )

    _run(ctx, go)


@click.command(name="replace-section")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="The heading whose body to replace (heading:N). The heading itself is kept.",
)
@click.option(
    "--body", "body", default=None, help="JSON array of new body paragraphs, or '-' for stdin."
)
@click.option(
    "--markdown",
    "markdown",
    default=None,
    help="New body as constrained Markdown, or '-' for stdin.",
)
@click.pass_context
def replace_section_cmd(
    ctx: click.Context, anchor_id: str, body: str | None, markdown: str | None
) -> None:
    """Rewrite a heading's body, preserving the heading paragraph.

    Clears the span under `--anchor-id` (up to the next same-or-higher heading)
    and inserts the new body after the heading. Give exactly one of `--body`
    (insert-block items) or `--markdown` (constrained Markdown).
    """
    if (body is None) == (markdown is None):
        raise click.UsageError("give exactly one of --body or --markdown")
    if markdown is not None:
        new_body = click.get_text_stream("stdin").read() if markdown == "-" else markdown
    else:
        assert body is not None
        raw = click.get_text_stream("stdin").read() if body == "-" else body
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--body must be a JSON array: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--body must be a JSON array of paragraphs")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            if not hasattr(anchor, "replace_section_body"):
                raise click.UsageError(
                    f"replace-section needs a heading anchor; {anchor_id} is a {anchor.kind}"
                )
            with doc.edit(f"CLI: replace section {anchor_id}"):
                if markdown is not None:
                    rng = anchor.replace_section_body(new_body, markdown=True)
                else:
                    rng = anchor.replace_section_body(parsed)
            emit(
                {"ok": True, "anchor_id": rng.anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"replaced section body of {anchor_id} → {rng.anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-break")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the break relative to (e.g. heading:1, para:3, end).",
)
@click.option(
    "--kind",
    "kind",
    type=click.Choice(["page", "column", "section_next", "section_continuous"]),
    default="page",
    show_default=True,
    help="Break kind. Page is the common case; section breaks pair with `section`.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the break before the anchor instead of after it.",
)
@click.pass_context
def insert_break_cmd(ctx: click.Context, anchor_id: str, kind: str, before: bool) -> None:
    """Insert a page / column / section break at an anchor (atomic-undo).

    The explicit one-off break, the clean alternative to a literal form-feed
    paragraph. To make a *style* (e.g. every Heading 1) open a new page without
    a stray break character, prefer
    `format-paragraph --anchor-id ID --page-break-before` instead.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert {kind} break {where} {anchor_id}"):
                anchor.insert_break(kind, where=where)
            emit(
                {"ok": True, "anchor_id": anchor_id, "kind": kind, "where": where},
                as_text=not ctx.obj["as_json"],
                text=f"inserted {kind} break {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-field")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to insert the field at (e.g. footer:1:primary, end).",
)
@click.option(
    "--kind",
    "kind",
    type=click.Choice(["page", "numpages", "date", "time", "filename", "author", "title", "field"]),
    required=True,
    help="Field kind. Use 'field' with --text for a raw field code.",
)
@click.option(
    "--text",
    "text",
    default=None,
    help="Raw field code, required when --kind field (e.g. 'REF myBookmark').",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the field before the anchor instead of after it.",
)
@click.pass_context
def insert_field_cmd(
    ctx: click.Context, anchor_id: str, kind: str, text: str | None, before: bool
) -> None:
    """Insert a self-updating field (page number, date, …) at an anchor (atomic-undo).

    Page numbers belong in a footer/header: `insert-field --anchor-id
    footer:1:primary --kind page`. Refresh stale fields with `update-fields`.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert {kind} field {where} {anchor_id}"):
                anchor.insert_field(kind, text=text, where=where)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "applied": {"kind": kind, "text": text, "where": where},
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {kind} field {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="update-fields")
@click.pass_context
def update_fields_cmd(ctx: click.Context) -> None:
    """Refresh the document's fields — recompute page numbers, refs, dates (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: update fields"):
                doc.update_fields()
            emit(
                {"ok": True, "updated": True},
                as_text=not ctx.obj["as_json"],
                text="updated fields",
            )

    _run(ctx, go)


@click.command(name="insert-image")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor to insert the image relative to."
)
@click.option(
    "--path", "path", default=None, type=click.Path(path_type=Path), help="Path to the image file."
)
@click.option(
    "--base64", "b64", default=None, help="Base64 image data, or '-' to read base64 from stdin."
)
@click.option(
    "--wrap",
    "wrap",
    required=True,
    type=click.Choice(_WRAP_CHOICES),
    help="Layout / text-wrap (required). 'inline' stays in the text flow; "
    "'auto' floats Square when small, else top-bottom.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.option(
    "--block/--no-block",
    "block",
    default=False,
    show_default="--no-block",
    help="Place the image on its own new (Normal) line instead of in the anchor's text run.",
)
@click.option("--width", "width", type=float, default=None, help="Width in points (optional).")
@click.option("--height", "height", type=float, default=None, help="Height in points (optional).")
@click.option("--alt-text", "alt_text", default=None, help="Alternative (accessibility) text.")
@click.option(
    "--lock-aspect/--no-lock-aspect",
    "lock_aspect",
    default=True,
    show_default=True,
    help="Keep the image's aspect ratio when resizing.",
)
@click.pass_context
def insert_image_cmd(
    ctx: click.Context,
    anchor_id: str,
    path: Path | None,
    b64: str | None,
    wrap: str,
    before: bool,
    block: bool,
    width: float | None,
    height: float | None,
    alt_text: str | None,
    lock_aspect: bool,
) -> None:
    """Insert an image at any anchor, from a file or base64 (atomic-undo).

    Exactly one of --path / --base64 is required. --path is best for large
    images; base64 (or '--base64 -' from stdin) suits an LLM holding image
    data in memory. --wrap is required so layout is always explicit.
    """
    if (path is None) == (b64 is None):
        raise click.UsageError("pass exactly one of --path or --base64")
    if b64 == "-":
        b64 = click.get_text_stream("stdin").read()
    image: str | Path = path if path is not None else (b64 or "")
    where = "before" if before else "after"

    def go() -> None:
        # Screen a --path source against the policy *before* the COM/filesystem
        # probe: a UNC path's own existence check would authenticate to a remote
        # SMB server. Inside go() so _run() maps a denial to the right exit code.
        if path is not None:
            ctx.obj["policy"].screen_image_path(path)
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert image {where} {anchor_id}"):
                shape = anchor.insert_image(
                    image,
                    wrap=wrap,
                    where=where,
                    block=block,
                    width=width,
                    height=height,
                    alt_text=alt_text,
                    lock_aspect=lock_aspect,
                )
            # A floating image returns its shape:N handle (inline stays None).
            shape_id = shape.anchor_id if shape is not None else None
            text = f"inserted image {where} {anchor_id} (wrap={wrap})"
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "anchor": {"kind": anchor.kind, "name": anchor.name},
                    "shape": shape_id,
                    "wrap": wrap,
                    "where": where,
                    "block": block,
                },
                as_text=not ctx.obj["as_json"],
                text=(f"{text} -> {shape_id}" if shape_id else text),
            )

    _run(ctx, go)


@click.command(name="insert-text-box")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to attach the text box to.")
@click.option("--text", "text", required=True, help="Text box body.")
@click.option("--width", "width", default="200", show_default=True, help="Width (pt or '3in').")
@click.option("--height", "height", default="100", show_default=True, help="Height (pt or '2cm').")
@click.option(
    "--wrap",
    "wrap",
    type=click.Choice(["square", "tight", "through", "top-bottom", "front", "behind"]),
    default="square",
    show_default=True,
    help="How body text flows around the box.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Anchor before the anchor instead of after it.",
)
@click.option("--font", "font", default=None, help="Text font.")
@click.option("--size", "size", default=None, help="Font size (pt or unit string).")
@click.option("--bold/--no-bold", "bold", default=None, help="Bold the text.")
@click.option("--italic/--no-italic", "italic", default=None, help="Italicise the text.")
@click.option(
    "--align",
    "alignment",
    type=click.Choice(["left", "center", "right", "justify"]),
    default=None,
    help="Paragraph alignment.",
)
@click.option("--fill", "fill", default=None, help="Background colour (e.g. '#eeeeff' / 'navy').")
@click.option("--border-color", "border_color", default=None, help="Outline colour.")
@click.option("--no-border", "no_border", is_flag=True, default=False, help="No outline.")
@click.pass_context
def insert_text_box_cmd(
    ctx: click.Context,
    anchor_id: str,
    text: str,
    width: str,
    height: str,
    wrap: str,
    before: bool,
    font: str | None,
    size: str | None,
    bold: bool | None,
    italic: bool | None,
    alignment: str | None,
    fill: str | None,
    border_color: str | None,
    no_border: bool,
) -> None:
    """Insert a floating text box / pull quote at an anchor (atomic-undo)."""
    if no_border and border_color is not None:
        raise click.UsageError("pass either --no-border or --border-color (not both)")
    border: str | bool | None = False if no_border else border_color
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert text box {where} {anchor_id}"):
                shape = anchor.insert_text_box(
                    text,
                    width=width,
                    height=height,
                    wrap=wrap,
                    where=where,
                    font=font,
                    size=size,
                    bold=bold,
                    italic=italic,
                    alignment=alignment,
                    fill=fill,
                    border=border,
                )
            emit(
                {"ok": True, "anchor_id": shape.anchor_id, "wrap": wrap},
                as_text=not ctx.obj["as_json"],
                text=f"inserted text box {where} {anchor_id} -> {shape.anchor_id}",
            )

    _run(ctx, go)
