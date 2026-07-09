"""Reference apparatus: notes, TOC/index, citations, cross-refs, bookmarks, hyperlinks."""

from __future__ import annotations

from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_fields,
    _fmt_hyperlinks,
    _fmt_notes,
)


@click.command(name="insert-footnote")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor the footnote's reference mark attaches to (e.g. range:120-140, para:3).",
)
@click.option("--text", "text", required=True, help="The footnote body text.")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Place the reference mark before the anchor instead of after it.",
)
@click.pass_context
def insert_footnote_cmd(ctx: click.Context, anchor_id: str, text: str, before: bool) -> None:
    """Insert a footnote at an anchor (atomic-undo). Reports the new footnote:N."""
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert footnote {where} {anchor_id}"):
                note = anchor.insert_footnote(text, where=where)
                note_id = note.anchor_id
                index = note.index
            emit(
                {"ok": True, "anchor_id": anchor_id, "footnote": index, "note_id": note_id},
                as_text=not ctx.obj["as_json"],
                text=f"inserted {note_id} {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-endnote")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor the endnote's reference mark attaches to (e.g. range:120-140, para:3).",
)
@click.option("--text", "text", required=True, help="The endnote body text.")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Place the reference mark before the anchor instead of after it.",
)
@click.pass_context
def insert_endnote_cmd(ctx: click.Context, anchor_id: str, text: str, before: bool) -> None:
    """Insert an endnote at an anchor (atomic-undo). Reports the new endnote:N."""
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert endnote {where} {anchor_id}"):
                note = anchor.insert_endnote(text, where=where)
                note_id = note.anchor_id
                index = note.index
            emit(
                {"ok": True, "anchor_id": anchor_id, "endnote": index, "note_id": note_id},
                as_text=not ctx.obj["as_json"],
                text=f"inserted {note_id} {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-toc")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="start",
    show_default=True,
    help="Where to insert the TOC (default: the document start).",
)
@click.option(
    "--levels",
    "levels",
    default="1-3",
    show_default=True,
    help="Heading levels to include, as 'upper-lower' (e.g. 1-3).",
)
@click.option(
    "--heading-styles/--no-heading-styles",
    "heading_styles",
    default=True,
    show_default=True,
    help="Source entries from the built-in Heading styles.",
)
@click.option(
    "--hyperlinks/--no-hyperlinks",
    "hyperlinks",
    default=True,
    show_default=True,
    help="Make each entry a clickable link (and a real link in exported PDFs).",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert the TOC before the anchor instead of after it.",
)
@click.pass_context
def insert_toc_cmd(
    ctx: click.Context,
    anchor_id: str,
    levels: str,
    heading_styles: bool,
    hyperlinks: bool,
    before: bool,
) -> None:
    """Insert a table of contents (atomic-undo).

    Page numbers populate after repagination — run `update-fields` (or take a
    `snapshot`) before reading them.
    """
    upper_str, sep, lower_str = levels.partition("-")
    if not sep:
        raise click.UsageError("--levels must be 'upper-lower', e.g. 1-3")
    try:
        level_pair = (int(upper_str), int(lower_str))
    except ValueError as e:
        raise click.UsageError("--levels must be two integers, e.g. 1-3") from e
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert TOC {where} {anchor_id}"):
                anchor.insert_toc(
                    levels=level_pair,
                    use_heading_styles=heading_styles,
                    hyperlinks=hyperlinks,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "levels": list(level_pair),
                        "use_heading_styles": heading_styles,
                        "hyperlinks": hyperlinks,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted TOC {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="footnotes")
@click.pass_context
def footnotes_cmd(ctx: click.Context) -> None:
    """List the document's footnotes with their footnote:N id, text, and para:N."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.footnotes.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_notes(rows, "footnote"))

    _run(ctx, go)


@click.command(name="endnotes")
@click.pass_context
def endnotes_cmd(ctx: click.Context) -> None:
    """List the document's endnotes with their endnote:N id, text, and para:N."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.endnotes.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_notes(rows, "endnote"))

    _run(ctx, go)


@click.command(name="hyperlinks")
@click.pass_context
def hyperlinks_cmd(ctx: click.Context) -> None:
    """List the document's hyperlinks (text, destination, range:START-END id).

    The read mirror of `link`: each link's visible text, external `address` or
    internal `sub_address` bookmark, screen tip, and the range/para it sits in.
    Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.hyperlinks.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_hyperlinks(rows))

    _run(ctx, go)


@click.command(name="set-hyperlink")
@click.option(
    "--index", "index", type=int, required=True, help="1-based hyperlink index (see `hyperlinks`)."
)
@click.option("--address", "address", default=None, help="External URL to retarget to.")
@click.option(
    "--sub-address", "sub_address", default=None, help='In-document bookmark ("" clears it).'
)
@click.option("--text", "text", default=None, help="Visible link text.")
@click.option("--screen-tip", "screen_tip", default=None, help='Hover tooltip ("" clears it).')
@click.pass_context
def set_hyperlink_cmd(
    ctx: click.Context,
    index: int,
    address: str | None,
    sub_address: str | None,
    text: str | None,
    screen_tip: str | None,
) -> None:
    """Retarget or relabel an existing hyperlink in place (atomic-undo).

    Address the link by its 1-based --index (from `hyperlinks`). Pass at least
    one field; omitting one leaves it untouched. These retarget a link, they
    don't unlink it: --sub-address / --screen-tip clear with "", but --address /
    --text cannot be emptied (Word keeps a link pointing somewhere).
    """
    raw: dict[str, Any] = {
        "address": address,
        "sub_address": sub_address,
        "text": text,
        "screen_tip": screen_tip,
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one of --address/--sub-address/--text/--screen-tip")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: set hyperlink {index}"):
                doc.hyperlinks[index].update(**kwargs)
            emit(
                {"ok": True, "index": index, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"updated hyperlink {index}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="fields")
@click.pass_context
def fields_cmd(ctx: click.Context) -> None:
    """List the document's fields (kind, code, rendered result, range:START-END id).

    The read mirror of `insert-field`: each field's `kind` (the code's leading
    keyword — PAGE, REF, TOC, …), raw `code`, and last-rendered `result`. Run
    `update-fields` first to refresh stale results. Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.fields.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_fields(rows))

    _run(ctx, go)


@click.group(name="bookmark", hidden=True)
def bookmark() -> None:
    """Deprecated: use `write bookmark NAME --create --anchor-id ID`.

    Kept as a hidden alias for one release.
    """


@bookmark.command(name="add")
@click.argument("name")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor whose range the bookmark covers (e.g. heading:2, range:120-140).",
)
@click.pass_context
def bookmark_add(ctx: click.Context, name: str, anchor_id: str) -> None:
    """Deprecated alias for `write bookmark NAME --create --anchor-id ID`.

    Create a bookmark NAME over an anchor's range (atomic-undo). The prerequisite
    for internal links (`link --bookmark NAME`) and cross-references
    (`cross-ref --target bookmark:NAME`). NAME must start with a letter and
    contain only letters, digits, and underscores.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: add bookmark {name}"):
                doc.bookmarks.add(name, anchor_id)
            emit(
                {"ok": True, "bookmark": name, "anchor_id": anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"added bookmark:{name} over {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="link")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to turn into a hyperlink.")
@click.option("--url", "url", default=None, help="External link target (URL, mailto:, file path).")
@click.option(
    "--bookmark", "bookmark_target", default=None, help="Internal target: a bookmark name."
)
@click.option(
    "--text", "text", default=None, help="Visible link text (replaces the range content)."
)
@click.option("--screen-tip", "screen_tip", default=None, help="Hover tooltip.")
@click.pass_context
def link_cmd(
    ctx: click.Context,
    anchor_id: str,
    url: str | None,
    bookmark_target: str | None,
    text: str | None,
    screen_tip: str | None,
) -> None:
    """Turn an anchor into a hyperlink — external `--url` or internal `--bookmark` (atomic-undo)."""
    if (url is None) == (bookmark_target is None):
        raise click.UsageError("pass exactly one of --url or --bookmark")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: link {anchor_id}"):
                anchor.link_to(
                    address=url, bookmark=bookmark_target, text=text, screen_tip=screen_tip
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {"url": url, "bookmark": bookmark_target, "text": text},
                },
                as_text=not ctx.obj["as_json"],
                text=f"linked {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="cross-ref")
@click.option("--anchor-id", "anchor_id", required=True, help="Where to insert the reference.")
@click.option(
    "--target",
    "target",
    required=True,
    help="Anchor id to reference: bookmark:NAME | heading:N | footnote:N | endnote:N.",
)
@click.option(
    "--kind",
    "kind",
    type=click.Choice(["text", "page", "number", "above_below"]),
    default="text",
    show_default=True,
    help="What the reference shows.",
)
@click.option(
    "--hyperlink/--no-hyperlink",
    "hyperlink",
    default=True,
    show_default=True,
    help="Make the inserted reference a clickable jump.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def cross_ref_cmd(
    ctx: click.Context, anchor_id: str, target: str, kind: str, hyperlink: bool, before: bool
) -> None:
    """Insert a cross-reference to another anchor (atomic-undo).

    `--target` resolves a bookmark by name, a heading/footnote/endnote by its id.
    Refresh stale references (page numbers move) with `update-fields`.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: cross-ref {target} {where} {anchor_id}"):
                anchor.insert_cross_reference(target, kind=kind, hyperlink=hyperlink, where=where)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "target": target,
                        "kind": kind,
                        "hyperlink": hyperlink,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted cross-reference to {target} {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="caption")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to caption (e.g. a figure).")
@click.option(
    "--label", "label", default="Figure", show_default=True, help="Caption label (Figure/Table/…)."
)
@click.option("--text", "text", default=None, help="Caption title after the label and number.")
@click.option(
    "--position",
    "position",
    type=click.Choice(["above", "below"], case_sensitive=False),
    default=None,
    help="Place the caption above or below the anchor "
    "(default: above for a Table, below otherwise).",
)
@click.pass_context
def caption_cmd(
    ctx: click.Context, anchor_id: str, label: str, text: str | None, position: str | None
) -> None:
    """Insert a numbered caption (Figure 1, Table 2, …) as its own paragraph (atomic-undo).

    The caption always becomes its own Caption-styled paragraph; on a table cell
    it is placed above/below the whole table. Table captions default to above,
    figures to below — pass --position to override.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: caption {label} {anchor_id}"):
                anchor.insert_caption(label, text=text, position=position)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {"label": label, "text": text, "position": position},
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {label} caption at {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="mark-index-entry")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor whose range to index.")
@click.option(
    "--entry",
    "entry",
    required=True,
    help="Index text; use 'main:sub' for a subentry.",
)
@click.option(
    "--cross-reference",
    "cross_reference",
    default=None,
    help="Replace the page number with a 'see …' pointer.",
)
@click.option("--bold", "bold", is_flag=True, help="Bold the entry's page number.")
@click.option("--italic", "italic", is_flag=True, help="Italicise the entry's page number.")
@click.pass_context
def mark_index_entry_cmd(
    ctx: click.Context,
    anchor_id: str,
    entry: str,
    cross_reference: str | None,
    bold: bool,
    italic: bool,
) -> None:
    """Mark an anchor's range as a back-of-book index entry (atomic-undo).

    The per-term step; build the index itself with `insert-index`.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: mark index entry {anchor_id}"):
                anchor.mark_index_entry(
                    entry, cross_reference=cross_reference, bold=bold, italic=italic
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "entry": entry,
                        "cross_reference": cross_reference,
                        "bold": bold,
                        "italic": italic,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"marked index entry {entry!r} at {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-index")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="end",
    show_default=True,
    help="Where to insert the index (default: the document end).",
)
@click.option("--columns", "columns", type=int, default=2, show_default=True, help="Column count.")
@click.option("--run-in", "run_in", is_flag=True, help="Pack subentries into one paragraph.")
@click.option(
    "--right-align-page-numbers",
    "right_align",
    is_flag=True,
    help="Flush page numbers to the right margin.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def insert_index_cmd(
    ctx: click.Context,
    anchor_id: str,
    columns: int,
    run_in: bool,
    right_align: bool,
    before: bool,
) -> None:
    """Insert a back-of-book index from the marked entries (atomic-undo).

    Mark entries first with `mark-index-entry`. Page numbers populate after
    repagination — run `update-fields` (or take a `snapshot`) before reading them.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert index {where} {anchor_id}"):
                anchor.insert_index(
                    columns=columns,
                    run_in=run_in,
                    right_align_page_numbers=right_align,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "columns": columns,
                        "run_in": run_in,
                        "right_align_page_numbers": right_align,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted index {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="table-of-figures")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="start",
    show_default=True,
    help="Where to insert the table of figures (default: the document start).",
)
@click.option(
    "--label",
    "label",
    default="Figure",
    show_default=True,
    help="Caption label to gather (Figure/Table/Equation/…).",
)
@click.option(
    "--no-label",
    "no_label",
    is_flag=True,
    help="Drop the 'Figure 1' label prefix from each entry.",
)
@click.option(
    "--hyperlinks/--no-hyperlinks",
    "hyperlinks",
    default=True,
    show_default=True,
    help="Make each entry a clickable link.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def table_of_figures_cmd(
    ctx: click.Context,
    anchor_id: str,
    label: str,
    no_label: bool,
    hyperlinks: bool,
    before: bool,
) -> None:
    """Insert a table of figures built from captions of one label (atomic-undo).

    Page numbers populate after repagination — run `update-fields` (or take a
    `snapshot`) before reading them.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: table of figures {label} {where} {anchor_id}"):
                anchor.insert_table_of_figures(
                    label=label,
                    include_label=not no_label,
                    hyperlinks=hyperlinks,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "label": label,
                        "include_label": not no_label,
                        "hyperlinks": hyperlinks,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted table of figures ({label}) {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="bibliography-style")
@click.option(
    "--style",
    "style",
    required=True,
    help="Citation style, e.g. APA/MLA/Chicago/IEEE (build-dependent).",
)
@click.pass_context
def bibliography_style_cmd(ctx: click.Context, style: str) -> None:
    """Set the document's citation/bibliography style (atomic-undo).

    Refresh existing citations/bibliography with `update-fields` afterwards.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: bibliography style {style}"):
                doc.bibliography_style = style
            emit(
                {"ok": True, "applied": {"style": style}},
                as_text=not ctx.obj["as_json"],
                text=f"set bibliography style to {style!r}",
            )

    _run(ctx, go)


@click.command(name="add-source")
@click.option(
    "--type",
    "source_type",
    default="book",
    show_default=True,
    help="Source type (book/journal_article/conference_proceedings/case/…).",
)
@click.option(
    "--tag",
    "tag",
    default=None,
    help="Citation tag; auto-derived from first author + year if omitted.",
)
@click.option("--author", "authors", multiple=True, help="Author 'Last, First' (repeatable).")
@click.option("--title", "title", default=None, help="Source title.")
@click.option("--year", "year", default=None, help="Publication year.")
@click.option("--publisher", "publisher", default=None, help="Publisher.")
@click.option("--city", "city", default=None, help="City of publication.")
@click.option("--journal-name", "journal_name", default=None, help="Journal/periodical name.")
@click.option("--volume", "volume", default=None, help="Volume.")
@click.option("--issue", "issue", default=None, help="Issue.")
@click.option("--pages", "pages", default=None, help="Page range.")
@click.option("--url", "url", default=None, help="URL.")
@click.option("--edition", "edition", default=None, help="Edition.")
@click.option("--doi", "doi", default=None, help="DOI.")
@click.option(
    "--xml",
    "xml",
    default=None,
    help="Raw <b:Source> XML (escape hatch; supersedes the typed fields).",
)
@click.pass_context
def add_source_cmd(
    ctx: click.Context,
    source_type: str,
    tag: str | None,
    authors: tuple[str, ...],
    title: str | None,
    year: str | None,
    publisher: str | None,
    city: str | None,
    journal_name: str | None,
    volume: str | None,
    issue: str | None,
    pages: str | None,
    url: str | None,
    edition: str | None,
    doi: str | None,
    xml: str | None,
) -> None:
    """Add a bibliography source to the document's store (atomic-undo).

    Cite it with `insert-citation --tag TAG`, then list cited sources with
    `insert-bibliography`.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: add source"):
                if xml:
                    src = doc.sources.add_xml(xml)
                else:
                    src = doc.sources.add(
                        source_type,
                        tag=tag,
                        author=list(authors) or None,
                        title=title,
                        year=year,
                        publisher=publisher,
                        city=city,
                        journal_name=journal_name,
                        volume=volume,
                        issue=issue,
                        pages=pages,
                        url=url,
                        edition=edition,
                        doi=doi,
                    )
            emit(
                {
                    "ok": True,
                    "source": src.tag,
                    "applied": {"source_type": source_type, "tag": src.tag},
                },
                as_text=not ctx.obj["as_json"],
                text=f"added source {src.tag!r}",
            )

    _run(ctx, go)


@click.command(name="insert-citation")
@click.option("--anchor-id", "anchor_id", required=True, help="Where to insert the citation.")
@click.option("--tag", "tag", required=True, help="Source tag to cite.")
@click.option("--pages", "pages", default=None, help="Page locator, e.g. '15'.")
@click.option("--prefix", "prefix", default=None, help="Text before the citation, e.g. 'see '.")
@click.option("--suffix", "suffix", default=None, help="Text after the citation, e.g. ', at 12'.")
@click.option("--volume", "volume", default=None, help="Volume.")
@click.option("--suppress-author", "suppress_author", is_flag=True, help="Drop the author.")
@click.option("--suppress-year", "suppress_year", is_flag=True, help="Drop the year.")
@click.option("--suppress-title", "suppress_title", is_flag=True, help="Drop the title.")
@click.option("--locale", "locale", type=int, default=1033, show_default=True, help="Format LCID.")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def insert_citation_cmd(
    ctx: click.Context,
    anchor_id: str,
    tag: str,
    pages: str | None,
    prefix: str | None,
    suffix: str | None,
    volume: str | None,
    suppress_author: bool,
    suppress_year: bool,
    suppress_title: bool,
    locale: int,
    before: bool,
) -> None:
    """Insert an in-text citation referencing a source by tag (atomic-undo)."""
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: citation {tag} {where} {anchor_id}"):
                anchor.insert_citation(
                    tag,
                    pages=pages,
                    prefix=prefix,
                    suffix=suffix,
                    volume=volume,
                    suppress_author=suppress_author,
                    suppress_year=suppress_year,
                    suppress_title=suppress_title,
                    locale=locale,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "citation": tag,
                    "applied": {
                        "tag": tag,
                        "pages": pages,
                        "prefix": prefix,
                        "suffix": suffix,
                        "volume": volume,
                        "suppress_author": suppress_author,
                        "suppress_year": suppress_year,
                        "suppress_title": suppress_title,
                        "locale": locale,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted citation {tag!r} {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="insert-bibliography")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="end",
    show_default=True,
    help="Where to insert the bibliography (default: the document end).",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def insert_bibliography_cmd(ctx: click.Context, anchor_id: str, before: bool) -> None:
    """Insert a bibliography of the cited sources (atomic-undo).

    Page numbers/entries populate after repagination — run `update-fields` (or take
    a `snapshot`) before reading them.
    """
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert bibliography {where} {anchor_id}"):
                anchor.insert_bibliography(where=where)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": {"where": where}},
                as_text=not ctx.obj["as_json"],
                text=f"inserted bibliography {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="mark-citation")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor whose range to mark.")
@click.option(
    "--long",
    "long_citation",
    required=True,
    help="Full citation as it appears in the table.",
)
@click.option(
    "--short",
    "short_citation",
    default=None,
    help="Abbreviated form Word matches elsewhere (defaults to --long).",
)
@click.option(
    "--category",
    "category",
    default="cases",
    show_default=True,
    help="cases/statutes/other/rules/treatises/regulations/constitutional, or 1-16.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def mark_citation_cmd(
    ctx: click.Context,
    anchor_id: str,
    long_citation: str,
    short_citation: str | None,
    category: str,
    before: bool,
) -> None:
    """Mark an anchor's range as a table-of-authorities citation (atomic-undo).

    The per-authority step; build the table with `table-of-authorities`.
    """
    where = "before" if before else "after"
    cat: str | int = int(category) if category.isdigit() else category

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: mark citation {anchor_id}"):
                anchor.mark_citation(
                    long_citation, short_citation=short_citation, category=cat, where=where
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "long_citation": long_citation,
                        "short_citation": short_citation,
                        "category": category,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"marked citation {long_citation!r} at {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="table-of-authorities")
@click.option(
    "--anchor-id",
    "anchor_id",
    default="end",
    show_default=True,
    help="Where to insert the table (default: the document end).",
)
@click.option(
    "--category",
    "category",
    default="all",
    show_default=True,
    help="all/cases/statutes/other/rules/treatises/regulations/constitutional, or 1-16.",
)
@click.option("--no-passim", "no_passim", is_flag=True, help="Don't collapse 5+ refs to 'passim'.")
@click.option(
    "--no-keep-formatting",
    "no_keep_formatting",
    is_flag=True,
    help="Don't preserve each entry's character formatting.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def table_of_authorities_cmd(
    ctx: click.Context,
    anchor_id: str,
    category: str,
    no_passim: bool,
    no_keep_formatting: bool,
    before: bool,
) -> None:
    """Insert a table of authorities from the marked citations (atomic-undo).

    Mark citations first with `mark-citation`. Page numbers populate after
    repagination — run `update-fields` (or take a `snapshot`) before reading them.
    """
    where = "before" if before else "after"
    cat: str | int = int(category) if category.isdigit() else category

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: table of authorities {where} {anchor_id}"):
                anchor.insert_table_of_authorities(
                    category=cat,
                    passim=not no_passim,
                    keep_entry_formatting=not no_keep_formatting,
                    where=where,
                )
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "applied": {
                        "category": category,
                        "passim": not no_passim,
                        "keep_entry_formatting": not no_keep_formatting,
                        "where": where,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted table of authorities {where} {anchor_id}",
            )

    _run(ctx, go)
