"""The wordlive MCP server: four dispatch tools over a COM worker thread.

Tools (all prefixed `word_`):
  - `word_read`     — every read, dispatched on `command`.
  - `word_write`    — every single atomic-undo write, dispatched on `command`.
  - `word_exec`     — apply a batch of ops as one atomic undo (the power tool).
  - `word_snapshot` — render page(s) to PNG so the model can *see* the layout.

Plus a `wordlive://guide` resource holding the full agent guide.

A handful of dispatch tools (rather than one tool per verb) keeps the client's
tool list — and the context cost of its schemas — small. The op vocabulary for
`word_exec` and the anchor model are taught by the guide resource, not by dozens
of schemas.

All Word access funnels through a single `ComWorker` thread (see `_worker`), so
COM stays on one apartment-initialised thread and concurrent calls serialise.
`WordliveError`s are translated to MCP tool errors carrying a stable `code` and
a `retryable` hint (the same taxonomy as the CLI's exit codes).
"""

from __future__ import annotations

import base64
import json
from typing import TYPE_CHECKING, Any, Literal

from .._guide import skill_body
from .._paths import PathPolicy
from ..exceptions import WordliveError
from ._common import _INSTRUCTIONS, _error_payload, _image_format, _tool_error
from ._exec import _exec_impl
from ._read import _read_impl
from ._snapshot import _snapshot_impl
from ._worker import ComWorker, Worker
from ._write import _build_write_op, _write_impl

if TYPE_CHECKING:
    from mcp.server.fastmcp import FastMCP

__all__ = [
    "build_server",
    "main",
    "_read_impl",
    "_write_impl",
    "_build_write_op",
    "_exec_impl",
    "_snapshot_impl",
    "_error_payload",
    "_image_format",
    "_tool_error",
]


def build_server(worker: Worker | None = None) -> FastMCP:
    """Build the FastMCP server. `worker` is injectable for tests (defaults to a
    real COM worker thread). Importing the `mcp` extra is deferred to here."""
    try:
        from mcp.server.fastmcp import FastMCP, Image
        from mcp.types import TextContent
    except ImportError as e:  # pragma: no cover - exercised only without the extra
        raise RuntimeError(
            "the wordlive MCP server requires the 'mcp' extra: "
            'pip install "wordlive[mcp]" (or "wordlive[mcp,snapshot]" for snapshots)'
        ) from e

    w: Worker = worker if worker is not None else ComWorker()
    # Saving is default-deny: the operator opts in by configuring WORDLIVE_SAVE_DIRS
    # at launch (image-source paths optionally restricted via WORDLIVE_IMAGE_DIRS).
    policy = PathPolicy.from_env()
    mcp = FastMCP("wordlive", instructions=_INSTRUCTIONS)

    @mcp.tool()
    def word_read(
        command: Literal[
            "status",
            "guide",
            "outline",
            "paragraphs",
            "find",
            "read_bookmark",
            "read_cc",
            "read_section",
            "to_markdown",
            "to_html",
            "digest",
            "between",
            "nearest_heading",
            "find_paragraphs",
            "table_list",
            "table_read",
            "table_records",
            "styles",
            "comments",
            "revisions",
            "read_text",
            "track",
            "sections",
            "footnotes",
            "endnotes",
            "images",
            "read_image",
            "equations",
            "charts",
            "shapes",
            "hyperlinks",
            "fields",
            "properties",
            "watermark",
            "variables",
            "proofing",
            "lint",
            "checkpoint",
            "diff",
            "format_info",
            "list_levels",
            "location",
            "stats",
            "theme",
            "themes",
        ],
        doc: str | None = None,
        name: str | None = None,
        text: str | None = None,
        in_anchor: str | None = None,
        mode: str | None = None,
        anchor_id: str | None = None,
        heading: str | None = None,
        table: int | None = None,
        all_paragraphs: bool = False,
        start: int | None = None,
        count: int | None = None,
        view: str | None = None,
        start_anchor: str | None = None,
        end_anchor: str | None = None,
        inclusive: bool = False,
        direction: str | None = None,
        limit: int | None = None,
        min_score: float | None = None,
        budget: int | None = None,
        depth: int | None = None,
        rules: Any = None,
        within: str | None = None,
        profile: Any = None,
        include: str | None = None,
        checkpoint: Any = None,
        cp_a: Any = None,
        cp_b: Any = None,
    ) -> Any:
        """Read from the open Word document. Dispatch on `command`:

        guide (no Word needed — returns the full agent guide: anchor model, the
        word_exec op vocabulary, and every field; read this first) ·
        status (no doc needed; reports name/path/saved/is_active per open doc) ·
        outline [all_paragraphs] · paragraphs [start,count] ·
        find {text,[in_anchor],[mode=fuzzy|literal|regex]} · read_bookmark {name} · read_cc {name} ·
        read_section {heading | anchor_id} (body under a heading) ·
        to_markdown {[within]} (serialise the document — or one anchor's range — to
        clean Markdown: headings, lists, **bold**/*italic*, GFM tables, ![alt](image:N),
        [text](url); the read mirror of insert_markdown, lossy by design; returns
        {markdown}) ·
        to_html {[within]} (same as to_markdown but an HTML fragment; returns {html}) ·
        digest {[budget=6000],[depth]} (a token-budgeted, structure-aware read of the
        WHOLE document — headings verbatim (each tagged with its heading:N anchor),
        tables as one-line shape stubs, body sampled to fit budget; every anchor stays
        addressable so you can drill in with to_markdown(within=…). Loads a large doc
        into context cheaply; returns {digest} markdown) ·
        between {start_anchor,end_anchor,[inclusive]} (content spanning two anchors —
        e.g. the block between two headings; default excludes both heading lines,
        inclusive covers them; returns a range:START-END id + text) ·
        nearest_heading {anchor_id,[direction=before|after]} (the heading nearest a
        position — before=enclosing/preceding, after=next; an outline row or null) ·
        find_paragraphs {text,[limit=5],[min_score=0.6]} (fuzzy-rank paragraphs by
        similarity to text — typo/paraphrase tolerant, unlike the exact-substring find;
        returns para:N candidates with scores) ·
        table_list · table_read {table} ·
        table_records {table} (body rows as dicts keyed by the header row — the
        read mirror of building a table from records) ·
        styles · comments · revisions (tracked changes: type/author/text/range per
        change) · read_text {anchor_id,[view]} (an anchor's text; view=raw|final|
        original|segments resolves tracked changes — final=as if accepted,
        original=as if rejected, segments=per-run insert/delete breakdown) ·
        track (is Track Changes on?) · sections · footnotes · endnotes ·
        images (embedded pictures: image:N id, mime, size, crop, alt, para) ·
        read_image {anchor_id} (SEE an embedded picture: returns it as an inline
        image block — like word_snapshot — plus a {anchor_id,mime,bytes} label;
        pass an image:N id or any single-image anchor) ·
        equations (math zones: equation:N id, type, linear preview, para) ·
        charts (Excel-backed charts: chart:N id, kind, title, chart_style, has_legend, para) ·
        shapes (floating shapes: shape:N id, shape_type=text_box|picture|wordart, size, wrap,
        wrap_side, crop, alt_text, para — text boxes / floating images / WordArt; the restyle handles) ·
        hyperlinks (links: text, address/sub_address, range:START-END id — the read
        mirror of add_hyperlink) · fields (PAGE/REF/TOC/…: kind, code, rendered
        result, range id — the read mirror of insert_field) ·
        properties (document metadata: {builtin, custom} name→value bags) ·
        watermark ({text, sections}, or null — the text watermark behind the pages;
        the read mirror of watermark set/remove) ·
        variables (invisible named storage: {name: value}) ·
        proofing (spelling/grammar errors with counts + flagged runs, and readability
        statistics — heavier than stats; it (re)checks the document) ·
        lint {[rules],[within],[profile]} (audit publishing-quality defects: dangling
        headings, multi-page tables with no repeating header, split numbered lists, direct
        formatting drifted from the style — severity-ranked findings, each fixable one
        carrying the op regularize would run, with adds_content=true marking fixes that
        insert/delete content (regularize withholds those unless allow_content); rules
        selects ids/tags or {exclude:[…]};
        profile is a house-style config — a path or inline object — that enables policy
        rules (body-justified, body-line-spacing, table-numeric-right-align) + their
        targets) ·
        checkpoint {[include=text|text+style|text+format],[within]} (an opaque,
        serialisable structural fingerprint of the document now — store it, edit,
        then diff; the only way to answer "what changed" since Word has no
        content-change event) ·
        diff {checkpoint | cp_a,cp_b} (content-aligned change list: pass a stored
        `checkpoint` to diff it against the document now, or `cp_a`+`cp_b` to diff
        two stored checkpoints — each change is replace/insert/delete/restyle/reformat
        carrying the current para:N) ·
        format_info {anchor_id} (effective paragraph + character formatting at an anchor,
        each field with its style baseline and an override flag, plus font.mixed — the
        read mirror of format_paragraph/format_run, and the linter's substrate) ·
        list_levels {anchor_id} (the per-level format of the list at an anchor: one
        {level,kind,format,style,trailing,number_position,text_position,font} per template
        level — the read mirror of the list `format` action) ·
        location {anchor_id} (where an anchor sits in the laid-out document:
        page/end_page span, line, column, in_table — "what page is this on"
        without a snapshot) ·
        stats (one-shot document summary: page/word/character/paragraph/line
        counts plus section/heading/table/image/equation/comment/revision counts
        and saved) ·
        theme (the document theme: 12 brand colours as #RRGGBB + major/minor fonts —
        the read mirror of apply_theme/set_theme_*) ·
        themes (the built-in themes, colour schemes, and font schemes Office ships —
        the names apply_theme/set_theme_* accept).
        `doc` targets a document by name (default: active).
        """
        params = {
            "doc": doc,
            "name": name,
            "text": text,
            "in_anchor": in_anchor,
            "mode": mode,
            "anchor_id": anchor_id,
            "heading": heading,
            "table": table,
            "all_paragraphs": all_paragraphs,
            "start": start,
            "count": count,
            "view": view,
            "start_anchor": start_anchor,
            "end_anchor": end_anchor,
            "inclusive": inclusive,
            "direction": direction,
            "limit": limit,
            "min_score": min_score,
            "budget": budget,
            "depth": depth,
            "rules": rules,
            "within": within,
            "profile": profile,
            "include": include,
            "checkpoint": checkpoint,
            "cp_a": cp_a,
            "cp_b": cp_b,
        }
        try:
            result = _read_impl(w, command, params)
        except WordliveError as exc:
            raise _tool_error(exc) from exc
        if command == "read_image" and isinstance(result, dict) and result.get("base64"):
            # Hand the model the actual pixels (an ImageContent block, like
            # word_snapshot) plus a compact metadata block — not base64 text it
            # can't see. The raw base64 stays out of the returned content.
            meta = {k: result[k] for k in ("anchor_id", "mime", "bytes") if k in result}
            data = base64.b64decode(result["base64"])
            return [
                TextContent(type="text", text=json.dumps(meta)),
                Image(data=data, format=_image_format(result.get("mime"))).to_image_content(),
            ]
        return result

    @mcp.tool()
    def word_write(
        command: Literal[
            "insert",
            "insert_block",
            "insert_section",
            "insert_markdown",
            "replace_section",
            "delete_paragraph",
            "append",
            "prepend",
            "replace",
            "write_bookmark",
            "write_cc",
            "apply_style",
            "format_paragraph",
            "format_run",
            "set_shading",
            "set_borders",
            "cell_valign",
            "drop_cap",
            "add_tab_stop",
            "add_style",
            "set_style",
            "list",
            "comment",
            "revision",
            "table",
            "header",
            "footer",
            "track",
            "watermark",
            "text_box",
            "insert_image",
            "insert_equation",
            "insert_chart",
            "format_chart",
            "format_axis",
            "add_trendline",
            "set_series_color",
            "format_series",
            "add_error_bars",
            "set_shape_wrap",
            "set_shape_crop",
            "set_shape_position",
            "set_shape_size",
            "format_shape",
            "set_shape_alt_text",
            "set_shape_text",
            "set_shape_rotation",
            "set_shape_z_order",
            "set_shape_text_frame",
            "replace_shape_image",
            "delete_shape",
            "group_shapes",
            "ungroup_shape",
            "set_image_alt_text",
            "set_image_size",
            "set_image_crop",
            "insert_break",
            "insert_field",
            "update_fields",
            "regularize",
            "set_property",
            "delete_property",
            "set_variable",
            "delete_variable",
            "insert_footnote",
            "insert_endnote",
            "insert_toc",
            "add_bookmark",
            "pin",
            "pin_outline",
            "add_hyperlink",
            "set_hyperlink",
            "insert_cross_reference",
            "insert_caption",
            "create_content_control",
            "set_cc_properties",
            "set_cc_items",
            "mark_index_entry",
            "insert_index",
            "insert_table_of_figures",
            "set_bibliography_style",
            "add_source",
            "insert_citation",
            "insert_bibliography",
            "mark_citation",
            "insert_table_of_authorities",
            "apply_theme",
            "set_theme_colors",
            "set_theme_fonts",
            "page_setup",
            "save",
            "save_as",
            "export_pdf",
        ],
        doc: str | None = None,
        anchor_id: str | None = None,
        text: str | None = None,
        runs: list[Any] | None = None,
        items: list[Any] | None = None,
        name: str | None = None,
        bind: str | bool | None = None,
        style: str | None = None,
        before: bool = False,
        paragraph: bool = True,
        find: str | None = None,
        all: bool = False,
        occurrence: int | None = None,
        in_anchor: str | None = None,
        action: str | None = None,
        type: str | None = None,
        author: str | list[str] | None = None,
        index: int | None = None,
        table: int | None = None,
        row: int | None = None,
        col: int | None = None,
        rows: int | None = None,
        cols: int | None = None,
        data: list[Any] | None = None,
        header: bool | None = None,
        heading: str | bool | None = None,
        body: list[Any] | str | None = None,
        markdown: str | None = None,
        level: int | None = None,
        allow_break: bool | None = None,
        first_row: bool | None = None,
        last_row: bool | None = None,
        first_column: bool | None = None,
        last_column: bool | None = None,
        banded_rows: bool | None = None,
        banded_columns: bool | None = None,
        values: list[Any] | dict[str, Any] | None = None,
        record: dict[str, Any] | None = None,
        key: str | None = None,
        column: str | None = None,
        value: str | int | float | bool | None = None,
        custom: bool | None = None,
        mode: str | None = None,
        section: int | None = None,
        which: str = "primary",
        on: bool | None = None,
        alignment: str | None = None,
        left_indent: float | None = None,
        right_indent: float | None = None,
        first_line_indent: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        line_spacing: str | float | None = None,
        page_break_before: bool | None = None,
        keep_together: bool | None = None,
        keep_with_next: bool | None = None,
        widow_control: bool | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikethrough: bool | None = None,
        font: str | None = None,
        size: str | float | None = None,
        color: str | None = None,
        highlight: str | None = None,
        subscript: bool | None = None,
        superscript: bool | None = None,
        small_caps: bool | None = None,
        all_caps: bool | None = None,
        spacing: str | float | None = None,
        fill: str | None = None,
        pattern: str | None = None,
        sides: str | None = None,
        line_style: str | None = None,
        weight: float | None = None,
        lines: int | None = None,
        distance: str | float | None = None,
        position: str | float | None = None,
        align: str | None = None,
        leader: str | None = None,
        based_on: str | None = None,
        next_style: str | None = None,
        kind: str | None = None,
        wrap: str | None = None,
        image_base64: str | None = None,
        path: str | None = None,
        block: bool | None = None,
        width: float | None = None,
        height: float | None = None,
        alt_text: str | None = None,
        lock_aspect: bool | None = None,
        unicodemath: str | None = None,
        latex: str | None = None,
        mathml: str | None = None,
        display: bool | None = None,
        margins: str | float | None = None,
        top_margin: str | float | None = None,
        bottom_margin: str | float | None = None,
        left_margin: str | float | None = None,
        right_margin: str | float | None = None,
        gutter: str | float | None = None,
        orientation: str | None = None,
        paper_size: str | None = None,
        columns: int | None = None,
        column_spacing: str | float | None = None,
        levels: list[int] | None = None,
        use_heading_styles: bool | None = None,
        hyperlinks: bool | None = None,
        url: str | None = None,
        bookmark: str | None = None,
        screen_tip: str | None = None,
        target: str | None = None,
        hyperlink: bool | None = None,
        label: str | None = None,
        title: str | None = None,
        tag: str | None = None,
        lock_contents: bool | None = None,
        lock_control: bool | None = None,
        entry: str | None = None,
        cross_reference: str | None = None,
        run_in: bool | None = None,
        right_align_page_numbers: bool | None = None,
        include_label: bool | None = None,
        where: str | None = None,
        source_type: str | None = None,
        year: str | None = None,
        publisher: str | None = None,
        city: str | None = None,
        journal_name: str | None = None,
        volume: str | None = None,
        issue: str | None = None,
        pages: str | None = None,
        edition: str | None = None,
        doi: str | None = None,
        xml: str | None = None,
        prefix: str | None = None,
        suffix: str | None = None,
        suppress_author: bool | None = None,
        suppress_year: bool | None = None,
        suppress_title: bool | None = None,
        locale: int | None = None,
        long_citation: str | None = None,
        short_citation: str | None = None,
        category: str | int | None = None,
        passim: bool | None = None,
        keep_entry_formatting: bool | None = None,
        entry_separator: str | None = None,
        page_range_separator: str | None = None,
        theme: str | None = None,
        scheme: str | None = None,
        major: str | None = None,
        minor: str | None = None,
        colors: dict[str, str] | None = None,
        overwrite: bool = False,
        from_page: int | None = None,
        to_page: int | None = None,
        layout: str | None = None,
        semitransparent: bool | None = None,
        remove: bool | None = None,
        border: str | bool | None = None,
        border_weight: str | float | None = None,
        left: str | float | None = None,
        top: str | float | None = None,
        relative_to: str | None = None,
        degrees: float | None = None,
        order: str | None = None,
        margin_left: str | float | None = None,
        margin_right: str | float | None = None,
        margin_top: str | float | None = None,
        margin_bottom: str | float | None = None,
        word_wrap: bool | None = None,
        side: str | None = None,
        distance_top: str | float | None = None,
        distance_bottom: str | float | None = None,
        distance_left: str | float | None = None,
        distance_right: str | float | None = None,
        crop_left: str | float | None = None,
        crop_top: str | float | None = None,
        crop_right: str | float | None = None,
        crop_bottom: str | float | None = None,
        shapes: list[str] | None = None,
        rules: Any = None,
        within: str | None = None,
        profile: Any = None,
        dry_run: bool | None = None,
        allow_content: bool | None = None,
    ) -> dict[str, Any]:
        """Make one atomic-undo edit to the open Word document. Dispatch on `command`:

        insert {anchor_id, text|runs, [before,style,bind]} — text is literal; runs is
            [{text,bold?,italic?,underline?,code?,style?}] for inline-formatted spans;
            bind ("slug" or true) mints a durable pin on the new paragraph ·
        insert_block {anchor_id, items, [before,bind]} — a contiguous run of styled
            paragraphs in one op; each item is "plain text" or {text|runs, style?}
            (text carries **bold**/*italic* markdown); returns the block's
            range:START-END (and pin:CODE when bind is set) ·
        insert_section {anchor_id, heading, body, [level=1, before]} — a Heading {level}
            paragraph plus its body (body = insert_block items shape) in one op ·
        insert_markdown {anchor_id, markdown, [before]} — a constrained-Markdown block
            as real Word structure: #/##/### headings, -/* bullets, 1. numbers,
            blank-line paragraphs, inline **bold**/*italic* (a subset, not CommonMark) ·
        replace_section {anchor_id=heading:N, body | markdown} — rewrite a heading's body
            (up to the next same-or-higher heading), keeping the heading; one of body/markdown ·
        delete_paragraph {anchor_id} — remove the paragraph(s) at an anchor, mark included ·
        append/prepend {text,[style]} — new final/first paragraph; pass paragraph=false
            to continue the adjacent paragraph inline (an inline append takes no style) ·
        replace {text, find|anchor_id, [all,occurrence,in_anchor,mode=fuzzy|literal|regex]} ·
        write_bookmark/write_cc {name,text} · apply_style {anchor_id,name} ·
        format_paragraph {anchor_id,[alignment,*_indent,space_*,line_spacing,page_break_before,keep_together,keep_with_next,widow_control]} ·
        format_run {anchor_id,[bold,italic,underline,strikethrough,font,size,color,
            highlight,subscript,superscript,small_caps,all_caps,spacing]} — colour is a
            name/hex; highlight is a named palette colour; size/spacing accept unit strings ·
        set_shading {anchor_id,fill} — fill colour of a range/cell ·
        set_borders {anchor_id,[sides=all|box|top|bottom|left|right|horizontal|vertical,
            line_style=single|double|dot|dash|none,weight,color]} — anchor_id may be a
            cell (table:N:R:C), a whole row (table:N:row:R), or a whole column
            (table:N:col:C); shading/borders/apply_style/format_run all take those too ·
        cell_valign {anchor_id=table:N:R:C,align=top|center|bottom} — vertical alignment of a cell ·
        drop_cap {anchor_id,[position=dropped|margin|none,lines=3,distance,font]} —
            an oversized initial letter on the anchor's paragraph (a real Word DropCap) ·
        add_tab_stop {anchor_id,position,[align=left|center|right|decimal|bar,
            leader=dots|dashes|lines|none]} ·
        add_style {name,[type=paragraph|character|table|list,based_on,next_style]} —
            define a new style ·
        set_style {name,[bold,italic,underline,font,size,color,alignment,space_*,line_spacing,
            based_on,next_style]} — set an existing style's font/paragraph defaults ·
        list {anchor_id,action=apply|remove|restart|indent|outdent|format,[type,levels]} —
            action=format authors a custom multi-level list from a `levels` array (see word_exec apply_list_format) ·
        comment {action=add|resolve|delete,...} ·
        revision {action=accept|reject (index) | accept_all|reject_all ([anchor_id] scopes to that
            range, else whole doc)} — resolve tracked changes; accept/reject renumber the rest ·
        table {action=set_cell|add_row|delete_row|add_column|delete_column|merge_cells|split_cell|
               append_record|update_row|set_heading_row|autofit|
               set_style|set_alignment|set_borders|set_banding|create|delete,
               create needs anchor_id and [rows,cols] (optional when data is given —
               inferred from it),[style,header,data,before]; data is a 2-D array OR
               records (a list of objects whose keys become a header row);
               append_record {table,record} — append a row from a {header: value} object;
               update_row {table,key,values,[column]} — set cells (values={header: value})
               on the first row whose key-column (column=, default first) equals key;
               set_heading_row {table,[row=1,heading=true,allow_break]} — repeating header row;
               add_column {table,[values]} — append a column (values fill top-to-bottom);
               delete_column {table,column} — fails on a merged/mixed-width table (delete its cells via table:N:R:C);
               merge_cells {table,from,to} — merge the rectangle between two cells ([row,col] or "R:C"); makes the table non-uniform;
               split_cell {table,cell,[rows=1,cols=2]} — split one cell ([row,col] or "R:C") into a grid;
               autofit {table,[mode=content|window|fixed]} — resize columns to content/window or pin them;
               set_style {table,style} — restyle an existing table (restyle first, then cell overrides);
               set_alignment {table,alignment=left|center|right} — the whole table across the page;
               set_borders {table,[sides,style|line_style,weight,color]} — the whole grid in one call;
               set_banding {table,[first_row,last_row,first_column,last_column,banded_rows,banded_columns]} —
               toggle table-style options (needs a real table style applied to show)} ·
        insert_break {anchor_id,[kind=page|column|section_next|section_continuous,before]} ·
        insert_field {anchor_id,kind=page|numpages|date|time|filename|author|title|field,[text,before]} —
            a self-updating field; put page numbers in a footer; kind=field takes a raw code in text ·
        update_fields {} — recompute the document's fields (page numbers, refs, dates) ·
        set_property {name,value,[custom]} — set a built-in document property (Title/Author/
            Subject/Keywords/…) or, with custom=true, a custom one (created if absent) ·
        delete_property {name} — delete a custom property (built-ins can't be removed) ·
        set_variable {name,value} — create/update a document variable (DOCVARIABLE storage) ·
        delete_variable {name} — delete a document variable ·
        insert_footnote {anchor_id,text,[before]} / insert_endnote {anchor_id,text,[before]} —
            a note anchored to a range; the new footnote:N/endnote:N is in result ·
        insert_toc {[anchor_id=start],levels=[upper,lower],use_heading_styles,hyperlinks,[before]} —
            a table of contents; run update_fields after to populate page numbers ·
        add_bookmark {name,anchor_id} — create a named bookmark over an anchor's range ·
        pin {anchor_id,[name]} — plant a DURABLE handle on an anchor; returns pin:CODE that
            survives the inserts/deletes that renumber para:N / heading:N (name = a readable
            slug like budget-intro, else a random code) ·
        pin_outline {[levels]} — pin every heading at once; returns {heading:N: pin:CODE}
            (idempotent; levels = an inclusive [lo,hi] band) ·
        add_hyperlink {anchor_id, url | bookmark, [text,screen_tip]} — external URL or internal
            bookmark jump; text sets the visible link text ·
        set_hyperlink {index, [url, bookmark, text, screen_tip]} — retarget/relabel an existing link
            in place (index is 1-based, from word_read hyperlinks); url=external, bookmark=in-document;
            pass at least one; bookmark/screen_tip clear with "", but url/text can't be emptied
            (delete the link to unlink) ·
        insert_cross_reference {anchor_id,target,[kind=text|page|number|above_below,hyperlink,before]} —
            target is a bookmark:/heading:/footnote:/endnote: id ·
        insert_caption {anchor_id,[label=Figure,text,position=above|below]} — a numbered
            caption in its own paragraph (Table defaults above, else below) ·
        create_content_control {anchor_id,[kind=rich_text|text|picture|combo_box|dropdown|date|
            checkbox|building_block|group|repeating_section, title, tag, items, where=wrap|before|after,
            lock_contents, lock_control]} — a form control; where=wrap surrounds the anchor's range
            (else insert an empty one); title makes it addressable as cc:TITLE; items=[..] fills a
            combo_box/dropdown (each a string or {text,value}) ·
        set_cc_properties {anchor_id=cc:NAME, [title, tag, lock_contents, lock_control]} — re-set a
            control's metadata in place; pass at least one; "" clears title/tag; a rename changes
            its cc:NAME id ·
        set_cc_items {anchor_id=cc:NAME, items} — replace a combo_box/dropdown's choice list (items
            replaces, not appends; each a string or {text,value}) ·
        mark_index_entry {anchor_id, entry, [cross_reference, bold, italic]} — mark a range as a
            back-of-book index entry (use "main:sub" for a subentry); build the list with insert_index ·
        insert_index {[anchor_id=end], columns=2, [run_in, right_align_page_numbers, before]} —
            gather marked entries into an index; run update_fields after to populate page numbers ·
        insert_table_of_figures {[anchor_id=start], label=Figure, [include_label, hyperlinks,
            right_align_page_numbers, before]} — a table of captions of one label; run update_fields after ·
        set_bibliography_style {style} — set the citation style (APA/MLA/Chicago/IEEE/…;
            build-dependent); run update_fields after to re-render citations/bibliography ·
        add_source {source_type=book|journal_article|conference_proceedings|report|web_site|case|…,
            [tag, author, title, year, publisher, city, journal_name, volume, issue, pages, url,
            edition, doi] | xml} — register a bibliography source; author is "Last, First" (or a
            list); tag auto-derives from author+year; xml is a raw <b:Source> escape hatch ·
        insert_citation {anchor_id, tag, [pages, prefix, suffix, volume, suppress_author,
            suppress_year, suppress_title, locale=1033, before]} — an in-text citation of a source ·
        insert_bibliography {[anchor_id=end], [before]} — the reference list of cited sources;
            run update_fields after to populate it ·
        mark_citation {anchor_id, long_citation, [short_citation, category=cases|statutes|other|
            rules|treatises|regulations|constitutional|1-16, before]} — mark a range as a
            table-of-authorities citation; build the table with insert_table_of_authorities ·
        insert_table_of_authorities {[anchor_id=end], category=all|cases|…|1-16, [passim,
            keep_entry_formatting, entry_separator, page_range_separator, before]} — gather marked
            citations into a table; run update_fields after to populate page numbers ·
        apply_theme {theme} — apply a whole document theme (colours+fonts+effects); theme is a
            built-in name (read 'themes') or a .thmx path ·
        set_theme_colors {[scheme, colors]} — scheme is a built-in colour-scheme name or .xml path;
            colors is {accent1|text1|background1|…: name/hex} to override individual brand colours ·
        set_theme_fonts {[scheme, major, minor]} — scheme is a built-in font-scheme name or .xml
            path; major/minor override the heading/body font names ·
        page_setup {[section=1],margins|top_margin|bottom_margin|left_margin|right_margin|gutter,
            orientation=portrait|landscape,paper_size=letter|legal|tabloid|a3|a4|a5,columns,column_spacing} —
            section page geometry; lengths accept unit strings ·
        header/footer {section,text,[which]} · track {on} ·
        watermark {text,[font,color,layout=diagonal|horizontal,semitransparent]} | {remove:true} —
            a text watermark behind every page (DRAFT/CONFIDENTIAL); remove:true clears it ·
        text_box {anchor_id,text,[width,height,wrap=square|tight|through|top-bottom|front|behind,
            before,font,size,bold,italic,alignment,fill,border]} — a floating text box / pull quote;
            border=false for no outline, a colour for a coloured one; returns its shape:N handle ·
        insert_image {anchor_id,wrap, image_base64|path,
            [before,block,width,height,alt_text,lock_aspect]} — block puts the image on its
            own new line instead of in the anchor's text run; a floating wrap returns the
            image's shape:N handle (inline stays image:N) ·
        insert_equation {anchor_id, unicodemath|latex|mathml, [display=true,before]} — a math
            equation on its own paragraph; unicodemath is native, latex needs the server's
            latex extra, mathml uses Office's transform; display centres it, else inline ·
        insert_chart {anchor_id, kind=bar|pie|line|scatter, data, [title,before]} — an Excel-backed
            chart; data is {label:value} for bar/pie/line or [[x,y],…] pairs for scatter (numeric
            axes); needs Excel installed (else error code excel_not_available); data is then static ·
        format_chart {anchor_id=chart:N, [title,legend,legend_position,chart_style,background,
            plot_background,font,font_size,font_color,data_labels,data_label_format,chart_type,
            gap_width,overlap,data_table]} — whole-chart/design formatting (no Excel needed;
            tri-state — only fields you pass apply) ·
        format_axis {anchor_id=chart:N, which=value|y|category|x,
            [title,minimum,maximum,scale=linear|log,number_format,gridlines]} — format one axis ·
        add_trendline {anchor_id=chart:N, [series=1,kind=linear|exponential|logarithmic|
            moving_average|polynomial|power,display_equation,display_r_squared,forward,backward,
            order,period]} — order=polynomial degree, period=moving-average window ·
        set_series_color {anchor_id=chart:N, color, [series=1,point]} — recolour a series or one
            1-based point/slice (color = name, hex, or [r,g,b]) ·
        format_series {anchor_id=chart:N, [series=1,point,marker=circle|square|diamond|triangle|x|
            star|dot|dash|plus|none|auto,marker_size,smooth,explosion,data_labels,data_label_size,
            data_label_color]} — series/point markers, line smoothing, pie explosion, label font ·
        add_error_bars {anchor_id=chart:N, [series=1,kind=fixed|percent|stdev|sterror,amount,
            include=both|plus|minus,axis=y|value|x|category]} — amount required unless kind=sterror ·
        set_shape_wrap {anchor_id=shape:N, [wrap=square|tight|through|top-bottom|front|behind,
            side=both|left|right|largest,distance_top,distance_bottom,distance_left,distance_right]}
            — wrap style / which sides text flows past (side honoured by square/tight/through) /
            standoff gaps; pass at least one ·
        set_shape_crop {anchor_id=shape:N, [crop_left,crop_top,crop_right,crop_bottom]} — trim a
            floating PICTURE shape in from its edges (lengths; shrinks displayed size) ·
        set_shape_position {anchor_id=shape:N, [left,top,relative_to=margin|page]} — left/top are
            lengths or "center" · set_shape_size {anchor_id=shape:N, [width,height,lock_aspect]} ·
        format_shape {anchor_id=shape:N, [fill,border,border_weight]} — fill/outline a text box or
            shape (border=false none / true default / a colour) ·
        set_shape_rotation {anchor_id=shape:N, degrees} — rotate clockwise ·
        set_shape_z_order {anchor_id=shape:N, order=front|back|forward|backward} — restack in the
            float layer · set_shape_text_frame {anchor_id=shape:N,
            [margin_left,margin_right,margin_top,margin_bottom,word_wrap]} — a text box's insets/wrap ·
        group_shapes {shapes=[shape:N,…]} — group two or more into one (returns the group's shape:N) ·
        ungroup_shape {anchor_id=shape:N} — dissolve a group into its members ·
        set_image_alt_text {anchor_id=image:N, text} · set_image_size {anchor_id=image:N,
            [width,height,lock_aspect]} — alt text / resize an INLINE picture (re-wrap via insert_image) ·
        set_image_crop {anchor_id=image:N, [crop_left,crop_top,crop_right,crop_bottom]} — trim an
            inline picture in from its edges ·
        set_shape_alt_text {anchor_id=shape:N, text} · set_shape_text {anchor_id=shape:N, text} —
            replace a text box's contents ·
        replace_shape_image {anchor_id=shape:N, image_base64|path} — swap a floating picture's
            image in place (preserves wrap/position/size) · delete_shape {anchor_id=shape:N} ·
        save {} — save to the document's existing file (must already be saved) ·
        save_as {path,[overwrite]} — save a .docx to path ·
        export_pdf {path,[from_page,to_page]} — export a PDF (the deliverable path).
            save/save_as/export_pdf are GATED: they only write inside the server's
            configured save directories (WORDLIVE_SAVE_DIRS); with none set, saving is off.

        For several edits in one undo step, use word_exec instead. Call
        word_read(command="guide") for the full anchor model and field reference.
        """
        params = {
            "doc": doc,
            "anchor_id": anchor_id,
            "text": text,
            "runs": runs,
            "items": items,
            "name": name,
            "bind": bind,
            "style": style,
            "before": before,
            "paragraph": paragraph,
            "find": find,
            "all": all,
            "occurrence": occurrence,
            "in_anchor": in_anchor,
            "action": action,
            "type": type,
            "author": author,
            "index": index,
            "table": table,
            "row": row,
            "col": col,
            "rows": rows,
            "cols": cols,
            "data": data,
            "header": header,
            "heading": heading,
            "body": body,
            "markdown": markdown,
            "level": level,
            "allow_break": allow_break,
            "first_row": first_row,
            "last_row": last_row,
            "first_column": first_column,
            "last_column": last_column,
            "banded_rows": banded_rows,
            "banded_columns": banded_columns,
            "values": values,
            "record": record,
            "key": key,
            "column": column,
            "value": value,
            "custom": custom,
            "mode": mode,
            "section": section,
            "which": which,
            "on": on,
            "alignment": alignment,
            "left_indent": left_indent,
            "right_indent": right_indent,
            "first_line_indent": first_line_indent,
            "space_before": space_before,
            "space_after": space_after,
            "line_spacing": line_spacing,
            "page_break_before": page_break_before,
            "keep_together": keep_together,
            "keep_with_next": keep_with_next,
            "widow_control": widow_control,
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "strikethrough": strikethrough,
            "font": font,
            "size": size,
            "color": color,
            "highlight": highlight,
            "subscript": subscript,
            "superscript": superscript,
            "small_caps": small_caps,
            "all_caps": all_caps,
            "spacing": spacing,
            "fill": fill,
            "pattern": pattern,
            "sides": sides,
            "line_style": line_style,
            "weight": weight,
            "lines": lines,
            "distance": distance,
            "position": position,
            "align": align,
            "leader": leader,
            "based_on": based_on,
            "next_style": next_style,
            "kind": kind,
            "wrap": wrap,
            "image_base64": image_base64,
            "path": path,
            "block": block,
            "width": width,
            "height": height,
            "alt_text": alt_text,
            "lock_aspect": lock_aspect,
            "unicodemath": unicodemath,
            "latex": latex,
            "mathml": mathml,
            "display": display,
            "margins": margins,
            "top_margin": top_margin,
            "bottom_margin": bottom_margin,
            "left_margin": left_margin,
            "right_margin": right_margin,
            "gutter": gutter,
            "orientation": orientation,
            "paper_size": paper_size,
            "columns": columns,
            "column_spacing": column_spacing,
            "levels": levels,
            "use_heading_styles": use_heading_styles,
            "hyperlinks": hyperlinks,
            "url": url,
            "bookmark": bookmark,
            "screen_tip": screen_tip,
            "target": target,
            "hyperlink": hyperlink,
            "label": label,
            "title": title,
            "tag": tag,
            "lock_contents": lock_contents,
            "lock_control": lock_control,
            "entry": entry,
            "cross_reference": cross_reference,
            "run_in": run_in,
            "right_align_page_numbers": right_align_page_numbers,
            "include_label": include_label,
            "where": where,
            "source_type": source_type,
            "year": year,
            "publisher": publisher,
            "city": city,
            "journal_name": journal_name,
            "volume": volume,
            "issue": issue,
            "pages": pages,
            "edition": edition,
            "doi": doi,
            "xml": xml,
            "prefix": prefix,
            "suffix": suffix,
            "suppress_author": suppress_author,
            "suppress_year": suppress_year,
            "suppress_title": suppress_title,
            "locale": locale,
            "long_citation": long_citation,
            "short_citation": short_citation,
            "category": category,
            "passim": passim,
            "keep_entry_formatting": keep_entry_formatting,
            "entry_separator": entry_separator,
            "page_range_separator": page_range_separator,
            "theme": theme,
            "scheme": scheme,
            "major": major,
            "minor": minor,
            "colors": colors,
            "overwrite": overwrite,
            "from_page": from_page,
            "to_page": to_page,
            "layout": layout,
            "semitransparent": semitransparent,
            "remove": remove,
            "border": border,
            "border_weight": border_weight,
            "left": left,
            "top": top,
            "relative_to": relative_to,
            "degrees": degrees,
            "order": order,
            "margin_left": margin_left,
            "margin_right": margin_right,
            "margin_top": margin_top,
            "margin_bottom": margin_bottom,
            "word_wrap": word_wrap,
            "side": side,
            "distance_top": distance_top,
            "distance_bottom": distance_bottom,
            "distance_left": distance_left,
            "distance_right": distance_right,
            "crop_left": crop_left,
            "crop_top": crop_top,
            "crop_right": crop_right,
            "crop_bottom": crop_bottom,
            "shapes": shapes,
            "rules": rules,
            "within": within,
            "profile": profile,
            "dry_run": dry_run,
            "allow_content": allow_content,
        }
        try:
            return _write_impl(w, command, params, policy=policy)
        except WordliveError as exc:
            raise _tool_error(exc) from exc

    @mcp.tool()
    def word_exec(
        # `| str` keeps the array in the JSON schema (as an anyOf) while letting a
        # JSON-encoded payload reach the body, where `coerce_ops` can answer with
        # an actionable message instead of pydantic's raw `list_type` error.
        ops: list[dict[str, Any]] | str,
        doc: str | None = None,
        label: str | None = None,
        tracked: bool = False,
    ) -> dict[str, Any]:
        """Apply a batch of ops to the open document as a SINGLE atomic undo.

        Each op is `{"op": "<kind>", ...}` — e.g.
        {"op":"write_bookmark","name":"Addr","text":"…"},
        {"op":"insert_paragraph","anchor_id":"heading:2","text":"…","style":"Body Text"},
        {"op":"find_replace","find":"Q3","text":"Q4","all":true}. Set `tracked`
        true to record the batch as tracked changes. Stops at the first failing
        op and reports `failure` (its `index`, `error`, `type`). Fields an op
        doesn't use are reported in `warnings`, not silently dropped.

        Anchor ids (the `anchor_id` of placement ops):
          heading:N · para:N (any paragraph) · bookmark:NAME · cc:NAME ·
          table:N:R:C (a cell) · range:START-END (what `find` emits — for
          replace/comments, NOT a placement target) · header:S:WHICH ·
          footer:S:WHICH · start · end. heading:N / para:N are positional and
          renumber on structural inserts — re-read outline/paragraphs after one.
          bookmark:/cc: are name-based and survive edits.

        Ops (required fields → behaviour):
          write_bookmark {name,text} · write_cc {name,text} ·
          insert_paragraph {anchor_id, text|runs, [style,before]} — new paragraph by an anchor;
            text is literal, runs is [{text,bold?,italic?,underline?,code?,style?}] for inline spans ·
          insert_block {anchor_id, items, [before]} — a contiguous run of styled paragraphs in one
            op; items are "plain text" or {text|runs, style?} (text takes **bold**/*italic*);
            returns the block's range:START-END in outputs ·
          insert_section {anchor_id, heading, body, [level=1,before]} — a Heading {level} + its body
            (body = insert_block items) in one op; returns the section's range:START-END ·
          insert_markdown {anchor_id, markdown, [before]} — a constrained-Markdown block as Word
            structure (#/##/### headings, -/* bullets, 1. numbers, paragraphs, inline **bold**/*italic*) ·
          replace_section {anchor_id=heading:N, body|markdown} — rewrite a heading's body, keep the heading ·
          delete_paragraph {anchor_id} — remove the paragraph(s) at an anchor, mark included ·
          append {text,[style]} / prepend {text,[style]} — new final/first paragraph ·
          append_inline {text} / prepend_inline {text} — continue the last/first paragraph (NO style) ·
          append_paragraph / prepend_paragraph — explicit synonyms of append/prepend ·
          replace {anchor_id,text} · find_replace {find,text,[all,occurrence,in]} ·
          apply_style {anchor_id,name} · format_paragraph {anchor_id,[alignment,*_indent,space_*,line_spacing,page_break_before,keep_together,keep_with_next,widow_control]} ·
          insert_image {anchor_id,wrap, path|base64, [before,block,width,height,alt_text,lock_aspect]} ·
          insert_equation {anchor_id, unicodemath|latex|mathml, [display=true,before]} — own-paragraph
            math; unicodemath native, latex needs the latex extra, mathml via Office; returns equation:N in outputs ·
          insert_chart {anchor_id, kind=bar|pie|line|scatter, data, [title,before]} — Excel-backed chart;
            data is {label:value} (bar/pie/line) or [[x,y],…] pairs (scatter); needs Excel; returns chart:N in outputs ·
          format_chart {anchor_id=chart:N,[title,legend,legend_position,chart_style,background,plot_background,font,font_size,font_color,data_labels,data_label_format,chart_type,gap_width,overlap,data_table]} — chart design/format (no Excel) ·
          format_axis {anchor_id=chart:N,which=value|y|category|x,[title,minimum,maximum,scale=linear|log,number_format,gridlines]} ·
          add_trendline {anchor_id=chart:N,[series=1,kind=linear|exponential|logarithmic|moving_average|polynomial|power,display_equation,display_r_squared,forward,backward,order,period]} ·
          set_series_color {anchor_id=chart:N,color,[series=1,point]} — recolour a series or one 1-based point/slice ·
          format_series {anchor_id=chart:N,[series=1,point,marker=circle|square|diamond|triangle|x|star|dot|dash|plus|none|auto,marker_size,smooth,explosion,data_labels,data_label_size,data_label_color]} — markers/smoothing/explosion/label font ·
          add_error_bars {anchor_id=chart:N,[series=1,kind=fixed|percent|stdev|sterror,amount,include=both|plus|minus,axis=y|value|x|category]} — amount required unless kind=sterror ·
          set_shape_wrap/set_shape_position/set_shape_size/format_shape/set_shape_alt_text/set_shape_text/replace_shape_image/delete_shape {anchor_id=shape:N,…} —
            restyle a floating shape (text box / image): wrap (style/side/distance_*), position (left/top/relative_to), size (width/height/lock_aspect), fill/border, alt text, text-box contents, picture swap (path|base64), or delete ·
          set_shape_crop {anchor_id=shape:N,[left,top,right,bottom]} — trim a floating PICTURE shape in from its edges (crop_* aliases also accepted) ·
          set_shape_rotation {anchor_id=shape:N,degrees} · set_shape_z_order {anchor_id=shape:N,order=front|back|forward|backward} ·
          set_shape_text_frame {anchor_id=shape:N,[margin_left,margin_right,margin_top,margin_bottom,word_wrap]} — a text box's insets / word-wrap ·
          group_shapes {shapes=[shape:N,…]} — group two or more floats into one (returns the group's shape:N) · ungroup_shape {anchor_id=shape:N} — dissolve a group ·
          set_image_alt_text {anchor_id=image:N,text} · set_image_size {anchor_id=image:N,[width,height,lock_aspect]} · set_image_crop {anchor_id=image:N,[left,top,right,bottom]} — alt text / resize / crop an INLINE picture (crop_* aliases also accepted; re-wrap floats it via insert_image) ·
          insert_break {anchor_id,[kind=page|column|section_next|section_continuous,before]} ·
          insert_field {anchor_id,kind,[text,before]} · update_fields {} · set_page_setup {section,[margins,*_margin,gutter,orientation,paper_size,columns,column_spacing]} ·
          regularize {[rules],[within],[profile],[dry_run],[allow_content]} — apply the fixable word_read lint findings in one atomic step (targeted, idempotent); profile enables policy-rule fixes (justify, line-spacing, numeric-column alignment); formatting fixes apply by default, content-changing fixes (insert caption/notice, delete stray para, strip watermark) are withheld into `deferred` unless allow_content=true; returns {applied,skipped,deferred,findings} ·
          insert_footnote/insert_endnote {anchor_id,text,[before]} — returns the new footnote:N/endnote:N in outputs ·
          insert_toc {anchor_id,[levels=[upper,lower],use_heading_styles,hyperlinks,before]} — update_fields after to fill page numbers ·
          add_bookmark {name,anchor_id} · pin {anchor_id,[name]} — durable pin:CODE handle that
            survives renumbering · pin_outline {[levels]} — pin every heading, returns {heading:N: pin:CODE} ·
          add_hyperlink {anchor_id, url|bookmark, [text,screen_tip]} ·
          set_hyperlink {index, [url,bookmark,text,screen_tip]} — retarget/relabel an existing link
            in place (index is 1-based, from word_read hyperlinks); url=external, bookmark=in-document;
            pass at least one; bookmark/screen_tip clear with "", url/text can't be emptied ·
          insert_cross_reference {anchor_id,target,[kind,hyperlink,before]} — target is a bookmark:/heading:/footnote:/endnote: id ·
          insert_caption {anchor_id,[label,text,position=above|below]} — own-paragraph caption ·
          create_content_control {anchor_id,[kind=rich_text|text|picture|combo_box|dropdown|date|checkbox|
            building_block|group|repeating_section,title,tag,items,where=wrap|before|after,lock_contents,lock_control]} —
            a form control; where=wrap surrounds the anchor's range; title makes it cc:TITLE; items fills a list; returns the cc: id in outputs ·
          set_cc_properties {anchor_id=cc:NAME,[title,tag,lock_contents,lock_control]} — re-set a control's
            metadata in place; pass at least one; "" clears title/tag; a rename changes its cc:NAME id ·
          set_cc_items {anchor_id=cc:NAME, items} — replace a combo_box/dropdown's choice list (items
            replaces, not appends; each is a string or {text,value}) ·
          mark_index_entry {anchor_id,entry,[cross_reference,bold,italic]} — mark a range as a back-of-book index entry ("main:sub" for a subentry) ·
          insert_index {[anchor_id=end],[columns=2,run_in,right_align_page_numbers,before]} — gather marked entries; update_fields after to fill page numbers ·
          insert_table_of_figures {[anchor_id=start],[label=Figure,include_label,hyperlinks,right_align_page_numbers,before]} — a table of captions of one label; update_fields after ·
          create_table {anchor_id, [rows,cols] (optional when data given — inferred),[style,data,header,before]} —
            data is a 2-D array OR records (objects whose keys become a header row); cells default to Normal; returns the new index in outputs ·
          set_cell {table,row,col,text} · add_row {table,[values]} · delete_row {table,row} ·
          add_column {table,[values]} — append a column (values fill top-to-bottom) ·
          delete_column {table,column} — fails on a merged/mixed-width table (delete its cells via table:N:R:C) ·
          merge_cells {table,from,to} — merge the rectangle between two cells (from/to are [row,col] or "R:C"); makes the table non-uniform ·
          split_cell {table,cell,[rows=1,cols=2]} — split one cell (cell is [row,col] or "R:C") into a grid ·
          append_record {table,record} — append a row from a {header: value} object ·
          update_row {table,key,values,[column]} — set cells (values={header: value}) on the first row whose key-column equals key ·
          set_heading_row {table,[row=1,heading,allow_break]} — repeating header row on a multi-page table ·
          autofit_table {table,[mode=content|window|fixed]} — resize a table's columns · delete_table {table} ·
          set_property {name,value,[custom]} / delete_property {name} — document metadata (built-in or custom) ·
          set_variable {name,value} / delete_variable {name} — invisible DOCVARIABLE storage ·
          add_comment {anchor_id,text,[author]} · resolve_comment {index} · delete_comment {index} ·
          accept_revision/reject_revision {index} — resolve one tracked change (renumbers the rest) ·
          accept_all_revisions/reject_all_revisions {[anchor_id]} — resolve every tracked change ([anchor_id] scopes to that range) ·
          set_watermark {text,[font,color,layout=diagonal|horizontal,semitransparent]} / remove_watermark {} — text watermark behind every page ·
          insert_text_box {anchor_id,text,[width,height,wrap,before,font,size,bold,italic,alignment,fill,border]} — a floating pull quote ·
          apply_list {anchor_id,[type=bulleted|numbered|outline,continue_previous]} · remove_list/restart_numbering/indent_list/outdent_list {anchor_id} ·
          apply_list_format {anchor_id,levels,[continue_previous]} — author + apply a custom list: levels is a 1-based array of per-level
            {[kind=number|bullet,format="%1.",style=arabic|upper-roman|lower-roman|upper-letter|lower-letter|…,bullet,font,start_at,
            number_position,text_position,trailing=tab|space|none,alignment,bold,italic,color]} (a bullet level needs a glyph; read back via word_read list_levels) ·
          write_header/write_footer {section,text,[which=primary|first|even]}.

        Durable handles: add bind:"slug" (or true) to insert/insert_block/insert_section/
        insert_markdown/create_table to mint a pin: on the new content (returned in that op's
        outputs). Any op field of the exact form $ops[N].field is replaced with an earlier op's
        output before it runs (e.g. create_table at op 0, then set_cell table:"$ops[0].table").

        Call word_read(command="guide") for the full field reference.
        """
        try:
            result, exc = _exec_impl(w, ops, doc=doc, label=label, tracked=tracked, policy=policy)
        except WordliveError as setup_exc:
            raise _tool_error(setup_exc) from setup_exc
        if exc is not None:
            raise _tool_error(exc, result=result)
        return result

    # structured_output=False: this tool returns MCP content blocks (image + text
    # labels), not structured data. Without it FastMCP infers a wrapped output
    # schema from the `-> list[Any]` annotation and *re-serialises every block —
    # including the base64 PNG bytes — into structuredContent*, sending each image
    # on the wire twice (a large, silent token cost on hosts that forward
    # structuredContent). Suppressing the schema sends the image exactly once.
    @mcp.tool(structured_output=False)
    def word_snapshot(
        doc: str | None = None,
        pages: str | None = None,
        anchor: str | None = None,
        dpi: int = 150,
        max_dim: int | None = None,
        markup: str = "none",
    ) -> list[Any]:
        """Render page(s) of the open document to PNG so you can SEE the layout.

        Pick at most one target: `anchor` (the page(s) an anchor occupies — a
        heading expands to its whole section), or `pages` ("4" or "2-5"). With
        neither, the whole document renders. `max_dim` caps each page's long edge
        to that many pixels (only ever lowering resolution) — pair it with no page
        target to check the WHOLE document's layout cheaply: a vision model is
        billed on pixel area, so the cap is a predictable per-page token budget
        regardless of paper size (~1000 stays legible for "did my styling land").
        `markup` is "none" (the final document) or "all" (show tracked changes and
        comments as visible revision marks / balloons — pair with
        word_read(command="revisions") for the structured list). Returns image
        content (and a "page N" label per page) inline, so a vision model sees the
        render directly — no filesystem path that a remote/sandboxed host couldn't
        open. Needs the snapshot extra (PyMuPDF).
        """
        try:
            rendered = _snapshot_impl(
                w, doc=doc, pages=pages, anchor=anchor, dpi=dpi, max_dim=max_dim, markup=markup
            )
        except WordliveError as exc:
            raise _tool_error(exc) from exc
        content: list[Any] = []
        for page, png in rendered:
            content.append(TextContent(type="text", text=f"page {page}"))
            # Convert to an ImageContent block explicitly: FastMCP won't serialise
            # a bare Image when it's one element of a mixed content list.
            content.append(Image(data=png, format="png").to_image_content())
        return content

    @mcp.resource("wordlive://guide", mime_type="text/markdown")
    def guide() -> str:
        """The full wordlive agent guide: anchor model, every verb, the op vocabulary."""
        return skill_body()

    return mcp


def main() -> None:
    """Launch the server over stdio (the transport Claude Desktop spawns)."""
    build_server().run()
