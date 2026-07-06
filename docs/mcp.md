# MCP server

`wordlive` ships an [MCP](https://modelcontextprotocol.io) server so MCP clients —
**Claude Desktop**, Cursor, and other agent hosts — can read and edit the Word
document you have open right now, including *seeing* rendered pages.

It talks to the same running Word instance the CLI and Python API do, over COM, on
Windows. Edits stay polite (your cursor, selection, and scroll are preserved; each
write is a single Ctrl-Z), and failures come back with a stable error `code` and a
`retryable` hint.

## Install

Three ways, easiest first. All three end with the same `word_*` tools in your
client — pick whichever fits.

### 1. One-click bundle (`.mcpb`)

The repo ships an **MCP bundle** in
[`mcpb/`](https://github.com/thomas-villani/wordlive/tree/main/mcpb) — a single
`wordlive.mcpb` file you drag onto Claude Desktop → **Settings → Extensions** to
install in one click, no JSON editing. The bundle carries no code: it declares a
dependency on the published `wordlive[mcp,snapshot]` package and Claude Desktop's
`uv` runtime resolves it on first run. Needs [`uv`](https://docs.astral.sh/uv/)
on `PATH`. See [`mcpb/README.md`](https://github.com/thomas-villani/wordlive/blob/main/mcpb/README.md)
to download or rebuild it.

### 2. `wordlive install-mcp`

Let the CLI write the config entry for you (it launches the server with `uvx`,
so there's no separate install step):

```
wordlive install-mcp                       # merge into Claude Desktop's config
wordlive install-mcp --client claude-code  # write a project-local ./.mcp.json
wordlive install-mcp --print               # just print the JSON snippet
wordlive install-mcp --directory .         # dev: run a local checkout via uv run
```

It merges an `mcpServers.wordlive` entry (using `uvx --from "wordlive[mcp,snapshot]"
wordlive-mcp`) into the target config, refusing to clobber an existing entry
without `--force`. Offline — it never touches Word. Restart the client afterward.

### 3. By hand

Install the optional extra, then add the entry yourself:

```
pip install "wordlive[mcp,snapshot]"   # snapshot extra adds the vision tool
uv tool install "wordlive[mcp,snapshot]"
```

```json
{
  "mcpServers": {
    "wordlive": {
      "command": "wordlive-mcp"
    }
  }
}
```

If `wordlive-mcp` isn't on Claude Desktop's `PATH`, point at the interpreter in
your environment explicitly (note the doubled backslashes in JSON):

```json
{
  "mcpServers": {
    "wordlive": {
      "command": "C:\\Users\\you\\project\\.venv\\Scripts\\python.exe",
      "args": ["-m", "wordlive.mcp"]
    }
  }
}
```

Claude Desktop's config lives at *Settings → Developer → Edit Config*. Restart
Claude Desktop, open a `.docx` in Word, and the `word_*` tools appear.

## Run

```
wordlive-mcp            # console script (stdio transport)
python -m wordlive.mcp  # equivalent
```

The server speaks MCP over **stdio** — the transport Claude Desktop spawns. Word
must already be running on the same machine (wordlive *attaches*; it never launches
or closes Word).

## Tools

A small set of **dispatch tools** keeps the client's tool list (and its context
cost) lean. Every tool takes an optional `doc` (target a document by name; default
is the active one).

| Tool | What it does |
| --- | --- |
| `word_read` | Every read, dispatched on `command`: `guide`, `status`, `outline`, `paragraphs`, `find` `{text,[in_anchor],[mode=fuzzy\|literal\|regex]}`, `read_bookmark`, `read_cc`, `read_section`, `to_markdown` `{[within]}` (serialise the document — or one anchor's range via `within` — to clean Markdown: headings, lists, `**bold**`/`*italic*`, GFM tables, `![alt](image:N)`, `[text](url)`; the read mirror of `insert_markdown`, lossy by design; returns `{markdown}`), `to_html` `{[within]}` (the same but an HTML fragment; returns `{html}`), `digest` `{[budget=6000],[depth]}` (a token-budgeted, structure-aware read of the **whole** document — headings verbatim (each tagged with its `heading:N` anchor), tables as one-line shape stubs, body sampled to fit `budget`; every anchor stays addressable, so drill into elided regions with `to_markdown` `{within}`; loads a large doc into context cheaply; returns `{digest}`), `between` `{start_anchor,end_anchor,[inclusive]}` (content spanning two anchors — e.g. the block between two headings; default excludes both heading lines, `inclusive` covers them; returns a `range:START-END` id + text), `nearest_heading` `{anchor_id,[direction=before\|after]}` (the heading nearest a position — `before`=enclosing/preceding, `after`=next; an outline row or `null`), `find_paragraphs` `{text,[limit=5],[min_score=0.6]}` (fuzzy-rank paragraphs by similarity to `text` — typo/paraphrase tolerant, unlike the exact-substring `find`; returns `para:N` candidates with scores), `table_list`, `table_read`, `table_records` `{table}` (body rows as dicts keyed by the header row — the read mirror of building a table from records), `styles`, `comments`, `revisions` (tracked changes), `read_text` `{anchor_id,[view]}` (an anchor's text; `view`=`raw`/`final`/`original`/`segments` resolves tracked changes — `final` as if accepted, `original` as if rejected, `segments` the per-run insert/delete breakdown), `track` (is Track Changes on?), `sections`, `footnotes`, `endnotes`, `images` (embedded pictures: `image:N` id, mime, size, crop, alt, para), `read_image` `{anchor_id}` (**SEE** an embedded picture — returns it as an inline image content block, like `word_snapshot`, plus a compact `{anchor_id,mime,bytes}` label; pass an `image:N` id or any single-image anchor), `equations` (math zones: `equation:N` id, type, linear preview, para), `charts` (Excel-backed charts: `chart:N` id, kind, title, chart_style, has_legend, para), `shapes` (floating shapes: `shape:N` id, `shape_type`=text_box/picture/wordart/group, size, rotation, z_order, wrap, wrap_side, crop, alt_text, para — text boxes / floating images / WordArt, the restyle handles; header-story watermarks excluded), `hyperlinks` (links: text, address/sub_address, `range:START-END` — the read mirror of `add_hyperlink`), `fields` (PAGE/REF/TOC/…: kind, code, rendered result, `range` — the read mirror of `insert_field`), `properties` (document metadata: `{builtin, custom}` name→value bags), `watermark` (the text watermark behind the pages: `{text, sections}` or `null` — the read mirror of the `watermark` write), `variables` (invisible named storage: `{name: value}`), `proofing` (spelling/grammar errors with counts + flagged runs, and readability statistics — heavier than `stats`; it (re)checks the document), `format_info` `{anchor_id}` (the effective formatting at an anchor: `style` name plus paragraph/font fields as `{value, style, override}` — `override` flags a direct override of the style's baseline, `mixed` lists font fields that vary across runs; font also carries `hidden` and `highlight` (a keyword or `"none"`, effective-only); the read mirror of `format_paragraph`/`format_run`), `list_levels` `{anchor_id}` (the per-level format of the list at an anchor — one `{level, kind, format, style, trailing, number_position, text_position, font}` per template level; the read mirror of the list `format` action), `lint` `{[rules],[within],[profile]}` (audit the document for formatting/structural/policy issues — a severity-ranked list of `{rule, kind, severity, anchor_id, message, fixable, fix, …}`; `fix` is an op-shaped dict you can hand to `regularize`; `rules` selects/excludes by id or tag — e.g. `["typography"]` for the whitespace/punctuation/heading-hygiene cluster or `["finalization"]` for the off-by-default "is-this-ready-to-send?" cluster (leftover comments/revisions, track-changes-on, hidden text, highlight, updatable fields), or `["academia"]` for the field-code cluster (`broken-cross-reference` + `caption-manual-numbering`, both on by default, plus `xref-as-literal-text` — a figure/table mentioned by literal number with no `REF` field, heuristic/off; `page-numbers-present` is off, tag `layout`), or `["hyperlinks"]` for the link cluster (`hyperlink-broken-internal` — a dead internal `\l` jump to a missing bookmark, on by default; plus the off-by-default `hyperlink-bare-for-print` and `hyperlink-display-is-raw-url`, which `["print"]` selects on their own), or `["layout"]` for the §H page-layout/document-level cluster (all off by default, report-only: `header-footer-consistent`, `draft-watermark-present`, `document-properties-filled`, and the profile-driven `confidentiality-notice`/`copyright-notice`, which `["notices"]` selects on their own), or `["structure"]` for the §B heading & document-structure cluster (`heading-level-skip` and `empty-heading` on by default, plus the off-by-default `adjacent-headings`, `heading-numbering-manual`, `heading-trailing-period` — the one fixable rule, stripping the period in place — and `toc-present-and-current`; `["headings"]` also selects it), the first two clusters of which also pull in their off-by-default members; `within` scopes to an anchor; `profile` is a house-style config — a path or inline object — that enables **policy** rules (`body-justified`, `body-line-spacing`, `table-numeric-right-align`) and supplies their targets/thresholds, or overrides a rule's severity / disables a default), `checkpoint` `{[include=text\|text+style\|text+format],[within]}` (an opaque, serialisable structural fingerprint of the document now — store it, edit, then `diff`; the only way to answer "what changed" since Word has no content-change event), `diff` `{checkpoint \| cp_a,cp_b}` (a content-aligned change list — pass a stored `checkpoint` to diff it against the document **now**, or `cp_a`+`cp_b` to diff two stored checkpoints; each change is `replace`/`insert`/`delete`/`restyle`/`reformat` carrying the current `para:N`), `location` `{anchor_id}` (where an anchor sits in the laid-out document: `page`/`end_page` span, `line`, `column`, `in_table` — "what page is this on" without a snapshot), `stats` (one-shot document summary: page/word/character/paragraph/line counts plus section/heading/table/image/equation/comment/revision counts and `saved`), `theme` (the document theme: 12 brand colours as `#RRGGBB` + major/minor fonts — the read mirror of `apply_theme`/`set_theme_*`), `themes` (the built-in themes, colour schemes, and font schemes Office ships — the names the theme write commands accept). |
| `word_write` | Every single atomic-undo edit, dispatched on `command`: `insert`, `insert_block`, `insert_section`, `insert_markdown`, `replace_section`, `delete_paragraph`, `insert_break`, `append`, `prepend`, `replace`, `write_bookmark`, `write_cc`, `apply_style`, `format_paragraph`, `format_run`, `regularize`, `set_shading`, `set_borders`, `drop_cap`, `add_tab_stop`, `add_style`, `set_style`, `list`, `comment`, `revision` (`action` = `accept`/`reject`/`accept_all`/`reject_all`), `table` (`action` = `set_cell`/`add_row`/`delete_row`/`add_column`/`delete_column`/`merge_cells`/`split_cell`/`append_record`/`update_row`/`set_heading_row`/`autofit`/`set_style`/`set_alignment`/`set_borders`/`set_banding`/`create`/`delete`), `cell_valign`, `header`, `footer`, `track`, `watermark`, `text_box`, `insert_image`, `insert_equation`, `insert_chart`, `format_chart`, `format_axis`, `add_trendline`, `set_series_color`, `format_series`, `add_error_bars`, `set_shape_wrap`, `set_shape_crop`, `set_shape_position`, `set_shape_size`, `format_shape`, `set_shape_alt_text`, `set_shape_text`, `set_shape_rotation`, `set_shape_z_order`, `set_shape_text_frame`, `replace_shape_image`, `delete_shape`, `group_shapes`, `ungroup_shape`, `set_image_alt_text`, `set_image_size`, `set_image_crop`, `insert_field`, `update_fields`, `set_property`, `delete_property`, `set_variable`, `delete_variable`, `insert_footnote`, `insert_endnote`, `insert_toc`, `insert_table_of_figures`, `mark_index_entry`, `insert_index`, `set_bibliography_style`, `add_source`, `insert_citation`, `insert_bibliography`, `mark_citation`, `insert_table_of_authorities`, `apply_theme`, `set_theme_colors`, `set_theme_fonts`, `create_content_control`, `set_cc_properties`, `set_cc_items`, `add_bookmark`, `pin`, `pin_outline`, `add_hyperlink`, `set_hyperlink`, `insert_cross_reference`, `insert_caption`, `page_setup`. `insert {anchor_id, text \| runs, [style,before]}` adds one paragraph — `text` is literal, `runs` is `[{text,bold?,italic?,underline?,style?}]` for inline-formatted spans. `insert_block {anchor_id, items, [before]}` inserts a contiguous run of styled paragraphs in one op (each item a string or `{text \| runs, style?}`, where `text` takes `**bold**`/`*italic*` markdown) and returns the block's `range:START-END`. `insert_section {anchor_id, heading, body, [level,before]}` places a `Heading {level}` paragraph plus its body (the `insert_block` items shape) in one op; `insert_markdown {anchor_id, markdown, [before]}` maps a constrained-Markdown subset (`#`/`##`/`###` headings, `-`/`*` bullets, `1.` numbers, blank-line paragraphs, inline `**bold**`/`*italic*` — not CommonMark) to real Word structure; `replace_section {anchor_id, body \| markdown}` rewrites a `heading:N`'s body up to the next same-or-higher heading while keeping the heading. `delete_paragraph {anchor_id}` removes the paragraph(s) at an anchor, mark included. `append`/`prepend` add a new paragraph (optional `style`); pass `paragraph: false` to continue the adjacent paragraph inline (no `style`). `table action="create"` needs `anchor_id` and `rows`/`cols` — both optional when `data` is given (inferred from it); `data` is a 2-D array **or** records (objects whose keys become a header row). `format_paragraph` also carries `line_spacing` (a multiple like `1.5`, `single`/`1.5`/`double`, or an exact length like `14pt`) and the pagination controls `keep_together`/`keep_with_next`/`widow_control` (multi-page layout); `table action="set_heading_row"` `{table,[row=1,heading,allow_break]}` marks a repeating header row; `table action="append_record"` `{table,record}` appends a row from a `{header: value}` object, and `table action="update_row"` `{table,key,values,[column]}` sets cells (`values` = `{header: value}`) on the first row whose key-column (`column`, default first) equals `key`. `table action="add_column"` `{table,[values]}` appends a column (the mirror of `add_row`, `values` filling top-to-bottom) and `table action="delete_column"` `{table,column}` removes one — `delete_column` fails on a merged / mixed-width table (delete those cells via `table:N:R:C`). `table action="merge_cells"` `{table,from,to}` merges the rectangle between two cells and `table action="split_cell"` `{table,cell,[rows=1,cols=2]}` divides one (`from`/`to`/`cell` are `[row,col]` or `"R:C"`); both leave the table **non-uniform** — `table:N:R:C` then indexes physical cells (the `table` read reports `"uniform": false`). `format_run` styles characters (bold/italic/colour/highlight/…); `replace {text, anchor_id \| find, [all, occurrence, in_anchor, mode=fuzzy\|literal\|regex]}` overwrites an anchor's whole range or find-replaces text (`mode=regex` makes `find` a Python regex and lets `text` use `\1` backreferences, expanded per match); `regularize {[rules],[within],[profile],[dry_run],[allow_content]}` applies the fixable `word_read(command="lint")` findings in one atomic-undo edit (each writes the style's own value back, so it's idempotent — a second run is a no-op; the typography rules in `rules=["typography"]` fix via regex `replace`; a `profile` — a path or inline house-style object — enables the policy-rule fixes: justify, line-spacing, numeric-column alignment; `dry_run=true` plans without writing; content-changing fixes — insert a caption/notice, delete a stray paragraph, strip a watermark — are flagged `adds_content` and withheld into `deferred` unless `allow_content=true`) and returns its `{applied, skipped, deferred, findings}` report; `set_shading`/`set_borders`/`add_tab_stop` add cell/range shading, borders, and tab stops; `drop_cap {anchor_id, [position=dropped|margin|none, lines=3, distance, font]}` turns the first letter of the anchor's paragraph into an editorial drop cap (`position=none` removes one); `add_style`/`set_style` create and configure styles (the border line style is the `line_style` param, to avoid colliding with `style`). `insert_equation {anchor_id, unicodemath \| latex \| mathml, [display,before]}` inserts a math equation on its own paragraph (UnicodeMath is native, LaTeX needs the server's `latex` extra, MathML uses Office's transform; `display=true` gives it the centred `Equation` style, `display=false` is Normal + left — the paragraph style is pinned either way so it never inherits a neighbouring heading) and returns the new `equation:N` in `result` — a positional id (OMaths order) that renumbers when an earlier equation is inserted, so re-list rather than caching it. `insert_chart {anchor_id, kind, data, [title,before]}` embeds an Excel-backed chart (`kind` = `bar`/`pie`/`line`/`scatter`; `data` is a `{label: value}` object for bar/pie/line or an array of `[x, y]` pairs for scatter — numeric axes, duplicate x kept) and returns the new `chart:N`; it needs Excel installed (else the `excel_not_available` error, document untouched), and the chart's data is made **static** (no embedded workbook ships, and the series data isn't read back — `command="charts"` is metadata only). The chart **formatting & design** commands all take an `anchor_id=chart:N` and operate on the static chart (**no Excel needed**, tri-state — only fields you pass apply): `format_chart {[title,legend,legend_position,chart_style,background,plot_background,font,font_size,font_color,data_labels,data_label_format,chart_type,gap_width,overlap,data_table]}` is the whole-chart/design surface (`chart_type` re-types in place, `chart_style` is the design-gallery id, `gap_width`/`overlap` tune bar spacing, `data_table` toggles the data-table grid), `format_axis {which=value|y|category|x,[title,minimum,maximum,scale=linear|log,number_format,gridlines]}` formats one axis (`scale=log` for order-of-magnitude data), `add_trendline {[series=1,kind=linear|exponential|logarithmic|moving_average|polynomial|power,display_equation,display_r_squared,forward,backward,order,period]}` fits a trendline (a `power` fit with `display_equation` draws the law of best fit; `order` is the polynomial degree, `period` the moving-average window), `set_series_color {color,[series=1,point]}` recolours a series or one 1-based point/slice (colour = name/hex/`[r,g,b]`), `format_series {[series=1,point,marker=circle|square|diamond|triangle|x|star|dot|dash|plus|none|auto,marker_size,smooth,explosion,data_labels,data_label_size,data_label_color]}` styles a series' markers / line smoothing / pie explosion / data-label font (`point` narrows marker / explosion / label to one point), and `add_error_bars {[series=1,kind=fixed|percent|stdev|sterror,amount,include=both|plus|minus,axis=y|value|x|category]}` draws error bars (`amount` required unless `kind=sterror`). `insert_field` drops a self-updating field (page numbers, dates, refs) — put `kind: page` in a footer; `update_fields` refreshes them; `page_setup` sets a section's margins/orientation/paper size/`columns`. `set_property {name,value,[custom]}` / `delete_property {name}` write the document's metadata (a built-in like Title/Author, or a custom property with `custom=true`); `set_variable {name,value}` / `delete_variable {name}` manage the invisible `DOCVARIABLE` storage; `table action="autofit" {table,[mode=content|window|fixed]}` resizes a table's columns (fit to content/window or pin them). `table action="set_style" {table,style}` restyles an existing table (applying a style overwrites direct cell shading — restyle first, then layer cell overrides); `table action="set_alignment" {table,alignment=left|center|right}` positions the whole table across the page; `table action="set_borders" {table,[sides,style|line_style,weight,color]}` rules the whole grid in one call (interior gridlines via `horizontal`/`vertical`); `table action="set_banding" {table,[first_row,last_row,first_column,last_column,banded_rows,banded_columns]}` toggles the table-style options (needs a real table style applied to show); `cell_valign {anchor_id=table:N:R:C, align=top|center|bottom}` sets a cell's vertical alignment. A whole **row** (`table:N:row:R`) or **column** (`table:N:col:C`) is itself an anchor, so `set_shading`/`set_borders`/`apply_style`/`format_run` style the strip in one call — a column op on a merged / mixed-width table raises an error (style those cells via `table:N:R:C`). `insert_footnote`/`insert_endnote` attach a note (the new `footnote:N`/`endnote:N` comes back in `result`); `insert_toc` inserts a table of contents (run `update_fields` after to fill its page numbers). `insert_table_of_figures {anchor_id (default start), [label=Figure, include_label, hyperlinks, right_align_page_numbers, before]}` lists the document's captions of one label as a table of figures. `mark_index_entry {anchor_id, entry, [cross_reference, bold, italic]}` marks a range as a back-of-book index entry (`entry` uses `main:sub` to nest), then `insert_index {anchor_id (default end), [columns, run_in, right_align_page_numbers, before]}` builds the index from those marks — both field blocks, like the TOC, so `update_fields` after to fill page numbers. Citations are a three-step flow: `set_bibliography_style {style}` picks the scheme (`APA`/`MLA`/`Chicago`/`IEEE`/… — build-dependent), `add_source {source_type, [tag, author, title, year, publisher, city, journal_name, volume, issue, pages, url, edition, doi] | xml}` registers a source (`source_type` is `book`/`journal_article`/`web_site`/`case`/…; `author` is `Last, First`; `tag` auto-derives from surname+year if omitted; `xml` is the raw `<b:Source>` escape hatch) and returns its tag in `result`, then `insert_citation {anchor_id, tag, [pages, prefix, suffix, volume, suppress_author, suppress_year, suppress_title, locale, before]}` cites it (an unknown tag renders "Invalid source specified.") and `insert_bibliography {anchor_id (default end), [before]}` inserts the works-cited block — a field block, so `update_fields` after. A table of authorities is the mark-then-build pattern: `mark_citation {anchor_id, long_citation, [short_citation, category, before]}` marks a `TA` entry (`category` is `cases`/`statutes`/`other`/… or an int 1–16; `short_citation` defaults to `long_citation`), then `insert_table_of_authorities {anchor_id (default end), [category (default all), passim, keep_entry_formatting, entry_separator, page_range_separator, before]}` builds the table from those marks (a field block, so `update_fields` after to fill page numbers). The document **theme** is the brand primitive: `apply_theme {theme}` applies a whole theme (colours+fonts+effects) by built-in name (read `command="themes"`) or `.thmx` path, `set_theme_colors {[scheme, colors]}` loads a named colour scheme and/or overrides individual brand colours (`colors` is `{accent1|text1|background1|…: name/hex}`), and `set_theme_fonts {[scheme, major, minor]}` sets the heading/body fonts. `create_content_control {anchor_id, [kind=rich_text, title, tag, items, where=wrap, lock_contents, lock_control]}` creates a content control — `where=wrap` surrounds the anchor's range (`before`/`after` insert a fresh empty one), `items` populates a `combo_box`/`dropdown`, and a `title` (or `tag`) makes it addressable as `cc:TITLE`. `set_cc_properties {anchor_id=cc:NAME, [title, tag, lock_contents, lock_control]}` re-sets a control's metadata in place (pass ≥1; `""` clears `title`/`tag`; a rename changes its `cc:NAME` id), and `set_cc_items {anchor_id=cc:NAME, items}` replaces a combo_box/dropdown's choice list — both edit-in-place instead of delete-and-reinsert. `add_bookmark` names a range; `add_hyperlink` links it (external `url` or internal `bookmark`); `set_hyperlink {index, [url, bookmark, text, screen_tip]}` retargets/relabels an existing link in place (index is 1-based, from `word_read` `hyperlinks`; `url`→external, `bookmark`→in-document; `bookmark`/`screen_tip` clear with `""`, but `url`/`text` can't be emptied — delete the link to unlink); `insert_cross_reference` references a `bookmark:`/`heading:`/`footnote:`/`endnote:` target; `insert_caption` adds a numbered caption in its own paragraph (`position` = `above`/`below`, default above for a `Table` and below otherwise; on a `table:N:R:C` anchor it captions the whole table). `revision {action=accept|reject (index) | accept_all|reject_all ([anchor_id] scopes to that range, else the whole document)}` resolves tracked changes — accept/reject renumber the rest, so the bulk forms are safer for several. `watermark {text,[font,color,layout=diagonal|horizontal,semitransparent]}` (or `{remove:true}`) stamps / clears a text watermark behind every page; `text_box {anchor_id,text,[width,height,wrap,before,font,size,bold,italic,alignment,fill,border]}` drops a floating pull quote (`border=false` for no outline) and returns its `shape:N` handle. **Floating shapes** (text boxes, floating images from a non-`inline` `insert_image`, WordArt) are addressed `shape:N` (list them with `word_read command="shapes"`; header-story watermarks excluded) and restyled in place: `set_shape_wrap {anchor_id=shape:N, [wrap, side=both|left|right|largest, distance_top, distance_bottom, distance_left, distance_right]}` (the wrap style, which sides text flows past — `square`/`tight`/`through` honour `side` — and the standoff gaps; pass any one), `set_shape_crop {anchor_id=shape:N, [crop_left, crop_top, crop_right, crop_bottom]}` (trim a floating **picture** shape in from its edges), `set_shape_position {anchor_id=shape:N, [left, top, relative_to=margin|page]}` (`left`/`top` are lengths or `"center"`), `set_shape_size {anchor_id=shape:N, [width, height, lock_aspect]}`, `format_shape {anchor_id=shape:N, [fill, border, border_weight]}` (`border` = `false`/`true`/a colour), `set_shape_alt_text {anchor_id=shape:N, text}`, `set_shape_text {anchor_id=shape:N, text}` (text boxes), `replace_shape_image {anchor_id=shape:N, image_base64|path}` (swap a floating picture's image in place, preserving wrap/position/size), and `delete_shape {anchor_id=shape:N}`. Deeper layout: `set_shape_rotation {anchor_id=shape:N, degrees}` (absolute angle), `set_shape_z_order {anchor_id=shape:N, order=front|back|forward|backward}` (restack within the float layer — distinct from wrap's front/behind-text), and `set_shape_text_frame {anchor_id=shape:N, [margin_left, margin_right, margin_top, margin_bottom, word_wrap]}` (a text box's insets). `group_shapes {shapes=[shape:N, …]}` collapses two or more floats into one group `shape:N` (returns it) and `ungroup_shape {anchor_id=shape:N}` dissolves it back into its members. `set_image_alt_text {anchor_id=image:N, text}` / `set_image_size {anchor_id=image:N, [width, height, lock_aspect]}` / `set_image_crop {anchor_id=image:N, [crop_left, crop_top, crop_right, crop_bottom]}` restyle an **inline** picture without floating it (re-wrapping floats it via `insert_image`). `shape:N` is positional — it renumbers as shapes are added/removed, so re-list rather than cache it (`textbox:N` is an alias onto a text box's canonical `shape:N`). There is no autosize — Word's resize-to-fit-text doesn't expose cleanly over COM. Plus the **gated** persistence commands `save` (to the existing file), `save_as {path,[overwrite]}` (a `.docx`), and `export_pdf {path,[from_page,to_page]}` (a PDF deliverable) — terminal side-effects, not undoable; see [Saving](#saving). |
| `word_exec` | Apply a batch of `ops` as a **single** atomic undo — the power tool for multi-step intents. |
| `word_snapshot` | Render page(s) to PNG so the model can *see* the layout (`markup: "all"` shows tracked changes / comments as revision marks). `max_dim` caps each page's long edge in pixels — pair it with no page target to check the whole document's layout cheaply (a predictable per-page token budget; ~1000 stays legible). Returns image content. Needs the `snapshot` extra. |

The anchor model (`heading:N`, `para:N`, `bookmark:NAME`, `pin:CODE`, `cc:NAME`,
`footnote:N`, `endnote:N`, `image:N`, `table:N:R:C`, `table:N:row:R`,
`table:N:col:C`, `range:START-END`,
`header:S:WHICH` / `footer:S:WHICH`, `start` /
`end`) and the full `word_exec` op vocabulary are documented in the one-page
guide. Fetch it as a tool call with **`word_read(command="guide")`** (it needs
neither Word nor a document) — the most reliable path, since not every MCP
client surfaces resources. The same text is also the **`wordlive://guide`**
resource and `wordlive llm-help`. See also [CLI](cli.md) for each op's fields.

`heading:N` / `para:N` are **positional** and renumber when a structural edit
shifts the document, so re-read `outline` / `paragraphs` after an insert before
reusing ids (a missed positional id raises with a recovery hint); `bookmark:NAME`
/ `cc:NAME` are name-based and survive edits. For ad-hoc content with no name,
`word_write(command="pin", anchor_id=…)` mints a durable `pin:CODE` handle (and
`pin_outline` pins every heading at once) that Word keeps attached to the same
content across edits. Bulk-pin is write-side only — `word_read(command="outline")`
stays a pure read.

### Batches

`word_exec` mirrors the CLI's [`exec`](cli.md):

```json
{
  "ops": [
    {"op": "write_bookmark",   "name": "Address",      "text": "123 Main St"},
    {"op": "insert_paragraph", "anchor_id": "heading:2", "text": "New section.", "style": "Body Text"},
    {"op": "find_replace",     "find": "Q3", "text": "Q4", "all": true}
  ],
  "tracked": false
}
```

It stops at the first failing op and reports its `index`, `op`, and error `type`;
the successful prefix still rolls back as one undo step. A field an op doesn't
use (a typo, or `style` on an inline append) is reported in a `warnings` array on
the result rather than silently dropped.

**Durable handles in a batch.** Add `bind: "slug"` (or `bind: true`) to an
`insert` / `insert_block` / `insert_section` / `insert_markdown` / `create_table`
op to mint a `pin:` on the new content — returned as `pin` in that op's `outputs`
entry. And any op field of the exact form `$ops[N].field` is replaced with an
earlier op's recorded output before the op runs, so a batch can create then target
without a round-trip (e.g. `create_table` at op 0, then `set_cell` with
`"table": "$ops[0].table"`).

## Saving

The `word_write` commands `save` / `save_as` / `export_pdf` are **gated and
default-deny**: they only write inside directories the operator whitelists when
launching the server, via the **`WORDLIVE_SAVE_DIRS`** environment variable (an
`os.pathsep`-separated list). With none configured, saving is off and every
attempt returns a `path_not_allowed` error. The target is resolved first (so
`..`/symlinks can't escape) and must then sit inside an allowed directory.
`export_pdf` is the recommended way to hand back a deliverable — a pixel-faithful
PDF via the same engine as `word_snapshot`.

The same launch config optionally restricts image-source paths with
**`WORDLIVE_IMAGE_DIRS`**; regardless of that allowlist, `insert_image` with a
non-local `path` (a UNC `\\host\share\…`, a `file://`, or any URL) is always
rejected before any filesystem access — prefer `image_base64` for
LLM-supplied images. Configure both in the client's server entry:

```json
{
  "mcpServers": {
    "wordlive": {
      "command": "wordlive-mcp",
      "env": {
        "WORDLIVE_SAVE_DIRS": "C:\\Users\\you\\Documents\\wordlive-out",
        "WORDLIVE_IMAGE_DIRS": "C:\\Users\\you\\Pictures"
      }
    }
  }
}
```

## Errors

A failed tool call comes back flagged as an error whose message is a JSON object:

```json
{"error": "bookmark not found: 'Addr'", "code": "anchor_not_found", "retryable": false, "type": "AnchorNotFoundError"}
```

`code` / `retryable` mirror the CLI's [exit-code taxonomy](errors.md):

| `code` | Meaning | `retryable` |
| --- | --- | --- |
| `anchor_not_found` | anchor missing, or `find` matched zero | no — re-read first |
| `style_not_found` | a named style isn't defined in the document | no — read `styles` first |
| `ambiguous_match` | `find` matched several | yes — pass `occurrence`/`all` |
| `replace_verification` | a `find_replace` target couldn't be verified (e.g. a whole-doc match inside a table) | no — scope to the cell anchor (`table:N:R:C`) |
| `word_busy` | a modal dialog is open | **yes** — back off and retry |
| `word_not_running` | Word isn't running | no — until Word is opened |
| `excel_not_available` | `insert_chart` needs Excel installed | no — install Excel |
| `document_not_found` | no document by that name | no |
| `path_not_allowed` | a save/image path was outside the whitelist (or saving is off) | no — configure `WORDLIVE_SAVE_DIRS` / use a local image |
| `error` | bad input / other | no — fix the request |

## How it works

All Word access funnels through a single dedicated, COM-initialised worker thread,
so COM stays on one apartment and concurrent tool calls serialise — matching
Word's single-threaded reality. Each call attaches to the running instance fresh,
exactly like the CLI. The server is in-process: it calls the wordlive Python API
directly rather than shelling out, which is also how `word_snapshot` returns native
image content.
