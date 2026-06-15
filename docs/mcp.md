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
| `word_read` | Every read, dispatched on `command`: `guide`, `status`, `outline`, `paragraphs`, `find`, `read_bookmark`, `read_cc`, `read_section`, `table_list`, `table_read`, `table_records` `{table}` (body rows as dicts keyed by the header row — the read mirror of building a table from records), `styles`, `comments`, `revisions` (tracked changes), `track` (is Track Changes on?), `sections`, `footnotes`, `endnotes`, `images` (embedded pictures: `image:N` id, mime, size, alt, para), `read_image` `{anchor_id}` (an embedded picture's bytes as base64 + mime — pass an `image:N` id or any single-image anchor), `equations` (math zones: `equation:N` id, type, linear preview, para), `hyperlinks` (links: text, address/sub_address, `range:START-END` — the read mirror of `add_hyperlink`), `fields` (PAGE/REF/TOC/…: kind, code, rendered result, `range` — the read mirror of `insert_field`), `properties` (document metadata: `{builtin, custom}` name→value bags), `variables` (invisible named storage: `{name: value}`), `proofing` (spelling/grammar errors with counts + flagged runs, and readability statistics — heavier than `stats`; it (re)checks the document), `location` `{anchor_id}` (where an anchor sits in the laid-out document: `page`/`end_page` span, `line`, `column`, `in_table` — "what page is this on" without a snapshot), `stats` (one-shot document summary: page/word/character/paragraph/line counts plus section/heading/table/image/equation/comment/revision counts and `saved`), `theme` (the document theme: 12 brand colours as `#RRGGBB` + major/minor fonts — the read mirror of `apply_theme`/`set_theme_*`), `themes` (the built-in themes, colour schemes, and font schemes Office ships — the names the theme write commands accept). |
| `word_write` | Every single atomic-undo edit, dispatched on `command`: `insert`, `insert_block`, `insert_section`, `insert_markdown`, `replace_section`, `delete_paragraph`, `insert_break`, `append`, `prepend`, `replace`, `write_bookmark`, `write_cc`, `apply_style`, `format_paragraph`, `format_run`, `set_shading`, `set_borders`, `add_tab_stop`, `add_style`, `set_style`, `list`, `comment`, `table` (`action` = `set_cell`/`add_row`/`append_record`/`update_row`/`delete_row`/`set_heading_row`/`autofit`/`create`/`delete`), `header`, `footer`, `track`, `insert_image`, `insert_equation`, `insert_field`, `update_fields`, `set_property`, `delete_property`, `set_variable`, `delete_variable`, `insert_footnote`, `insert_endnote`, `insert_toc`, `insert_table_of_figures`, `mark_index_entry`, `insert_index`, `set_bibliography_style`, `add_source`, `insert_citation`, `insert_bibliography`, `mark_citation`, `insert_table_of_authorities`, `apply_theme`, `set_theme_colors`, `set_theme_fonts`, `create_content_control`, `add_bookmark`, `add_hyperlink`, `insert_cross_reference`, `insert_caption`, `page_setup`. `insert {anchor_id, text \| runs, [style,before]}` adds one paragraph — `text` is literal, `runs` is `[{text,bold?,italic?,underline?,style?}]` for inline-formatted spans. `insert_block {anchor_id, items, [before]}` inserts a contiguous run of styled paragraphs in one op (each item a string or `{text \| runs, style?}`, where `text` takes `**bold**`/`*italic*` markdown) and returns the block's `range:START-END`. `insert_section {anchor_id, heading, body, [level,before]}` places a `Heading {level}` paragraph plus its body (the `insert_block` items shape) in one op; `insert_markdown {anchor_id, markdown, [before]}` maps a constrained-Markdown subset (`#`/`##`/`###` headings, `-`/`*` bullets, `1.` numbers, blank-line paragraphs, inline `**bold**`/`*italic*` — not CommonMark) to real Word structure; `replace_section {anchor_id, body \| markdown}` rewrites a `heading:N`'s body up to the next same-or-higher heading while keeping the heading. `delete_paragraph {anchor_id}` removes the paragraph(s) at an anchor, mark included. `append`/`prepend` add a new paragraph (optional `style`); pass `paragraph: false` to continue the adjacent paragraph inline (no `style`). `table action="create"` needs `anchor_id` and `rows`/`cols` — both optional when `data` is given (inferred from it); `data` is a 2-D array **or** records (objects whose keys become a header row). `format_paragraph` also carries `line_spacing` (a multiple like `1.5`, `single`/`1.5`/`double`, or an exact length like `14pt`) and the pagination controls `keep_together`/`keep_with_next`/`widow_control` (multi-page layout); `table action="set_heading_row"` `{table,[row=1,heading,allow_break]}` marks a repeating header row; `table action="append_record"` `{table,record}` appends a row from a `{header: value}` object, and `table action="update_row"` `{table,key,values,[column]}` sets cells (`values` = `{header: value}`) on the first row whose key-column (`column`, default first) equals `key`. `format_run` styles characters (bold/italic/colour/highlight/…); `set_shading`/`set_borders`/`add_tab_stop` add cell/range shading, borders, and tab stops; `add_style`/`set_style` create and configure styles (the border line style is the `line_style` param, to avoid colliding with `style`). `insert_equation {anchor_id, unicodemath \| latex \| mathml, [display,before]}` inserts a math equation on its own paragraph (UnicodeMath is native, LaTeX needs the server's `latex` extra, MathML uses Office's transform; `display=true` gives it the centred `Equation` style, `display=false` is Normal + left — the paragraph style is pinned either way so it never inherits a neighbouring heading) and returns the new `equation:N` in `result` — a positional id (OMaths order) that renumbers when an earlier equation is inserted, so re-list rather than caching it. `insert_field` drops a self-updating field (page numbers, dates, refs) — put `kind: page` in a footer; `update_fields` refreshes them; `page_setup` sets a section's margins/orientation/paper size/`columns`. `set_property {name,value,[custom]}` / `delete_property {name}` write the document's metadata (a built-in like Title/Author, or a custom property with `custom=true`); `set_variable {name,value}` / `delete_variable {name}` manage the invisible `DOCVARIABLE` storage; `table action="autofit" {table,[mode=content|window|fixed]}` resizes a table's columns (fit to content/window or pin them). `insert_footnote`/`insert_endnote` attach a note (the new `footnote:N`/`endnote:N` comes back in `result`); `insert_toc` inserts a table of contents (run `update_fields` after to fill its page numbers). `insert_table_of_figures {anchor_id (default start), [label=Figure, include_label, hyperlinks, right_align_page_numbers, before]}` lists the document's captions of one label as a table of figures. `mark_index_entry {anchor_id, entry, [cross_reference, bold, italic]}` marks a range as a back-of-book index entry (`entry` uses `main:sub` to nest), then `insert_index {anchor_id (default end), [columns, run_in, right_align_page_numbers, before]}` builds the index from those marks — both field blocks, like the TOC, so `update_fields` after to fill page numbers. Citations are a three-step flow: `set_bibliography_style {style}` picks the scheme (`APA`/`MLA`/`Chicago`/`IEEE`/… — build-dependent), `add_source {source_type, [tag, author, title, year, publisher, city, journal_name, volume, issue, pages, url, edition, doi] | xml}` registers a source (`source_type` is `book`/`journal_article`/`web_site`/`case`/…; `author` is `Last, First`; `tag` auto-derives from surname+year if omitted; `xml` is the raw `<b:Source>` escape hatch) and returns its tag in `result`, then `insert_citation {anchor_id, tag, [pages, prefix, suffix, volume, suppress_author, suppress_year, suppress_title, locale, before]}` cites it (an unknown tag renders "Invalid source specified.") and `insert_bibliography {anchor_id (default end), [before]}` inserts the works-cited block — a field block, so `update_fields` after. A table of authorities is the mark-then-build pattern: `mark_citation {anchor_id, long_citation, [short_citation, category, before]}` marks a `TA` entry (`category` is `cases`/`statutes`/`other`/… or an int 1–16; `short_citation` defaults to `long_citation`), then `insert_table_of_authorities {anchor_id (default end), [category (default all), passim, keep_entry_formatting, entry_separator, page_range_separator, before]}` builds the table from those marks (a field block, so `update_fields` after to fill page numbers). The document **theme** is the brand primitive: `apply_theme {theme}` applies a whole theme (colours+fonts+effects) by built-in name (read `command="themes"`) or `.thmx` path, `set_theme_colors {[scheme, colors]}` loads a named colour scheme and/or overrides individual brand colours (`colors` is `{accent1|text1|background1|…: name/hex}`), and `set_theme_fonts {[scheme, major, minor]}` sets the heading/body fonts. `create_content_control {anchor_id, [kind=rich_text, title, tag, items, where=wrap, lock_contents, lock_control]}` creates a content control — `where=wrap` surrounds the anchor's range (`before`/`after` insert a fresh empty one), `items` populates a `combo_box`/`dropdown`, and a `title` (or `tag`) makes it addressable as `cc:TITLE`. `add_bookmark` names a range; `add_hyperlink` links it (external `url` or internal `bookmark`); `insert_cross_reference` references a `bookmark:`/`heading:`/`footnote:`/`endnote:` target; `insert_caption` adds a numbered caption in its own paragraph (`position` = `above`/`below`, default above for a `Table` and below otherwise; on a `table:N:R:C` anchor it captions the whole table). Plus the **gated** persistence commands `save` (to the existing file), `save_as {path,[overwrite]}` (a `.docx`), and `export_pdf {path,[from_page,to_page]}` (a PDF deliverable) — terminal side-effects, not undoable; see [Saving](#saving). |
| `word_exec` | Apply a batch of `ops` as a **single** atomic undo — the power tool for multi-step intents. |
| `word_snapshot` | Render page(s) to PNG so the model can *see* the layout (`markup: "all"` shows tracked changes / comments as revision marks). `max_dim` caps each page's long edge in pixels — pair it with no page target to check the whole document's layout cheaply (a predictable per-page token budget; ~1000 stays legible). Returns image content. Needs the `snapshot` extra. |

The anchor model (`heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`,
`footnote:N`, `endnote:N`, `image:N`, `table:N:R:C`, `range:START-END`,
`header:S:WHICH` / `footer:S:WHICH`, `start` /
`end`) and the full `word_exec` op vocabulary are documented in the one-page
guide. Fetch it as a tool call with **`word_read(command="guide")`** (it needs
neither Word nor a document) — the most reliable path, since not every MCP
client surfaces resources. The same text is also the **`wordlive://guide`**
resource and `wordlive llm-help`. See also [CLI](cli.md) for each op's fields.

`heading:N` / `para:N` are **positional** and renumber when a structural edit
shifts the document, so re-read `outline` / `paragraphs` after an insert before
reusing ids; `bookmark:NAME` / `cc:NAME` are name-based and survive edits.

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
