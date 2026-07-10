---
name: wordlive-mcp
description: Read and edit the Microsoft Word document the user has open right now, over MCP. Four dispatch tools — word_read (inspect structure: outline, paragraphs, tables, formatting), word_write (one atomic-undo edit), word_exec (a batch of edits as one undo), word_snapshot (render a page to an image a vision model can see). All target stable anchors, never the live cursor; edits preserve the user's selection and scroll and collapse into a single Ctrl-Z. Use when the user wants to read, edit, or visually render a .docx open in Word on Windows through an MCP client (Claude Desktop, Cursor, …).
---

# wordlive (MCP)

`wordlive` drives a **running** Microsoft Word instance over COM (Windows only).
Unlike `python-docx`, it edits the document the user has **open right now** — and
politely: their cursor, selection, and scroll position are preserved, and every
edit collapses into a single Ctrl-Z.

You reach it through **four dispatch tools**. Most work is choosing a `command`
(or, for a batch, an `op`) and its fields — the tool list stays tiny so its
schemas cost little context:

| Tool | What it does |
| ---- | ------------ |
| `word_read`     | Every read, dispatched on `command` (`outline`, `paragraphs`, `find`, `format_info`, `lint`, …). Never mutates. |
| `word_write`    | One **atomic-undo** edit, dispatched on `command` (`insert`, `replace`, `apply_style`, `table`, …). |
| `word_exec`     | A batch of `ops` applied as a **single** undo step — the power tool for multi-step intents. |
| `word_snapshot` | Render page(s) to PNG so you can *see* the layout — real fonts, spacing, page geometry. |

Every tool takes an optional `doc` (target an open document by name; default is
the active one).

## First, orient yourself
1. `word_read(command="status")` — confirm Word is reachable; lists open documents (no doc needed).
2. `word_read(command="outline")` — the heading tree with `heading:N` ids (`all_paragraphs=true` = every paragraph).
3. `word_read(command="paragraphs")` — every paragraph with `para:N` ids and char offsets.
4. `word_read(command="find", text="phrase")` — fuzzy locate (tolerant of whitespace + smart quotes); emits `range:START-END` ids.
5. `word_read(command="guide")` — this guide, as a tool result (no Word needed). Also the `wordlive://guide` resource.

## Anchors — how you address things
Operations target stable **anchors**, never the live cursor. An anchor id is a
short string you pass as `anchor_id` (or `in_anchor` / `within` / `start_anchor`,
per command):

| Anchor id | Refers to |
| --------- | --------- |
| `heading:N` | the Nth *paragraph*, which must be a heading — same index space as `para:N`, so the first heading is rarely `heading:1` (copy the id from `outline`) |
| `para:N` | the Nth paragraph, any kind (see `paragraphs`) |
| `bookmark:NAME` | a bookmark |
| `pin:CODE` | a **durable handle** minted by the `pin` / `pin_outline` write — survives renumbering |
| `cc:NAME` | a content control (by title) |
| `footnote:N` / `endnote:N` | the Nth note's body (1-based; see `footnotes` / `endnotes`) |
| `image:N` | the Nth embedded picture (1-based; see `images`) |
| `shape:N` | the Nth floating shape — text box, floating image, WordArt (see `shapes`) |
| `chart:N` / `equation:N` | the Nth chart / equation (see `charts` / `equations`) |
| `table:N:R:C` | row R, col C of the Nth table (all 1-based) |
| `table:N:row:R` / `table:N:col:C` | a whole row / column of the Nth table — a styling handle |
| `range:START-END` | a raw character span (what `find` emits) |
| `header:S:WHICH` / `footer:S:WHICH` | header/footer of section S (`primary` / `first` / `even`) |
| `start` / `end` | the position before the first / past the last paragraph (prepend / append targets) |

`heading:N` and `para:N` are **positional** — they renumber when a structural
edit (a new paragraph, an inserted table) shifts the document, so re-read
`outline` / `paragraphs` after one before reusing ids downstream (an
`anchor_not_found` error is the signal you skipped that; it carries a recovery
hint). `bookmark:NAME` and `cc:NAME` are name-based and survive edits. For ad-hoc
content with no name, `word_write(command="pin", anchor_id=…)` mints a `pin:CODE`
that Word keeps attached to the same content across later inserts/deletes
(`command="pin_outline"` pins every heading at once). `shape:N`, `chart:N`,
`equation:N`, and `image:N` are positional too — re-list rather than cache an id.

## Coming from the CLI or the docs? The name map
Command and field names on these tools differ from the `wordlive` CLI. If you know
the CLI (or read the CLI/Python docs), translate:

| CLI | MCP |
| --- | --- |
| `wordlive read format --anchor-id ID` | `word_read(command="format_info", anchor_id=ID)` |
| `wordlive read markdown --within ID` | `word_read(command="to_markdown", within=ID)` |
| `wordlive read html` | `word_read(command="to_html")` |
| `wordlive read text --anchor-id ID --view final` | `word_read(command="read_text", anchor_id=ID, view="final")` |
| `wordlive read bookmark NAME` / `read cc NAME` | `word_read(command="read_bookmark", name=NAME)` / `command="read_cc"` |
| `wordlive read section "Heading"` | `word_read(command="read_section", heading="Heading")` |
| `wordlive find-paragraph --text "…"` | `word_read(command="find_paragraphs", text="…")` |
| `wordlive table read N` / `table records N` | `word_read(command="table_read", table=N)` / `command="table_records"` |
| `wordlive list show` / `list info --anchor-id ID` | `word_read(command="list_levels", anchor_id=ID)` (there is **no** `list info`; `list_levels` is the per-level read) |
| `wordlive lint --rule typography --within ID --profile PATH` | `word_read(command="lint", rules=["typography"], within=ID, profile=PATH_or_object)` |
| `wordlive insert --anchor-id ID --text "…"` | `word_write(command="insert", anchor_id=ID, text="…")` |
| `wordlive apply_list --anchor-id ID --type bulleted` | `word_write(command="list", action="apply", anchor_id=ID, type="bulleted")` |
| `wordlive comment add` / `revision accept` | `word_write(command="comment", action="add", …)` / `command="revision", action="accept"` |
| `wordlive table add-row --table N` | `word_write(command="table", action="add_row", table=N)` |
| `wordlive style apply --anchor-id ID --name "Heading 2"` | `word_write(command="apply_style", anchor_id=ID, name="Heading 2")` |
| `wordlive exec --script ops.json` | `word_exec(ops=[…])` |
| `wordlive snapshot --anchor-id ID` | `word_snapshot(anchor=ID)` |

Three shifts drive most of the friction. **Reads that are noun-like on the CLI
gain a verb prefix** (`read format` → `format_info`, `read markdown` →
`to_markdown`). **The CLI's flags become plain fields** — `--within` → `within`,
`--rule` → `rules`, `--profile PATH` → `profile` (which here also accepts an
**inline object**, not just a path). And **sub-commands become an `action`
field**: `apply_list`, `comment add`, `table add-row` are all one `command` with
an `action` (see below). If you reach for an `op` name (`add_row`) as a `command`,
the error tells you the `(command, action)` pair to use instead.

## word_read — inspect the document
`word_read(command=…)`. Grouped by what you want:

- **Structure:** `outline` `[all_paragraphs]` · `paragraphs` `[start,count]` · `sections` · `stats` · `table_list` · `table_read {table}` · `table_records {table}` (body rows as header-keyed dicts).
- **Locate:** `find {text,[in_anchor],[mode=fuzzy|literal|regex]}` (exact substring → `range:` ids) · `find_paragraphs {text,[limit,min_score]}` (fuzzy, typo-tolerant → `para:N` + scores) · `nearest_heading {anchor_id,[direction]}` · `location {anchor_id}` (page/line/column — "what page is this on" without a snapshot).
- **Content:** `read_text {anchor_id,[view=raw|final|original|segments]}` (`final`/`original` resolve tracked changes) · `read_bookmark {name}` · `read_cc {name}` · `read_section {heading|anchor_id}` · `between {start_anchor,end_anchor,[inclusive]}` · `to_markdown {[within]}` / `to_html {[within]}` (serialise the doc or one anchor's range; the read mirror of `insert_markdown`) · `digest {[budget,depth]}` (a token-budgeted whole-document read — headings verbatim, tables as stubs, body sampled; drill in with `to_markdown {within}`).
- **Objects:** `footnotes` · `endnotes` · `images` · `read_image {anchor_id}` (**SEE** a picture — returns an inline image block) · `equations` · `charts` · `shapes` · `hyperlinks` · `fields` · `comments` · `revisions` · `properties` · `variables` · `watermark`.
- **Formatting:** `format_info {anchor_id}` (effective paragraph/char formatting, each field with its style baseline + an `override` flag — the linter's substrate) · `list_levels {anchor_id}` (per-level list format) · `styles` · `theme` · `themes` · `proofing` (spelling/grammar + readability).
- **Audit & verify:** `lint {[rules],[within],[profile]}` (below) · `checkpoint {[include],[within]}` then `diff {checkpoint | cp_a,cp_b}` — the way to answer "what changed" (Word fires no content-change event) and to confirm your own edits landed.

The `word_read` tool description carries the full per-command field list.

## word_write — one atomic-undo edit
`word_write(command=…, …)`. Every call is a single Word `UndoRecord`: one Ctrl-Z
reverts the whole intent. The tool description spells out every command's fields;
the shape to internalise is the four **sub-dispatchers**, which take an `action`:

- `command="list"`, `action=` `apply` | `remove` | `restart` | `indent` | `outdent` | `format`
- `command="comment"`, `action=` `add` | `resolve` | `delete`
- `command="revision"`, `action=` `accept` | `reject` | `accept_all` | `reject_all`
- `command="table"`, `action=` `create` | `set_cell` | `add_row` | `delete_row` | `add_column` | `delete_column` | `merge_cells` | `split_cell` | `append_record` | `update_row` | `set_heading_row` | `autofit` | `delete` | `set_style` | `set_alignment` | `set_borders` | `set_banding`

A missing or misspelled `action` lists the valid ones for that command.

Everything else is a flat command: `insert`, `insert_block`, `insert_section`,
`insert_markdown`, `replace_section`, `delete_paragraph`, `append`, `prepend`,
`replace`, `write_bookmark`, `write_cc`, `apply_style`, `format_paragraph`,
`format_run`, `set_shading`, `set_borders`, `cell_valign`, `drop_cap`,
`add_tab_stop`, `add_style`, `set_style`, `header`, `footer`, `track`,
`watermark`, `text_box`, `insert_image`, `insert_equation`, `insert_chart` and
its formatting siblings (`format_chart`, `format_axis`, `add_trendline`,
`set_series_color`, `format_series`, `add_error_bars`), the `set_shape_*` /
`set_image_*` / `replace_shape_image` / `group_shapes` / `ungroup_shape` shape
verbs, `insert_break`, `insert_field`, `update_fields`, `regularize`,
`set_property`, `delete_property`, `set_variable`, `delete_variable`,
`insert_footnote`, `insert_endnote`, `insert_toc`, `add_bookmark`, `pin`,
`pin_outline`, `add_hyperlink`, `set_hyperlink`, `insert_cross_reference`,
`insert_caption`, `create_content_control`, `set_cc_properties`, `set_cc_items`,
`mark_index_entry`, `insert_index`, `insert_table_of_figures`,
`set_bibliography_style`, `add_source`, `insert_citation`, `insert_bibliography`,
`mark_citation`, `insert_table_of_authorities`, `apply_theme`, `set_theme_colors`,
`set_theme_fonts`, `page_setup`, and the **gated** `save` / `save_as` /
`export_pdf` (see [Saving](#saving)).

A few contracts worth knowing up front:
- **`insert`** takes `text` (literal) **or** `runs` (`[{text,bold?,italic?,underline?,code?,style?}]`) for inline-formatted spans in one op; `before=true` inserts above the anchor (default is after).
- **`insert_block` / `insert_section` / `insert_markdown`** build a run of styled paragraphs at once. A body item with **no `style` gets `Normal`**, never the insertion point's style — so a body paragraph can't silently come out as a heading. Pass `style` to match the surroundings. To *start* a blank document, lead with one of these (they reuse the doc's lone empty paragraph); `append` can't — it promises a *new* final paragraph, stranding the empty one above your content.
- **`replace {text, anchor_id | find, [all,occurrence,in_anchor,mode]}`** overwrites a range, or find-replaces. Default `mode=fuzzy` is whitespace/smart-quote tolerant; `regex` makes `find` a Python regex and lets `text` use `\1`. To edit **inside a table**, scope with `in_anchor="table:N:R:C"` — an unverifiable whole-document match fails rather than risk the wrong cell.
- Field/reference blocks (`insert_toc`, `insert_index`, `insert_bibliography`, cross-refs, captions) fill in page numbers only after `update_fields` (or a `word_snapshot`, which repaginates).

## word_exec — many edits, one undo
`word_exec(ops=[…], [label], [tracked])` applies a batch as a **single** undo
step. This is the right path for multi-step intents and for inline base64 images
(no argv limits). `ops` is a JSON **array** of `{"op": "<kind>", …}` objects:

```json
{
  "ops": [
    {"op": "write_bookmark",   "name": "Address",      "text": "123 Main St"},
    {"op": "insert_paragraph", "anchor_id": "heading:2", "text": "New section.", "style": "Body Text"},
    {"op": "insert_image",     "anchor_id": "heading:2", "base64": "<base64…>", "wrap": "auto"},
    {"op": "find_replace",     "find": "Q3", "text": "Q4", "all": true}
  ],
  "tracked": false
}
```

**The `op` names are not the `word_write` command names** — a batch speaks the op
vocabulary. The most common shift: a single insert is `word_write(command="insert")`,
but in a batch it's `{"op": "insert_paragraph"}`. Likewise the sub-dispatcher
actions have flat op names — `{"op": "add_row"}`, `{"op": "apply_list"}`,
`{"op": "add_comment"}` — no `action` field inside a batch.

The full op vocabulary (every `word_write` capability has one):
`write_bookmark`, `write_cc`, `insert_paragraph`, `insert_block`,
`insert_section`, `insert_markdown`, `replace_section`, `delete_paragraph`,
`append`, `append_paragraph`, `append_inline`, `prepend`, `prepend_paragraph`,
`prepend_inline`, `insert_image`, `insert_equation`, `insert_chart`,
`format_chart`, `format_axis`, `add_trendline`, `set_series_color`,
`format_series`, `add_error_bars`, `set_shape_wrap`, `set_shape_crop`,
`set_shape_position`, `set_shape_size`, `format_shape`, `set_shape_alt_text`,
`set_shape_text`, `set_shape_rotation`, `set_shape_z_order`,
`set_shape_text_frame`, `replace_shape_image`, `delete_shape`, `group_shapes`,
`ungroup_shape`, `set_image_alt_text`, `set_image_size`, `set_image_crop`,
`replace`, `find_replace`, `apply_style`, `format_paragraph`, `format_run`,
`set_shading`, `set_borders`, `drop_cap`, `add_tab_stop`, `add_style`,
`set_style`, `insert_field`, `set_page_setup`, `update_fields`, `regularize`,
`insert_footnote`, `insert_endnote`, `insert_toc`, `add_bookmark`, `pin`,
`pin_outline`, `add_hyperlink`, `set_hyperlink`, `insert_cross_reference`,
`insert_caption`, `create_content_control`, `set_cc_properties`, `set_cc_items`,
`mark_index_entry`, `insert_index`, `insert_table_of_figures`,
`set_bibliography_style`, `add_source`, `insert_citation`, `insert_bibliography`,
`mark_citation`, `insert_table_of_authorities`, `apply_theme`, `set_theme_colors`,
`set_theme_fonts`, `set_property`, `delete_property`, `set_variable`,
`delete_variable`, `set_cell`, `add_row`, `add_column`, `append_record`,
`update_row`, `delete_row`, `delete_column`, `merge_cells`, `split_cell`,
`set_heading_row`, `autofit_table`, `set_table_style`, `set_table_alignment`,
`set_table_borders`, `set_table_banding`, `set_cell_vertical_alignment`,
`create_table`, `delete_table`, `insert_break`, `add_comment`, `resolve_comment`,
`delete_comment`, `accept_revision`, `reject_revision`, `accept_all_revisions`,
`reject_all_revisions`, `set_watermark`, `remove_watermark`, `insert_text_box`,
`apply_list`, `apply_list_format`, `remove_list`, `restart_numbering`,
`indent_list`, `outdent_list`, `write_header`, `write_footer`.
(`append` / `prepend` add a new final / first **paragraph** and take `text` +
optional `style`; `append_paragraph` / `prepend_paragraph` are explicit synonyms.
`append_inline` / `prepend_inline` instead **continue** the last / first paragraph
and take `text` only.) The `word_exec` tool description carries each op's fields.

**Atomicity and ordering — read this once:**
- The batch is **one undo entry**, not a database transaction. It **stops at the first failing op** and reports `failure` with that op's `index`, `error`, and `type` — but the ops **before** it already ran and are **not rolled back**; they simply share the one undo step. So "atomic" means *one Ctrl-Z reverts the whole intent*, not *all-or-nothing*.
- Each op resolves its `anchor_id` **fresh against the live document at the moment it runs** — there is **no pre-batch snapshot**. A positional id (`heading:N` / `para:N`) therefore sees the shifts that **earlier ops in the same batch** made, exactly as it would across separate calls. When inserting several paragraphs after one fixed anchor, either insert in reverse order or anchor each to the previous insert; name-based ids (`bookmark:` / `cc:`) are stable across the batch.
- A field an op doesn't use (a typo, or `style` on an inline append) comes back in a `warnings` array on the result rather than being silently dropped — so a clean-looking success can't hide a payload you got wrong.
- `tracked=true` records the whole batch as tracked changes (Track Changes is restored to its prior state afterward).

**Durable handles in a batch.** Add `bind: "slug"` (or `true` for a random code)
to an `insert` / `insert_block` / `insert_section` / `insert_markdown` /
`create_table` op to mint a `pin:` handle on the new content — it comes back as
`pin` in that op's `outputs` entry. And any op field of the exact form
`$ops[N].field` is replaced with an earlier op's recorded output before the op
runs, so a batch can create then target without a round-trip:
`[{"op":"create_table",…}, {"op":"set_cell","table":"$ops[0].table",…}]`.

## word_snapshot — see the layout
`word_snapshot([anchor], [pages], [dpi=150], [max_dim], [markup="none"])` exports
a pixel-faithful PDF of the live document and returns the requested page(s) as
inline **image content** — a true WYSIWYG render (real fonts, spacing, page
geometry), ideal for judging or iterating on style and formatting.
- Pick **at most one** target: `anchor` (the page(s) it occupies — a `heading:` expands to its **whole section**) or `pages` (`"4"` or `"2-5"`). With neither, the whole document renders.
- `max_dim` caps each page's long edge to that many pixels (only ever lowering resolution) — pair it with no page target to eyeball the whole document's layout cheaply. A vision model is billed on pixel area, so the cap is a predictable per-page token budget (~1000 stays legible).
- `markup="all"` renders tracked changes and comments as visible revision marks (default `"none"` shows the final document); the structured list is `word_read(command="revisions")`.
- Needs the server's `snapshot` extra (PyMuPDF).

## Linting & house style
`word_read(command="lint", [rules], [within], [profile])` audits publishing-quality
defects and returns severity-ranked findings; each `fixable` one carries the `fix`
op that `regularize` would run. Then
`word_write(command="regularize", [rules], [within], [profile], [dry_run], [allow_content])`
applies the fixable subset in **one atomic-undo step** (targeted + idempotent — a
second pass is a no-op), returning `{applied, skipped, deferred, findings}`.
Formatting fixes apply by default; content-changing fixes (insert a caption/notice,
delete a stray paragraph, strip a watermark) are flagged `adds_content` and held in
`deferred` unless you pass `allow_content=true`.

**`rules`** selects ids or tags, or `{exclude:[…]}`. The default set is the
on-by-default rules; naming a **tag** pulls in that cluster — *including its
off-by-default members*:
- `["typography"]` — text hygiene: trailing/leading/double spaces, space-before-punctuation, hyphen→en-dash, manual "headings" that aren't styled, tables on mismatched styles (the tag also enables its off-by-default opinion rules).
- `["structure"]` (or `["headings"]`) — the heading & document-structure cluster: `heading-level-skip`, `empty-heading` (on), plus off-by-default `adjacent-headings`, `heading-numbering-manual`, `heading-trailing-period` (fixable — strips it in place), `toc-present-and-current` (the tag also enables these off-by-default members).
- `["finalization"]` — the off-by-default "ready-to-send?" cluster: leftover comments/revisions, track-changes-on, hidden text, highlight, updatable fields.
- `["academia"]` — field-code cluster: `broken-cross-reference`, `caption-manual-numbering` (on), plus off-by-default `xref-as-literal-text`.
- `["hyperlinks"]` — `hyperlink-broken-internal` (on), plus off-by-default `hyperlink-bare-for-print`, `hyperlink-display-is-raw-url` (which `["print"]` selects on their own).
- `["layout"]` — the page-layout/document-level cluster, **all off by default** (naming the tag enables its off-by-default opinion rules): `page-numbers-present`, `header-footer-consistent`, `draft-watermark-present`, `document-properties-filled`; `["notices"]` selects the profile-driven `confidentiality-notice` / `copyright-notice` on their own.

**`profile`** is a house-style config that turns on the **policy** rules
(`body-justified`, `body-line-spacing`, `table-numeric-right-align`) and supplies
their targets — over MCP you may pass either a **path** or an **inline object**.
The JSON shape:

```json
{
  "extends": "default",
  "rules": {
    "body-justified":            { "enabled": true, "severity": "warning" },
    "body-line-spacing":         { "enabled": true, "target": "1.5" },
    "table-numeric-right-align": { "enabled": true, "threshold": 0.8 },
    "double-space":              { "enabled": false }
  }
}
```

| key | effect |
| --- | --- |
| `rules.<id>.enabled` | `true` opts a policy rule in (a bare `<id>: {}` mention also enables it); `false` disables a rule that's otherwise on |
| `rules.<id>.target` | the value a policy rule measures against (`body-line-spacing` needs `"single"`/`"1.5"`/`"double"`, or it no-ops) |
| `rules.<id>.threshold` | numeric cutoff (`table-numeric-right-align`: the fraction of a column's cells that must parse numeric) |
| `rules.<id>.severity` | override the finding severity (`error`/`warning`/`info`) |
| `extends` | accepted and recorded; only `"default"` is meaningful today |

Passed as `profile`, the same object (or path) drives both `lint` and
`regularize` — lint to see the findings, regularize to apply the fixable ones.

## Saving — gated, hand back a deliverable
`save`, `save_as {path,[overwrite]}` (a `.docx`), and `export_pdf {path,[from_page,to_page]}`
(a PDF — the recommended deliverable, a pixel-faithful render) are **default-deny**:
they only write inside directories the operator whitelisted at server launch via
`WORDLIVE_SAVE_DIRS`. With none set, every save fails (`path_not_allowed`). Image
sources may likewise be restricted with `WORDLIVE_IMAGE_DIRS`; regardless,
`insert_image` with a non-local `path` (UNC / `file://` / URL) is always rejected —
prefer `image_base64` for LLM-supplied images.

## Errors — branch on `code`
A failed tool call comes back flagged as an error whose message is a JSON object:

```json
{"error": "bookmark not found: 'Addr'", "code": "anchor_not_found", "retryable": false, "type": "AnchorNotFoundError"}
```

| `code` | Meaning | `retryable` |
| ------ | ------- | ----------- |
| `anchor_not_found` | anchor missing, or `find` matched zero | no — re-read `outline`/`paragraphs` first |
| `style_not_found` | a named style isn't defined in the document | no — read `styles` first |
| `ambiguous_match` | `find` matched several | yes — pass `occurrence`/`all` |
| `replace_verification` | a `find_replace` target couldn't be verified (e.g. a whole-doc match inside a table) | no — scope to the cell (`in_anchor="table:N:R:C"`) |
| `word_busy` | a modal dialog is open | **yes** — back off and retry |
| `word_not_running` | Word isn't running | no — until the user opens Word |
| `excel_not_available` | `insert_chart` needs Excel installed | no — install Excel |
| `document_not_found` | no open document by that name | no |
| `path_not_allowed` | a save/image path was outside the whitelist (or saving is off) | no — configure `WORDLIVE_SAVE_DIRS` / use a local image |
| `error` | bad input / other | no — fix the request |

**Misspelled a `command`, `op`, or `action`?** The error suggests rather than
reciting the whole vocabulary. A near-miss names the likely target
(*"did you mean `format_info`?"*). Three cases are handled exactly:
- An **`op` name used as a `command`** (`add_row`, `apply_list`) is answered with the `(command, action)` pair to use: *"it is the `add_row` action of command `table` — call `word_write(command='table', action='add_row')`"*.
- A **missing/wrong `action`** lists every valid action for that command.
- **`table:2`** (a whole table — a collection, not one anchor) points you at `table:2:R:C`, `table:2:row:R`, `table:2:col:C`, and `table_read`.

A name with no near neighbour gets no guess — just a count and a pointer here.

## Typical workflow
1. `word_read(command="status")` → confirm Word and the target document.
2. `word_read(command="outline")` / `"paragraphs"` / `"find"` → get anchor ids (for a large doc, `"digest"` first, then drill with `"to_markdown"` `{within}`).
3. For a multi-step session, `word_write(command="pin"|"pin_outline")` the blocks you'll revisit so positional ids can't drift under your own edits.
4. Edit with `word_write`, or batch related changes with `word_exec` (one user-visible intent = one batch = one undo). Suggest rather than overwrite where it fits — `word_write(command="comment", action="add")`, or `command="track", on=true` for accept/reject-able tracked changes.
5. Verify — Word fires **no** content-change event, so `word_read(command="checkpoint")` first and `"diff"` it afterward to confirm exactly which paragraphs changed; `word_snapshot` when the change is visual.

For the CLI or the Python API instead, see the `wordlive-cli` / `wordlive-python`
skills. Full docs: https://thomas-villani.github.io/wordlive/
