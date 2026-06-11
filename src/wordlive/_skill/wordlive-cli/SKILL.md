---
name: wordlive-cli
description: Read and edit the Microsoft Word document the user has open right now, from the command line. Inspect structure (outline, paragraphs, tables), make polite edits (text, styles, lists, images, comments, headers/footers), render a page or section to a PNG so a vision model can see the layout, and batch changes into a single atomic undo — all JSON-in / JSON-out with deterministic exit codes. Use when the user wants to read, edit, or visually render a .docx that is currently open in Word on Windows.
---

# wordlive (CLI)

`wordlive` drives a **running** Microsoft Word instance over COM (Windows only).
Unlike `python-docx`, it edits the document the user has **open right now** — and
politely: their cursor, selection, and scroll position are preserved, and every
edit collapses into a single Ctrl-Z.

Prefer the **CLI**. Every command prints exactly one JSON object on stdout and
returns a deterministic exit code, so you branch on failures without parsing
prose. JSON is the default; `--text` (human-readable), `--json`, and
`--doc NAME` (default: the active document) are **global flags — put them
before the subcommand**: `wordlive --text outline`, not `wordlive outline --text`.

## First, orient yourself
1. `wordlive status` — confirm Word is reachable and see open documents.
2. `wordlive outline` — heading tree with `heading:N` ids (`--all` = every paragraph).
3. `wordlive paragraphs` — every paragraph with `para:N` ids and char offsets.
4. `wordlive find --text "phrase"` — fuzzy locate (tolerant of whitespace + smart quotes); emits `range:START-END` ids.

## Anchors — how you address things
Operations target stable **anchors**, never the live cursor. An anchor id is a
short string you pass as `--anchor-id`:

| Anchor id            | Refers to |
| -------------------- | --------- |
| `heading:N`          | the Nth *paragraph*, which must be a heading — same index space as `para:N`, so the first heading is rarely `heading:1` (copy the id from `outline`) |
| `para:N`             | the Nth paragraph, any kind (see `paragraphs`) |
| `bookmark:NAME`      | a bookmark |
| `cc:NAME`            | a content control (by title) |
| `footnote:N` / `endnote:N` | the Nth note's body (1-based; see `footnotes` / `endnotes`) |
| `image:N`            | the Nth embedded picture (1-based; see `images`) |
| `table:N:R:C`        | row R, col C of the Nth table (all 1-based) |
| `range:START-END`    | a raw character span (what `find` emits) |
| `header:S:WHICH` / `footer:S:WHICH` | header/footer of section S (`primary` / `first` / `even`) |
| `start` / `end`      | the position before the first / past the last paragraph (prepend / append targets) |

`heading:N` and `para:N` are **positional** — they renumber when a structural
edit (a new paragraph, an inserted table) shifts the document, so re-read
`outline` / `paragraphs` after one before reusing ids downstream (an exit `2`
`anchor_not_found` is the signal you skipped that). `bookmark:NAME` and `cc:NAME`
are name-based and survive edits — reach for them when you need a durable handle.

## Reading
- `wordlive read bookmark NAME` (or `read bookmark --list` for every bookmark name) · `read cc NAME` · `read section "Heading Text"`
- `wordlive find --text "phrase"` — locate before editing; returns `range:` ids.
- `wordlive table list` · `wordlive table read N` · `wordlive table records N` — body rows as `{header: value}` dicts (row 1 is the header), the read mirror of building a table from records.
- `wordlive footnotes` · `wordlive endnotes` — each note's `footnote:N`/`endnote:N` id, text, and `para:N`.
- `wordlive images` — each embedded picture's `image:N` id, MIME, size, alt text, and `para:N`. Pull one out with `wordlive read-image --anchor-id image:N [--out FILE]` (`--out` writes the raw bytes; otherwise base64 + mime inline) — the path for handing a picture to a vision model.
- `wordlive revisions` — tracked changes as structured data (`type`/`author`/`text`/`range`); the readable counterpart to `snapshot --markup all`. `wordlive track status` reports whether Track Changes is on.
- `wordlive locate --anchor-id ID` — where an anchor sits in the laid-out doc: `page`/`end_page` span, `line`, `column`, `in_table`. Answers "what page is this on" without a `snapshot`. `wordlive stats` — one-shot summary: page/word/char/paragraph/line counts plus section/heading/table/image/comment/revision counts and `saved`. Both repaginate first (selection/scroll untouched); reads, non-mutating.

## Writing — each command is one atomic undo
- `wordlive write bookmark NAME --text "…"` (set existing) · `write bookmark NAME --create --anchor-id ID` (create a new bookmark over a range — the prerequisite for internal links and cross-refs; NAME starts with a letter, letters/digits/underscores only) · `write cc NAME --text "…"`
- `wordlive insert --anchor-id ID (--text "…" | --runs JSON) [--before | --after] [--style "Body Text"]` — new paragraph relative to any anchor (`--after` is the default; appending after the document's last paragraph works too, so you can build a doc top-down). `--text` is literal; `--runs '[{"text":"Bold","bold":true},{"text":" rest"}]'` gives inline-formatted spans in one op.
- `wordlive insert-block --anchor-id ID --items JSON [--before | --after]` — insert a **contiguous run of styled paragraphs** in one op (a whole bulleted section, a heading + body) in reading order, instead of reverse-ordered `insert` calls. Each item is `"plain text"` or `{text|runs, style?}`; item `text` takes `**bold**`/`*italic*` markdown. Reports the block's `range:START-END` — feed it to `list apply --anchor-id range:… --type bulleted` to bullet the section you just inserted.
- `wordlive insert-section --anchor-id ID --heading "…" --body JSON [--level N] [--before | --after]` — a `Heading {level}` paragraph plus its body (`--body` is the same items shape as `insert-block`) in one op. The "add a whole section" shortcut.
- `wordlive insert-markdown --anchor-id ID --markdown "…"|-` — drop a chunk of **constrained Markdown** as real Word structure: `#`/`##`/`###` headings, `-`/`*` bullets, `1.` numbers, blank-line paragraphs, inline `**bold**`/`*italic*`. A documented subset, not CommonMark (no code fences/nested lists/tables). Use `--markdown -` for multi-line stdin.
- `wordlive replace-section --anchor-id heading:N (--body JSON | --markdown "…")` — rewrite a heading's body (everything up to the next same-or-higher heading), **keeping the heading**. The "rewrite section X" workflow.
- `wordlive delete-paragraph --anchor-id ID` — remove the paragraph(s) at an anchor, **mark included**, so the text closes up (no blank line left, unlike `replace --text ""`). Good for a stray leading empty `para:1`.
- `wordlive append --text "…" [--inline] [--style "Body Text"]` — add a new final paragraph at the very end of the document (`--inline` continues the last paragraph). The high-level "end of doc" helper; same as `insert --anchor-id end`.
- `wordlive prepend --text "…" [--inline] [--style "Body Text"]` — the mirror: add to the very start of the document (same as `insert --anchor-id start`).
- `wordlive replace --anchor-id ID --text "…"` — overwrite a range.
- `wordlive replace --find "old" --text "new" [--all | --occurrence N] [--in ID]` — fuzzy find + replace. To edit **inside a table**, scope it to the cell (`--in table:N:R:C`); an unverifiable whole-document match fails (exit 1, `replace_verification`) rather than risk overwriting the wrong cell.
- `wordlive style apply --anchor-id ID --name "Heading 2"` (names: `style list`).
- `wordlive style add NAME [--type paragraph|character] [--based-on NAME] [--next-style NAME]` then `wordlive style set NAME [--bold] [--color "#1F3864"] [--size 16pt] [--alignment center] [--space-before 12] …` — define and configure a style, then `style apply` it everywhere (the brand/template workflow).
- `wordlive format-paragraph --anchor-id ID [--alignment center] [--left-indent 36] [--space-before 6] [--line-spacing 1.5] [--page-break-before] [--keep-together] [--keep-with-next] [--widow-control] …` — `--line-spacing` sets the leading within the paragraph (a multiple `1`/`1.5`/`2`, `single`/`1.5`/`double`, or an exact length like `14pt`). `--page-break-before` is the clean, reflow-safe way to make a paragraph (e.g. a `Heading 1`) start a new page, leaving no stray break character. The three pagination flags keep a paragraph's lines together, keep it with the next paragraph, and suppress widows/orphans.
- `wordlive format-run --anchor-id ID [--bold] [--italic] [--underline] [--font NAME] [--size 12pt] [--color "#FF0000"] [--highlight yellow] …` — **character** formatting; pair with a `range:` id (from `find`) to style a phrase. Colours take a name/hex/`r,g,b`; sizes take points or a unit string.
- `wordlive shading --anchor-id ID --fill "#FFF2CC"` · `wordlive borders --anchor-id ID [--sides all|top|…] [--style single|double|dot|dash] [--weight 0.5] [--color black]` · `wordlive tab-stop --anchor-id ID --position 3in [--align right] [--leader dots]` — range/cell fill, borders, and tab stops (a cell is an anchor, so these shade/border table cells too).
- `wordlive drop-cap --anchor-id ID [--position dropped|margin|none] [--lines 3] [--distance 2pt] [--font Georgia]` — the editorial oversized initial letter on a paragraph (a real Word DropCap, so the body text wraps around it natively). `--position none` removes it.
- `wordlive insert-break --anchor-id ID [--kind page|column|section_next|section_continuous] [--before | --after]` — an **explicit** one-off page/column/section break (the discoverable replacement for a literal form-feed paragraph). `--kind` defaults to `page`. Section breaks start a new section that can carry its own headers/footers + page setup. For a break that follows a *style*, prefer `format-paragraph --page-break-before`.
- `wordlive insert-field --anchor-id ID --kind page|numpages|date|time|filename|author|title|field [--text "RAW CODE"] [--before | --after]` — a **self-updating** field; put page numbers in a footer (`--anchor-id footer:1:primary --kind page`). `--kind field` takes a raw field code via `--text`. Refresh stale fields with `wordlive update-fields`.
- `wordlive insert-footnote --anchor-id ID --text "…"` · `wordlive insert-endnote --anchor-id ID --text "…"` — attach a footnote/endnote to a range; reports the new `footnote:N`/`endnote:N`. List them with `wordlive footnotes` / `wordlive endnotes`.
- `wordlive insert-toc [--anchor-id start] [--levels 1-3] [--no-heading-styles] [--no-hyperlinks]` — insert a table of contents from the headings. Page numbers fill in after `wordlive update-fields` (or `snapshot`).
- `wordlive link --anchor-id ID (--url URL | --bookmark NAME) [--text "…"]` — turn an anchor into a hyperlink: external `--url` or internal `--bookmark`. `--text` inserts new linked text (no `--text` links the existing range).
- `wordlive cross-ref --anchor-id ID --target TARGET [--kind text|page|number|above_below] [--no-hyperlink]` — reference another anchor; `--target` is `bookmark:NAME` / `heading:N` / `footnote:N` / `endnote:N`. Refresh with `update-fields`.
- `wordlive caption --anchor-id ID [--label Figure] [--text "…"] [--position above|below]` — insert an auto-numbered caption (Figure/Table/…) as its own `Caption` paragraph (on a `table:N:R:C` anchor it captions the whole table). Default placement is above for a `Table`, below otherwise; `--position` overrides. Pairs with `cross-ref`.
- `wordlive page-setup [--section N] [--margins 1in] [--top-margin … --bottom-margin … --left-margin … --right-margin …] [--gutter …] [--orientation portrait|landscape] [--paper-size letter|legal|tabloid|a3|a4|a5] [--columns N] [--column-spacing …]` — set a section's page geometry. `--margins` sets all four (per-side flags override); lengths take points or a unit string; `--columns N` makes N equal columns. `--section` defaults to `1`.
- `wordlive list apply --anchor-id ID --type bulleted|numbered|outline` (+ `list remove|restart|indent|outdent`; `list show` enumerates every list in the doc, `list info --anchor-id ID` reports a paragraph's list type/level/number). **To number N paragraphs 1..N, scope one apply to a `range:` that spans them all** (offsets from `paragraphs`/`find`) — applying `numbered` per-paragraph makes N separate "1." lists.
- `wordlive table create --anchor-id ID --rows R --cols C [--style NAME] [--header] [--before | --after] [--data '[["…"],…]' | --data -]` — **build a new table** at a *position* anchor (`heading:`/`para:`/`start`/`end`); reports its 1-based `index`. Fill cells row-major with `--data` (use `--data -` for stdin to dodge Windows quoting); `--style` defaults to `Table Grid` (visible borders); `--header` bolds row 1. New cells default to the `Normal` paragraph style regardless of the anchor, so a table dropped under a heading doesn't inherit that heading style. Edit existing tables with `table add-row`/`delete-row`/`set-heading-row` and `table delete N`; `table set-heading-row --table N [--row R]` marks a row as a repeating header across pages. `table append-record --table N --record '{"Header":"val",…}'` appends a row by header name (the dict-keyed companion to positional `add-row`); `table update-row --table N --key VALUE --values '{"Header":"new",…}' [--column HEADER]` updates the first row whose key-column equals `--key`, addressing a row by content not index. Cells are anchors (`table:N:R:C`) you can `replace`/`style apply`/`format-paragraph`.
- `wordlive comment add --anchor-id ID --text "…"` (+ `comment list` · `comment resolve --index N` · `comment delete --index N`).
- `wordlive track on|off|status` — toggle / inspect Track Changes. Read the recorded edits structurally with `wordlive revisions`, or see them with `snapshot --markup all`.
- `wordlive header write --section S --text "…"` · `footer write --section S --text "…"` (`--which primary|first|even`); `header read` / `footer read --section S [--which …]` read them back.
- `wordlive sections` — list sections with their page setup (orientation, margins, page size).

## Saving — gated, hand back a deliverable
- `wordlive save` · `wordlive save-as PATH [--overwrite]` (a `.docx`) · `wordlive export-pdf PATH [--pages A-B]` (a PDF — the recommended deliverable, a pixel-faithful render).
- **Default-deny:** saving only writes inside a directory whitelisted with the global `--save-dir DIR` (repeatable) or `WORDLIVE_SAVE_DIRS`. With none set, every save fails (exit 1, `path_not_allowed`). `save` needs the doc already saved once; `save-as` refuses to overwrite without `--overwrite`.

## Images
```
wordlive insert-image --anchor-id ID (--path FILE | --base64 VALUE) --wrap WRAP \
    [--before | --after] [--block] [--width N] [--height N] [--alt-text "…"] [--lock-aspect | --no-lock-aspect]
```
- Supply the image as a file (`--path`) **or** base64 (`--base64`; use `--base64 -`
  to read base64 from stdin). Base64 is ideal when you hold image data in memory.
  A `--path` is screened: non-local sources (UNC `\\host\…`, `file://`, URLs) are
  rejected, and `--image-dir`/`WORDLIVE_IMAGE_DIRS` can restrict it further —
  prefer base64 for images you're holding rather than a path.
- `--wrap` is **required**: `inline` (stays in the text flow), `auto` (floats —
  Square if small, else top-and-bottom), or
  `square|tight|through|top-bottom|behind|front`.
- `--block` puts the image on its own new line (a `Normal` paragraph) instead of
  in the anchor's text run — use with `--before` at a heading to land it above
  the heading rather than mid-line.

## Equations
```
wordlive insert-equation --anchor-id ID (--unicodemath "…" | --latex "…" | --mathml "…") \
    [--display | --inline] [--before | --after]
wordlive equations          # list equation:N ids, type, linear preview, para:N
```
- Exactly one input dialect: `--unicodemath` (Word's native linear form, e.g.
  `"x=(-b±√(b^2-4ac))/(2a)"`, zero-dependency), `--latex` (needs the `latex`
  extra), or `--mathml` (`--mathml -` reads from stdin; uses Office's own
  transform, no extra). As an `exec`/`insert_equation` op the fields are
  `{anchor_id, unicodemath|latex|mathml, [display, before]}`.
- The equation lands on its own paragraph with a pinned style (so it never
  inherits a neighbouring heading's style): `--display` gives it the centred
  `Equation` paragraph style (auto-created, based on `Normal`); `--inline` makes
  it `Normal` and left-aligned (still its own paragraph, not mid-sentence).
- `equation:N` is **positional** (Word's `OMaths` document order), not a durable
  handle: inserting another equation *before* an existing one renumbers it.
  Re-list with `equations` (or re-resolve right before you reference it) rather
  than caching an id across further inserts. The result carries the new id.

## Snapshot — render page(s) to PNG so you can *see* the layout
```
wordlive snapshot [--anchor-id ID | --page N | --pages A-B] [--out FILE] [--dpi 150] [--max-dim N] [--markup none|all]
```
Word exports a pixel-faithful PDF of the live document and wordlive rasterises
the requested pages — a true WYSIWYG image (real fonts, spacing, page geometry),
ideal for judging or iterating on style and formatting.
- Pick **at most one** target: `--anchor-id` (the page(s) the anchor occupies —
  a `heading:` expands to its **whole section**), `--page N`, or `--pages A-B`.
  With none, the whole document renders.
- `--max-dim N` caps each page's long edge to `N` pixels (only ever lowering
  resolution) — pair it with no page target to eyeball the **whole document's**
  layout cheaply. A vision model is billed on pixel area, so the cap is a
  predictable per-page token budget (~1000 stays legible); `--dpi 72` is a
  coarser alternative.
- `--markup all` renders tracked changes and comments as visible revision marks
  (default `none` shows the final document); the structured list is `revisions`.
- With `--out FILE` the image is written to disk (multiple pages become
  `<stem>-pN<suffix>`); **without `--out`, base64 PNG data is returned inline**
  in the JSON (`images[].base64`) — feed it straight back to yourself as an image.
- Needs the optional `snapshot` extra (PyMuPDF). If it's missing the command
  exits `1` with an install hint (`pip install "wordlive[snapshot]"`).
```bash
$ wordlive snapshot --anchor-id heading:3 --out section.png
{"ok": true, "selector": "heading:3", "dpi": 150, "count": 1, "images": [{"page": 4, "bytes": 81234, "path": "section.png"}]}
```

## Batch many edits atomically
Put several ops in a JSON file; they apply as a **single** undo step. This is the
best path for multi-step intents and for inline base64 images (no argv limits):
```json
{
  "label": "Update quarterly report",
  "ops": [
    {"op": "write_bookmark",    "name": "Address",   "text": "123 Main St"},
    {"op": "insert_paragraph",  "anchor_id": "heading:2", "text": "New section.", "style": "Body Text"},
    {"op": "insert_image",      "anchor_id": "heading:2", "base64": "<base64…>", "wrap": "auto"},
    {"op": "find_replace",      "find": "Q3", "text": "Q4", "all": true}
  ]
}
```
Run with `wordlive exec --script ops.json`, or pass the JSON inline with
`wordlive exec --ops '{"ops": [...]}'` (or `--ops -` to pipe it via stdin — best
for large payloads like base64 images, which can blow the command-line length
limit). Add `"tracked": true` at the top level to record the whole batch as
tracked changes. On failure it stops at the first bad op and reports `failure`
with the op `index`, `error`, and `type`. A field an op doesn't use (a typo, or
`style` on an inline append) is reported in a `warnings` array on the result
rather than silently dropped — so a clean-looking success can't hide a payload
you got wrong.

`insert_paragraph` and `insert_image` default to inserting **after** the anchor;
pass `"before": true` to insert above it (mirrors the CLI's `--before`/`--after`).

Each op resolves its `anchor_id` **fresh against the live document at the moment
it runs** — there's no pre-batch snapshot. So a positional id (`heading:N` /
`para:N`) sees the shifts earlier ops in the same batch made. When inserting
several paragraphs after one fixed anchor, either insert in reverse order or
anchor each to the previous insert; name-based ids (`bookmark:` / `cc:`) are
stable across the batch.

Ops (the full vocabulary — every CLI verb above has one): `write_bookmark`,
`write_cc`, `insert_paragraph`, `insert_block`, `insert_section`,
`insert_markdown`, `replace_section`, `delete_paragraph`, `append`,
`append_inline`, `prepend`, `prepend_inline`, `insert_image`, `insert_equation`, `replace`,
`find_replace`, `apply_style`, `format_paragraph`, `format_run`, `set_shading`,
`set_borders`, `drop_cap`, `add_tab_stop`, `add_style`, `set_style`, `insert_field`,
`set_page_setup`, `update_fields`, `insert_footnote`, `insert_endnote`,
`insert_toc`, `add_bookmark`, `add_hyperlink`, `insert_cross_reference`,
`insert_caption`, `set_cell`, `add_row`, `append_record`, `update_row`,
`delete_row`, `set_heading_row`, `create_table`, `delete_table`, `insert_break`,
`add_comment`, `resolve_comment`, `delete_comment`, `apply_list`, `remove_list`,
`restart_numbering`, `indent_list`, `outdent_list`, `write_header`,
`write_footer`.
(`append` / `prepend` add a new final / first **paragraph** and take `text` +
optional `style`, no anchor — `append_paragraph` / `prepend_paragraph` are
explicit synonyms. `append_inline` / `prepend_inline` instead **continue** the
last / first paragraph and take `text` only — no `style`. `insert_paragraph`
takes `text` **or** `runs` (`[{text,bold?,italic?,underline?,style?}]`) for
inline formatting; `insert_block` takes `anchor_id` + `items` (each a string or
`{text|runs, style?}`, item `text` carrying `**bold**`/`*italic*` markdown) and
returns the block's `range:START-END` in `outputs`. `insert_section` takes
`anchor_id` + `heading` + `body` (items shape) + optional `level`;
`insert_markdown` takes `anchor_id` + `markdown` (the constrained block subset);
`replace_section` takes a heading `anchor_id` + one of `body`/`markdown` and keeps
the heading — each returns the new content's `range:START-END`. `create_table` takes
`anchor_id` + optional `rows`/`cols` (inferred from `data` when omitted),
optional `style` / `data` / `header`; `data` is a row-major 2-D array **or**
records (objects whose keys become a header row). New cells default to the
`Normal` paragraph style regardless of the anchor, so they don't inherit a
heading style from the paragraph above. A successful batch returns an `outputs`
array reporting each new table's `index` and each block's `range`.
`insert_break` takes `anchor_id`, optional `kind` (default `page`) and `before`;
`format_paragraph`'s `page_break_before` bool is the reflow-safe alternative for
breaking before a styled paragraph, alongside the `keep_together` /
`keep_with_next` / `widow_control` pagination bools. `set_heading_row` takes
`table`, optional `row` (default 1) / `heading` / `allow_break` — a repeating
header row. `append_record` takes `table` + `record` (a `{header: value}` object
→ a new row); `update_row` takes `table` + `key` + `values` (`{header: value}`)
and optional `column` — it sets cells on the first row whose key-column equals
`key`.)

## Exit codes — branch on these
| Code | Meaning | Retry? |
| ---- | ------- | ------ |
| 0 | success | — |
| 1 | other / bad input (e.g. a missing/malformed image, or an unverifiable in-table `replace --find`) | fix the input (for a table match, scope with `--in table:N:R:C`) |
| 2 | anchor or style not found, or `find` matched zero | re-read with `outline` / `paragraphs`, then retry |
| 3 | Word busy (a modal dialog is open) | **yes** — back off and retry |
| 4 | Word not running | only after the user opens Word |
| 5 | ambiguous match (`replace --find` hit several) | re-run with `--occurrence N` or `--all` |

## Typical workflow
1. `wordlive status` → confirm Word and the target document.
2. `wordlive outline` / `paragraphs` / `find` → get the anchor ids you need.
3. Edit with the verbs above, or batch related changes with `exec`.
4. Read back to confirm. Edits are atomic and leave the user's cursor untouched.

For Python instead of the CLI, `import wordlive as wl` exposes the same model
(`wl.attach()`, `doc.edit("label")`, anchors with `.set_text()`,
`.insert_image()`, etc.) — see the `wordlive-python` skill (`wordlive llm-help
--python`). Full docs: https://thomas-villani.github.io/wordlive/
