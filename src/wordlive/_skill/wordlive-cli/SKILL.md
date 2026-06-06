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
- `wordlive read bookmark NAME` · `read cc NAME` · `read section "Heading Text"`
- `wordlive find --text "phrase"` — locate before editing; returns `range:` ids.
- `wordlive table list` · `wordlive table read N`

## Writing — each command is one atomic undo
- `wordlive write bookmark NAME --text "…"` · `write cc NAME --text "…"`
- `wordlive insert --anchor-id ID --text "…" [--before | --after] [--style "Body Text"]` — new paragraph relative to any anchor (`--after` is the default; appending after the document's last paragraph works too, so you can build a doc top-down).
- `wordlive append --text "…" [--inline] [--style "Body Text"]` — add a new final paragraph at the very end of the document (`--inline` continues the last paragraph). The high-level "end of doc" helper; same as `insert --anchor-id end`.
- `wordlive prepend --text "…" [--inline] [--style "Body Text"]` — the mirror: add to the very start of the document (same as `insert --anchor-id start`).
- `wordlive replace --anchor-id ID --text "…"` — overwrite a range.
- `wordlive replace --find "old" --text "new" [--all | --occurrence N] [--in ID]` — fuzzy find + replace. To edit **inside a table**, scope it to the cell (`--in table:N:R:C`); an unverifiable whole-document match fails (exit 1, `replace_verification`) rather than risk overwriting the wrong cell.
- `wordlive style apply --anchor-id ID --name "Heading 2"` (names: `style list`).
- `wordlive style add NAME [--type paragraph|character] [--based-on NAME] [--next-style NAME]` then `wordlive style set NAME [--bold] [--color "#1F3864"] [--size 16pt] [--alignment center] [--space-before 12] …` — define and configure a style, then `style apply` it everywhere (the brand/template workflow).
- `wordlive format-paragraph --anchor-id ID [--alignment center] [--left-indent 36] [--space-before 6] [--page-break-before] …` — `--page-break-before` is the clean, reflow-safe way to make a paragraph (e.g. a `Heading 1`) start a new page, leaving no stray break character.
- `wordlive format-run --anchor-id ID [--bold] [--italic] [--underline] [--font NAME] [--size 12pt] [--color "#FF0000"] [--highlight yellow] …` — **character** formatting; pair with a `range:` id (from `find`) to style a phrase. Colours take a name/hex/`r,g,b`; sizes take points or a unit string.
- `wordlive shading --anchor-id ID --fill "#FFF2CC"` · `wordlive borders --anchor-id ID [--sides all|top|…] [--style single|double|dot|dash] [--weight 0.5] [--color black]` · `wordlive tab-stop --anchor-id ID --position 3in [--align right] [--leader dots]` — range/cell fill, borders, and tab stops (a cell is an anchor, so these shade/border table cells too).
- `wordlive insert-break --anchor-id ID [--kind page|column|section_next|section_continuous] [--before | --after]` — an **explicit** one-off page/column/section break (the discoverable replacement for a literal form-feed paragraph). `--kind` defaults to `page`. Section breaks start a new section that can carry its own headers/footers + page setup. For a break that follows a *style*, prefer `format-paragraph --page-break-before`.
- `wordlive list apply --anchor-id ID --type bulleted|numbered|outline` (+ `list remove|restart|indent|outdent`).
- `wordlive table create --anchor-id ID --rows R --cols C [--style NAME] [--header] [--before | --after] [--data '[["…"],…]' | --data -]` — **build a new table** at a *position* anchor (`heading:`/`para:`/`start`/`end`); reports its 1-based `index`. Fill cells row-major with `--data` (use `--data -` for stdin to dodge Windows quoting); `--style` defaults to `Table Grid` (visible borders); `--header` bolds row 1. New cells default to the `Normal` paragraph style regardless of the anchor, so a table dropped under a heading doesn't inherit that heading style. Edit existing tables with `table add-row`/`delete-row` and `table delete N`; cells are anchors (`table:N:R:C`) you can `replace`/`style apply`/`format-paragraph`.
- `wordlive comment add --anchor-id ID --text "…"` (+ `comment list` · `comment resolve --index N` · `comment delete --index N`).
- `wordlive track on|off|status` — tracked changes.
- `wordlive header write --section S --text "…"` · `footer write --section S --text "…"` (`--which primary|first|even`).

## Images
```
wordlive insert-image --anchor-id ID (--path FILE | --base64 VALUE) --wrap WRAP \
    [--before | --after] [--block] [--width N] [--height N] [--alt-text "…"] [--lock-aspect | --no-lock-aspect]
```
- Supply the image as a file (`--path`) **or** base64 (`--base64`; use `--base64 -`
  to read base64 from stdin). Base64 is ideal when you hold image data in memory.
- `--wrap` is **required**: `inline` (stays in the text flow), `auto` (floats —
  Square if small, else top-and-bottom), or
  `square|tight|through|top-bottom|behind|front`.
- `--block` puts the image on its own new line (a `Normal` paragraph) instead of
  in the anchor's text run — use with `--before` at a heading to land it above
  the heading rather than mid-line.

## Snapshot — render page(s) to PNG so you can *see* the layout
```
wordlive snapshot [--anchor-id ID | --page N | --pages A-B] [--out FILE] [--dpi 150]
```
Word exports a pixel-faithful PDF of the live document and wordlive rasterises
the requested pages — a true WYSIWYG image (real fonts, spacing, page geometry),
ideal for judging or iterating on style and formatting.
- Pick **at most one** target: `--anchor-id` (the page(s) the anchor occupies —
  a `heading:` expands to its **whole section**), `--page N`, or `--pages A-B`.
  With none, the whole document renders.
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

Ops: `write_bookmark`, `write_cc`, `insert_paragraph`, `append`,
`append_inline`, `prepend`, `prepend_inline`, `insert_image`, `replace`,
`find_replace`, `apply_style`, `format_paragraph`, `set_cell`, `add_row`,
`delete_row`, `create_table`, `delete_table`, `insert_break`, `add_comment`, `resolve_comment`,
`delete_comment`, `apply_list`, `remove_list`, `restart_numbering`,
`indent_list`, `outdent_list`, `write_header`, `write_footer`.
(`append` / `prepend` add a new final / first **paragraph** and take `text` +
optional `style`, no anchor — `append_paragraph` / `prepend_paragraph` are
explicit synonyms. `append_inline` / `prepend_inline` instead **continue** the
last / first paragraph and take `text` only — no `style`. `create_table` takes
`anchor_id` + `rows` + `cols`, optional `style` / `data` (row-major 2-D) /
`header`; new cells default to the `Normal` paragraph style regardless of the
anchor, so they don't inherit a heading style from the paragraph above. A
successful batch returns an `outputs` array reporting each new table's `index`.
`insert_break` takes `anchor_id`, optional `kind` (default `page`) and `before`;
`format_paragraph`'s `page_break_before` bool is the reflow-safe alternative for
breaking before a styled paragraph.)

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
