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
| `pin:CODE`           | a **durable handle** minted by `pin` / `pin-outline` — survives renumbering |
| `cc:NAME`            | a content control (by title) |
| `footnote:N` / `endnote:N` | the Nth note's body (1-based; see `footnotes` / `endnotes`) |
| `image:N`            | the Nth embedded picture (1-based; see `images`) |
| `table:N:R:C`        | row R, col C of the Nth table (all 1-based) |
| `table:N:row:R` / `table:N:col:C` | a whole row / column of the Nth table — a styling handle (`shading` / `borders` / `apply-style` / `format-run`) |
| `range:START-END`    | a raw character span (what `find` emits) |
| `header:S:WHICH` / `footer:S:WHICH` | header/footer of section S (`primary` / `first` / `even`) |
| `start` / `end`      | the position before the first / past the last paragraph (prepend / append targets) |

`heading:N` and `para:N` are **positional** — they renumber when a structural
edit (a new paragraph, an inserted table) shifts the document, so re-read
`outline` / `paragraphs` after one before reusing ids downstream (an exit `2`
`anchor_not_found` is the signal you skipped that). `bookmark:NAME` and `cc:NAME`
are name-based and survive edits — reach for them when you need a durable handle.
For ad-hoc content with no name, `wordlive pin ANCHOR_ID` mints a `pin:CODE`
handle that Word keeps pinned to the same content across later inserts/deletes
(`wordlive pin-outline` pins every heading at once). When a stale `para:N` /
`heading:N` misses, the exit-2 error carries a recovery hint (out-of-range vs
not-a-heading, the nearest heading, and a nudge to pin).

## Reading
- `wordlive read bookmark NAME` (or `read bookmark --list` for every bookmark name) · `read cc NAME` · `read section "Heading Text"`
- `wordlive read markdown [--within ID]` (or `read html`) — serialise the document (or one anchor's range) to clean Markdown / HTML: headings, lists, `**bold**`/`*italic*`, GFM tables, `![alt](image:N)`, `[text](url)`. The read mirror of `insert-markdown`; lossy by design. `--text` pipes the raw Markdown out (e.g. `wordlive --text read markdown > doc.md`).
- `wordlive read digest [--budget 6000] [--depth D]` — a token-budgeted whole-document read: headings verbatim (each tagged with its `heading:N` anchor), tables as one-line stubs, body sampled to fit the budget. Loads a large doc into context cheaply while every anchor stays addressable; drill into an elided region with `read markdown --within …`.
- `wordlive read text --anchor-id ID [--view raw|final|original|segments]` — an anchor's text; `final`/`original` resolve tracked changes (final = as if accepted, original = as if rejected — the deleted wording, which `raw`/`Range.Text` omits), `segments` is the per-run insert/delete breakdown.
- `wordlive read between --start ID --end ID [--inclusive]` — content between two anchors (e.g. the block between two headings; default excludes both heading lines). `wordlive read nearest-heading --anchor-id ID [--direction before|after]` — the heading nearest a position (`before` = enclosing/preceding, `after` = next).
- `wordlive find --text "phrase"` — exact locate before editing; returns `range:` ids.
- `wordlive find-paragraph --text "approx text" [--limit N] [--min-score F]` — **fuzzy** paragraph search (typo/paraphrase tolerant), returns ranked `para:N` candidates with scores. Use this when you remember the gist; use `find` for exact substrings.
- `wordlive table list` · `wordlive table read N` · `wordlive table records N` — body rows as `{header: value}` dicts (row 1 is the header), the read mirror of building a table from records.
- `wordlive footnotes` · `wordlive endnotes` — each note's `footnote:N`/`endnote:N` id, text, and `para:N`.
- `wordlive images` — each embedded picture's `image:N` id, MIME, size, alt text, and `para:N`. Pull one out with `wordlive read-image --anchor-id image:N [--out FILE]` (`--out` writes the raw bytes; otherwise base64 + mime inline) — the path for handing a picture to a vision model.
- `wordlive revisions` — tracked changes as structured data (`type`/`author`/`text`/`range`); the readable counterpart to `snapshot --markup all`. `wordlive track status` reports whether Track Changes is on. Resolve them with `wordlive revision accept|reject --index N` or `revision accept-all|reject-all [--anchor-id ID]` (scope to one anchor's range).
- `wordlive hyperlinks` — every link's text, external `address`/internal `sub_address`, and `range:`/`para:` (the read mirror of `link`). `wordlive fields` — every field's `kind` (`PAGE`/`REF`/`TOC`/…), raw `code`, and rendered `result` (the read mirror of `insert-field`; run `update-fields` first to refresh).
- `wordlive properties list` — built-in + custom document properties (metadata). `wordlive variables list` — invisible `{ DOCVARIABLE }` storage as `{name: value}`.
- `wordlive proofing` — spelling/grammar errors (count + flagged runs with `range:` ids) and readability statistics (Flesch Reading Ease, Flesch-Kincaid Grade Level, …). Heavier than `stats` (it (re)checks the doc).
- `wordlive lint [--rule ID/tag] [--exclude ID/tag] [--within ID]` — audit publishing-quality defects (dangling headings, multi-page tables with no repeating header, numbered lists Word split into independent runs, direct formatting drifted from the applied style). Severity-ranked findings, each `fixable` one carrying the op `regularize` would run. Then `wordlive regularize [--dry-run] [--rule …] [--within ID]` applies the fixable subset in one atomic-undo step (targeted + idempotent — a second pass is a no-op). `read format --anchor-id ID` is the underlying probe: effective paragraph/character formatting with each field's style baseline and an `override` flag (the read mirror of `format-paragraph`/`format-run`).
- `wordlive checkpoint [--include text|text+style|text+format] [--within ID] [--out FILE]` — fingerprint the document's structure now → a checkpoint token (pure read; `--out` writes it to a file). Then `wordlive diff --since FILE` (token vs the document now) or `wordlive diff --from A --to B` (two stored tokens) → a content-aligned change list: each change is `replace`/`insert`/`delete`/`restyle`/`reformat` (plus `table_change`/`table_insert`/`table_delete` for tables) carrying the current `para:N`/`table:N`. This is how you answer "what changed" (Word fires no content-change event) and verify your own edits landed; an unchanged doc returns `[]`. `--include text+format` also catches a direct-formatting change as a `reformat`. (`--since` needs a checkpoint scoped to a stable anchor, not a `range:`.)
- `wordlive locate --anchor-id ID` — where an anchor sits in the laid-out doc: `page`/`end_page` span, `line`, `column`, `in_table`. Answers "what page is this on" without a `snapshot`. `wordlive stats` — one-shot summary: page/word/char/paragraph/line counts plus section/heading/table/image/comment/revision counts and `saved`. Both repaginate first (selection/scroll untouched); reads, non-mutating.

## Writing — each command is one atomic undo
- `wordlive write bookmark NAME --text "…"` (set existing) · `write bookmark NAME --create --anchor-id ID` (create a new bookmark over a range — the prerequisite for internal links and cross-refs; NAME starts with a letter, letters/digits/underscores only) · `write cc NAME --text "…"`
- `wordlive pin ANCHOR_ID [--name SLUG]` — plant a durable `pin:CODE` handle on any anchor (random code, or a readable `--name budget-intro`); resolve it later with `--anchor-id pin:CODE`. `wordlive pin-outline [--levels LO HI]` — pin every heading at once, printing the `{heading:N → pin:CODE}` map (idempotent). Durable handles are the fix for positional ids that renumber under edits.
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
- `wordlive table-of-figures [--anchor-id start] [--label Figure] [--no-label] [--no-hyperlinks]` — a table of figures from the captions of one `--label` (`Figure`/`Table`/`Equation`/custom). Page numbers fill in after `update-fields`.
- **Back-of-book index** (two steps): `wordlive mark-index-entry --anchor-id ID --entry "topic"` (use `"main:sub"` to nest; `--cross-reference X` for a "see X") marks the entries, then `wordlive insert-index [--anchor-id end] [--columns 2] [--run-in]` builds the index from them. Page numbers fill in after `update-fields`.
- **Citations & bibliography** (source → cite → build): pick the scheme once with `wordlive bibliography-style --style APA` (also `MLA`/`Chicago`/`IEEE`/`Turabian`, build-dependent), register a source with `wordlive add-source --type book --author "Smith, Jane" --title "…" --year 2020` (`--type` is `book`/`journal_article`/`web_site`/`case`/…; `--author` is `"Last, First"`, repeatable; `--tag` auto-derives from surname+year if omitted; `--xml RAW` is the OOXML escape hatch) — it reports the source `tag`. Then cite it where you want with `wordlive insert-citation --anchor-id ID --tag Smith2020 [--pages 15] [--suppress-author] …` (an unknown tag renders "Invalid source specified."), and build the works-cited block with `wordlive insert-bibliography [--anchor-id end]`. Page numbers/entries fill in after `update-fields` (or `snapshot`).
- **Table of authorities** (legal-citation index; two steps like the back-of-book index): `wordlive mark-citation --anchor-id ID --long "Brown v. Board, 347 U.S. 483 (1954)" [--short Brown] [--category cases]` marks each entry (`--category` is `cases`/`statutes`/`other`/`rules`/`treatises`/`regulations`/`constitutional`, or an int 1-16; `--short` defaults to `--long`), then `wordlive table-of-authorities [--anchor-id end] [--category all|cases|…] [--no-passim] [--no-keep-formatting]` builds the table from those marks. Page numbers fill in after `update-fields`.
- **Document theme / branding** (the document-wide brand primitive): apply a whole theme with `wordlive apply-theme --theme Facet` (a built-in name or a `.thmx` path — see `wordlive list-themes` for the names), then brand it with `wordlive set-theme-colors [--scheme Blue] [--accent1 "#1A73E8"] [--text1 navy] …` (`--scheme` loads a named colour scheme; the per-colour flags — `--text1`/`--background1`/`--text2`/`--background2`/`--accent1`..`--accent6`/`--hyperlink`/`--followed-hyperlink` — take a name or hex and override individuals) and `wordlive set-theme-fonts [--scheme Garamond] [--major Arial] [--minor Calibri]` (heading/body fonts). Read the current theme with `wordlive theme` and the built-in library with `wordlive list-themes`.
- `wordlive link --anchor-id ID (--url URL | --bookmark NAME) [--text "…"]` — turn an anchor into a hyperlink: external `--url` or internal `--bookmark`. `--text` inserts new linked text (no `--text` links the existing range).
- `wordlive cross-ref --anchor-id ID --target TARGET [--kind text|page|number|above_below] [--no-hyperlink]` — reference another anchor; `--target` is `bookmark:NAME` / `heading:N` / `footnote:N` / `endnote:N`. Refresh with `update-fields`.
- `wordlive caption --anchor-id ID [--label Figure] [--text "…"] [--position above|below]` — insert an auto-numbered caption (Figure/Table/…) as its own `Caption` paragraph (on a `table:N:R:C` anchor it captions the whole table). Default placement is above for a `Table`, below otherwise; `--position` overrides. Pairs with `cross-ref`.
- `wordlive page-setup [--section N] [--margins 1in] [--top-margin … --bottom-margin … --left-margin … --right-margin …] [--gutter …] [--orientation portrait|landscape] [--paper-size letter|legal|tabloid|a3|a4|a5] [--columns N] [--column-spacing …]` — set a section's page geometry. `--margins` sets all four (per-side flags override); lengths take points or a unit string; `--columns N` makes N equal columns. `--section` defaults to `1`.
- `wordlive list apply --anchor-id ID --type bulleted|numbered|outline` (+ `list remove|restart|indent|outdent`; `list show` enumerates every list in the doc, `list info --anchor-id ID` reports a paragraph's list type/level/number). **To number N paragraphs 1..N, scope one apply to a `range:` that spans them all** (offsets from `paragraphs`/`find`) — applying `numbered` per-paragraph makes N separate "1." lists.
- `wordlive table create --anchor-id ID --rows R --cols C [--style NAME] [--header] [--before | --after] [--data '[["…"],…]' | --data -]` — **build a new table** at a *position* anchor (`heading:`/`para:`/`start`/`end`); reports its 1-based `index`. Fill cells row-major with `--data` (use `--data -` for stdin to dodge Windows quoting); `--style` defaults to `Table Grid` (visible borders); `--header` bolds row 1. New cells default to the `Normal` paragraph style regardless of the anchor, so a table dropped under a heading doesn't inherit that heading style. Edit existing tables with `table add-row`/`delete-row`/`add-column`/`delete-column`/`merge-cells`/`split-cell`/`set-heading-row` and `table delete N`; `table set-heading-row --table N [--row R]` marks a row as a repeating header across pages. `table add-column --table N [--values '["…",…]']` / `table delete-column --table N --column C` mirror the row ops; `table merge-cells --table N --from R:C --to R:C` / `table split-cell --table N --cell R:C [--rows R --cols C]` reshape cells (which makes the table non-uniform, so `table:N:R:C` then indexes *physical* cells). `table append-record --table N --record '{"Header":"val",…}'` appends a row by header name (the dict-keyed companion to positional `add-row`); `table update-row --table N --key VALUE --values '{"Header":"new",…}' [--column HEADER]` updates the first row whose key-column equals `--key`, addressing a row by content not index. Cells are anchors (`table:N:R:C`) you can `replace`/`style apply`/`format-paragraph`. `table autofit --table N [--mode content|window|fixed]` resizes columns — `content` fits cells, `window` stretches to the page, `fixed` pins current widths. Restyle a whole table with `table set-style`/`set-alignment`/`set-borders`/`set-banding`, and set a cell's vertical alignment with `cell-valign --anchor-id table:N:R:C --align top|center|bottom`. A whole **row** (`table:N:row:R`) or **column** (`table:N:col:C`) is itself an anchor, so `shading`/`borders`/`style apply`/`format-run` style the strip in one call (a column op raises on a merged / mixed-width table — style those cells via `table:N:R:C`).
- `wordlive properties set --name Title --value "…"` (built-in metadata) · `properties set --name Project --value Apollo --custom` (custom property, created if absent) · `properties delete --name Project` — read with `properties list`.
- `wordlive variables set --name ClientName --value "Acme"` · `variables delete --name ClientName` — create/update/remove a `{ DOCVARIABLE }` variable; read with `variables list`.
- `wordlive comment add --anchor-id ID --text "…"` (+ `comment list` · `comment resolve --index N` · `comment delete --index N`).
- `wordlive track on|off|status` — toggle / inspect Track Changes. Read the recorded edits structurally with `wordlive revisions`, or see them with `snapshot --markup all`.
- `wordlive watermark --text "DRAFT" [--layout diagonal|horizontal] [--color C] [--solid]` (or `--remove`) — a text watermark behind every page (replaces a prior one, doesn't stack). `wordlive insert-text-box --anchor-id ID --text "…" [--width 2.5in] [--height H] [--wrap square|…] [--fill C] [--no-border]` — a floating pull quote anchored to the paragraph.
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
- A **floating** wrap (anything but `inline`) leaves the text flow, so the image
  is no longer an `image:N` — `insert-image` then reports a **`shape:N`** handle
  (an `inline` image stays `image:N`). Use that `shape:N` to re-wrap / reposition
  / resize / replace the floating image (see **Floating shapes**).

## Floating shapes (`shape:N`) — text boxes, floating images, WordArt
```
wordlive shapes                       # list shape:N ids, kind, size, wrap, para:N
wordlive set-shape-wrap   --anchor-id shape:N [--wrap square|tight|through|top-bottom|front|behind] [--side both|left|right|largest] [--distance-top D] [--distance-bottom D] [--distance-left D] [--distance-right D]
wordlive set-shape-crop   --anchor-id shape:N [--left L] [--top T] [--right R] [--bottom B]   # picture shapes only
wordlive set-shape-position --anchor-id shape:N [--left L] [--top T] [--relative-to margin|page]
wordlive set-shape-size   --anchor-id shape:N [--width W] [--height H] [--lock-aspect|--no-lock-aspect]
wordlive format-shape     --anchor-id shape:N [--fill C] [--border-color C | --no-border | --default-border] [--border-weight W]
wordlive set-shape-alt-text --anchor-id shape:N --text "…"
wordlive set-shape-text   --anchor-id shape:N --text "…"     # text boxes only
wordlive set-shape-rotation --anchor-id shape:N --degrees 30
wordlive set-shape-z-order --anchor-id shape:N --order front|back|forward|backward
wordlive set-shape-text-frame --anchor-id shape:N [--margin-left L] [--margin-right R] [--margin-top T] [--margin-bottom B] [--word-wrap|--no-word-wrap]   # text boxes only
wordlive replace-shape-image --anchor-id shape:N (--path FILE | --base64 VALUE)   # picture shapes only
wordlive group-shapes     --anchor-id shape:N --anchor-id shape:M [...]   # two or more → one group
wordlive ungroup-shape    --anchor-id shape:N                            # group shapes only
wordlive delete-shape     --anchor-id shape:N
# inline pictures (image:N) — alt text / resize without floating them:
wordlive set-image-alt-text --anchor-id image:N --text "…"
wordlive set-image-size   --anchor-id image:N [--width W] [--height H] [--lock-aspect|--no-lock-aspect]
wordlive set-image-crop   --anchor-id image:N [--left L] [--top T] [--right R] [--bottom B]
```
- `shape:N` addresses the document's **floating** shapes (a text box from
  `insert-text-box`, a floating image from `insert-image`, WordArt) in document
  order — header-story watermarks are excluded. It's **positional**: inserting a
  shape earlier renumbers it, so re-list with `shapes` rather than caching an id.
  `--left`/`--top` are lengths (`2in`) or `center`.
- `replace-shape-image` swaps a floating picture's image **in place** (delete +
  reinsert at the same anchor), preserving wrap / position / size / alt text.
- `group-shapes` collapses two or more floating shapes into one group `shape:N`
  (move / size / delete as a unit); `ungroup-shape` dissolves it back, returning
  the members' `shape:N` ids. `set-shape-z-order` restacks within the float layer
  (distinct from wrap's front/behind-text) — and because `Shapes` orders by
  z-order, it **renumbers `shape:N`**, so re-run `shapes` after one.
  `set-shape-rotation` is an absolute angle. There is no autosize knob — Word's
  "resize-to-fit-text" doesn't expose cleanly over COM.
- `set-image-alt-text` / `set-image-size` restyle an **inline** picture (`image:N`)
  without floating it. *Re-wrapping* an image (floating it) is `insert-image
  --wrap`, which converts it to a `shape:N`.
- `textbox:N` is an addressing alias onto a text box's canonical `shape:N` (so
  `anchor-id textbox:1` ≡ the first text box's `shape:M`).
- These are the post-insert restyle handles `insert-text-box` and a floating
  `insert-image` hand back — the exec ops are `set_shape_wrap`, `set_shape_crop`,
  `set_shape_position`, `set_shape_size`, `format_shape`, `set_shape_alt_text`,
  `set_shape_text`, `set_shape_rotation`, `set_shape_z_order`,
  `set_shape_text_frame`, `replace_shape_image`, `delete_shape`, `group_shapes`,
  `ungroup_shape`, `set_image_alt_text`, `set_image_size`, `set_image_crop`.

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

## Charts (Excel-backed)
```
wordlive insert-chart --anchor-id ID --kind bar|pie|line|scatter --data JSON \
    [--title "…"] [--before | --after]
wordlive charts             # list chart:N ids, kind, title, chart_style, has_legend, para:N
# Formatting & design — operate on an existing chart:N (no Excel needed):
wordlive format-chart --anchor-id chart:N [--title …] [--legend|--no-legend] \
    [--legend-position right|left|top|bottom|corner] [--chart-style INT] \
    [--background COLOR] [--plot-background COLOR] [--font …] [--font-size …] \
    [--font-color COLOR] [--data-labels|--no-data-labels] [--data-label-format …] \
    [--chart-type bar|pie|line|scatter]
wordlive format-axis --anchor-id chart:N --which value|y|category|x \
    [--title …] [--minimum N] [--maximum N] [--scale linear|log] \
    [--number-format …] [--gridlines|--no-gridlines]
wordlive add-trendline --anchor-id chart:N [--series N] \
    [--kind linear|exponential|logarithmic|moving_average|polynomial|power] \
    [--display-equation] [--display-r-squared] [--forward N] [--backward N]
wordlive set-series-color --anchor-id chart:N --color COLOR [--series N] [--point N]
```
- `--data` is JSON (or `--data -` to read from stdin): an object
  `{"Q1": 10, "Q2": 25}` for `bar`/`pie`/`line`, or an array of `[x, y]` pairs
  `[[1.2, 3.4], [2.5, 6.1]]` for `scatter` (both axes numeric; duplicate x kept)
  — `line` accepts either. As an `exec`/`insert_chart` op the fields are
  `{anchor_id, kind, data, [title, before]}`.
- Charts are **Excel-backed**: the data lives in a hidden Excel workbook, so
  **Excel must be installed** — if it isn't, the command exits **`6`**
  (`excel_not_available`) and the document is left untouched. wordlive then breaks
  the data link, so the chart's data is **static** (no embedded workbook ships in
  the doc, and the series data isn't read back — `charts` is metadata only).
- `chart:N` is **positional** (document order); inserting another chart earlier
  renumbers it. The result carries the new id.
- **Formatting & design** (`format-chart` / `format-axis` / `add-trendline` /
  `set-series-color` / `format-series` / `add-error-bars`, or the `format_chart` /
  `format_axis` / `add_trendline` / `set_series_color` / `format_series` /
  `add_error_bars` ops) work on the **static** post-insert chart — **no Excel
  needed**. All fields are tri-state (only what you pass is written); colours are a
  name / hex / `r,g,b`. A `power`/`exponential` trendline with `--display-equation`
  draws the law of best fit; `--scale log` suits order-of-magnitude data;
  `polynomial` takes `--order`, `moving_average` takes `--period`. `format-series`
  styles markers (`--marker circle`, `--marker-size`), line `--smooth`, pie
  `--explosion`, and per-point data-label fonts; `add-error-bars` draws
  `fixed`/`percent`/`stdev`/`sterror` bars. `format-chart` also tunes bar
  `--gap-width`/`--overlap` and toggles the `--data-table`.

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
`append_inline`, `prepend`, `prepend_inline`, `insert_image`, `insert_equation`,
`insert_chart`, `format_chart`, `format_axis`, `add_trendline`, `set_series_color`,
`format_series`, `add_error_bars`,
`set_shape_wrap`, `set_shape_crop`, `set_shape_position`, `set_shape_size`,
`format_shape`, `set_shape_alt_text`, `set_shape_text`, `set_shape_rotation`,
`set_shape_z_order`, `set_shape_text_frame`, `replace_shape_image`,
`delete_shape`, `group_shapes`, `ungroup_shape`, `set_image_alt_text`,
`set_image_size`, `set_image_crop`,
`replace`,
`find_replace`, `apply_style`, `format_paragraph`, `format_run`, `set_shading`,
`set_borders`, `drop_cap`, `add_tab_stop`, `add_style`, `set_style`, `insert_field`,
`set_page_setup`, `update_fields`, `regularize`, `insert_footnote`, `insert_endnote`,
`insert_toc`, `add_bookmark`, `pin`, `pin_outline`, `add_hyperlink`, `set_hyperlink`,
`insert_cross_reference`,
`insert_caption`, `create_content_control`, `set_cc_properties`, `set_cc_items`,
`mark_index_entry`, `insert_index`,
`insert_table_of_figures`, `set_bibliography_style`, `add_source`, `insert_citation`,
`insert_bibliography`, `mark_citation`, `insert_table_of_authorities`,
`apply_theme`, `set_theme_colors`, `set_theme_fonts`,
`set_cell`, `add_row`, `append_record`, `update_row`,
`delete_row`, `add_column`, `delete_column`, `merge_cells`, `split_cell`,
`set_heading_row`, `autofit_table`, `create_table`,
`delete_table`, `set_table_style`, `set_table_alignment`, `set_table_borders`,
`set_table_banding`, `set_cell_vertical_alignment`,
`set_property`, `delete_property`, `set_variable`,
`delete_variable`, `insert_break`,
`add_comment`, `resolve_comment`, `delete_comment`, `accept_revision`,
`reject_revision`, `accept_all_revisions`, `reject_all_revisions`,
`set_watermark`, `remove_watermark`, `insert_text_box`, `apply_list`,
`apply_list_format`, `remove_list`, `restart_numbering`, `indent_list`,
`outdent_list`, `write_header`, `write_footer`.
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
`key`. `autofit_table` takes `table` + optional `mode` (`content` default /
`window` / `fixed`). `set_table_style` takes `table` + `style` — restyle an
existing table (restyle **first**, then layer cell-level shading, which a style
reapply overwrites). `set_table_alignment` takes `table` + `alignment`
(`left`/`center`/`right`) — the whole table across the page. `set_table_borders`
takes `table` + optional `sides`/`style`/`line_style`/`weight`/`color` — the whole
grid in one call (interior gridlines via `horizontal`/`vertical`).
`set_table_banding` takes `table` + optional `first_row`/`last_row`/
`first_column`/`last_column`/`banded_rows`/`banded_columns` bools — toggle the
table-style options (needs a real table style applied to show).
`set_cell_vertical_alignment` takes `anchor_id` (a `table:N:R:C` cell) + `align`
(`top`/`center`/`bottom`). **Row / column styling:** address a whole row as
`table:N:row:R` or a whole column as `table:N:col:C` — these are anchors, so the
`shading` / `borders` / `apply-style` / `format-run` verbs (and `set_shading` /
`set_borders` ops) style the whole strip in one call. A column op on a table with
merged / mixed-width cells has no per-column model in Word and raises an error —
style those cells individually via `table:N:R:C`. `set_property` takes `name` +
`value` + optional `custom`
(a custom property when true, else a built-in like Title/Author); `delete_property`
takes `name` (custom only). `set_variable` takes `name` + `value`;
`delete_variable` takes `name`. `pin` takes `anchor_id` + optional `name` (slug)
and returns its `pin:CODE`; `pin_outline` takes optional `levels` and returns the
`{heading:N: pin:CODE}` map.)

**Durable handles in a batch.** Add `bind: "slug"` (or `true` for a random code)
to `insert` / `insert_block` / `insert_section` / `insert_markdown` / `create_table`
to mint a `pin:` handle on the new content — it comes back in that op's `outputs`
entry as `pin`. And any op field of the exact form `$ops[N].field` is replaced with
an earlier op's recorded output before the op runs, so a batch can create then
target without a round-trip: `[{"op":"create_table",...},{"op":"set_cell","table":"$ops[0].table",...}]`.

## Exit codes — branch on these
| Code | Meaning | Retry? |
| ---- | ------- | ------ |
| 0 | success | — |
| 1 | other / bad input (e.g. a missing/malformed image, or an unverifiable in-table `replace --find`) | fix the input (for a table match, scope with `--in table:N:R:C`) |
| 2 | anchor or style not found, or `find` matched zero | re-read with `outline` / `paragraphs`, then retry |
| 3 | Word busy (a modal dialog is open) | **yes** — back off and retry |
| 4 | Word not running | only after the user opens Word |
| 5 | ambiguous match (`replace --find` hit several) | re-run with `--occurrence N` or `--all` |
| 6 | Excel not installed (only `insert-chart` — charts are Excel-backed) | only after installing Excel |

## Typical workflow
1. `wordlive status` → confirm Word and the target document.
2. `wordlive outline` / `paragraphs` / `find` → get the anchor ids you need.
3. Edit with the verbs above, or batch related changes with `exec`.
4. Read back to confirm. Edits are atomic and leave the user's cursor untouched.

For Python instead of the CLI, `import wordlive as wl` exposes the same model
(`wl.attach()`, `doc.edit("label")`, anchors with `.set_text()`,
`.insert_image()`, etc.) — see the `wordlive-python` skill (`wordlive llm-help
--python`). Full docs: https://thomas-villani.github.io/wordlive/
