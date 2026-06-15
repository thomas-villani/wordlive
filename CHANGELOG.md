# Changelog

All notable changes to **wordlive** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.16.0] — 2026-06-15

### Added
- **Content-control creation — `anchor.insert_content_control(...)`.** Closes the
  read/write-but-couldn't-*create* gap: wordlive could read (`read_cc`) and write
  (`write_cc`) an existing content control, but not make one. `anchor.insert_content_control(
  kind="rich_text", title=…, tag=…, items=…, where="wrap", lock_contents=…,
  lock_control=…)` wraps the anchor's existing range in a new control (or inserts a
  fresh empty one with `where="before"`/`"after"`) and returns the
  `ContentControl`. `kind` is `rich_text` (default) / `text` / `picture` /
  `combo_box` / `dropdown` / `date` / `checkbox` / `building_block` / `group` /
  `repeating_section`; `items` populates a combo_box/dropdown (strings or
  `{text, value}`); a `title` (falling back to `tag`) names it, so it's addressable
  later as `cc:TITLE`. `doc.content_controls.add(anchor, kind=…, **kwargs)` takes an
  `Anchor` or an anchor-id string. Across the `create_content_control` exec op, the
  CLI (`create-content-control`), and MCP (`word_write command="create_content_control"`).
  New `WdContentControlType` constant in `wordlive.constants`.
- **Back-of-book index — `mark_index_entry` + `insert_index`.** Two steps, like
  Word's own: `anchor.mark_index_entry(entry, cross_reference=…, bold=…, italic=…)`
  marks the anchor's range as an `XE` index field (`entry` uses `"main:sub"` for a
  subentry), then `anchor.insert_index(columns=2, run_in=…, right_align_page_numbers=…,
  where="after")` builds the index from those marks and returns a new
  `Index` — a field block like the TOC, so `index.update()` repopulates it and
  page numbers fill only after repagination (`update_fields` / `snapshot`).
  `doc.add_index(...)` is the sugar for one at the document end. Across the
  `mark_index_entry` / `insert_index` exec ops, the CLI (`mark-index-entry` /
  `insert-index`), and MCP (`word_write command="mark_index_entry"` /
  `command="insert_index"`). New public `Index` class and `WdIndexType` constant.
- **Table of figures — `anchor.insert_table_of_figures(...)`.** Consumes the
  captions wordlive already ships: `anchor.insert_table_of_figures(label="Figure",
  include_label=True, hyperlinks=True, right_align_page_numbers=True, where="after")`
  lists every caption of one `label` (`Figure`/`Table`/`Equation`/custom) with page
  numbers, and returns a `TableOfFigures`. It's a field block reusing the TOC
  pattern — `.update()` / `.update_page_numbers()`. Across the
  `insert_table_of_figures` exec op, the CLI (`table-of-figures`), and MCP
  (`word_write command="insert_table_of_figures"`). New public
  `TableOfFigures` class.
- **Citations & bibliography — `doc.sources` + `anchor.insert_citation` +
  `anchor.insert_bibliography`.** The academic-writing workflow end to end.
  `doc.sources.add("book", author="Smith, John", title=…, year=2020, …)` registers
  a source in the document's store (a friendly typed API over Word's `<b:Source>`
  XML — `book` / `journal_article` / `conference_proceedings` / `report` /
  `web_site` / `case` / …; `author` is `"Last, First"` or a list; `tag`
  auto-derives from author + year), with `doc.sources.add_xml(...)` as the raw
  escape hatch and the collection subscriptable/iterable by tag.
  `anchor.insert_citation(tag, pages=…, prefix=…, suffix=…, volume=…,
  suppress_author=…, suppress_year=…, suppress_title=…, locale=1033)` inserts an
  in-text citation (returns a `Citation`); `anchor.insert_bibliography()` /
  `doc.add_bibliography()` inserts the reference list of cited sources (returns a
  `Bibliography`). `doc.bibliography_style` (read/write — APA/MLA/Chicago/IEEE/…)
  sets the rendering style. Across the `set_bibliography_style` / `add_source` /
  `insert_citation` / `insert_bibliography` exec ops, the CLI (`bibliography-style`
  / `add-source` / `insert-citation` / `insert-bibliography`), and MCP. New public
  `Source`, `Citation`, and `Bibliography` classes.
- **Table of authorities — `mark_citation` + `insert_table_of_authorities`.** The
  legal mark-then-build workflow, mirroring the index: `anchor.mark_citation(
  long_citation, short_citation=…, category="cases")` marks the anchor's range as
  a `TA` field (`category` is `cases`/`statutes`/`other`/`rules`/`treatises`/
  `regulations`/`constitutional`, or `1`-`16`), then
  `anchor.insert_table_of_authorities(category="all", passim=…,
  keep_entry_formatting=…, entry_separator=…, page_range_separator=…)` builds the
  table from those marks and returns a `TableOfAuthorities` (a field block;
  `.update()`). `doc.add_table_of_authorities(...)` is the sugar for one at the
  document end. Across the `mark_citation` / `insert_table_of_authorities` exec
  ops, the CLI (`mark-citation` / `table-of-authorities`), and MCP. New public
  `TableOfAuthorities` class; new `CITATION`/`BIBLIOGRAPHY`/`TOA`/`TOA_ENTRY`
  members on `WdFieldType`.
- **Document themes — `doc.theme`.** The document-wide brand primitive for
  producing themed / branded documents. `doc.theme.apply("Facet")` applies a whole
  theme (colours + fonts + effects) by built-in name or `.thmx` path;
  `doc.theme.set_colors(scheme="Blue", accent1="#1A73E8", text1="navy", …)` sets the
  colour scheme and/or overrides individual brand colours (keys `text1` /
  `background1` / `text2` / `background2` / `accent1`–`accent6` / `hyperlink` /
  `followed_hyperlink`; values take a colour name, hex, or `(r, g, b)`);
  `doc.theme.set_fonts(scheme="Garamond", major="Arial", minor="Calibri")` sets the
  heading/body fonts. `doc.theme.colors` / `.major_font` / `.minor_font` /
  `.to_dict()` read the current theme, and `doc.theme.list_available()` lists the
  built-in themes, colour schemes, and font schemes Office ships. Across the
  `apply_theme` / `set_theme_colors` / `set_theme_fonts` exec ops, the CLI
  (`theme` / `list-themes` / `apply-theme` / `set-theme-colors` / `set-theme-fonts`),
  and MCP (`word_read command="theme"` / `"themes"`; `word_write` apply/set
  commands). New public `DocumentTheme` class and a `bgr_to_hex` colour helper.

## [0.15.0] — 2026-06-13

### Added
- **Document metadata — `doc.properties`.** Read and write the file's built-in
  properties (Title, Author, Subject, Keywords, Comments, Category, Manager,
  Company, …) and free-form custom properties. `doc.properties.read()` returns
  `{builtin, custom}`; `set(name, value)` writes a built-in, `set(name, value,
  custom=True)` a custom one (created if absent); `delete(name)` removes a custom
  one. Across the Python API, the `set_property` / `delete_property` exec ops, the
  CLI (`properties list|set|delete`), and MCP (`word_read command=properties`,
  `word_write command=set_property|delete_property`).
- **Document variables — `doc.variables`.** Invisible named string storage (the
  backing store for `{ DOCVARIABLE }` fields). `doc.variables.list()` returns
  `{name: value}`; `set` / `get` / `delete` manage them. Across the Python API,
  the `set_variable` / `delete_variable` exec ops, the CLI (`variables
  list|set|delete`), and MCP.
- **`doc.hyperlinks` — the read mirror of `link_to`/`add_hyperlink`.** A
  read-only, indexable collection reporting each link's visible text, external
  `address` or internal `sub_address` bookmark, screen tip, and a
  `range:START-END` / `para:N`. CLI `hyperlinks`, MCP `word_read
  command=hyperlinks`.
- **`doc.fields` — the read mirror of `insert_field`.** A read-only collection
  reporting each field's `kind` (the code's leading keyword — PAGE/REF/TOC/…),
  raw `code`, rendered `result`, `locked`, and a `range:START-END` / `para:N`.
  CLI `fields`, MCP `word_read command=fields`.
- **`doc.proofing()` — spelling, grammar, and readability.** Runs Word's proofing
  tools and returns `{spelling, grammar, readability}`: spelling/grammar give a
  count plus a (capped) list of flagged runs with `range:START-END` ids, and
  readability gives Flesch Reading Ease, Flesch-Kincaid Grade Level,
  passive-sentence %, and averages. A heavier read than `stats` (it (re)checks the
  document). CLI `proofing`, MCP `word_read command=proofing`.
- **Table autofit — `Table.autofit(mode)`.** Resize a table's columns to fit
  their contents (`"content"`), stretch to the page (`"window"`), or pin the
  current widths (`"fixed"`). Across the Python API, the `autofit_table` exec op,
  the CLI (`table autofit`), and MCP (`word_write command=table action=autofit`).
- **`drop_cap` — the editorial oversized initial letter.** `anchor.drop_cap(lines=3,
  position="dropped"|"margin"|"none", distance=…, font=…)` turns the first letter
  of the anchor's paragraph into a real Word `DropCap` (the body text wraps around
  it natively, not a faked big-font run); `position="none"` removes one. Across
  the Python API, the `drop_cap` exec op, the CLI (`drop-cap`), and MCP.
- **`line_spacing` on `format_paragraph` / `set_style`.** Sets the leading
  *within* a paragraph (distinct from `space_before`/`space_after`, which space
  paragraphs apart): a number is a multiple of single spacing (`1`, `1.5`, `2`),
  the keywords `"single"`/`"1.5"`/`"double"` map to Word's named rules, and a
  length string (`"14pt"`, `"1.5cm"`) sets an exact line height. Wired through
  the Python API, the `format_paragraph` / `set_style` exec ops, the CLI
  (`format-paragraph --line-spacing` / `style set --line-spacing`), and MCP.
- **A dedicated `Equation` paragraph style.** Display equations now land on a
  centred, `Normal`-based `Equation` paragraph style (created on first use), so
  an equation is styled consistently regardless of where it was inserted — and
  there's a stable, named hook for future equation numbering / cross-references.

### Fixed
- **Equations no longer inherit a neighbouring heading's style.** An equation
  inserted before a `Heading 2` was written at the paragraph boundary and
  adopted the *following* paragraph's style — coming out styled `Heading 2` and
  polluting the navigation outline / TOC (it appeared as a heading entry). The
  equation's paragraph style is now pinned after insertion: `display=True` gets
  the centred `Equation` style; `display=False` is reset to `Normal` and
  left-aligned (it remains its own paragraph but reads as body text). The
  returned `equation:N` is documented as a positional id (Word's `OMaths` order)
  that renumbers when an earlier equation is inserted — re-list, don't cache it.
- **Composing at the end of a document no longer merges into the last
  paragraph.** `insert_block` (and so `insert_section` / `insert_markdown`)
  targeting `doc.end` wrote the block *before* the final paragraph mark, so when
  the last paragraph already had text the first inserted paragraph fused into it
  — `…last line.` + `## Heading` became one `…last line.Heading` paragraph,
  stealing the heading's style. The end-of-document case now detects the
  terminal mark correctly (`doc.end`'s range ends one short of it) and either
  fills an empty final paragraph (no stray trailing empty) or opens a fresh one
  after a non-empty one (no merge, no style theft).
- **A pure read no longer dirties the document.** `doc.stats()` and
  `anchor.location()` repaginate first (for print-layout-truth `pages`/`lines`),
  which flips Word's dirty bit — so a read of a freshly-saved document used to
  report (and leave) a spurious unsaved-changes star. Both now snapshot and
  restore `Document.Saved` around the repaginate, honouring their "nothing is
  mutated" contract.

### Changed
- **`set_borders` reconciles its line-style field name across surfaces.** The
  MCP `set_borders` command and its `word_write` schema name the line style
  `line_style` (to avoid colliding with the paragraph-`style` param), but a
  hand-built `word_exec`/`exec` batch reusing that name had it warned-and-ignored
  — the op only read `style`. The exec op now accepts `line_style` as an alias
  for `style`, so the same name works on every surface. (CLI `--style` and the
  Python `set_borders(style=…)` keyword are unchanged.)

## [0.14.0] — 2026-06-11

### Added
- **Equations — insert math from UnicodeMath, LaTeX, or MathML; read it back.**
  `anchor.insert_equation(*, unicodemath= | latex= | mathml=, where="after",
  display=True)` places a built-up Office Math equation on its own paragraph and
  returns an `EquationAnchor` (`equation:N`). Three input dialects, exactly one
  per call:
  - `unicodemath=` — Word's native linear form (`"a^2+b^2=c^2"`); typed into a
    math zone and *built up* by Word. Zero dependencies.
  - `mathml=` — a `<math>…</math>` string, converted to OMML through Office's own
    shipped `MML2OMML.XSL` (via MSXML). Zero dependencies.
  - `latex=` — a LaTeX math string; the LaTeX→MathML hop uses the **optional
    `latex` extra** (`pip install "wordlive[latex]"`, `latex2mathml`), then the
    same MathML→OMML→Word path. A missing backend raises a clear `EquationError`.

  `display=True` centres the equation; `display=False` marks it inline. The read
  side is `doc.equations` (a discovery collection: `equation:N` id, type, linear
  preview, `para:N`) and `EquationAnchor.mathml` — a **non-mutating** round-trip
  back to MathML via Office's `OMML2MML.XSL`. New `equation:N` anchor id (in
  `anchor_by_id` and `doc.stats()`); the `insert_equation` exec op; CLI
  `insert-equation` + `equations`; MCP `word_write` `insert_equation` and
  `word_read` `equations`. New `EquationError` (exit `1`, like `ImageSourceError`).
- **Compose helpers — add a whole section, or a chunk of Markdown, in one op.**
  A thin layer over `insert_block` (and the `**bold**`/`*italic*` run parser) so
  an agent composes structure instead of issuing a storm of single inserts:
  - `anchor.insert_section(heading, body, *, level=1, where="after")` places a
    `Heading {level}` paragraph plus its body (the `insert_block` items shape, or
    a bare string) atomically and returns the section's `range:START-END`.
  - `anchor.insert_markdown(md, *, where="after")` maps a **constrained-Markdown
    subset** to real Word structure — `#`/`##`/`###` → `Heading 1/2/3`, `-`/`*`
    → a bulleted list, `1.` → a numbered list (numbered 1..N over its own span),
    blank-line-separated text → `Normal` paragraphs, inline `**bold**`/`*italic*`
    honoured. Explicitly a subset, not CommonMark: no code fences, nested lists,
    block quotes, or tables in v1 — anything unrecognised stays literal text.
  - `heading.replace_section_body(body, *, markdown=False)` clears the body under
    a heading (up to the next same-or-higher heading) and inserts a replacement,
    keeping the heading — the "rewrite section X" workflow. `body` is the items
    shape, or a Markdown string with `markdown=True`.

  All three return the new content's `range:START-END`. New CLI commands
  `insert-section`, `insert-markdown`, `replace-section`; the `insert_section` /
  `insert_markdown` / `replace_section` exec ops; and the matching `word_write` /
  `word_exec` MCP commands. Block parsing lives in a new COM-free `_markdown.py`.
- **Document introspection — reason about layout without a snapshot.** Two cheap
  read surfaces so an agent can answer "what page is this on" / "how long is
  this" deterministically, no vision pass:
  - `anchor.location()` → `{page, end_page, line, column, in_table}`: where an
    anchor sits in the laid-out document. `page`/`end_page` are the pages its
    first and last characters fall on (its page *span* — equal for a single-line
    anchor), so a table/section/image that straddles a boundary reports both;
    scan `paragraphs` and watch `page` step up to find "which paragraph starts
    page 2". CLI `locate --anchor-id ID`, MCP `word_read location`.
  - `doc.stats()` → `{pages, words, characters, paragraphs, lines, sections,
    headings, tables, images, comments, revisions, saved}`: the "what am I
    looking at before I act" read. Text counts come from Word's
    `ComputeStatistics`; the structural counts from wordlive's own collections
    (so they agree with `doc.tables` / `outline` / …). CLI `stats`, MCP
    `word_read stats`.

  Both are pure reads with **no exec op**. Page/line numbers are print-layout
  truth, so each **repaginates first** — content-neutral, so the user's
  selection, scroll, and view are left untouched (the same guarantee a snapshot
  gives). Backed by a widened `WdInformation` and a new `WdStatistic`
  constant enum.
- **Table-as-records — read/update a table by its header row.** The read/update
  mirror of v0.13.0's "tables from records" write side, header-name indexed
  throughout:
  - `Table.records()` reads the body rows back as a list of `{header: value}`
    dicts (row 1 is the header). CLI `table records N`, MCP `word_read
    table_records`.
  - `Table.append_record({...})` appends a row from a dict (keys mapped to
    header columns; missing → empty, extra → ignored, like the create path).
    CLI `table append-record`, the `append_record` exec op, MCP `table` action.
  - `Table.update_row(key, {...}, column=None)` sets cells by header name on the
    first row whose key-column (the first column, or the header named by
    `column`) equals `key` — addressing a row by content instead of a fragile
    1-based index. Validates against the header before mutating (unknown column
    / values key → exit 1; no matching row → exit 2). CLI `table update-row`,
    the `update_row` exec op, MCP `table` action.

### Fixed
- **Documentation: the `exec` op vocabulary is now listed in full.** The CLI
  agent skill (`wordlive-cli/SKILL.md`) and the `exec --help` docstring each
  enumerated only a subset of the batch ops, silently omitting 16 that have long
  been supported (`format_run`, `set_shading`, `set_borders`, `add_tab_stop`,
  `add_style`, `set_style`, `insert_field`, `set_page_setup`, `update_fields`,
  `insert_footnote`, `insert_endnote`, `insert_toc`, `add_bookmark`,
  `add_hyperlink`, `insert_cross_reference`, `insert_caption`) — so an agent
  reading the skill would wrongly conclude they could not be batched. Both lists
  are completed, the CLI skill now also points at the `list show` / `list info`
  and `header read` / `footer read` read commands, and a new
  `tests/test_skill_consistency.py` pins both enumerations to
  `_ops.OP_REQUIRED_FIELDS` so the surfaces can no longer drift apart silently.

## [0.13.0] — 2026-06-09

### Added
- **Block insert — drop a contiguous run of styled paragraphs in one op.**
  `anchor.insert_block(items, where="after")` places a whole styled section (a
  feature list, a heading plus its body) at a single point in natural reading
  order — no more reverse-ordering single inserts to dodge positional-anchor
  renumbering. Each item is a plain string or `{text|runs, style?}`; it returns
  a `RangeAnchor` (`range:START-END`) spanning the block, so a follow-up op can
  target the whole run (e.g. bullet it with `apply_list`). New CLI command
  `insert-block --anchor-id ID --items JSON` (or `--items -` for stdin), the
  `insert_block` exec op, and the `word_write` / `word_exec` MCP command.
- **Inline runs — formatted spans within an inserted paragraph.** Inserted text
  can now carry character formatting in one shot, so the standard "**Bold
  lead** — rest" bullet no longer needs a second find→style pass. Two forms,
  both normalising to the same runs: a tiny inline **markdown** (`**bold**`,
  `*italic*`, `***both***`, with `\*` / `\\` escapes) wherever an item's `text`
  is given, and a **structured** `runs: [{text, bold?, italic?, underline?,
  style?}]` for unambiguous/precise control. Exposed on `insert_block` items,
  the `insert_paragraph` op's `runs` field, and the CLI `insert --runs JSON`.
  Plain `insert --text` stays literal (markdown lives in block/`runs`).
- **Tables from tabular data — build a table straight from your data.**
  `insert_table` / `table create` / the `create_table` op now accept **records**
  (a list of objects, `[{"Item":"Travel","Cost":"$400"}, …]`) whose keys become
  a bolded header row, in addition to the existing row-major 2-D array. When
  `data` is given, `rows`/`cols` are **optional** — inferred from its shape — so
  the common case is just `table create --anchor-id end --data …` (or
  `doc.end.insert_table(data=…)`). Pass explicit `rows`/`cols` to pad the grid
  larger than the data; without `data`, both stay required.
- **Persistence — save the document, or export a PDF deliverable.** New ungated
  Python-API methods: `doc.save()` (to the existing file), `doc.save_as(path,
  fmt="docx", overwrite=False)`, `doc.export_pdf(path, from_page=None,
  to_page=None)`, and a `doc.saved` property. The CLI gains `save`, `save-as
  PATH [--format docx] [--overwrite]`, and `export-pdf PATH [--pages A-B]`; MCP
  gains `word_write` commands `save` / `save_as` / `export_pdf`. **The CLI/MCP
  surfaces are gated** (the Python API is not): saving is *default-deny* and only
  writes inside directories whitelisted with `--save-dir` (repeatable) /
  `WORDLIVE_SAVE_DIRS` — with none configured, saving is off. Containment
  resolves the target first (so `..`/symlinks can't escape) then requires it
  inside the whitelist. `save_as` writes `.docx`; PDF goes through `export_pdf`
  (the recommended hand-back-a-deliverable path, a pixel-faithful render via the
  same engine as `snapshot`). `PathNotAllowedError` (exit 1) is the policy-denial
  type. **Not an exec op** — a terminal side-effect with no undo.
- **Low-resolution snapshots — `max_dim` (cheap whole-document layout checks).**
  `snapshot(..., max_dim=N)` caps each rendered page's **long edge** to `N`
  pixels (only ever lowering resolution). A vision model is billed on an image's
  pixel area, not its dpi, and that area depends on the page geometry — so a
  long-edge cap gives a predictable per-page token budget regardless of paper
  size, the right lever for "render the whole doc and check my styling landed"
  without the token cost of full-resolution pages (~1000 stays legible; e.g. a
  Letter page drops from 1275×1650 to a capped size). On
  `Document.snapshot` / `Document.snapshot_anchor` / `Anchor.snapshot`, the
  `wordlive snapshot --max-dim N` CLI flag, and `word_snapshot`'s `max_dim` param.
  `dpi` is unchanged (default 150) and composes with `max_dim` (the cap wins when
  it implies a lower resolution).
- **Image extraction — read embedded pictures back out.** The read mirror of
  `insert_image`, for handing a document's images to a vision model.
  `anchor.read_image()` returns `(bytes, mime_type)` for the single picture in an
  anchor's range; the new `image:N` anchor (1-based over Word's `InlineShapes`)
  targets one directly, and `doc.images` is a read-only discovery collection
  whose `list()` emits `{index, anchor_id, mime, width, height, alt_text, para}`.
  Extraction goes through `Range.WordOpenXML` (Flat OPC) — no clipboard, no
  save-to-temp, pure stdlib. CLI `wordlive images` (list) and `wordlive
  read-image --anchor-id ID [--out FILE]` (`--out` writes the raw bytes and
  reports `{path, mime, bytes}`; otherwise base64 + mime inline), and
  `word_read command="images"` / `command="read_image"` over MCP. A range with no
  image — or more than one — raises `ImageSourceError`. No exec op (extraction is
  a read, off the `doc.edit()` surface).
- **Paragraph pagination controls.** `format_paragraph` gains three tri-state
  flags for clean multi-page layout: `keep_together` (keep all lines of a
  paragraph on one page), `keep_with_next` (keep a paragraph with the following
  one — e.g. a heading with its first body line), and `widow_control` (prevent a
  lone first/last line stranded at a page boundary). They join the existing
  `page_break_before`, write to `ParagraphFormat.KeepTogether` /
  `.KeepWithNext` / `.WidowControl`, and are available on the `format_paragraph`
  exec op, `wordlive format-paragraph` CLI (`--keep-together` / `--keep-with-next`
  / `--widow-control`), and `word_write command="format_paragraph"` — plus
  `style.format_paragraph(…)` for a style's defaults.
- **Repeating table heading rows.** `Table.set_heading_row(row=1, heading=True,
  allow_break=None)` marks a row as a heading that repeats at the top of every
  page the table spans (`Row.HeadingFormat`); `allow_break` controls
  `Row.AllowBreakAcrossPages` and defaults to keeping a heading row intact. Wired
  through the `set_heading_row` exec op, `wordlive table set-heading-row` CLI, and
  `word_write command="table" action="set_heading_row"`.
- **`OpError` is now part of the public API** (`wordlive.OpError`). The
  malformed-op / bad-input exception the `exec` batch and dispatched writes
  already raised (exit code 1) was previously importable only from
  `wordlive.exceptions`; it now lives in `__all__` alongside the rest of the
  taxonomy and is documented in [Errors & exit codes](docs/errors.md).

### Security
- **Image-source path hardening (read-side gate, pairs with persistence).** On
  the CLI / MCP surfaces, `insert-image --path` (and the `insert_image` exec op's
  `path`) now **reject a non-local source** — a UNC path (`\\host\share\…`), a
  `file://`, or any URL — *before* the filesystem `is_file()` probe, which on a
  UNC path would itself authenticate to a remote SMB server and leak NTLM
  credentials (URLs were an SSRF / local-file-disclosure vector). An optional
  `--image-dir` / `WORDLIVE_IMAGE_DIRS` allowlist further restricts which local
  directories a path may come from. The Python API is unchanged (trusted);
  base64 / bytes image sources are unaffected.

### Changed
- **CLI bookmark ops consolidated (verb-first).** Creating a bookmark moved from
  `bookmark add NAME --anchor-id ID` to **`write bookmark NAME --create
  --anchor-id ID`** (bookmark creation is semantically a write, keeping the
  `read`/`write` dispatch groups whole and parallel with `cc`). `read bookmark`
  gained **`--list`** (every bookmark name; `--include-hidden` for Word's
  internal ones), surfacing `doc.bookmarks.list()`. Section listing moved from
  the one-verb `section list` group to a top-level **`sections`** verb
  (flat-first). The old `bookmark add` and `section list` spellings remain as
  **hidden, deprecated aliases for one release**. The Python API and `exec` ops
  are unchanged.

### Fixed
- **`insert_caption` now produces a real standalone caption.** Previously it
  collapsed the anchor to a point and let Word fuse the `SEQ` field + title
  **inline into the host paragraph** (restyling it `Caption`); on a table cell it
  could raise a COM "end of a table row" error. The caption now always lands in
  its **own `Caption`-styled paragraph**, leaving the target paragraph untouched,
  and a table-cell anchor captions the **whole table** (above/below it) rather
  than a cell. Placement follows convention — a `Table` caption goes **above**, a
  figure **below** — overridable with the new `position` argument
  (`"above"`/`"below"`); the CLI gains `--position` and MCP a `position` param
  (the old `before` flag is still honoured on the exec op for back-compat).
- **Docs build no longer breaks on the `OpError` cross-reference.** `Document.save`'s
  docstring linked to `OpError`, which wasn't rendered anywhere, failing the
  strict (`mkdocs build --strict`) docs CI. `OpError` and the already-public
  `PathNotAllowedError` are now both documented in the Python API reference and
  the [Errors & exit codes](docs/errors.md) hierarchy / exit-code table.

### Docs
- **PyPI project links + keywords.** `pyproject.toml` now declares
  `[project.urls]` (Homepage, Documentation, Repository, Changelog, Issues) and
  `keywords`, so the PyPI sidebar links out and the package is discoverable.
- **README badges** (PyPI version, Python versions, license, CI, docs).
- **New "Agents & LLM tools" guide** with copy-paste setup per client (Claude
  Code, Claude Desktop, Cursor, generic MCP) — consolidating the skill / MCP /
  `llm-help` paths.
- **`CONTRIBUTING.md` and `SECURITY.md`.** A contributor guide (uv dev setup, the
  four invariants, the four-surfaces-must-agree rule, testing / lint / docs
  gates, commit & PR conventions) and a security policy (private vulnerability
  reporting plus the gated-surface threat model — trusted Python API vs.
  prompt-injection-aware CLI/MCP path policy).

## [0.12.0] — 2026-06-08

### Added
- **Character formatting — `format_run`.** Direct run-level formatting on any
  anchor: `anchor.format_run(bold=…, italic=…, underline=…, strikethrough=…,
  font=…, size=…, color=…, highlight=…, subscript=…, superscript=…, small_caps=…,
  all_caps=…, spacing=…)`, tri-state like `format_paragraph`. Pairs with a
  `range:START-END` id to style a phrase. Colours accept a name (`"red"`), hex
  (`"#FF0000"`), or `(r, g, b)`; `highlight` is a named palette colour;
  `size`/`spacing` accept points or a unit string (`"12pt"`, `"1.5mm"`). Wired
  through the `format_run` exec op, `wordlive format-run` CLI, and
  `word_write command="format_run"`.
- **Borders, shading & tab stops.** `anchor.set_shading(fill=…)`,
  `anchor.set_borders(sides=…, style=…, weight=…, color=…)`, and
  `anchor.add_tab_stop(position, align=…, leader=…)` — range- and cell-level (a
  `Cell` is an `Anchor`). Border weight snaps to Word's discrete line widths.
  Exec ops `set_shading`/`set_borders`/`add_tab_stop`, CLI `shading`/`borders`/
  `tab-stop`, and the matching `word_write` commands (the border line style is
  the `line_style` param there, to avoid colliding with `style`).
- **Style creation & modification — styles are now writable.**
  `doc.styles.add(name, type=…, based_on=…, next_style=…)` defines a new style
  and returns a writable `Style`; `style.format_run(…)` / `style.format_paragraph(…)`
  set its font / paragraph defaults (the same kwarg vocabulary as the anchor
  methods, minus `highlight`), and `style.base_style` / `style.next_paragraph_style`
  chain styles. Exec ops `add_style`/`set_style`, CLI `style add`/`style set`, and
  the matching `word_write` commands. The brand/template primitive: define a
  house style once, then `apply_style` it everywhere.
- **Internal colour/units helper** (`_format.py`) underpinning the above:
  colours → Word's byte-swapped BGR long; lengths (`pt`/`in`/`cm`/`mm`) → points.
- **Page setup writes & multi-column layout.** `section.set_page_setup(margins=…,
  top_margin=…, …, gutter=…, orientation=…, paper_size=…, columns=…,
  column_spacing=…)` — the write mirror of `page_setup()`. `margins` sets all four
  at once (per-side kwargs override); lengths take points or a unit string;
  `columns=N` lays the section out in N equal newspaper columns (the section half
  of `insert_break("column")`). Exec op `set_page_setup`, CLI `page-setup`, and
  `word_write command="page_setup"`. Per-section; `doc.sections[1]` is the whole
  document for a single-section file.
- **Fields & page numbers.** `anchor.insert_field(kind, text=…)` inserts a
  self-updating field — `page`/`numpages`/`date`/`time`/`filename`/`author`/
  `title`, or `field` with a raw field code in `text`. `HeaderFooter.insert_page_number()`
  is the footer sugar for `insert_field("page")`, and `doc.update_fields()`
  recomputes the document's fields. Fields land in the anchor's own story, so
  page numbers in headers/footers work. Exec ops `insert_field`/`update_fields`,
  CLI `insert-field`/`update-fields`, and the matching `word_write` commands.
- **Footnotes & endnotes.** `anchor.insert_footnote(text)` /
  `anchor.insert_endnote(text)` attach a note to any anchor's range and return a
  `Footnote` / `Endnote` (addressed `footnote:N` / `endnote:N`) whose `set_text`
  edits the body and `delete()` removes the mark and body together. Word
  auto-numbers the reference mark. Read-only discovery via `doc.footnotes` /
  `doc.endnotes` (`list()` reports each note's number, text, and anchoring
  `para:N`). Exec ops `insert_footnote`/`insert_endnote` (the new id comes back
  in `outputs`), CLI `insert-footnote`/`insert-endnote` + `footnotes`/`endnotes`
  listings, and the matching `word_read`/`word_write` commands.
- **Table of contents.** `anchor.insert_toc(levels=(1, 3), use_heading_styles=…,
  hyperlinks=…)` inserts a TOC built from the document's headings and returns a
  `Toc` with `update()` / `update_page_numbers()`; `doc.add_toc(...)` is the sugar
  for one at the document start. Page numbers populate after repagination — call
  `update()`, `doc.update_fields()`, or take a `snapshot`. Exec op `insert_toc`,
  CLI `insert-toc`, and `word_write command="insert_toc"`.
- **Anchoring & linking — bookmarks, hyperlinks, cross-references, captions.**
  `doc.bookmarks.add(name, anchor)` creates a bookmark over a range (name
  validated against Word's rules first) — the prerequisite for the rest.
  `anchor.link_to(address=… | bookmark=…, text=…, screen_tip=…)` makes an anchor
  a hyperlink (external URL or internal bookmark jump); with `text` it inserts
  new linked text rather than overwriting the range.
  `anchor.insert_cross_reference(target, kind=…)` references another anchor —
  `target` is a `bookmark:NAME`, `heading:N`, `footnote:N`, or `endnote:N` id,
  `kind` is `text`/`page`/`number`/`above_below`. `anchor.insert_caption(label=…,
  text=…)` adds an auto-numbered caption. Exec ops `add_bookmark`/`add_hyperlink`/
  `insert_cross_reference`/`insert_caption`, the matching CLI verbs
  (`bookmark add`, `link`, `cross-ref`, `caption`), and `word_write` commands.
  All four features are exercised by live-Word smoke tests.
- **Tracked-changes visibility — `doc.revisions` and `snapshot(markup=…)`.** An
  agent making tracked edits can now *see* them, structurally and visually.
  `doc.revisions.list()` reports each tracked change as
  `{index, type, author, text, anchor_id, start, end, date}` (`type` is
  `"insert"` / `"delete"` / `"format"` / …) — read via `wordlive revisions`,
  `word_read command="revisions"`, and indexable as `doc.revisions[N]`.
  `doc.snapshot(markup="all")` (and the `--markup all` / `markup="all"` CLI/MCP
  options) renders revision marks and comment balloons into the image instead of
  the final text — via the export's `Item` parameter, so the user's on-screen
  markup mode is left untouched. Track-changes status is now also readable over
  MCP (`word_read command="track"`).
- **`delete_paragraph` — remove a paragraph, mark and all.**
  `doc.delete_paragraph(anchor)` deletes the paragraph(s) at an anchor including
  the trailing paragraph mark, so the surrounding text closes up (no empty line,
  unlike `replace`-ing with `""`) — for that stray leading empty paragraph.
  Deleting the document's last paragraph clears it but keeps Word's mandatory
  final mark. Exec op `delete_paragraph`, CLI `delete-paragraph`, and
  `word_write command="delete_paragraph"`.

### Changed
- **Bad formatting input now raises `OpError` (exit 1, bad-input) instead of a
  raw `ValueError`.** `format_paragraph` (and the new formatting methods) catch
  colour/length/enum coercion errors and re-raise as `OpError`, so an exec batch
  reports the failure cleanly instead of crashing the op loop. Indents/spacing on
  `format_paragraph` now also accept unit strings.
- **In-cell `find` / `find_replace` no longer overruns the cell boundary.** A
  cell's text ends with CR + the cell mark (`\r\x07`), which occupy a single
  document position — so a match at a cell's tail mapped its end *past* the cell
  into the next one, tripping the write-verification guard on essentially every
  in-cell find (the old `'Opus\r\x072'` error). The find/replace segmenter now
  drops those trailing markers from each cell segment, so a cell-scoped
  (`scope=table:N:R:C`) or whole-document find resolves inside the cell. The
  `ReplaceVerificationError` message is reworded too — it means the document
  shifted under the match (an earlier edit, or Track Changes leaving both runs),
  not specifically a table cell.
- **Numbered lists: apply over a span to number 1..N.** Applying a numbered list
  to paragraphs *one at a time* makes N independent "1." lists (and
  `continue_previous` only chains a clean in-order apply — it can't repair an
  already-split list). Applying `apply_list("numbered")` over a single
  `range:START-END` (or a heading's section) that spans all the items numbers
  them 1, 2, 3 as one list — now the documented, tested path. To repair a split
  list, `remove_list` the span then re-apply over it.

### Deferred
- Table-wide (`Table.Borders`) and page (`Section.Borders`) borders, shading
  patterns/textures, highlight on a style's font, and font kerning/character-scale
  on `format_run`.
- Page-setup: unequal column widths, line numbering, vertical alignment,
  different-first-page toggles, and an all-sections convenience (iterate
  `doc.sections`). `update_fields` refreshes the main story only (header/footer
  and other-story fields self-render on repagination — take a `snapshot`).
- The rest of the publishing flourishes (watermark, drop cap, text box / pull
  quote) — only the fields/page-number slice of that grab-bag landed here.
- Footnote/endnote polish: custom reference marks, note separators, numbering
  format/restart, and footnote↔endnote conversion. TOC: table of figures/
  authorities, custom TOC field codes, and explicit per-style level mapping.
  Cross-references and captions (which target footnotes/bookmarks) are the next
  cluster, not in this release.
- Anchoring & linking: hyperlink read-back (`doc.hyperlinks`) and edit/removal;
  cross-references to numbered-list items / equations and
  `IncludePositionInformation` combos; caption numbering format / chapter-style
  and a table of figures. `kind="text"` on a footnote/endnote cross-reference
  falls back to the note number (Word has no text content for a note mark).
- From the LLM-ergonomics feedback, still open: a multi-paragraph block insert
  and inline runs in insert ops (`insert_block` / `runs:[…]`); intra-batch output
  references (`$ops[N]`) and minting durable bookmark handles on insert
  (`bind:`); accepting/rejecting individual revisions (reads ship here, the write
  side stays on `.com`); and revision-aware text reads (a tracked `find_replace`
  on the *same* paragraph still drifts because both runs are present — re-read
  between tracked edits, or take a `markup="all"` snapshot).

## [0.11.1] — 2026-06-04

### Fixed
- **find/replace no longer crashes on the final paragraph.** A match in the
  document's last paragraph (and `add_table` / `create_table` anchored at the very
  end) wrote a range that straddled Word's undeletable terminal paragraph mark,
  raising COM `0x80020009`. The replace target is now clamped off that mark, and
  table insertion opens a trailing paragraph first.
- **find/replace inside table cells no longer corrupts neighbouring cells.**
  `Range.Text` offsets don't line up with Word document positions across table
  structure, so a whole-document `occurrence`/`all` replace could silently
  overwrite the wrong cell while returning success. Matching is now segmented at
  table-cell boundaries so offsets stay exact, and every write is verified against
  the located text — a mismatch raises the new `ReplaceVerificationError`
  (`code: "replace_verification"`, exit 1) instead of corrupting the document.
- **Inline images read back as a `[image]` token** instead of a phantom control
  character that polluted text reads and diffs (`heading.text`, paragraph / cell /
  header / footer / comment reads).
- **`section_continuous` / `section_next` breaks no longer pollute the outline.**
  The break paragraph inherited the anchor's style (a heading-anchored break
  became an empty heading in the navigation outline / TOC); it's now reset to
  `Normal`, matching `create_table`'s cell reset. Page/column breaks are
  unaffected.
- `StyleNotFoundError` now surfaces a distinct `code: "style_not_found"` to MCP
  clients (the CLI exit code stays `2`), instead of reusing `anchor_not_found`.
- A malformed anchor scheme (e.g. `banana:7`) now reports "unknown anchor type"
  and lists the valid types, distinguishing it from a valid-scheme-but-missing
  target.

### Added
- `insert_image(block=True)` (CLI `--block`, `word_write` / `word_exec`
  `insert_image` `block` field) places the image in its own new `Normal`
  paragraph instead of embedding it in the anchor's text run — so an inline image
  anchored `before` a heading lands on its own line above it rather than mid-line.
- **`paragraphs` now reports each paragraph's applied `style` name** (e.g.
  `"List Number"`, `"Normal"`) alongside `level`. Word's `OutlineLevel` is `10`
  for every non-heading paragraph, so the style name is what lets a caller tell a
  list item from body text and mirror an existing document's formatting on the
  first write. Surfaces through `doc.paragraphs.list()`, CLI `paragraphs`, and
  `word_read(command="paragraphs")`.

## [0.11.0] — 2026-06-01

### Added
- `word_read(command="guide")` — the full agent guide (anchor model, the
  `word_exec` op vocabulary, every field) is now fetchable as a **tool call**,
  not only the `wordlive://guide` resource. Resources aren't surfaced by every
  MCP client, so the guide the tool descriptions point at was unreachable in
  practice; the command needs neither Word nor a document. The `word_exec` /
  `word_write` tool docstrings also now inline the anchor taxonomy and op list so
  the essentials survive the projection into the MCP tool surface.
- `wordlive status` (and `word_read(command="status")`) now reports a `saved`
  flag and always a non-empty `name` (`Document1` for an unsaved document), so a
  caller can reliably confirm which document it is about to edit. The active
  document is matched by full path, robust when several unsaved documents share a
  blank path.
- `exec` / `word_exec` batches now return a `warnings` array flagging any field
  an op doesn't use (a typo, or a `style` handed to an inline append). The op
  still applies, but the ignored field is surfaced instead of silently dropped —
  closing the "successful-looking response hiding a wrong payload" footgun.

### Changed
- **Breaking (op vocabulary):** the `append` and `prepend` exec ops now add a new
  **paragraph** (taking `text` + optional `style`), matching their description
  and the `append_paragraph` / `prepend_paragraph` synonyms. The inline
  "continue the adjacent paragraph" behaviour moved to the new `append_inline` /
  `prepend_inline` ops (`text` only — no `style`). Previously a bare `append`
  concatenated inline and silently ignored any `style`, so a batch meant to build
  a styled document could collapse into one paragraph with no warning. The CLI
  `append` / `prepend` commands (with `--inline`) and the Python API
  (`Document.append` vs `append_paragraph`) are unchanged.
- New table cells created by `create_table` / `insert_table` now default to the
  `Normal` paragraph style regardless of the insertion anchor, instead of
  inheriting the anchor paragraph's style. A table dropped under a `Heading 2` no
  longer renders its cells as heading text or pollutes the navigation outline.
- CI: the release workflow's `actions/setup-node` is bumped `v4` → `v5` (off the
  deprecated Node 20 action runtime; GitHub forces Node 24 after 2026-06-16), and
  the bundle build now uses Node 22 LTS instead of Node 20.

## [0.10.2] — 2026-05-31

### Fixed
- `insert_image` now resolves a relative path to an absolute one before handing
  it to `InlineShapes.AddPicture`. Word resolves a relative filename against
  *its own* working directory, not the caller's, so a relative `--path` (or
  `image=` argument) previously failed with COM `0x80020009` ("not a valid file
  name"). Relative paths from the CLI's working directory now embed correctly.

## [0.10.1] — 2026-05-29

### Fixed
- `word_snapshot` no longer double-encodes its rendered pages. The tool returns
  each page as an MCP image content block, but its `-> list[Any]` return made
  FastMCP infer a structured-output schema and re-serialise the base64 PNG bytes
  into `structuredContent` as well — sending every page twice (a large, silent
  token cost on hosts that forward `structuredContent`). Marked the tool
  `structured_output=False` so the image is sent exactly once.

### Changed
- CI: the release workflow now packs `mcpb/` into `wordlive.mcpb` and attaches it
  to the GitHub Release (built outside the PyPI upload, so it never reaches PyPI).

## [0.10.0] — 2026-05-29

### Added
- Runnable example scripts under `examples/` (Python + PowerShell) and an
  **Examples** docs page, linked from the README and getting-started.
- **Python-API agent skill** (`wordlive-python`) alongside the existing CLI
  skill (now `wordlive-cli`). `install-skill` installs the CLI skill by default;
  `--python` installs just the Python one, `--both` installs both. `llm-help
  --python` prints the Python guide.
- **MCP bundle** (`mcpb/`) — a one-click `.mcpb` for Claude Desktop, kept in
  version lock-step with the package via `bump-my-version`.
- **`wordlive install-mcp`** — register the MCP server in Claude Desktop or
  Claude Code (`--client`, `--directory`, `--config`, `--print`, `--force`).
- `wordlive-mcp` console script (`[project.scripts]`), which the MCP docs and
  bundle already reference.
- MIT `LICENSE`, with `license` / `license-files` declared in `pyproject.toml`
  and the bundle manifest.

### Changed
- CI: bumped GitHub Actions off the deprecated Node 20 runtime to current
  Node 24 majors (`checkout` v6, `setup-python` v6, `setup-uv` v8,
  `upload-artifact` v7, `download-artifact` v8, `upload-pages-artifact` v5,
  `deploy-pages` v5).
- Added this changelog.
- Docs audit for v0.9.0: corrected the documented Python floor (3.10+), the
  README exit-code list (added `5`), the MCP `word_write` command list (added
  `insert_break`), the `design.md` roadmap (snapshot/MCP/tables/breaks shipped),
  and populated the previously-empty `CLAUDE.md`.

## [0.9.0] — 2026-05-29

First release since 0.8.3. Bundles four features that were developed earlier but
had not yet been published.

### Added
- **Snapshots** — `Document.snapshot(...)` / `Anchor.snapshot(...)` and the
  `wordlive snapshot` command render page(s) or a section to PNG (Word exports a
  pixel-faithful PDF, PyMuPDF rasterises it) so a vision model can *see* the
  layout. Requires the optional `snapshot` extra (PyMuPDF).
- **MCP server** (`wordlive-mcp`) — four dispatch tools (`word_read`,
  `word_write`, `word_exec`, `word_snapshot`) plus a `wordlive://guide` resource,
  for Claude Desktop and other agents. Requires the optional `mcp` extra.
- **Table creation / deletion** — `Document.add_table(...)`,
  `Anchor.insert_table(...)`, and `Table.delete()`; the `wordlive table create`
  / `table delete` commands; and the `create_table` / `delete_table` exec ops.
  Populates cells from a row-major `data` grid, defaults to the `Table Grid`
  style, and separates appended tables so Word doesn't silently merge adjacent
  ones.
- **Page / column / section breaks** — `Anchor.insert_break(kind=...)` and
  `format_paragraph(page_break_before=...)`; the `wordlive insert-break` command
  and a `--page-break-before` flag on `format-paragraph`; the `insert_break` exec
  op and a `page_break_before` field on `format_paragraph`.

## [0.8.3] — 2026-05-26

### Added
- `llm-help` command that dumps the full agent guide to stdout.

## [0.8.2] — 2026-05-21

### Added
- `append` / `prepend` helpers and `start` / `end` anchors for the document
  edges, so a document can be built top-down from a blank page.

### Changed
- CI: lint + test workflow across Python 3.10–3.15.

## [0.8.1] — 2026-05-21

### Added
- Inline `exec` JSON via `--ops` (and `--ops -` for stdin), terminal-paragraph
  append, and `before`/`after` placement on exec insert ops.

### Changed
- Tooling: `bump-my-version` configuration.

## [0.8.0] — 2026-05-21

Initial PyPI release. Drives a running Microsoft Word instance over COM
(Windows), with a JSON-in / JSON-out CLI built for LLM agents. Highlights of the
v0–v0.8 development line bundled here:

### Added
- Live automation core: `attach()` / `connect()`, anchors (bookmark, content
  control, heading, paragraph, range, cell, header/footer, start/end),
  `Document.edit()` atomic-undo with cursor/selection preservation, typed
  exceptions, and deterministic CLI exit codes.
- Reading: `status`, `outline`, `paragraphs`, `find` (fuzzy), `read section`.
- Writing: `replace`, `insert`, fuzzy find/replace, styles + `format-paragraph`,
  tables (read/edit), comments, track changes, lists & numbering, sections /
  headers / footers, and image insertion (path / bytes / base64).
- Batch edits via `exec` (single atomic undo), the bundled agent skill +
  `wordlive install-skill`, an mkdocs Material docs site, and PyPI
  trusted-publishing on tag push.

[0.16.0]: https://github.com/thomas-villani/wordlive/compare/v0.15.0...v0.16.0
[0.15.0]: https://github.com/thomas-villani/wordlive/compare/v0.14.0...v0.15.0
[0.14.0]: https://github.com/thomas-villani/wordlive/compare/v0.13.0...v0.14.0
[0.13.0]: https://github.com/thomas-villani/wordlive/compare/v0.12.0...v0.13.0
[0.12.0]: https://github.com/thomas-villani/wordlive/compare/v0.11.1...v0.12.0
[0.11.1]: https://github.com/thomas-villani/wordlive/compare/v0.11.0...v0.11.1
[0.11.0]: https://github.com/thomas-villani/wordlive/compare/v0.10.2...v0.11.0
[0.10.2]: https://github.com/thomas-villani/wordlive/compare/v0.10.1...v0.10.2
[0.10.1]: https://github.com/thomas-villani/wordlive/compare/v0.10.0...v0.10.1
[0.10.0]: https://github.com/thomas-villani/wordlive/compare/v0.9.0...v0.10.0
[0.9.0]: https://github.com/thomas-villani/wordlive/compare/v0.8.3...v0.9.0
[0.8.3]: https://github.com/thomas-villani/wordlive/compare/v0.8.2...v0.8.3
[0.8.2]: https://github.com/thomas-villani/wordlive/compare/v0.8.1...v0.8.2
[0.8.1]: https://github.com/thomas-villani/wordlive/compare/v0.8.0...v0.8.1
[0.8.0]: https://github.com/thomas-villani/wordlive/releases/tag/v0.8.0
