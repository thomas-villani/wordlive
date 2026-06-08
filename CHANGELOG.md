# Changelog

All notable changes to **wordlive** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
