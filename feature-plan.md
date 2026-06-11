# wordlive — feature roadmap

How to read this document, after the 2026-06-01 reorg:

- **Status-first, not version-numbered.** Earlier drafts labelled tranches
  `v0.9`–`v0.15`, but those numbers decoupled from real releases long ago (the
  repo is at **v0.11.0**, yet "v0.12" shipped back in release **0.9.0**). The
  fake labels are gone. Work is now bucketed by **status** and, within "next
  up", by **priority**. The authoritative release history lives in
  `CHANGELOG.md`; shipped items below carry the real release they landed in.
- Three parts: **I — Shipped**, **II — Approved / next up** (priority order),
  **III — Deferred & declined**. A short **cross-cutting** section and the
  **open papercuts** close it out.
- Ordering principle throughout: **LLM-agent leverage**, not spec order.

---

# Part I — Shipped

Quick index (capability → real release):

| Capability | Release |
|---|---|
| Core: `attach`/`connect`, anchors, `doc.edit()` atomic undo, typed errors, CLI | v0.8.0 |
| `replace`/`go_to`/`exec` by anchor-id; `append`/`prepend`; `start`/`end` | v0.8.0–0.8.2 |
| Find/replace (fuzzy) | v0.8.0 |
| Styles (read/apply) + `format_paragraph` | v0.8.0 |
| Tables (read/edit: cells, rows, grid) | v0.8.0 |
| Comments, track changes, `RangeAnchor` | v0.8.0 |
| Lists / numbering; sections / headers / footers | v0.8.0 |
| `para:N` addressing + cursor surface | v0.8.0 |
| Image **insertion** (path / bytes / base64, wrap modes) | v0.8.0 |
| Snapshots (render page/section to PNG) | v0.9.0 |
| MCP server (`word_read`/`word_write`/`word_exec`/`word_snapshot`) | v0.9.0 |
| Table **creation / deletion** | v0.9.0 |
| Page / column / section **breaks** | v0.9.0 |
| Python-API skill, MCP bundle, `install-mcp`, examples | v0.10.0 |
| Guide-as-tool-call, `status.saved`, exec `warnings`, paragraph-op split | v0.11.0 |
| find/replace boundary + table-cell fixes; inline-image `[image]` token; `insert_image(block=)` | v0.11.1 |
| Character formatting (`format_run`); borders / shading / tab-stops; style creation; colour/units helper | v0.12.0 |
| PageSetup writes + multi-column; fields / page numbers; `update_fields` | v0.12.0 |
| Footnotes / endnotes; table of contents | v0.12.0 |
| Bookmark creation, hyperlinks, cross-references, captions | v0.12.0 |
| Tracked-changes visibility (`doc.revisions`, `snapshot(markup=…)`); `delete_paragraph`; ergonomics fixes | v0.12.0 |
| Multi-page pagination (paragraph controls, repeating heading rows, caption fix) | v0.13.0 |
| Image **extraction** (`read_image`, `image:N`, `doc.images`); low-res `max_dim` snapshots | v0.13.0 |
| Persistence (save / save-as / export-pdf, gated) + image-source hardening | v0.13.0 |
| Block insert + inline runs; tables from records; verb-first bookmark CLI | v0.13.0 |
| Non-visual layout introspection (`anchor.location()`, `doc.stats()`) | v0.14.0 |
| Table-as-records read/update (`records`/`append_record`/`update_row`) | v0.14.0 |
| Compose helpers: `insert_section`, `insert_markdown`, `replace_section_body` | v0.14.0 |
| Equations (`insert_equation` UnicodeMath/LaTeX/MathML; `doc.equations`, `equation:N`, `.mathml`) | Unreleased |

The detail below preserves the **load-bearing reference facts** (addressing
schemes, gotchas a future change must respect). Deeper deliberation lives in git
history, `CHANGELOG.md`, and `spec.md`.

## Core loop — anchors, edit scope, exec (v0.8.x)

- `replace --anchor-id <id>`, `go_to --anchor-id <id>`: unified
  `doc.anchor_by_id(id)` resolver. Anchor-id taxonomy is the stable LLM-visible
  addressing scheme: `heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`,
  `table:N:R:C`, `range:START-END`, `header:S:WHICH` / `footer:S:WHICH`,
  `start`, `end`.
- `exec --script ops.json` / `--ops -`: batches N anchor ops in one `doc.edit()`
  / `UndoRecord`, so a multi-step edit reverts with one Ctrl-Z. **Failure
  semantics:** if op K fails, ops 1..K-1 are already applied; the UndoRecord
  closes and failure is reported — one Ctrl-Z reverts the partial work. A
  successful batch carries an `outputs` array (per structure-creating op) and a
  `warnings` array (fields an op ignored) since v0.11.0.

## Styles + paragraph formatting (v0.8.0)

- `doc.styles["Body Text"]`, `doc.styles.list()`, `style.exists`;
  `anchor.apply_style(name)` on all anchor types; `insert --style` validates via
  `doc.styles[]` before mutating (exit 2 on miss).
- `anchor.format_paragraph(**kwargs)` — alignment, indent, spacing,
  `page_break_before` (added with breaks) — + `wordlive format-paragraph` CLI.
- `StyleNotFoundError` subclasses `AnchorNotFoundError` → reuses exit code 2.
- **Note:** styles are **read-only** — wordlive consumes existing styles, it
  does not create or modify them. (Promotion candidate in Part II.)

## Tables — read/edit (v0.8.0)

- `doc.tables[N]` (1-based or by `Title`); `Table` wrapper: `row_count`,
  `column_count`, `cell(row, col)`, iteration, `read()`/`grid()`/`to_dict()`.
- `Cell` **is** an `Anchor` (inherits `apply_style`/`format_paragraph`/
  `set_text`); `Table.add_row`/`delete_row`.
- **Addressing:** cells are `table:N:R:C`; bare `table:N` is **not** an anchor
  (a whole table is a collection), addressed via `doc.tables[N]` / the `table`
  CLI group. Cells don't appear in `doc.outline()` (heading-only).
- Bookmarks inside cells round-trip through `set_text` (E2E-tested).

## Collaboration — comments, track changes, RangeAnchor (v0.8.0)

- `doc.comments.add(anchor, text, author=…)`, `.list()`, `doc.comments[N]`,
  `comment.resolve()`/`reopen()`/`delete()`, `comment.scope_text`. Comments
  attach to any anchor's range without mutating text. Addressed 1-based,
  matching Word's `Comments(n)`; `resolve()` uses the `Done` flag (Word 2013+).
- `with doc.tracked_changes(): …` (self-restoring scope) + `doc.track_changes`
  property + `wordlive track on|off|status` + `"tracked": true` on exec scripts.
- `doc.range(start, end)` → `RangeAnchor`, addressed `range:START-END` — what
  `find()` emits *and* what `anchor_by_id` resolves, so a find hit feeds straight
  into `replace`/`comments.add`.

## Lists & structure — numbering, sections, headers/footers (v0.8.0)

- List verbs on the base `Anchor`: `apply_list("bulleted"|"numbered"|"outline",
  continue_previous=…)`, `remove_list()`, `list_info()`, `restart_numbering()`,
  `indent_list()`/`outdent_list()`. `doc.lists` is a read-only discovery
  collection yielding a `RangeAnchor` per list.
- `doc.sections` collection; `HeaderFooter` **is** an `Anchor`, addressed
  `header:S:WHICH` / `footer:S:WHICH` (WHICH = primary/first/even) — so
  `set_text`/`apply_style`/`format_paragraph` work on it; `Section.page_setup()`
  **reads** (the write mirror `set_page_setup()` shipped in v0.12.0 — see below).

## Paragraph addressing & cursor (v0.8.0)

- `Paragraph(Anchor)` + `doc.paragraphs`; `para:N` shares its index space with
  `heading:N` (a heading is both). `wordlive paragraphs` listing emits offsets;
  `outline --all` is an alias.
- `insert --anchor-id ID --text … [--before|--after] [--style …]` on any anchor
  (the old `--after-heading` flag is gone; exec op is `insert_paragraph`).
- `cursor read` / `cursor write` (+ `Selection.write`) — deliberately **not** an
  `anchor_by_id` scheme. `cursor write` opts into `EditScope.allow_cursor_move()`
  — the one op that intentionally moves the cursor; `cursor read` reports the
  containing `para:N`.

## Image insertion (v0.8.0)

- `anchor.insert_image(image, *, wrap, where="after", width=None, height=None,
  alt_text=None, lock_aspect=True)` — `AddPicture(LinkToFile=False,
  SaveWithDocument=True)` (always embed). `image` is a path, raw `bytes`, or
  base64 (priority case: an LLM holding base64); classification + temp-file
  handling in `_images.image_on_disk`.
- **`wrap` is required, no default** — `inline|auto|square|tight|through|
  top-bottom|behind|front`. Non-inline calls `InlineShape.ConvertToShape()` +
  sets `WrapFormat.Type`. `wrap="auto"` = square if width ≤ half the section's
  usable text width, else top-bottom.
- `ImageSourceError(WordliveError)` → exit 1 (bad-input, **not** an
  AnchorNotFound; six-code contract untouched). CLI `insert-image` + `insert_image`
  exec op. Relative paths resolved to absolute before COM (fix in v0.10.2).

## Document construction — breaks + table creation (v0.9.0)

- **Breaks:** `anchor.insert_break(kind="page"|"column"|"section_next"|
  "section_continuous", where="after")` → `Range.InsertBreak`. Plus the clean,
  reflow-safe primitive `format_paragraph(page_break_before=True)`. `WdBreakType`
  subset (internal). exec op `insert_break` + CLI + MCP. Two ways to page-break
  by design: explicit one-off mark vs. paragraph property (guide steers agents to
  the latter for "every Heading 1 starts a page").
- **Table creation:** `doc.add_table(rows, cols, *, style=None, data=None,
  header=False)` + `anchor.insert_table(rows, cols, *, where="after", …)` +
  `Table.delete()`. `data` is row-major, underfill allowed / overflow rejected;
  `header=True` bolds row 1; `style=None` → built-in `Table Grid`. exec ops
  `create_table`/`delete_table`; CLI `table create`/`table delete`.
- **Gotcha (live-Word only):** adjacent tables silently **merge** if they touch
  with no paragraph mark between; `insert_table` probes
  `Range.Information(wdWithInTable)` on each side and drops a separator paragraph
  only where it abuts an existing table.
- v0.11.0: new table cells default to `Normal` style (don't inherit a heading
  anchor's style and pollute the outline).

## Snapshots + MCP + packaging (v0.9.0–0.11.0)

- `Document.snapshot(...)` / `Anchor.snapshot(...)` + `wordlive snapshot` render
  page(s)/section to PNG (Word → PDF, PyMuPDF rasterises). Optional `snapshot`
  extra. This is the "let a vision model see the layout" path.
  - **`max_dim` low-res cap (2026-06-09).** `snapshot(max_dim=N)` (CLI `--max-dim`,
    MCP `word_snapshot` `max_dim`) caps each page's long edge to `N` pixels, only
    ever lowering resolution. The lever for a cheap *whole-document* layout check:
    a vision model is billed on pixel area, so a long-edge cap gives a predictable
    per-page token budget regardless of paper size (a live probe measured ~17×
    fewer tokens/page at `max_dim=400` vs. default 150 dpi on a Letter page).
    Composes with `dpi`; the cap wins when it implies a lower resolution.
- MCP server: four dispatch tools + `wordlive://guide` resource (also fetchable
  as `word_read(command="guide")` since v0.11.0, because resources aren't
  surfaced by every client). Optional `mcp` extra.
- Python + CLI agent skills, `.mcpb` bundle, `install-mcp`, `examples/`.

## Boundary / corruption fixes + read polish (v0.11.1)

- find/replace no longer crashes on the **final paragraph** (target clamped off
  Word's undeletable terminal mark; `add_table`/`create_table` at `end` open a
  trailing paragraph first) and no longer corrupts **neighbouring cells** —
  matching is segmented at cell boundaries so `Range.Text` offsets stay exact,
  and every write is verified (mismatch → `ReplaceVerificationError`, exit 1).
- Inline images read back as a `[image]` token (not a phantom control char) in
  every text read. `section_continuous`/`section_next` break paragraphs reset to
  `Normal` (no phantom heading in the outline). `StyleNotFoundError` carries a
  distinct `code`; a malformed anchor scheme (`banana:7`) reports "unknown anchor
  type". `insert_image(block=True)` places an image on its own `Normal` line.
- `paragraphs` reports each paragraph's applied **`style` name** alongside
  `level` (Word's `OutlineLevel` is `10` for all non-headings, so the style name
  is what distinguishes a list item from body text).

## Publishing-quality styling & layout (v0.12.0)

The "make it look designed" cluster. Built on an internal `_format.py` colour/
units helper (colours → Word's **byte-swapped BGR long**; lengths pt/in/cm/mm →
points) — not in `__all__`, mirrors the `WdStyleType` pattern.

- **`format_run`** — direct character formatting on any anchor (`bold`, `italic`,
  `underline`, `strikethrough`, `font`, `size`, `color`, `highlight`, `subscript`,
  `superscript`, `small_caps`, `all_caps`, `spacing`), tri-state like
  `format_paragraph`. Pairs with `range:START-END` to style a phrase. Colours
  accept name/hex/`(r,g,b)`; bad colour/length/enum input → `OpError` (exit 1).
- **Borders / shading / tab stops** — `set_shading(fill=…)`, `set_borders(sides,
  style, weight, color)` (weight snaps to Word's discrete widths),
  `add_tab_stop(position, align, leader)`. Range- and cell-level (a `Cell` is an
  `Anchor`). MCP names the border line style `line_style` (avoids colliding with
  `style`).
- **Style creation / modification** — styles are now **writable**:
  `doc.styles.add(name, type=…, based_on=…, next_style=…)` returns a writable
  `Style`; `style.format_run(…)` / `style.format_paragraph(…)` set its defaults
  (same kwarg vocab, minus `highlight`); `style.base_style` /
  `.next_paragraph_style` chain. The brand/template primitive.
- **PageSetup writes + multi-column** — `section.set_page_setup(margins=…,
  per-side overrides, gutter, orientation, paper_size, columns=N, column_spacing)`
  — the write mirror of `page_setup()`. `columns=N` = N equal newspaper columns
  (the section half of `insert_break("column")`). Per-section; `doc.sections[1]`
  is the whole document for a single-section file.
- **Fields / page numbers** — `anchor.insert_field(kind, text=…)` inserts a
  self-updating field (`page`/`numpages`/`date`/`time`/`filename`/`author`/
  `title`, or raw `field` code); `HeaderFooter.insert_page_number()` is the footer
  sugar; `doc.update_fields()` recomputes (main story only — header/footer fields
  self-render on repagination). Fields land in the anchor's own story, so
  header/footer page numbers work.
- **Paragraph pagination controls (v0.13.0)** — `format_paragraph` gains
  tri-state `keep_together` / `keep_with_next` / `widow_control`, joining
  `page_break_before` → `ParagraphFormat.KeepTogether`/`.KeepWithNext`/
  `.WidowControl`. Also on `style.format_paragraph(…)`.
- **Deferred:** table-wide / page borders, shading patterns, highlight-on-a-style,
  kerning/character-scale; unequal column widths, line numbering, vertical
  alignment, different-first-page, all-sections convenience. **Flourishes still
  open** (Part II): watermark, drop cap, text box / pull quote.

## Reference apparatus — footnotes / endnotes, TOC (v0.12.0)

- **Footnotes / endnotes** — `anchor.insert_footnote(text)` / `insert_endnote(text)`
  attach a note to a range and return a `Footnote`/`Endnote` (addressed
  `footnote:N` / `endnote:N`, resolving to the **note-body range**, its own story
  `StoryType == 2`); `set_text` edits the body, `delete()` removes mark + body.
  Read-only discovery via `doc.footnotes` / `doc.endnotes` (`list()` → `{index,
  marker, text, para}`). **Spike gotcha:** the `Add` args must be **positional** —
  the `Text=` keyword is silently dropped under pywin32 late binding; empty
  `Reference` auto-numbers.
- **Table of contents** — `anchor.insert_toc(levels=(1,3), use_heading_styles=…,
  hyperlinks=…)` returns a `Toc` with `update()` / `update_page_numbers()`;
  `doc.add_toc(...)` = one at the document start. **Page numbers populate only
  after repagination** — call `toc.update()`, `doc.update_fields()`, or take a
  `snapshot` (which forces print layout).
- **Deferred:** custom marks, separators, numbering format/restart,
  footnote↔endnote conversion; table of figures/authorities, custom TOC field
  codes, per-style level mapping.

## Anchoring & linking — bookmarks, hyperlinks, cross-refs, captions (v0.12.0)

Creating a named anchor is the prerequisite for the rest; live testing caught
three issues the fake-based unit tests missed (noted below). All four are
covered by **live-Word smoke tests**.

- **Bookmark creation** — `doc.bookmarks.add(name, anchor)` over the resolved
  anchor; `name` validated against Word's rules (letters/digits/underscores,
  leading letter, no spaces) → typed error before mutating.
- **Hyperlinks** — `anchor.link_to(address=… | bookmark=…, text=…,
  screen_tip=…)` (`address` XOR `bookmark`). `text=None` links the existing
  range; `text=…` **inserts** new linked text (live-testing fix — it used to
  overwrite the range). Deferred: `doc.hyperlinks` read-back, edit/remove.
- **Cross-references** — `anchor.insert_cross_reference(target, kind=…)`; `target`
  is `bookmark:NAME` / `heading:N` / `footnote:N` / `endnote:N`. **Mapping layer:**
  `ReferenceItem` is a 1-based index into `GetCrossReferenceItems(ReferenceType)`
  (ordered strings — heading *text* / bookmark *names*), so a bookmark cross-ref
  takes the bookmark **name** (live-testing fix). `kind="text"` is invalid for
  footnote/endnote refs (falls back to the note number). **Gotcha:**
  `IncludePositionInformation=` as a keyword raises under pywin32 — omit/positional.
- **Captions** — `anchor.insert_caption(label="Figure", text=…,
  position="above"|"below")`. Always lands in its **own `Caption`-styled
  paragraph** (text/para/range anchors carve out a dedicated empty paragraph; a
  table **cell** anchor captions the **whole table** via `Cell._caption_object_range()`,
  so Word's native above/below placement fires — fixes the old "end of a table
  row" `ComError`). Convention default: table caption **above**, figure **below**
  (the exec op still honours a legacy `before` bool). *(Caption fix shipped in the
  v0.13.0 pagination batch.)*

## Image extraction — read images out for vLLMs (v0.13.0)

The read mirror of `insert_image`, for handing a document's images to a vision
model. Spiked clean: extraction goes through **`Range.WordOpenXML`** (Flat OPC,
inlines each media part as base64 on a tight per-shape range) — no clipboard, no
save-to-temp, pure stdlib (`xml.etree` + `base64`).

- **`anchor.read_image() -> (bytes, mime)`** — the sole `image/*` part of an
  anchor's range, base64-decoded; **`image:N`** anchor (1-based over
  `InlineShapes`); **`doc.images`** discovery collection (`list()` → `{index,
  anchor_id, mime, width, height, alt_text, para}`). A range with 0 or >1 images
  → `ImageSourceError`. **No exec op** (it's a read). CLI default is
  base64-in-JSON with `--out FILE` for file-out; MCP `read_image` returns base64
  + mime (fits the four-tool `word_read` dispatch). **Carry-forward:** `WordOpenXML`
  always serializes the full package skeleton (~64 KB floor); rapid COM access can
  return `RPC_E_CALL_REJECTED` (already the `WordBusyError` retry class).
- **Deferred:** floating-shape / chart-image export, OLE extraction,
  whole-page-to-image (use `snapshot`); EMF/WMF + cropped images return raw
  bytes+mime untouched (not separately exercised). Possible follow-up: an MCP
  `ImageContent` block so the model *sees* the extracted original directly.

## Persistence — save / save-as / PDF export, gated (v0.13.0)

wordlive's one coherent **filesystem boundary**: the **Python API is
trusted/ungated**; the **CLI/MCP surfaces are gated** because their input can be
prompt-injected.

- **Python (ungated):** `doc.save()` (errors if no path yet),
  `doc.save_as(path, fmt="docx", overwrite=False)`, `doc.export_pdf(path,
  from_page=None, to_page=None)`, `doc.saved`, `doc.path`.
- **CLI/MCP (default-deny + directory whitelist):** enabled only by configuring
  allowed dirs (`--save-dir` repeatable / `WORDLIVE_SAVE_DIRS`; MCP at launch).
  **Containment:** `Path(target).resolve()` *then* `is_relative_to` each whitelist
  dir (resolve first so `..`/symlink escapes can't slip out). Plain `Save()` is
  gated too. MCP exposes save as `word_write` commands (special-cased like
  `track`, bypassing `run_batch`) to hold the four-tool invariant.
- `save_as` writes **`.docx` only** (rejects `fmt="pdf"` — use `export_pdf`, which
  reuses `_snapshot._export_pdf`). `PathNotAllowedError` → exit 1 (the neutral name
  that also guards image-source reads). **Not an exec op** (terminal side-effect,
  no undo). **Deferred:** `doc`/`rtf`/`txt`/`html` formats.

## Block insert, inline runs, tables-from-records; CLI cleanup (v0.13.0)

From the 2026-06-08 LLM-ergonomics probe; the still-open items from that triage
are in Part II, the full triage in `CHANGELOG.md`.

- **Multi-paragraph block insert** — `anchor.insert_block(items, where="after")`
  places a contiguous run of styled paragraphs atomically (no reverse-ordering to
  dodge positional-anchor renumbering) and returns the block's `range:START-END`
  for a follow-on `apply_list`/comment. CLI `insert-block`, the `insert_block` op,
  `word_write`/`word_exec`.
- **Inline runs** — formatted spans in one shot: a tiny `**bold**` / `*italic*` /
  `***both***` markdown (escapes `\*` / `\\`) in any item `text`, **and** structured
  `runs:[{text, bold?, italic?, underline?, style?}]` — on `insert_block` items, the
  `insert_paragraph` op, and CLI `insert --runs`. Plain `insert --text` stays
  literal. Parser lives in `_runs.py`.
- **Tables from records** — `insert_table` / `table create` / `create_table` accept
  a list of dicts (keys → bolded header row) alongside the 2-D array, and infer
  `rows`/`cols` from `data` when omitted.
- **Repeating table heading rows** — `Table.set_heading_row(row=1, heading=True,
  allow_break=None)` → `Row.HeadingFormat` / `.AllowBreakAcrossPages` (defaults to
  keeping a heading row intact). *(Shipped in the pagination batch.)*
- **Tracked-changes visibility + `delete_paragraph` (v0.12.0)** — `doc.revisions`
  structured reader (`{index, type, author, text, anchor_id, start, end, date}`;
  CLI `revisions`, `word_read revisions`), `snapshot(markup="all")` renders
  revision marks/comment balloons (via export `Item`, leaving the user's on-screen
  mode untouched), track status over MCP. `doc.delete_paragraph(anchor)` removes a
  paragraph incl. its mark (the surrounding text closes up). In-cell `find` strips
  trailing `\r\x07` cell markers so a cell-scoped find resolves inside the cell.
  Numbered lists: apply over one span to number 1..N (`continue_previous` can't
  repair an already-split list — remove + reapply).
- **CLI bookmark ops consolidated (verb-first)** — creation moved to `write
  bookmark NAME --create --anchor-id ID`; `read bookmark --list
  [--include-hidden]`; section listing flattened to a top-level `sections` verb.
  Old `bookmark add` / `section list` stay as **hidden deprecated aliases for one
  release**. CLI-only (Python API + exec ops unchanged).

## Non-visual layout introspection + table-as-records (v0.14.0)

The top two asks from the 2026-06-09 gpt-5.4 review, both promoted from Part II.

- **`anchor.location()`** → `{page, end_page, line, column, in_table}` off
  `Range.Information` — where an anchor sits in the laid-out document. `page` /
  `end_page` are the pages the first / last character fall on (the anchor's page
  *span*; equal for a single-line anchor), so a table/section/image straddling a
  boundary reports both, and scanning `paragraphs` for a `page` step-up answers
  "which paragraph starts page 2". CLI `locate --anchor-id ID`, MCP
  `word_read location`. **No exec op** (a read).
- **`doc.stats()`** → `{pages, words, characters, paragraphs, lines, sections,
  headings, tables, images, comments, revisions, saved}`. Text counts from
  `ComputeStatistics(wdStatistic*)`; structural counts from wordlive's own
  collections; `saved` from `doc.saved`. CLI `stats`, MCP `word_read stats`.
- Both are **reads** that **repaginate first** (`Document.Repaginate()` —
  content-neutral; selection/scroll/view untouched) so page/line numbers are
  print-layout truth, the same guarantee `snapshot` gives. New `WdStatistic`
  enum + widened `WdInformation` (line/column/adjusted-page selectors).
- **Table-as-records** — the read/update mirror of v0.13.0's tables-from-records,
  header-name indexed (row 1 = header): `Table.records()` (read body rows as
  `list[dict]`), `Table.append_record({...})` (append a row from a dict, lenient
  mapping like the create path), `Table.update_row(key, {...}, column=None)`
  (set cells by header on the first row whose key-column equals `key` — content,
  not index; validates against the header before mutating). CLI `table
  records`/`append-record`/`update-row`; the `append_record`/`update_row` exec
  ops; MCP `word_read table_records` + `table` write actions.

## Higher-level compose helpers — sections & block markdown (v0.14.0)

gpt-5.4's #10 ask, built as a thin layer over `insert_block` + the `_runs.py`
inline parser (block parsing in a new COM-free `_markdown.py`). All three return
the new content's `range:START-END`.

- **`anchor.insert_section(heading, body, *, level=1, where="after")`** — a
  `Heading {level}` paragraph plus its body (the `insert_block` items shape, or a
  bare string) in one atomic op. `heading` carries the same inline markdown an
  item's `text` does; `level` is 1–9.
- **`anchor.insert_markdown(md, *, where="after")`** — a **constrained-Markdown
  subset** → Word structure: `#`/`##`/`###` → `Heading 1/2/3`, `-`/`*` → a
  bulleted list, `1.` → a numbered list (numbered 1..N over its own span — each
  same-kind run is inserted then `apply_list`-ed over that span), blank-line text
  → `Normal` paragraphs, inline spans via `_runs.parse_markup`. **Explicitly a
  subset, not CommonMark** — no code fences / nested lists / block quotes /
  tables in v1; unrecognised lines stay literal. Holds the v0.13.0 decision that
  plain `insert --text` stays literal: markdown is an opt-in verb.
- **`heading.replace_section_body(body, *, markdown=False)`** — clears the body
  under a heading (`section_range`, up to the next same-or-higher heading) and
  inserts a replacement after the heading, keeping the heading. `body` is the
  items shape, or a Markdown string with `markdown=True`. The "rewrite section X"
  workflow.

CLI `insert-section` / `insert-markdown` / `replace-section`; the
`insert_section` / `insert_markdown` / `replace_section` exec ops; matching
`word_write` / `word_exec` MCP commands.

- **Deferred:** full CommonMark (tables, code fences, nested/mixed lists,
  images-by-URL); round-trip *export* back to markdown.

---

# Part II — Approved / next up (priority order)

Everything here is **specced but not yet implemented** (re-verified 2026-06-09 —
the bulk of the 2026-06-01 roadmap shipped in **v0.12.0**–**v0.13.0**, and the
review's #1/#2 asks — non-visual introspection and table-as-records — plus the
#10 compose helpers shipped in **v0.14.0**; all moved to Part I). What's left
below is the genuine backlog, ordered by leverage.

## Publishing flourishes — the rest of the grab-bag

The publishing-quality cluster (character formatting, borders/shading/tab-stops,
style creation, PageSetup writes, fields/page numbers, pagination controls) all
shipped — see Part I. What's left from item 5's grab-bag, individually small;
bundle whichever land cheap:

- **Watermark** — `doc.set_watermark(text, …)` via the header-story
  `Shapes.AddTextEffect` (WordArt) — DRAFT/CONFIDENTIAL stamps.
- **Drop cap** — `paragraph.set_drop_cap(lines=3, position="dropped")` →
  `Paragraph.DropCap`.
- **Text box / pull quote** — `anchor.insert_text_box(text, …)` →
  `Shapes.AddTextbox`.
- The shipped `insert_field` primitive already covers the general `Fields.Add`
  case these once leaned on, so they're now pure floating-shape work.

## Charts (Excel-backed)

Approved 2026-05-31 as the SmartArt substitute. `Range.InlineShapes.AddChart2`
embeds a chart whose data lives in an embedded Excel workbook. Heavier than
images on two axes (a new dependency + a much larger surface), hence lower
priority than the publishing cluster.

- **`anchor.insert_chart(kind, data, *, title=None)`** — `kind` → `XlChartType`
  (`bar`→`xlColumnClustered`, `pie`→`xlPie`, `scatter`→`xlXYScatter`,
  `line`→`xlLine`); `data` (flat label→value mapping) populates
  `ChartData.Workbook.Worksheets(1)`.
- **Transitive Excel dependency** — `AddChart2` spins up hidden Excel. Gate
  behind an "is Excel available?" probe + a typed error (clean exit, not exit 1).
- `XlChartType` subset in `constants.py` (internal). **Keep narrow:** common
  kinds + flat `data` only. **Deferred:** multi-series, secondary axes,
  axis/series formatting, `BreakLink` policy, reading existing charts back out.

## Intra-batch output refs + durable insert handles (from the ergonomics probe)

Deferred from the 2026-06-08 LLM-ergonomics triage (§5) — the rest of that batch
shipped (see Part I). Two related `exec`-batch ergonomics:

- **Intra-batch output references** — let an op reference a prior op's output,
  e.g. `anchor_id: "$ops[0].table"`, so a batch can create a table then target
  it without a second round-trip.
- **Durable insert handles** — an optional `bind: "name"` that mints a bookmark
  on inserted content, giving a stable handle instead of a fragile positional
  `para:N` that renumbers under later inserts.

**Paragraph IDs for *existing* content (design note, 2026-06-09).** The gpt-5.4
review's #1 repeated ask was durable IDs for paragraphs the agent *didn't*
insert. A live probe (throwaway doc, `WordOpenXML` + saved-zip inspection)
settles the mechanism:

- **Word's native `w14:paraId` is unusable for wordlive.** It's an 8-hex-digit
  value (observed `2754048E`, `5D3F89E2`, `69999C2F`) but (a) assigned
  **lazily** — a freshly authored doc has *none*, even after save; Word only
  stamps them when a feature needs them (adding a single comment triggered
  assignment on all three test paragraphs), and (b) **invisible through the COM
  surface** — both `Range.WordOpenXML` and `Content.WordOpenXML` strip it; it
  surfaces only in the saved-zip `word/document.xml`. Reading it means
  save-then-unzip, which fights the "drive a *running* instance" model.
- **Content-derived codes (hash or slug) are labels, not anchors.** A git-style
  6-char hash of the text, or a `heading1-project-budget` slug, is readable but
  breaks on the exact edits we want to survive: it changes when the paragraph's
  text is edited, collides across duplicate paragraphs (two empty cells, two
  identical headings), and still needs a full scan to resolve back to a live
  Range. Durability across *content* edits is impossible without a stored marker.
- **Decision — minted durable handle, same machinery as `bind:`.** A
  `pin`/`stamp` verb plants a hidden bookmark (`_wl_<code>`; the `_` prefix is
  Word-hidden) on an existing paragraph's range and returns a short id
  (`pin:a3f9c2`). Word maintains the Range association across insert / delete /
  text edits natively — the real source of durability — and resolve is O(1) via
  `Bookmarks`. The code is just the bookmark's *name*, so an agent may instead
  supply a readable slug (`pin:budget-intro`). A deleted paragraph's handle
  correctly vanishes → feeds the stale-anchor diagnostic below. This unifies
  existing-content IDs with the insert-time `bind:` handle under one
  bookmark-backed scheme.
- **Bulk-pin the outline — `doc.pin_outline(levels=…)`.** Walk the outline and
  stamp a handle on every heading (or section start) in one call, returning the
  `{anchor_id → pin}` map, so an agent gets a durable navigation scaffold *up
  front* instead of pinning paragraph-by-paragraph. Idempotent (skip a range
  that already carries a `_wl_` bookmark); the natural form is a `pin=True`
  option on the `outline` read that stamps-and-returns in the same call.

**Stale-anchor diagnostics (from the same review).** When a positional `para:N`
resolves to something the agent didn't expect after a mutation, return a recovery
hint — the op index that failed, whether the target likely moved vs. vanished,
and a nearest-match suggestion (closest heading / fuzzy text hit) — rather than a
bare `AnchorNotFoundError`. Pairs with the minted handles above (the durable
escape hatch the hint can recommend).

## Revision write surface (read side already shipped)

`doc.revisions` (structured read), `snapshot(markup="all")`, and MCP track-status
shipped in v0.12.0 (Part I). Still open:

- **Accept / reject individual revisions** — mutating a single `Revision` stays
  on `.com` for now; a typed `revision.accept()` / `.reject()` (and a doc-wide
  accept/reject-all) is the natural next step.
- **Revision-aware text reads** — a tracked `find_replace` on the *same*
  paragraph still drifts because both the inserted and deleted runs are present.
  Workaround documented (re-read between tracked edits, or take a `markup="all"`
  snapshot); a proper revision-aware read model is the real fix.

---

# Part III — Deferred & declined

## Declined

- **Native SmartArt** — declined 2026-05-31. `Shapes.AddSmartArt(Layout, …)` +
  driving the `Nodes` tree is heavy and brittle (GUID-indexed layouts from
  `Application.SmartArtLayouts`, floating-only, hard to read back,
  locale/version-sensitive) for an intent that **charts** (Part II) and
  **render-diagram-to-image** (mermaid/graphviz → `insert_image`) already serve.
  Revisit only on a concrete SmartArt-specific need.

## Deferred (no concrete trigger yet)

- **Events / sinks** — `WithEvents(word.com, Handler)` for `DocumentBeforeSave`,
  `WindowSelectionChange`. Wait for a use case before designing the marshalling
  layer.
- **`asyncio` wrapper** — natural once events land. Sync core stays.
- **Read-model caching** — premature; live reads are correct. Cache when events
  arrive to invalidate on `DocumentChange`.
- **Styles deep cuts** — character/list/linked styles, theme-aware fonts, style
  import from template, `UpdateStyles`. (Basic style *creation/modification*
  shipped in v0.12.0 — see Part I.)
- **Document themes** — `Document.DocumentTheme` / `ApplyTheme`, theme color/font
  schemes. Brand-consistency play; the natural next step now that style creation
  has shipped.
- ~~**Equations** — `Range.OMaths.Add` / build.~~ **Shipped** (Unreleased):
  `insert_equation(unicodemath= | latex= | mathml=)` builds Office Math, with
  `doc.equations` / `equation:N` reads (`.mathml` round-trips via Office's own
  XSLT). LaTeX is the optional `latex` extra. See Part I.
- **Table polish** — merged/split cells (addressing assumes rectangular),
  `add_column`/`delete_column`, AutoFit fit-window/fit-content policy.
- **List polish** — custom list-template authoring, per-level bullet/number
  format, multi-section `LinkToPrevious` editing.
- **Comment/revision polish** — comment replies (`comment.reply`), author/date
  filtering on `list()`. (Per-revision accept/reject is now active backlog —
  Part II.)
- **Image polish** — wrap *side* + text distance; absolute/relative positioning
  (`Left`/`Top`); cropping; replace-in-place; force-own-paragraph; EMF/WMF.
- **Cursor/inline polish** — raw inline `insert_before`/`insert_after` (no new
  paragraph) on the CLI; a `write_cursor` exec op (fights `EditScope`'s
  cursor-restore in a batch).

## Unexplored COM control surfaces (catalogue for prioritization)

Sketched 2026-06-01 — the broader sweep of "what else could an agent drive
through COM that's actually useful." **Nothing here is approved or designed**;
it's a parking lot to prioritize from. Each entry: the COM entry point, the
agent use-case, and a rough **weight** (leverage vs. cost/risk). Entries are
filtered for usefulness — Word exposes far more (page borders art, ink,
hyphenation dictionaries, etc.) that isn't worth an agent's attention.

Cross-cutting constraint: anything here must still honour the four invariants.
The **application/window/view** surfaces (group I) are the ones in real tension
with *politeness* — they change global UI state, not document content — so they
need a louder opt-in than an ordinary edit, or should stay reads.

### A. Proofing & language (mostly reads — high agent value, low risk)

- **Spelling / grammar as structured data** — `Range.SpellingErrors`,
  `GrammaticalErrors`, `Application.GetSpellingSuggestions`. An agent can
  proofread and propose fixes (pairs naturally with comments / track-changes).
  *Weight: high value, low risk — a read surface that fits the agent loop.*
- **Readability statistics** — `Document.ComputeStatistics(wdStatistic…)` /
  the readability scores. "How dense is section 3." *Low cost.*
- **AutoCorrect / AutoText entries** — `Application.AutoCorrect`. Niche; mostly
  a write to global app state. *Low priority.*

### B. Document metadata & properties (reads + light writes)

- **Built-in & custom document properties** — `BuiltInDocumentProperties`
  (Title/Author/Subject/Keywords/Company/Category), `CustomDocumentProperties`.
  Read for situational awareness; write to stamp a generated deliverable.
  *Weight: medium, cheap.*
- **Document variables** — `Document.Variables` (`{ DOCVARIABLE }` backing).
  Data-driven templates: set a variable, refresh fields. *Pairs with the field
  primitive.*
- **Statistics** — page/word/char/line counts via `ComputeStatistics`. *Cheap
  read; useful before/after an edit.* ✅ **shipped v0.14.0** as `doc.stats()`.

### C. Document lifecycle & safety

- **Protection / editing restrictions** — `Document.Protect(type, …)` /
  `Unprotect`, read-only enforcement, `Permission` (IRM). Hand back a locked
  deliverable. *Weight: medium; gated like persistence (policy surface).*
- **Encryption / passwords** — `Document.Password` / `WritePassword`. *Security-
  sensitive; gate hard or leave to the human.*
- **Compare / merge** — `Application.CompareDocuments`, `Document.Merge`. "Diff
  these two drafts" as tracked revisions — genuinely agent-shaped. *Weight:
  medium-high; needs a second-document handle in the model.*
- **Inspect / redact** — remove personal info / hidden data. *Niche; policy.*
- **Digital signatures** — `Document.Signatures`. *Low priority.*

### D. Mail merge & data-driven generation

- **`Document.MailMerge`** — data source, merge fields, `Execute` to a new doc /
  printer / email. The canonical "generate N letters from a table" workflow —
  directly in wordlive's document-generation wheelhouse, but a large surface
  (data-source binding, field mapping, output routing). *Weight: high value,
  high cost — its own multi-step design pass.*

### E. Structured & data-bound content

- **Content-control *creation*** — today wordlive reads/writes existing CCs;
  `ContentControls.Add(type, range)` would let an agent *build* form-like docs
  (rich-text / dropdown / date / checkbox / repeating-section types), optionally
  XML-mapped. *Weight: medium-high for template generation.*
- **Custom XML parts / data binding** — `Document.CustomXMLParts`, binding CCs to
  XML nodes. Powerful for structured generation; heavier and niche. *Defer below
  CC creation.*
- **Legacy form fields** — `Document.FormFields`. *Superseded by content
  controls; skip unless asked.*

### F. Reference apparatus — extensions (after the Part-II footnotes/TOC/captions)

- **Index** — `Document.Indexes.Add`, `Range.MarkIndexEntry`. Back-of-book index.
- **Table of figures / authorities** — `TablesOfFigures`, `TablesOfAuthorities`.
  Consume the Part-II captions.
- **Citations & bibliography** — `Document.Bibliography`, `Sources`,
  `Range.InsertCitation`. Academic-writing workflow. *Weight: medium; cohesive
  once footnotes ship.*

### G. Drawing layer (floating shapes & embedded objects)

- **General `Document.Shapes`** — lines, connectors, freeform, grouping, z-order,
  rotation, fill/gradient/picture-fill, shadow/3D. The watermark/text-box
  flourishes (Part II) are the thin edge; the full drawing layer is large and
  floating-only (hard to read back, fights the anchor model). *Weight: low —
  cherry-pick specific shapes (callout line, divider) rather than the whole API.*
- **Embedded / OLE objects** — `InlineShapes.AddOLEObject` (embed an Excel range,
  a PDF, another doc). *Medium-niche; pairs with charts.*

### H. Range / navigation read surface (cheap, high agent value)

- **`Range.Information(wd…)`** — computed reads: page number, line number,
  in-table, end-of-row, zoom, caps-lock, … Lets an agent answer "what page is
  this on" *without* a snapshot. *Weight: high value, low cost.* ✅ **shipped
  v0.14.0** as `anchor.location()` (`page`/`end_page`/`line`/`column`/`in_table`).
- **Story ranges** — `Document.StoryRanges` / `StoryType` iteration (main text,
  footnotes, headers, text frames, comments as distinct stories). Makes
  find/read *story-aware* instead of main-text-only. *Weight: medium; sharpens
  existing reads.*
- **Range unit navigation** — `Expand`/`Collapse`/`MoveStart/End`,
  sentence/word/paragraph units. Building blocks for finer anchors. *Low,
  incremental.*

### I. Application / window / view control (politeness-sensitive — gate or read-only)

- **`Application.Options`** — spelling/grammar-check toggles, display settings.
  *Global app state; write only behind a loud opt-in.*
- **View / window** — `ActiveWindow.View.Type`, zoom, split, `Windows.Arrange`,
  `ActivePane`. Could help a vision-model workflow (set print-layout before a
  snapshot) but mutates the user's UI. *Weight: low; mostly let `snapshot`
  handle layout implicitly.*
- **`DisplayAlerts` / `ScreenUpdating`** — performance/quiet-mode toggles for
  long batches. *Internal optimization, not an agent verb.*
- **Printing** — `Document.PrintOut`, print settings. *Real side-effect; gate
  like persistence if ever added.*

### J. Read-only discovery collections (round out "what's in this doc")

wordlive now has `doc.lists` / `doc.tables` / `doc.images` / `doc.footnotes` /
`doc.endnotes` / `doc.revisions` / `doc.bookmarks.list()` (all shipped). The same
one-call-to-see-everything pattern is still missing for:

- **`doc.hyperlinks`** (hyperlink read-back — deferred with the shipped
  hyperlinks), **`doc.fields`** (every field + its code/type/locked state),
  **`doc.properties`** (group B). *Weight: each is cheap and additive; bundle as
  a "discovery surface" pass.*

---

# Cross-cutting work (any release)

- **Image-source hardening (read-side gate, pairs with Persistence).** ✅
  **Shipped 2026-06-09** with the persistence cluster. CLI/MCP now reject a
  non-local `--path` / op `path` (UNC `\\…`, `file://`, URL) via
  `_paths.reject_nonlocal_image_path` *before* the `is_file()` probe, plus an
  optional `--image-dir` / `WORDLIVE_IMAGE_DIRS` allowlist; the Python API stays
  ungated; base64/bytes unaffected. Refusals raise the shared
  `PathNotAllowedError`. Original notes below. Raised
  2026-05-31. `insert_image --path …` hands `FileName` to `AddPicture`, which
  resolves more than local paths — making the *path* input (not bytes/base64) an
  attack surface a prompt injection can reach. Threats: **UNC → NTLM credential
  theft** (sharpest — `image_on_disk`'s own `Path(p).is_file()` probe triggers
  SMB auth *before* COM), **URL → SSRF** (incidentally closed today), **local
  image-file disclosure**. Mitigation mirrors the save whitelist: Python API
  ungated; **CLI/MCP reject non-local `FileName`s (UNC `\\…`, URL/`file://`)
  before the `is_file()` probe**; optional `--image-dir`/`WORDLIVE_IMAGE_DIRS`
  allowlist; steer MCP toward base64/bytes; reuse the `PathNotAllowedError`
  denial type. `LinkToFile=False` ⇒ one-shot fetch at insert time (blast radius
  = the single call, not a stored link).
- **HRESULT coverage** — `_BUSY_HRESULTS` is a starter set; widen as real
  `pywintypes.com_error` HRESULTs surface in smoke runs.
- **Smoke fixtures** — a real `.docx` checked in with known bookmarks / CCs /
  headings / tables, so smoke tests have a known target.
- **Docs** — `spec.md` is the design doc; a `cookbook.md` of end-to-end LLM-tool
  examples is probably more useful than API reference at this stage. The 2026-05
  agent-test build (blank doc → styled multi-page catalogue) is a ready-made
  cookbook entry.
- **CLI surface shape — flat-first (decided 2026-06-08).** The CLI is
  intentionally flat: top-level verbs are the default, mirroring the flat `exec`
  op vocabulary and keeping `SKILL.md` a single-lookup list for agents.
  Noun-groups are the *exception*, justified only by a stable object/collection
  with ≥3 verbs that don't read naturally as standalone verbs (grandfathered:
  `table`, `list`, `comment`, `style`, `track`, plus the `read`/`write`
  dispatch groups). New single-verb capabilities stay top-level. Do **not**
  deepen the tree for tidiness — the LLM's primary surface is MCP's four
  dispatch tools, not `--help`, and flat verbs track the op names most directly
  (`insert-break` ↔ `insert_break`). **Explicitly rejected:** folding the
  `insert-*` family into an `insert` group (would break that op parallel for six
  commands). The one outstanding violation — bookmark ops split across three
  places — was cleaned up verb-first (see Part I, v0.13.0).

## Open papercuts (from the 2026-05 agent test)

- ~~**`heading:N` is mis-described in the agent guide.**~~ ✅ Fixed 2026-06-01.
  Both `_skill/wordlive-python/SKILL.md` and `_skill/wordlive-cli/SKILL.md` now
  describe it as "the Nth *paragraph*, which must be a heading — same index space
  as `para:N`", matching the library docstrings (the first heading is rarely
  `heading:1`; copy the id from `outline`).
- **Inline JSON vs. Windows paths.** `exec --ops '{…}'` with backslash paths
  mangles under PowerShell quoting + JSON escaping. Add a one-line guide note
  that path-bearing batches should use `--script FILE` or `--ops -` to dodge
  shell-escaping.
- ~~**Relative image paths fail.**~~ ✅ Fixed in v0.10.2 (`Path(p).resolve()`
  before COM).

## LLM-ergonomics feedback (2026-06-08, from a live MCP probe session)

A Claude-in-Claude-Desktop session driving v0.11.1 end to end surfaced eight
ergonomics items. **All but two shipped** — in-cell `find` boundary fix, tracked-
changes visibility, the numbered-list span path, `delete_paragraph` + control-char
message normalisation (v0.12.0); multi-paragraph block insert, inline runs, and
table-from-records (v0.13.0) — all now in **Part I**, with the full per-item
triage in `CHANGELOG.md`. The two still open (intra-batch output refs + durable
insert handles §5; revision accept/reject + revision-aware text reads §1c) are in
**Part II**.

## gpt-5.4 review (2026-06-09, MCP test-drive)

A gpt-5.4 session driving the MCP over a toy doc (add table + footnote + comment
in one atomic op, snapshot with markup, read back). Most of what it praised is
**shipped behaviour** — batch/single-undo, snapshot-with-markup, the read-after-
write loop, tables-from-records, the structure/layout split — so the review is
mostly validation. Wishlist triage:

- **Non-visual layout introspection** + **document stats** — its #1/#2 repeated
  ask; ✅ **shipped v0.14.0** as `anchor.location()` / `doc.stats()` (promoted
  from the Part III catalogue §H/§B, now in Part I).
- **Table-as-records read + update-by-key** — ✅ **shipped v0.14.0** as
  `Table.records()` / `append_record()` / `update_row()` (completes the v0.13.0
  tables-from-records write side; now in Part I).
- **Higher-level compose helpers** — its #10 ask; ✅ **shipped v0.14.0** as
  `insert_section` / `insert_markdown` / `replace_section_body` (a thin layer over
  `insert_block`; now in Part I).
- **Stable paragraph IDs** — its top ask. Investigated with a live probe
  (2026-06-09); Word's native `w14:paraId` proved unusable (lazy + COM-invisible),
  so the answer is the minted bookmark-backed handle already in Part II
  ("durable insert handles"), extended with a `pin`/`stamp` verb for *existing*
  paragraphs. Full design note under that Part II item.
- **Stale-anchor diagnostics / recovery hints** — new; captured alongside the
  paragraph-ID note in Part II.
- **Structural query helpers** (content-under-heading, block-between-headings,
  nearest-heading-before/after, find-paragraph-by-approx-text) — new backlog;
  several reduce to thin compositions over `outline` + `find`, so lower priority
  than the introspection reads. Not yet ticketed.
- **Style usage inventory** ("used vs. defined styles", "style near this anchor")
  — Part III "Styles deep cuts"; cheap, bundle if a discovery-surface pass lands.
- Already in backlog: accept/reject revisions + section revision summary (Part II
  revision write surface); image position / floating-vs-inline metadata
  (Part III §G image polish).
- **Noise:** the first pass's "batch denied by the tool approval layer" was MCP
  permission-gating working as designed, not a defect.
