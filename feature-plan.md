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

The detail below preserves the **load-bearing reference facts** (addressing
schemes, gotchas a future change must respect). Deeper deliberation lives in git
history and `spec.md`.

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
  **reads** (writes deferred — see Part II/III).

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
- MCP server: four dispatch tools + `wordlive://guide` resource (also fetchable
  as `word_read(command="guide")` since v0.11.0, because resources aren't
  surfaced by every client). Optional `mcp` extra.
- Python + CLI agent skills, `.mcpb` bundle, `install-mcp`, `examples/`.

---

# Part II — Approved / next up (priority order)

Everything here is **specced but not yet implemented** (verified against source
2026-06-01 — none of these entry points exist yet). Ordered by leverage. The
top cluster (publishing-quality) was added in the 2026-06-01 reorg; the rest
were approved 2026-05-31 and carry their original design passes.

## ★ Publishing-quality styling & layout (new cluster, 2026-06-01)

The audit surfaced that wordlive can *place* content but barely *style* it: there
is **no run-level character formatting at all**, styles are read-only, and page
layout is read-only. This cluster closes the "make it look designed" gap. Build
the prerequisite first, then in order.

> **Status (2026-06-07):** items **0–4 implemented** plus the **fields/page-number
> slice of item 5** — colour/units helper, `format_run`, shading/borders/
> tab-stops, style creation/modification (items 0–3, shipped in #12); PageSetup
> writes + multi-column (item 4); and `insert_field` / `insert_page_number` /
> `update_fields` (item 5, fields slice). See CHANGELOG `[Unreleased]`. **Remaining
> in item 5:** watermark, drop cap, text box / pull quote.

### 0. Color / units helper (prerequisite, tiny)

- One internal module that coerces colors (named → RGB → Word's **BGR long**;
  Word stores `RGB(r,g,b)` byte-swapped) and lengths (pt / in / cm → points).
  Underpins font color, highlight, shading, and borders — build once, reuse.
  Not exported through `__all__` until asked (mirrors the `WdStyleType` pattern).

### 1. `format_run` — direct character formatting (highest leverage)

- **`anchor.format_run(*, bold=None, italic=None, underline=None,
  strikethrough=None, font=None, size=None, color=None, highlight=None,
  subscript=None, superscript=None, small_caps=None, all_caps=None,
  spacing=None)`** on the base `Anchor` — mirrors `format_paragraph`'s tri-state
  (`None` = leave untouched). Maps to `Range.Font.*` + `Range.HighlightColorIndex`.
- exec op `format_run` (anchor_id + the kwargs); CLI `wordlive format-run`;
  MCP `word_write format_run`. Pairs with `range:START-END` for "bold *this*
  phrase" without restyling the paragraph.
- **Deferred:** kerning, character scale, animation/text-effects, color via
  theme-color index (vs. RGB).

### 2. Borders, shading & tab stops (the "structured page" layer)

- **Shading:** `anchor.set_shading(fill=…, pattern=None)` →
  `Range.Shading.BackgroundPatternColor`. Closes the v0.4 deferred "cell shading"
  (a `Cell` is an `Anchor`, so it comes for free) and paragraph/range shading.
- **Borders:** `anchor.set_borders(...)` over `Range.Borders` (and `Table.Borders`
  / `Section.Borders` for table and page borders) — sides, weight, color, style.
- **Tab stops:** `anchor.add_tab_stop(position, *, align="left", leader=None)` →
  `ParagraphFormat.TabStops.Add` — dot-leader rows, price lists, aligned columns
  without a table.
- New internal `WdColorIndex` / `WdLineStyle` / `WdTabAlignment` / `WdTabLeader`
  enum subsets.
- **Deferred:** diagonal cell borders, art page borders, per-side leader mixing.

### 3. Style creation / modification

- **`doc.styles.add(name, *, type="paragraph", based_on=None, next_style=None)`**
  → `Styles.Add(Name, Type)`, returning a now-**writable** `Style`; setters for
  `.font` / `.paragraph_format` / `.base_style` / `.next_paragraph_style`.
  Removes the "read-only" caveat on `_styles.py`. The brand/template primitive:
  define a house style once, then `apply_style` everywhere.
- exec op `add_style` / `set_style`; CLI `wordlive style add|set`; MCP mirror.
- **Decide during build:** how much of `Font`/`ParagraphFormat` to expose on a
  style (reuse the `format_run`/`format_paragraph` kwarg vocab) vs. a thin
  passthrough. **Deferred:** list styles, linked styles, style import from a
  template, `Document.UpdateStyles`.

### 4. PageSetup writes + multi-column (promoted from Deferred)

- **`Section.set_page_setup(*, margins=…, orientation=…, paper_size=…,
  columns=…, gutter=…)`** — the write mirror of the v0.6 `page_setup()` read.
  Margins/gutter in points (via the units helper), orientation/paper-size via
  `WdOrientation`/`WdPaperSize`, **columns** via `TextColumns.SetCount` /
  `.Add` (newspaper layout — a real publishing primitive, and the section
  half of the `insert_break(kind="column")` story).
- CLI `wordlive page-setup`; MCP mirror. **Decide:** document-scope vs.
  per-section scope; unit defaulting. **Deferred:** line numbers, vertical
  alignment, page color/background, different-first-page toggles.

### 5. Publishing flourishes (grab-bag, ship as cheap)

- **Page-number fields** — `Footer.insert_page_number()` /
  `Footers(…).PageNumbers.Add()` (or the general field primitive below).
  Today headers/footers take only literal text; *any* multi-page deliverable
  needs `{ PAGE }` / `{ NUMPAGES }`.
- **General `Fields.Add` primitive** — `anchor.insert_field(kind, text=None)`
  over `Range.Fields.Add` (date, page, ref, TOC, …). Generalizes the
  field-refresh machinery the TOC/cross-ref specs already need
  (`doc.update_fields()`), and underpins page numbers above.
- **Watermark** — `doc.set_watermark(text, …)` via the header-story
  `Shapes.AddTextEffect` (WordArt) — DRAFT/CONFIDENTIAL stamps.
- **Drop cap** — `paragraph.set_drop_cap(lines=3, position="dropped")` →
  `Paragraph.DropCap`. **Text box / pull quote** — `anchor.insert_text_box(text,
  …)` → `Shapes.AddTextbox`.
- These are individually small; bundle whichever land cheap, defer the rest.

## Image extraction — read images out for vLLMs

Spiked clean (2026-05-31): `Range.WordOpenXML` returns the range as Flat OPC,
inlining each referenced media part as base64 on a tight per-shape range — no
clipboard, no save-to-temp, no fragile position→media mapping, pure stdlib
(`xml.etree` + `base64`).

- **`anchor.read_image() -> (bytes, str)`** — bytes + MIME type (from the part's
  `pkg:contentType`). Parses the Flat OPC fragment, takes the sole `image/*`
  part, base64-decodes it.
- **`doc.images`** — read-only discovery collection (mirrors `doc.lists`/
  `doc.tables`); `list()` emits `[{index, anchor_id, mime, width, height,
  alt_text, para}]` so an agent sees what's there before pulling bytes.
- CLI `wordlive images` (list) + `read-image --anchor-id ID [--out FILE]`
  (`--out` writes bytes + reports `{path, mime, bytes}`; else base64 + mime in
  JSON). **No exec op** (extraction is a read, off the `doc.edit()` surface).
- **Decide during build:** addressing scheme (`image:N` 1-based over
  `InlineShapes` vs. resolving any single-image text anchor); base64-in-JSON vs.
  file-out CLI default.
- **Carry into spec:** `WordOpenXML` always serializes the full package skeleton
  (~64 KB floor — negligible once images are real-sized); rapid COM access can
  return `RPC_E_CALL_REJECTED` (already the `WordBusyError` retry class — retry
  for free). **Untested:** EMF/WMF vector images and cropped-image behavior —
  exercise both before shipping.
- **Deferred:** floating-shape / chart-image export; OLE-object extraction;
  whole-page-to-image (use `snapshot`).

## Reference apparatus — footnotes / endnotes, TOC

Note/field structures that attach to a range, so they fit the anchor model like
`insert_table`/`insert_image` did. Spike-confirmed 2026-05-31 (live Word).

> **Status (2026-06-07): shipped.** Footnotes/endnotes (`insert_footnote`/
> `insert_endnote`, `footnote:N`/`endnote:N` anchors, `doc.footnotes`/
> `doc.endnotes`) and the table of contents (`insert_toc`/`add_toc`, `Toc.update()`)
> all landed across the four surfaces. `doc.update_fields()` (the TOC/field
> refresh companion) shipped earlier with the fields slice. See CHANGELOG
> `[Unreleased]`. **Deferred:** custom marks, note separators, numbering
> format/restart, footnote↔endnote conversion; table of figures/authorities,
> custom TOC field codes, per-style level mapping.

### Footnotes & endnotes

- **`anchor.insert_footnote(text, *, where="after")` / `insert_endnote(...)`** —
  `Range.Footnotes.Add(Range, Reference="", Text=text)` (endnotes mirror).
  **Spike gotcha:** args must be **positional** — the `Text=` keyword is silently
  dropped under pywin32 late binding. Empty `Reference` auto-numbers (mark `\x02`).
- **`doc.footnotes` / `doc.endnotes`** — read-only discovery; `list()` emits
  `[{index, marker, text, para}]`.
- **Addressing:** `footnote:N` / `endnote:N`, 1-based, resolving to the
  **note-body range** (its own story, `StoryType == 2`; `note.Range.Text`
  round-trips) — so `set_text` edits the note and `.delete()` removes mark + note.
- exec ops `insert_footnote`/`insert_endnote`; CLI `insert-footnote`/`-endnote`
  + `footnotes`/`endnotes` lists; MCP mirrors.
- **Deferred:** custom marks, separators, numbering format/restart,
  footnote↔endnote conversion.

### Table of contents

- **`anchor.insert_toc(*, levels=(1,3), use_heading_styles=True, hyperlinks=True,
  where="after")`** + **`doc.add_toc(...)`** sugar (= `doc.start.insert_toc`).
  COM `TablesOfContents.Add(...)`. Spike-confirmed: builds + renders entries;
  page numbers populate after `Update()`.
- **Page numbers need repagination** — expose **`toc.update()`** and steer the
  guide to call it (or `snapshot`, which forces print layout) before reading
  page numbers.
- **`doc.update_fields()`** — companion (`Fields.Update()`); one verb for TOC,
  cross-refs, `{ PAGE }`. CLI `update-fields`; MCP `word_write update_fields`.
  (Generalized by the Part-II field primitive above.)
- exec op `insert_toc`; CLI `insert-toc`; MCP mirror.
- **Deferred:** table of figures/authorities, custom field codes, explicit
  per-style level mapping, index generation.

## Anchoring & linking — bookmark creation, hyperlinks, cross-references, captions

These cluster because **creating a named anchor is the prerequisite for the
rest** — cross-refs and internal hyperlinks both target a bookmark. Build in
this order. Spike-confirmed 2026-05-31.

> **Status (2026-06-07): shipped.** `doc.bookmarks.add`, `Anchor.link_to`,
> `Anchor.insert_cross_reference`, and `Anchor.insert_caption` all landed across
> the four surfaces and are covered by **live-Word smoke tests**. Live testing
> caught three issues the fake-based unit tests missed: bookmark cross-refs take
> the bookmark *name* (not an index) as `ReferenceItem`; `kind="text"` is invalid
> for footnote/endnote cross-refs (falls back to the note number); and
> `link_to(text=…)` now *inserts* linked text instead of overwriting the anchor's
> range. See CHANGELOG `[Unreleased]`. **Deferred:** hyperlink read-back/edit,
> cross-refs to list items/equations, caption numbering format + table of figures.

### Bookmark creation (first)

- **`doc.bookmarks.add(name, anchor)`** — `Range.Bookmarks.Add(Name, Range)` over
  the resolved anchor; `anchor` is an id or an `Anchor`. Validates `name` against
  Word's rules (letters/digits/underscores, leading letter, no spaces) → typed
  error before mutating. Closes the gap where `bookmarks[name]` could only
  read/write *existing* bookmarks.
- exec op `add_bookmark`; CLI `bookmark add NAME --anchor-id ID`; MCP mirror.

### Hyperlinks

- **`anchor.link_to(address=None, *, bookmark=None, text=None, screen_tip=None)`**
  — `address` XOR `bookmark` (external URL vs. internal jump). `text=None` links
  the existing range; `text=…` inserts new linked text.
- exec op `add_hyperlink`; CLI `link --anchor-id ID (--url U | --bookmark B)
  [--text T]`; MCP mirror. **Deferred:** `doc.hyperlinks` read-back;
  editing/removing a link.

### Cross-references

- **`anchor.insert_cross_reference(target, *, kind="text", hyperlink=True,
  where="after")`** — `target` is an anchor id we map: `bookmark:NAME`,
  `heading:N`, `footnote:N`/`endnote:N`. `kind` = text/page/number/above_below.
- **The mapping layer (spike-confirmed):** `InsertCrossReference`'s
  `ReferenceItem` is a 1-based index into `GetCrossReferenceItems(ReferenceType)`,
  which returns ordered **strings** — heading *text* / bookmark *names* — so
  `bookmark:NAME` → `items.index(NAME)+1` and `heading:N` → its ordinal among
  headings map cleanly. Unresolvable target → `AnchorNotFoundError` (exit 2)
  before mutating.
- **Gotcha:** `IncludePositionInformation=` as a keyword raises under pywin32 —
  omit or pass positionally. `ReferenceKind`→field-switch (`\p` for page) wants a
  build-time check.
- New internal `WdReferenceType`/`WdReferenceKind` enum subsets. exec op
  `insert_cross_reference`; CLI `cross-ref`; MCP mirror. **Deferred:** cross-refs
  to numbered-list items / equations; `IncludePositionInformation` combos.

### Captions (rounds out the figure/table workflow)

- **`anchor.insert_caption(label="Figure", *, text=None, where="after")`** →
  `Range.InsertCaption(Label, Title=text)`. Built-in Figure/Table/Equation or
  custom; auto-numbers per label. Pairs with cross-references. Ship if cheap,
  else fold into a follow-up. **Deferred:** numbering format/chapter-style;
  table-of-figures.

## Persistence — save / save-as / PDF export (directory-whitelisted)

wordlive's one coherent **filesystem boundary**: the **Python API is
trusted/ungated**; the **CLI/MCP surfaces are gated** because their input can be
prompt-injected. (The read side — image-source hardening — is bundled with this;
see Deferred/Cross-cutting.)

- **Python API (ungated):** `doc.save()` (errors if no path yet),
  `doc.save_as(path, *, fmt="docx", overwrite=False)`,
  `doc.export_pdf(path, *, from_page=None, to_page=None)`, `doc.saved`,
  `doc.path`.
- **CLI/MCP (default-deny + directory whitelist):** saving enabled only by
  configuring **allowed directories** (`--save-dir DIR` repeatable /
  `WORDLIVE_SAVE_DIRS`; MCP configured at launch). Empty whitelist = no saving
  (default). **Containment:** `Path(target).resolve()` *then* `is_relative_to`
  each whitelist dir (resolve first so `..`/symlink escapes can't slip out).
  Plain `Save()` is gated too (its target is the doc's existing path, which must
  also sit inside the whitelist).
- **`SaveNotAllowedError(WordliveError)` → exit 1** (policy denial / bad-input,
  **not** AnchorNotFound — six-code contract untouched). Likely renamed to a
  neutral `PathNotAllowedError` once it also guards image-source reads.
- **Not an exec op** (terminal side-effect, no undo). **Formats:** `docx` + `pdf`
  first; `doc`/`rtf`/`txt`/`html` deferred. Guide: saving is **off** unless an
  operator whitelists dirs; PDF export is the recommended "hand back a
  deliverable" path (pairs with `snapshot`).

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
  import from template, `UpdateStyles`. (Basic style *creation* promoted to
  Part II.)
- **Document themes** — `Document.DocumentTheme` / `ApplyTheme`, theme color/font
  schemes. Brand-consistency play; revisit after style creation lands.
- **Equations** — `Range.OMaths.Add` / build. Scientific-publishing niche;
  heavier than the Part-II flourishes.
- **Table polish** — merged/split cells (addressing assumes rectangular),
  `add_column`/`delete_column`, AutoFit fit-window/fit-content policy.
- **List polish** — custom list-template authoring, per-level bullet/number
  format, multi-section `LinkToPrevious` editing.
- **Comment/revision polish** — comment replies (`comment.reply`), per-revision
  accept/reject (`doc.revisions`), author/date filtering on `list()`.
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
  read; useful before/after an edit.*

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
  this on" *without* a snapshot. *Weight: high value, low cost.*
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

wordlive already has `doc.lists` / `doc.tables` / (planned) `doc.images` /
`doc.footnotes`. The same one-call-to-see-everything pattern is missing for:

- **`doc.hyperlinks`** (already noted under Part-II hyperlinks),
  **`doc.fields`** (every field + its code/type/locked state),
  **`doc.bookmarks.list()`** (currently name-indexed only),
  **`doc.revisions`** (already deferred under collaboration),
  **`doc.properties`** (group B). *Weight: each is cheap and additive; bundle as
  a "discovery surface" pass once a couple of the producing features land.*

---

# Cross-cutting work (any release)

- **Image-source hardening (read-side gate, pairs with Persistence).** Raised
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
items (`wordlive-llm-ergonomics-feedback.md`). Triaged into 0.12.0 vs. the next
cycle:

- ✅ **Shipped in 0.12.0.**
  - *In-cell `find` overran the cell boundary* (§2): the segmenter now strips the
    trailing `\r\x07` cell markers, so a cell-scoped (`in: table:N:R:C`) or
    whole-doc find resolves inside the cell instead of tripping the verification
    guard (`'Opus\r\x072'`). The `ReplaceVerificationError` message is reworded
    (it means the doc shifted under the match — an edit or tracked change — not
    specifically a cell).
  - *Tracked changes made the agent blind* (§1, partial): `doc.revisions`
    structured reader (+ CLI `revisions`, `word_read revisions`), `track` status
    over MCP, and `snapshot(markup="all")` to render the marks. The hard part —
    revision-aware *text* reads (§1c) — stays deferred (a tracked `find_replace`
    on the same paragraph still drifts; re-read between tracked edits).
  - *Numbered-list footgun* (§4): applying `numbered` over one `range:` spanning
    the paragraphs numbers them 1..N (already worked; now the documented, tested
    path). `continue_previous` can't repair an already-split list — that's a Word
    dead-end; remove + reapply over the span.
  - *`delete_paragraph`* (§6): removes a paragraph incl. its mark across all four
    surfaces. *Error messages leaked control chars* (§6) — boundary message
    normalised. *Batch anchor-resolution model* (§6) — documented as per-op live
    resolution in both SKILLs.
- ⏳ **Deferred to a later cluster.**
  - **Multi-paragraph block insert + inline runs** (§3): `insert_block` /
    `insert_paragraphs {items:[{text, style}]}` placing a contiguous run
    atomically, and inline run support in insert ops (`runs:[{text, bold?,…}]` or
    lightweight `**markup**`). The frequent "bold lead-in bullet" pattern.
  - **Intra-batch output refs + durable insert handles** (§5): let an op
    reference a prior op's output (`anchor_id: "$ops[0].table"`), and an optional
    `bind: "name"` that mints a bookmark on inserted content (a durable handle vs.
    fragile positional `para:N`).
  - **Revision accept/reject** and **revision-aware text reads** (§1c) — the read
    side shipped; mutating individual revisions stays on `.com` for now.
