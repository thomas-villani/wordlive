# wordlive — feature roadmap

How to read this document (refreshed 2026-06-17, after the post-polish roadmap brainstorm):

- **Status-first, not version-numbered.** Work is bucketed by **status** and,
  within "next up", by **priority** — never by speculative version labels (an
  earlier draft's `v0.9`–`v0.15` tranche numbers had decoupled from real
  releases and are gone). The authoritative release history is `CHANGELOG.md`;
  shipped items below carry the **real** release they landed in.
- Three parts: **I — Shipped** (a condensed index + a load-bearing-facts
  digest), **II — Approved / next up** (priority order), **III — Deferred &
  declined**. A short **cross-cutting** section and the **open papercuts** close
  it out.
- Ordering principle throughout: **LLM-agent leverage**, not spec order.
- **The detail lives elsewhere.** Per-feature narrative is in `CHANGELOG.md`;
  design deliberation is in git history and `spec.md`. Part I here keeps only the
  **load-bearing reference facts** — addressing schemes and live-Word gotchas a
  future change must respect.

---

# Part I — Shipped

Quick index (capability → real release):

| Capability | Release |
|---|---|
| Core: `attach`/`connect`, anchors, `doc.edit()` atomic undo, typed errors, CLI | v0.8.0 |
| `replace`/`go_to`/`exec` by anchor-id; `append`/`prepend`; `start`/`end` | v0.8.0–0.8.2 |
| Find/replace (fuzzy); styles (read/apply) + `format_paragraph` | v0.8.0 |
| Tables (read/edit: cells, rows, grid) | v0.8.0 |
| Comments, track changes, `RangeAnchor` | v0.8.0 |
| Lists / numbering; sections / headers / footers; `para:N` + cursor | v0.8.0 |
| Image **insertion** (path / bytes / base64, wrap modes) | v0.8.0 |
| Snapshots (render page/section to PNG); MCP server (4 `word_*` tools) | v0.9.0 |
| Table **creation / deletion**; page / column / section **breaks** | v0.9.0 |
| Python-API skill, MCP bundle, `install-mcp`, examples | v0.10.0 |
| Guide-as-tool-call, `status.saved`, exec `warnings`, paragraph-op split | v0.11.0 |
| find/replace boundary + cell fixes; inline-image `[image]` token; `insert_image(block=)` | v0.11.1 |
| Character formatting (`format_run`); borders / shading / tab-stops; style creation | v0.12.0 |
| PageSetup writes + multi-column; fields / page numbers; `update_fields` | v0.12.0 |
| Footnotes / endnotes; table of contents | v0.12.0 |
| Bookmark creation, hyperlinks, cross-references, captions | v0.12.0 |
| Tracked-changes visibility (`doc.revisions`, `snapshot(markup=…)`); `delete_paragraph` | v0.12.0 |
| Multi-page pagination (paragraph controls, repeating heading rows) | v0.13.0 |
| Image **extraction** (`read_image`, `image:N`, `doc.images`); low-res `max_dim` snapshots | v0.13.0 |
| Persistence (save / save-as / export-pdf, gated) + image-source hardening | v0.13.0 |
| Block insert + inline runs; tables from records; verb-first bookmark CLI | v0.13.0 |
| Non-visual layout introspection (`anchor.location()`, `doc.stats()`) | v0.14.0 |
| Table-as-records read/update (`records`/`append_record`/`update_row`) | v0.14.0 |
| Compose helpers: `insert_section`, `insert_markdown`, `replace_section_body` | v0.14.0 |
| Equations (`insert_equation` UnicodeMath/LaTeX/MathML; `doc.equations`, `equation:N`) | v0.14.0 |
| Document-info collections: `doc.properties` / `.variables` / `.hyperlinks` / `.fields` / `.proofing()` | v0.15.0 |
| `Table.autofit`; `drop_cap`; `line_spacing`; dedicated `Equation` paragraph style | v0.15.0 |
| Content-control **creation** (`insert_content_control`); back-of-book index; table of figures | v0.16.0 |
| Citations & bibliography (`doc.sources`, `insert_citation`, `insert_bibliography`); table of authorities | v0.16.0 |
| Document themes (`doc.theme` — apply / brand colours / fonts) | v0.16.0 |
| Durable handles (`doc.pin`/`stamp`, `pin:`, `pin_outline`, insert `bind`, `$ops[N]` refs) + stale-anchor hints | Unreleased |
| Revision write surface (`revision.accept`/`reject`, `revisions.accept_all`/`reject_all` w/ `within=`) | Unreleased |
| Revision-aware reads (`anchor.text_final`/`text_original`/`revision_segments`) | Unreleased |
| Watermark (`doc.set_watermark`/`remove_watermark`); text box / pull quote (`anchor.insert_text_box`) | Unreleased |
| Charts (Excel-backed: `anchor.insert_chart`, `doc.charts`, `chart:N`; bar/pie/line/scatter; `ExcelNotAvailableError`, exit 6) | Unreleased |
| Chart formatting & design (`ChartAnchor.format`/`set_axis`/`add_trendline`/`set_series_color`; `chart_style`/`has_legend` reads) | Unreleased |
| Chart depth (`ChartAnchor.format_series`/`add_error_bars`; `format` gap/overlap/data-table; `add_trendline` order/period) | Unreleased |
| Structural query helpers (`doc.between`, `doc.nearest_heading`, `doc.find_paragraphs`; content-under-heading already shipped) | Unreleased |
| Format read mirror (`anchor.format_info` — effective vs style, per-field `override`, `mixed` runs) | Unreleased |
| Linter + regularizer foundation (`doc.lint`/`doc.regularize`; 3 structural + heading/font/spacing consistency rules; targeted idempotent fixes; `regularize` exec op) | Unreleased |
| Linter Batch 1 — typography hygiene (10 P2 text-scan rules: trailing/leading/double-space, space-before-punct, hyphen-as-range, em-dash, tabs, manual-line-break, manual-heading-formatting, table-style-consistent; `typography` tag, off-by-default `Rule.default_on`; `find_replace` `literal`/`regex` modes + `required=False`) | Unreleased |
| Linter Batch 2 — finalization (6 P3 review-state rules: comments-present, unaccepted-revisions, track-changes-on, hidden-text-present, stale-fields, leftover-highlight; `finalization` tag, all off-by-default; `format_info` gains `hidden`/`highlight`) | Unreleased |
| Linter Batch 3 — field-code backbone (3 P1 `Range.Fields`-walk rules: broken-cross-reference + caption-manual-numbering on-by-default & tagged `academia`, page-numbers-present off/`layout`; all report-only) | Unreleased |
| Linter Batch 3b — `xref-as-literal-text` (a figure/table mentioned by literal number with no `REF` field; heuristic, off-by-default, tags `crossref`/`academia`; report-only) | Unreleased |
| Linter Batch 4a — house-style profile loader (`profile=` on lint/regularize; `Profile` in `_lint_profile.py`) + 3 policy rules (`body-justified`, `body-line-spacing`, `table-numeric-right-align`) fixing idempotently via `format_paragraph`; threaded Python/CLI/exec/MCP; `house_style` deferred | Unreleased |
| Floating-shape anchor model (`shape:N`: `doc.shapes`/`doc.text_boxes`; `ShapeAnchor` set_wrap/position/size/format/alt_text/text/replace_image/delete; `insert_text_box`+floating `insert_image` return it) | Unreleased |
| Shape depth + inline restyle + `textbox:N` (`ShapeAnchor` set_rotation/set_z_order/set_text_frame; `doc.group_shapes`/`ungroup`; `ImageAnchor` set_alt_text/set_size; `textbox:N` alias) | Unreleased |
| Checkpoint + diff (`doc.checkpoint`/`changes_since`/`diff`; `Checkpoint` token, `include=text/+style/+format`; content-aligned `replace`/`insert`/`delete`/`restyle`/`reformat` w/ current `para:N`; `doc_hash` fast-path) | Unreleased |
| Table styling & polish (`Table.set_style`/`set_alignment`/`set_borders`/`set_banding`, `Cell.set_vertical_alignment`; row/column anchors `table:N:row:R` [`RowAnchor`] / `table:N:col:C` [`ColumnAnchor`] + `Table.row`/`column`) | Unreleased |
| Markdown/HTML export + budgeted read (`doc.to_markdown`/`to_html(within=)` — flat node walk, GFM tables, `![alt](image:N)`, `[text](url)`; `doc.read(budget=,depth=)` — heading-spine digest, depth-weighted body elision, addressable markers; `_export.py`; `read markdown`/`html`/`digest`, MCP `to_markdown`/`to_html`/`digest`) | Unreleased |
| List polish (`anchor.apply_list_format(levels)` authors a custom multi-level list template — per-level number/bullet format, style, indent, marker font; `read_list_levels` read mirror; bullet=glyph+symbol-font not NumberStyle; `WdListNumberStyle`/`WdTrailingCharacter`; CLI `list format`/`list levels`, exec `apply_list_format`, MCP `list` action `format` / read `list_levels`) | Unreleased |
| Table structural polish (`Table.add_column`/`delete_column` mirror the row ops; `Cell.merge`/`split` — merged-cell story; `Table.is_uniform` + physical-cell `read()`/`grid()`; `delete_column` raises clean `OpError` on a merged table; CLI `table add-column`/`delete-column`/`merge-cells`/`split-cell`, exec ops, MCP) | Unreleased |

## Load-bearing reference facts

The addressing scheme and the constraints every change must respect, distilled
from the shipped clusters above.

### Anchor-id taxonomy (the stable, LLM-visible addressing scheme)

`heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`, `table:N:R:C`,
`table:N:row:R`, `table:N:col:C`, `range:START-END`, `header:S:WHICH` /
`footer:S:WHICH`, `footnote:N`, `endnote:N`, `image:N`, `equation:N`, `chart:N`,
`shape:N`, `pin:CODE`, `start`, `end`. One resolver: `doc.anchor_by_id(id)`. A
malformed scheme (`banana:7`) reports "unknown anchor type".

- **`shape:N`** is positional over the document's **body-story floating shapes**
  (`Document.Shapes`, document order) — text boxes, **floating images**
  (post-`ConvertToShape`), and WordArt — header-story watermarks excluded. It's the
  restyle handle `insert_text_box` and a floating `insert_image` now **return**
  (inline `insert_image` stays `image:N` / returns `None`). `ShapeAnchor` carries
  `shape_type` and the in-place mutators (`set_wrap`/`set_position`/`set_size`/
  `format`/`set_alt_text`/`set_text`/`replace_image`/`delete`). `doc.shapes` is the
  full collection; `doc.text_boxes` is the text-box subset (a discovery filter, each
  keeping its canonical `shape:N` id); `textbox:N` is an addressing alias onto a text
  box's canonical `shape:N`. Positional ⇒ renumbers when a shape is added/removed —
  and a **`set_z_order` restack also renumbers** (`Document.Shapes` orders by
  z-order); re-list, don't cache (the `image:N`/`chart:N` rule). Depth knobs:
  `set_rotation`/`set_z_order`/`set_text_frame`; `doc.group_shapes(*ids)` collapses
  floats into a `group` shape (enables `AllowOverlap` first), `ShapeAnchor.ungroup`
  reverses it. **No autosize** (Word's resize-to-fit-text doesn't set over COM).
  Inline `image:N` gains `set_alt_text`/`set_size` (re-wrap still floats it). Helpers
  live in `_shapes.py` (mirrors `_charts.py`).

- **`chart:N`** is positional over the document's chart inline shapes
  (`HasChart`), in document order — it renumbers when an earlier chart is
  inserted; re-list (`doc.charts`), don't cache. Metadata only: `chart_type` /
  `title` / `chart_style` / `has_legend`. It's also the **write** target for
  formatting (`format`/`set_axis`/`add_trendline`/`set_series_color`/
  `format_series`/`add_error_bars`). Reading chart *series data* back is deferred
  (and the data link is broken at insert, so it's static anyway).

- **`pin:CODE`** is the **durable handle** (Unreleased): `doc.pin(anchor)` /
  `stamp` plants a Word-hidden bookmark `_wl_<code>` over a range and returns
  `pin:<code>` (random hex, or a `name=` slug stored `_wl_<slug,-→_>`). Word keeps
  the range association across edits — the actual source of durability — so it's
  the escape hatch for positional `para:N`/`heading:N` that renumber. `pin:`
  reuses the `Bookmark` class (a `_pin_code` makes `anchor_id` report `pin:` not
  `bookmark:`); minting bypasses `_validate_bookmark_name` (which forbids the
  leading `_`). `doc.pin_outline(levels=…)` bulk-pins headings, idempotent by
  range start. Insert ops take `bind:"slug"`; any op field `$ops[N].field` is
  substituted with an earlier op's output before it runs. A stale positional miss
  now carries a recovery `hint` (out-of-range vs not-a-heading, nearest heading).

- **`para:N` and `heading:N` share one index space** — a heading is both; copy
  the id from `outline` (the first heading is rarely `heading:1`). Word's
  `OutlineLevel` is `10` for all non-headings, so `paragraphs` reports each
  paragraph's **style name** to distinguish list item from body.
- **`table:N:R:C`** addresses a cell; **bare `table:N` is not an anchor** (a
  whole table is a collection — `doc.tables[N]` / the `table` CLI group). A
  `Cell` **is** an `Anchor` (inherits `apply_style`/`format_paragraph`/`set_text`
  + borders/shading + `set_vertical_alignment`); cells don't appear in
  `doc.outline()`. Whole-table restyle / alignment / grid-borders / banding live on
  the `Table` wrapper (`set_style`/`set_alignment`/`set_borders`/`set_banding`), not
  as anchors.
- **`table:N:row:R`** (`RowAnchor`) and **`table:N:col:C`** (`ColumnAnchor`) are
  styling handles for a whole row / column — both `Anchor`s, so the shipped
  `shading`/`borders`/`apply-style`/`format-run` verbs (and `set_shading`/
  `set_borders` ops) style the strip in one call; `Table.row(R)`/`Table.column(C)`
  return the same objects. A **row** is a contiguous `Rows(R).Range`. A **column**
  is **not** — `Column.Range` is absent under late binding and the whole `Columns(C)`
  collection raises **"mixed cell widths"** (`0x80020009`) on a merged / irregular
  table — so `ColumnAnchor` fans the op across `Columns(C).Cells` and re-raises that
  COM error as an `OpError` pointing at per-cell `table:N:R:C` styling
  (live-probed 2026-06-20). Restyle gotchas (live-probed 2026-06-20): `Table.Style =
  X` **overwrites direct cell shading** (restyle first, then cell overrides);
  `Cell.VerticalAlignment` takes 0/1/**3** (2 = `wdAlignVerticalJustify`, invalid for
  a cell, raises); the six banding `Table.ApplyStyle*` booleans read back as real
  Python bools and only show once a real table style is applied.
- **`range:START-END`** is what `find()` emits *and* what `anchor_by_id`
  resolves — a find hit feeds straight into `replace` / `comments.add` /
  `format_run`.
- **`header:S:WHICH` / `footer:S:WHICH`** — WHICH ∈ primary/first/even; a
  `HeaderFooter` **is** an `Anchor`. Fields inserted here land in the
  header/footer story, so page numbers work.
- **`footnote:N` / `endnote:N`** resolve to the **note-body range** (its own
  story, `StoryType == 2`); `set_text` edits the body, `delete()` removes mark +
  body.
- **`image:N`** is 1-based over `InlineShapes`; **`equation:N`** is positional
  over `OMaths` — both **renumber** when an earlier one is inserted; re-list,
  don't cache.
- **`cursor`** is deliberately *not* an `anchor_by_id` scheme. `cursor write`
  opts into `EditScope.allow_cursor_move()` — the one op that intentionally moves
  the user; `cursor read` reports the containing `para:N`.

### `exec` batch semantics

`exec --script ops.json` / `--ops -` batches N ops in one `doc.edit()` /
`UndoRecord` → one Ctrl-Z reverts the whole intent. **Failure:** if op K fails,
ops 1..K-1 are already applied; the UndoRecord closes and failure is reported
(one Ctrl-Z reverts the partial work). A successful batch carries an `outputs`
array (per structure-creating op) and a `warnings` array (fields an op ignored).
Inline `exec --ops '{…}'` with backslash paths mangles under PowerShell + JSON
escaping — path-bearing batches should use `--script FILE` or `--ops -`.

### Live-Word gotchas (hard-won; don't re-learn these)

- **Direct-override detection = effective vs `Range.ParagraphStyle` baseline**
  (the linter's substrate, live-validated 2026-06-19). `anchor.format_info()`
  compares `Range.Font`/`ParagraphFormat` (the rendered values) against
  `Range.ParagraphStyle.Font`/`.ParagraphFormat` (the applied style's resolved
  baseline — built-in styles resolve concrete values, no `BaseStyle` walk
  needed); they differ ⇒ a direct `override`. A font field reading **`9999999`
  (`WD_UNDEFINED`)** means the property *varies across runs* (mixed) — surfaced in
  `font.mixed`, never as a number. `regularize`'s default fix writes the **style's
  own value back as a direct property** (not `Font.Reset()`), which is why a
  second pass is a no-op — the idempotency invariant. Note Word 16's default
  `Heading 1` is **20pt with KeepWithNext on**, so `heading-keep-with-next` won't
  fire on a clean default heading (no false positive).

- **Tracked text is the FINAL view in `Range.Text`** (live-probed 2026-06-16):
  Word's `Range.Text` returns the *accepted* text — inserted runs present,
  **deleted runs absent**. The deleted characters survive only on the delete
  `Revision` (`.Range.Text`), and revision offsets are reported in a **markup**
  coordinate space where the deleted text still occupies positions (so a delete
  revision's range overlaps real unchanged text in the final stream). So
  reconstructing the *original* (reject-all) view can't filter `Range.Text` — it
  must walk the markup space and splice the deleted text back in from each delete
  revision. This is what `_revisions.segment_runs` does (`accept`/`reject` on a
  single `Revision` renumbers the rest; `Range.Revisions.AcceptAll`/`RejectAll`
  scope the bulk op). An **anchor's range is literal** — `accept_all(within=heading)`
  covers only the heading line, not its section body (use a range/paragraph that
  actually spans the revisions).
- **Floating-shape model (`shape:N`) — live-probed 2026-06-19.** `Shape.Type`
  discriminates the kind cleanly under late binding: picture = `13`, text box =
  `17`, WordArt = `15` (`MsoShapeType`). **Replacing a floating picture's image is
  delete + reinsert at the same anchor** (preserving wrap/position/size): probed
  `Shape.Fill.UserPicture` only *overlays* a second picture-fill on a picture shape
  (the range then holds **two** images) — it is **not** a true replace, so it's
  rejected. The `Shape.Left = wdShapeCenter` / `RelativeHorizontalPosition` position
  constants (the ones `set_watermark` uses on a *header* shape) work the same on a
  **body** shape, and `Shape.Anchor.StoryType` reads cleanly (`== 1` for the main
  text story — the body-vs-header guard). Watermarks are WordArt in the *header*
  story, so `Document.Shapes` (body) excludes them for free; `body_shapes` keeps a
  name-prefix + story-type guard anyway. A just-inserted shape is located by a unique
  temp `Shape.Name` (don't assume "last" — other floats can reorder). All restyle
  verbs operate on the `Shape` object directly (no `.Select()`), so they're polite
  (selection/scroll unmoved); the live smoke confirmed atomic single-undo too.
- **Shape depth + group/ungroup — live-probed 2026-06-19 (Slice 2).** `Shape.Rotation`
  (absolute degrees), `Shape.ZOrder(MsoZOrderCmd)` / `ZOrderPosition`, and
  `TextFrame.MarginLeft/Right/Top/Bottom` + `WordWrap` all set cleanly under late
  binding. **`ZOrder` renumbers `shape:N`** — `Document.Shapes.Item(i)` orders by
  z-order, so a restack moves the shape to a different `Item` index (re-list after
  `set_z_order`). **Grouping needs `WrapFormat.AllowOverlap = True` on each member**
  first (else `Shapes.Range([names]).Group()` raises "Grouping is disabled"); members
  must be named (live Word auto-names). `Shape.Ungroup()` returns a ShapeRange —
  capture child names from `GroupItems` *before* dissolving, then re-locate each
  `shape:N` by name. **Autosize is unsettable** and intentionally omitted:
  `TextFrame.AutoSize` silently no-ops, `TextFrame2.AutoSize` raises "value out of
  range". Inline `image:N` shares `Width`/`Height`/`LockAspectRatio`/`AlternativeText`
  with floating shapes, so `set_size`/`set_alt_text` reuse the same `_shapes` helper.
- **Floating shapes live on the HeaderFooter / Document, not on `Range`.** A
  `Range` has **no** `.Shapes` (an AttributeError); the watermark draws into
  `Section.Headers(wdHeaderFooterPrimary).Shapes.AddTextEffect(...)` (WordArt,
  named `PowerPlusWaterMarkObject<N>` to match — and replace — Word's own
  watermark feature), and a text box is `Document.Shapes.AddTextbox(..., Anchor=range)`.
  `Shape.WrapFormat.Type = behind` floats a watermark behind body text;
  `Shape.Left/Top = wdShapeCenter (-999995)` with `RelativeHorizontal/VerticalPosition
  = margin` centres it.
- **Styles** are writable since v0.12.0 (`doc.styles.add(...)` →
  `style.format_run`/`format_paragraph`); read-only before that. `StyleNotFound`
  subclasses `AnchorNotFound` → reuses exit 2.
- **`insert_image` `wrap` is required, no default** (`inline|auto|square|tight|
  through|top-bottom|behind|front`). Non-local `--path` (UNC/`file://`/URL) is
  rejected on CLI/MCP before the `is_file()` probe (NTLM-theft / SSRF guard);
  Python API ungated; base64/bytes unaffected. Bad source → `ImageSourceError`
  (exit 1, not AnchorNotFound).
- **Adjacent tables silently merge** if they touch with no paragraph mark
  between; `insert_table` probes `Information(wdWithInTable)` and drops a
  separator only where it abuts an existing table.
- **Final paragraph is undeletable** — find/replace clamps off Word's terminal
  mark; `add_table`/`create_table` at `end` opens a trailing paragraph first;
  composing at `doc.end` detects the terminal mark so a block doesn't fuse into
  the last paragraph.
- Inline images read back as a **`[image]` token** (not a phantom control char).
  `read_image` goes through **`Range.WordOpenXML`** (Flat OPC, base64 inline; no
  clipboard/temp) — a ~64 KB package-skeleton floor, and rapid COM access can
  return `RPC_E_CALL_REJECTED` (the `WordBusyError` retry class).
- **Footnote/endnote `Add` args must be positional** — the `Text=` keyword is
  silently dropped under pywin32 late binding.
- **TOC / index / table-of-figures page numbers populate only after
  repagination** — call `.update()`, `doc.update_fields()`, or take a `snapshot`.
  `location()` / `stats()` **repaginate first** (content-neutral) so page/line
  numbers are print-layout truth.
- **Cross-references:** a bookmark cross-ref takes the bookmark **name** (the
  mapping layer is a 1-based index into `GetCrossReferenceItems`, ordered
  strings); `kind="text"` is invalid for footnote/endnote refs;
  `IncludePositionInformation=` as a keyword raises — omit/positional.
- **Captions** land in their own `Caption`-styled paragraph; a **cell** anchor
  captions the **whole table**. Convention: table caption **above**, figure
  **below**.
- **Persistence** is wordlive's one filesystem boundary: Python API
  trusted/ungated; CLI/MCP default-deny + directory whitelist (`--save-dir` /
  `WORDLIVE_SAVE_DIRS`), containment via `resolve()` *then* `is_relative_to`.
  `save_as` writes `.docx` only (use `export_pdf` for PDF). Not exec ops
  (terminal side-effects, no undo).
- **Compose helpers** parse a **constrained-Markdown subset**, not CommonMark
  (`#`/`##`/`###`, `-`/`*`, `1.`, blank-line text; inline `**bold**`/`*italic*`);
  unrecognised lines stay literal. Plain `insert --text` is always literal —
  markdown is an opt-in verb.
- **Content controls** are addressable as `cc:TITLE` (falls back to `tag`);
  `kind` ∈ rich_text/text/picture/combo_box/dropdown/date/checkbox/
  building_block/group/repeating_section.
- **Index:** `Indexes.Add` `HeadingSeparator` must be the enum (`0`), not `""`,
  on a makepy-typed build. **Table of figures:** `TablesOfFigures.Add` needs
  **keyword** args for the flags (its optional string Variants reject positional
  `""`).
- **Table of authorities:** there is **no** `Range.MarkCitation` — mark via a raw
  `TA` field (`TA`=field type 74, `TOA`=73); `TablesOfAuthorities.Add` takes
  `Range, Category(int)` positional + keyword separators; Word's TOA has **no**
  `UpdatePageNumbers` (`.update()` only).
- **Citations & bibliography:** `CITATION`=field type 96, `BIBLIOGRAPHY`=97 (not
  119/34) — insert via the EMPTY raw-code path. `Sources.Add` ingests a single
  `<b:Source>`; `BibliographyStyle` is a plain string (APA/MLA/Chicago/IEEE/
  Turabian ok; GOST/ISO690 build-dependent).
- **Hidden bookmarks need `Bookmarks.ShowHidden`.** Word omits
  leading-underscore bookmarks (its own `_Toc`/`_Ref`, and wordlive's `_wl_`
  pins) from `Document.Bookmarks` *iteration* unless the collection's
  `ShowHidden` flag is set — but `Exists(name)` / `Bookmarks(name)` find them
  regardless. Enumerating pins (e.g. `pin_outline` idempotency, `list(
  include_hidden=True)`) must flip `ShowHidden=True` for the read and restore it
  (the `_bookmarks_including_hidden` helper). The fake COM fixture *does* yield
  hidden ones, so this is a smoke-only failure mode — caught live, not in units.
- **Themes:** Office 16 has **no `RemoveDocumentTheme`**; `.RGB` is a BGR
  `OLE_COLOR` int (12 friendly colour slots via `to_bgr`/`bgr_to_hex`).
- **Colours/units** go through the internal `_format.py` helper (colours →
  byte-swapped BGR long; lengths pt/in/cm/mm → points); bad colour/length/enum →
  `OpError` (exit 1).
- **Equations** land on a centred `Equation` paragraph style (created on first
  use) so they don't inherit a neighbouring heading's style and pollute the
  outline. LaTeX is the optional `latex` extra; `.mathml` round-trips via
  Office's own XSLT.
- **Charts are Excel-backed and the embedded-Excel lifecycle is the whole game**
  (live-probed 2026-06-16). `Range.InlineShapes.AddChart2` raises "Requested
  object is not available" — it **only works off `Selection`** (`doc.Range(pos,
  pos).Select()` then `Application.Selection.InlineShapes.AddChart2(-1, type)`;
  `doc.edit()` restores the user's selection). Populate by writing the data into
  the embedded workbook's **cells** and binding a `=SERIES(name, x, y, 1)`
  **formula string** — the `Series.XValues`/`.Values` array setters are flaky
  under pywin32 late binding ("Property XValues can not be set"), and a literal
  inline array stores x as **text**, so a scatter plots at category positions
  1,2,3 instead of its numeric x. **`ChartData.BreakLink()` before
  `Workbook.Close()`** is mandatory: touching `.Chart` opens an embedded-Excel
  *data grid* that is **global, persistent state surviving `doc.Close`** — leave
  it open and the hidden Excel orphans, and every later insert fails "the chart
  data grid is already open in DocumentN". BreakLink severs Word's link (data
  goes static) so the workbook closes and the Excel terminates; quit it only when
  `Workbooks.Count == 0`. Don't read a chart's *series* data back (it
  destabilises Word — RPC failures/crashes); metadata reads touch only
  `ChartType` / `ChartTitle`. The Excel-availability gate is a `winreg`
  `HKEY_CLASSES_ROOT\Excel.Application` lookup (`pythoncom.CLSIDFromProgID`
  doesn't exist in this pywin32) — non-invasive, never launches/disturbs a user's
  Excel.

### Exit codes (CLI)

`0` ok · `1` other/bad-input (incl. `ImageSourceError`, `SnapshotError`,
`PathNotAllowedError`, `DocumentNotFoundError`) · `2` anchor/style not found or
zero `find` matches · `3` Word busy (retryable) · `4` Word not running · `5`
ambiguous `find` match · `6` Excel not available (`ExcelNotAvailableError`, for
`insert-chart`).

---

# Part II — Approved / next up (priority order)

Everything here is **specced but not yet implemented**. Ordered by leverage. The
✅ blockquotes immediately below shipped since this section was first written (now
indexed in Part I); the live backlog — the **post-polish brainstorm wave**
(2026-06-17) — follows them.

> **Durable handles & stale-anchor diagnostics — ✅ shipped (Unreleased).** All
> five pieces landed: `doc.pin`/`stamp` + the `pin:` anchor, `doc.pin_outline` /
> `outline(pin=True)`, insert-op `bind:"name"`, `$ops[N].field` references, and
> stale-anchor recovery hints. See Part I's load-bearing facts (the `pin:`
> taxonomy entry and the `Bookmarks.ShowHidden` gotcha) and `CHANGELOG.md`. Native
> `w14:paraId` was rejected (live probe 2026-06-09): assigned lazily and
> COM-invisible (`Range.WordOpenXML` strips it), so minted bookmarks are the only
> mechanism that survives the edits we care about.

> **Revision write surface — ✅ shipped (Unreleased).** Both pieces landed:
> per-revision `accept()`/`reject()` + whole-doc/anchor-scoped `accept_all`/
> `reject_all(within=…)`, and the revision-aware reads
> `Anchor.text_final`/`text_original`/`revision_segments()`. See Part I's
> load-bearing facts (the tracked-text representation entry) and `CHANGELOG.md`.

> **Publishing flourishes (floating-shape remainder) — ✅ shipped (Unreleased).**
> `doc.set_watermark`/`remove_watermark` (header-story WordArt) and
> `anchor.insert_text_box` (floating `Shapes.AddTextbox`). The publishing-quality
> cluster (character/paragraph formatting, borders/shading/tab-stops, style
> creation, PageSetup writes, fields/page numbers, pagination controls, drop cap
> — v0.15.0) is now complete.

> **Charts (Excel-backed) — ✅ shipped (Unreleased).** `anchor.insert_chart(kind,
> data, *, title=None)` (bar/pie/line/scatter) → `ChartAnchor` (`chart:N`);
> `doc.charts` read collection; the `ExcelNotAvailableError` gate (CLI exit 6).
> `data` grew a second shape over the original "flat mapping" spec — an array of
> `[x, y]` pairs — so **scatter is first-class** (numeric axes, duplicate x). See
> Part I's load-bearing facts (the `chart:N` taxonomy entry and the Excel-backed
> live-Word gotcha — `Selection`-only insert, SERIES-formula population, the
> mandatory `BreakLink`-before-close to avoid orphan Excel) and `CHANGELOG.md`.

> **Chart formatting & design — ✅ shipped (Unreleased).** A curated formatting
> surface on `ChartAnchor`: `format(...)` (title/legend/`chart_style`/fills/font/
> data-labels/`chart_type`), `set_axis(which, ...)` (title/min/max/`scale=log`/
> number-format/gridlines), `add_trendline(...)` (linear…power + equation/R²), and
> `set_series_color(color, *, series, point)`. Read side gains `chart_style` /
> `has_legend`. All operate on the **post-insert static chart with no Excel respin**
> (live-re-probed 2026-06-17: 0 orphan EXCEL.EXE across the surface) — so these
> verbs need **no Excel gate**. Wired across CLI/exec/MCP. **Deferred as planned:**
> multi-series authoring, secondary axes, error bars, 3-D, `ApplyChartTemplate`
> (`.crtx`), exposing `BreakLink` as a policy knob, and reading existing charts'
> *data* back out (`doc.charts` stays metadata-only). One probe note: `Axis.Visible`
> is not settable under late binding, so axis show/hide is not exposed.

> **Structural query helpers — ✅ shipped (Unreleased).** The last Part II item.
> Three pure doc-level reads composing over `outline` + `find`:
> `doc.between(start, end, *, inclusive=False)` (a `RangeAnchor` between two
> anchors — the block between two headings), `doc.nearest_heading(where, *,
> direction="before"|"after")` (the enclosing/preceding or next heading nearest a
> position), and `doc.find_paragraphs(text, *, limit, min_score)` (**fuzzy**
> paragraph ranking via `difflib.SequenceMatcher` over `find()`'s normalization —
> the typo/paraphrase-tolerant counterpart to the exact-substring `find`, returning
> ranked `para:N`). Wired across CLI (`read between`, `read nearest-heading`,
> top-level `find-paragraph`) and MCP `word_read`. **Content-under-heading was
> already shipped** (`Heading.section_range()`/`section_text()` + `read section`),
> so these three complete the item. Pure reads — no `exec` ops.

## Priority 1 — Document quality & change tracking (the agent-publishing core)

### 1. Linter + formatting regularizer — `doc.lint()` / `doc.regularize()`

> **Foundation slice — ✅ shipped (Unreleased, 2026-06-19).** `anchor.format_info()`
> (the read mirror + direct-override detection), the three structural rules
> (`heading-keep-with-next`, `table-repeat-header`, `list-numbering-continuity`),
> the heading/font/spacing consistency rules + report-only `mixed-run-format`,
> and `doc.lint()` / `doc.regularize()` with the **targeted, idempotent** default
> fix — wired across Python / CLI (`lint`, `regularize`, `read format`) /
> `regularize` exec op / MCP, and live-validated.
>
> **Batch 1 — typography hygiene ✅ shipped (Unreleased, 2026-06-24).** The §5b·A
> text-scan cluster + `manual-heading-formatting` / `table-style-consistent` — 10
> rules in `_linting_typography.py` (6 on, 4 opinionated off-by-default behind the
> `typography` tag). Two enablers landed: `find_replace` gained **`literal` /
> `regex` modes** + `required=False` (the fix path — fuzzy find_replace can't
> express a literal whitespace edit, so the typography fixes are regex-mode
> `find_replace` ops scoped to a `para:N`, re-scanning live text), and `Rule`
> gained **`default_on`** for the default-off rules. Deferred to **1b**:
> `straight-quotes` / `nbsp-missing` / `sentence-spacing-consistent` (heuristic-heavy).
>
> **Batch 2 — finalization ✅ shipped (Unreleased, 2026-07-01).** The §5b·G
> "is-this-ready-to-send?" cluster — 6 rules in `_linting_finalization.py`, all
> **off by default** behind the `finalization` tag (`comments-present`,
> `unaccepted-revisions`, `track-changes-on`, `hidden-text-present`, `stale-fields`
> report-only; `leftover-highlight` the one fix). Reused the shipped revision/
> comment/field wrappers + `doc.track_changes`; added `format_info`'s `hidden` /
> `highlight` fields (the read mirror of `format_run`'s writes). `stale-fields` is a
> report-only nudge — no COM staleness flag (Batch 3 confirmed this is inherent, so
> it stays a nudge rather than an IOU).
>
> **Batch 3 — field-code backbone ✅ shipped (Unreleased, 2026-07-02).** The §5b·C
> cross-reference/caption cluster, built on a `Range.Fields` walk — 3 rules in
> `_linting_fields.py`. `broken-cross-reference` + `caption-manual-numbering` ship
> **on** by default (both also tagged `academia`); `page-numbers-present` is **off**,
> tag `layout`. All three report-only (their fixes add content or need target
> matching → deferred with the `adds_content` gate). No new COM write surface.
>
> **Batch 3b — `xref-as-literal-text` ✅ shipped (Unreleased, 2026-07-02).** The
> heuristic rule Batch 3 deferred (a figure/table mentioned by literal number in body
> text with no `REF` field) — added to `_linting_fields.py`, report-only, **off by
> default** (tags `crossref` / `academia`; a bare "Table 2" in prose is often
> legitimate). Caption/heading paragraphs skipped.
>
> **Batch 4a — profile loader + first policy rules ✅ shipped (Unreleased, 2026-07-02).**
> The §6 policy half: `profile=` on `lint`/`regularize` resolved by a new `Profile`
> (`_lint_profile.py`), and `body-justified` / `body-line-spacing` /
> `table-numeric-right-align` in `_linting_policy.py`, all fixing idempotently via
> `format_paragraph`. `Rule.check` widened to `(doc, span, profile)`; threaded through
> Python / CLI (`--profile`) / exec / MCP. `house_style` (§6's consistency-target pinning +
> `set_style`) deferred.
>
> **Still open (a follow-up pass):** **Batch 4b** — the §H/§I detection rules (hyperlinks,
> confidentiality/copyright notices, doc-properties, draft-watermark), some profile-driven;
> the `house_style` half of profiles; the opt-in `Font.Reset()` strip-to-style fix; the
> content-adding fixes (`stray-empty-paragraph`, `figure-caption-present` — these want the
> deferred `adds_content` Finding field); and the `docx-plus` cascade-provenance hybrid. A
> **v2 rule backlog** (~40 more rules for publishing/academia, primitive-driven
> batches: Batch 2 finalization ✅ · Batch 3 field-code backbone ✅ · Batch 3b
> heuristic xref-as-literal-text ✅ · Batch 4a profile loader + policy rules ✅ · Batch 4b
> layout/notices) — see `spec-linter.md` **§5b**. See also `spec-linter.md` (§6,
> §7c, §9) and `CHANGELOG.md`.

The highest-utility next feature: a declarative rule set that **audits** a document
for publishing-quality defects and **autofixes** the mechanical ones. Pure
composition over shipped write primitives (`format_paragraph`, `set_heading_row`,
styles, `fields`, `outline`) — **no new COM surface**. Seed the rule catalogue from
the `all2md` linter + the recurring by-hand edits. **Detailed design:
`spec-linter.md`** (rule catalogue, the consistency/structural/policy split,
profiles, and the direct-override detection probe).

- **`doc.lint(rules=None, within=anchor)`** → ranked findings
  `{rule, severity, anchor_id, message, fixable, fix_hint}`. Read-only.
- **`doc.regularize(rules=None, within=anchor, dry_run=False)`** → applies the
  fixable subset in one `doc.edit()` (atomic undo); returns the change list.
- **Seed rules:** heading not `keep_with_next` (dangling header at a page foot);
  missing widow/orphan control; a table broken across pages with no repeating
  heading row (`set_heading_row`); inconsistent line spacing / direct formatting
  vs. the paragraph style; split/orphaned list numbering (the "N independent 1.
  lists" footgun); empty heading polluting the outline/TOC; broken
  cross-ref / bookmark / field; mixed body fonts; missing image alt text; stray
  double-spaces / empty paragraphs.
- Rules are **declarative + individually selectable** (id + severity +
  enable/disable) so a caller can audit → review → apply, or later wire
  `regularize` into the save hook (Priority 7).
- Surfaces: Python (`doc.lint`/`doc.regularize`); CLI (`lint` JSON findings /
  `regularize`); `regularize` exec op (a write, for the atomic-undo batch — `lint`
  stays a read); MCP `word_read command=lint` / `word_write command=regularize`.
- **Foundational sub-item:** a new **`anchor.format_info()` read control** — the
  missing read mirror of `format_paragraph`/`format_run` (wordlive has the write
  side but no read side). Returns effective alignment/font/spacing per anchor,
  each annotated style-inherited vs direct-override, plus `mixed` for `wdUndefined`
  runs. The substrate every consistency rule consumes; useful standalone. CLI
  `read format`, MCP `word_read command=format_info`.
- **Direct-override detection live-validated 2026-06-17** (effective vs
  `ParagraphStyle` comparison, `wdUndefined` for mixed runs, `Font.Reset()`
  strip-to-style all confirmed — see `spec-linter.md` §7).
- **Probe before shipping:** each fix must be **idempotent** (re-running
  `regularize` is a no-op) so it doesn't "fix" intentional formatting.

### 2. Checkpoint + diff — `doc.checkpoint()` / `doc.changes_since()` / `doc.diff()`

> **✅ Shipped (Unreleased, 2026-06-20).** All of build-order steps 1–4 landed:
> `doc.checkpoint(include=…, within=…)` → a serialisable `Checkpoint`
> (`to_json`/`from_json`; `include` ∈ `text`/`text+style`/`text+format`),
> `doc.changes_since(cp)` (cp → now) and `doc.diff(cp_a, cp_b)` (two stored),
> emitting `replace`/`insert`/`delete`/`restyle`/`reformat` records that carry the
> **current** `para:N` — aligned by paragraph content via `SequenceMatcher`, with a
> whole-doc `doc_hash` fast-path to `[]`. **No new COM surface** (pure composition
> over `_findreplace._normalize` + `paragraph_text` + `format_info`). Wired across
> Python / CLI (`checkpoint`, `diff --since` / `--from`/`--to`) / MCP (`word_read
> command=checkpoint`/`diff`); not an `exec` op (the token round-trips through the
> caller). One refinement over the sketch: within a `replace` block, old→new
> paragraphs pair by **text similarity** (not position), so an edit beside an
> insert/delete classifies the way a human reads it (live-validated 2026-06-20).
> **Still deferred (steps 5–6):** pin-backed exact identity (`track=True`), move
> detection (`moves=True`), per-cell table diffing (v1 fingerprints a table as a
> single `cells_hash`), and the in-document checkpoint store (`doc.variables`). See
> `spec-checkpoint-diff.md` and `CHANGELOG.md`.

Promoted from nice-to-have to **load-bearing** by the 2026-06-17 event probe: Word
emits **no content-change event** (see Priority 7), so the only way to answer "what
changed" is a structural snapshot + diff. Also the substrate for compare/merge
(Priority 5) and the save-hook review (Priority 7). **Detailed design:
`spec-checkpoint-diff.md`** (the fingerprint shape, content-aligned diff via
`SequenceMatcher`, and the opt-in pin-backed tracked checkpoint).

- **`doc.checkpoint()`** → an opaque, serialisable structural snapshot (per
  paragraph: text + style + key format; tables/headings; a content hash). Cheap
  pure read; no Word state touched.
- **`doc.changes_since(cp)`** / **`doc.diff(cp_a, cp_b)`** → a structured change
  list `{op: insert|delete|replace|restyle, anchor_id, before, after}`, aligned by
  content (difflib over `find()`'s normalization).
- Lets a multi-step agent verify "did my edits land where I meant" without
  re-reading the whole document, and powers "review what the user changed since I
  last looked."
- Surfaces: Python + CLI (`checkpoint` emits a token the caller stores; `diff
  --since FILE`) + MCP `word_read`. **Not an exec op** (pure reads; the token
  round-trips through the caller, not Word).
- **Deferred:** persisting checkpoints inside the document (a `doc.variables`
  store) — start with caller-held tokens.

## Priority 2 — Read ergonomics & interchange

> **Both Priority-2 items — ✅ shipped (Unreleased, 2026-06-21).** Items 3 and 4
> landed together as one batch (new module `_export.py` holds one COM
> document-walk + the pure Markdown/HTML emitters + the budgeted elide). See the
> two ✅ notes below and Part I (the `_export` capability rows + `CHANGELOG.md`).

### 3. Token-budgeted whole-doc read — `doc.read(budget=N)`

> **✅ Shipped (Unreleased, 2026-06-21).** `doc.read(budget=6000, depth=None)`
> returns annotated Markdown: headings verbatim (each tagged `<!-- heading:N -->`,
> the navigation spine), tables as one-line shape stubs, body sampled to fit the
> budget — spread across sections **weighted by heading depth** (shallow keeps more
> than deep), overflow elided to markers naming the `para:` range so an agent drills
> in via `to_markdown(within=…)`. A section whose first block overflows keeps a
> bounded lead snippet (≤20 words) + a truncation marker, so `budget` bounds the
> output regardless of section count and no section vanishes; image-bearing blocks
> always survive (`image:N` stays addressable); `depth` caps how deep a section
> keeps body. **No new COM surface** — reuses Feature 4's `walk_blocks`, then a pure
> `build_digest`. Tuning knobs (`_DEPTH_WEIGHTS`/`_LEAD_WORDS`/`_CHARS_PER_TOKEN`)
> are `_export` constants; live-tuned against a real multi-section doc. Wired Python
> / CLI (`read digest [--budget] [--depth]`) / MCP (`word_read command=digest`). See
> `CHANGELOG.md`. **Deferred:** a real tokenizer (the char-based ~4/token estimate
> is intentionally cheap) and sub-block elision beyond the lead-snippet floor.

A structure-aware compressed representation of an **entire** document sized to a
token budget, so an agent can load an 80-page document into context cheaply with
every anchor still addressable. Headings verbatim; body summarised/elided by depth;
tables as shapes; anchors preserved. Pure read; Python + CLI + MCP `word_read`.
**Probe:** the eliding heuristic (how to spend the budget across a deep vs. flat
doc) wants a live tuning pass.

### 4. Markdown / HTML export — `doc.to_markdown(within=anchor)`

> **✅ Shipped (Unreleased, 2026-06-21).** `doc.to_markdown(within=None)` /
> `doc.to_html(within=None)` — the read mirror of `insert_markdown`. One COM
> document-walk (`_export.walk_blocks`) → a flat `Block`/`Span`/`TableNode` node
> list → two pure emitters (so md+html agree on structure). Emits headings, nested
> bullet/numbered lists, `**bold**`/`*italic*`/`***both***` (HTML keeps underline),
> GFM pipe tables (pipe-escaping + `:--`/`:-:`/`--:` alignment), inline images as
> `![alt](image:N)`, and hyperlinks as `[text](url)`; round-trips the constrained
> import subset and reads the rest richer (lossy by design). `within` scopes to an
> anchor's **literal range** (a `heading:N` is just its line — use `between`/`range:`
> for the section). **No new COM surface** — composes over `Range.Words` (per-word
> emphasis, whitespace hoisted out of markers), `ListFormat` (type/level, not style
> name), a table range-interval walk, and `Range.Hyperlinks` (word-offset matched) —
> all live-probed 2026-06-21. Mapping rules + escaping ported (conceptually) from the
> sibling `all2md`. Wired Python / CLI (`read markdown` / `read html`, `--within`,
> `--text` pipes raw) / MCP (`word_read command=to_markdown` / `to_html`). Reads, so
> **not** `exec` ops. Smoke-validated (insert→export round-trip, tables, images,
> links, scoping). See `CHANGELOG.md`. **Deferred:** intra-table-cell emphasis,
> footnote/comment emission, and a normalized AST (kept the flat node list — wordlive's
> import side is a flat block list, not a tree).

The **read** mirror of the shipped `insert_markdown` compose path — clean Markdown
(or HTML) *out* of a document or any anchor's range. Constant agent need ("give me
this section as markdown"). Port the docx→md mapping + learnings from `all2md`
(which already does this well over structured docx). Pure read; Python + CLI + MCP.
**Deferred:** full-fidelity HTML and round-trip guarantees (export is lossy by
design, like the constrained-subset import).

## Priority 3 — Polish (incremental; lands before the event loop)

Small, well-scoped cuts rounding out shipped clusters. COM detail for each is in
Part III's catalogue (promoted here, not re-derived).

### 5. Table styling & polish

> **Core slice — ✅ shipped (Unreleased, 2026-06-20).** The headline restyle gaps
> all landed, wired Python / CLI / `exec` op / MCP and live-Word validated:
> `Table.set_style` (restyle an existing table — `set_table_style` / `table
> set-style`), `Table.set_alignment` (whole table across the page), `Table.set_borders`
> (the **whole grid** in one call — the table-wide counterpart `set_borders` on a
> cell explicitly excluded), `Table.set_banding` (the six `ApplyStyle*` "Table Style
> Options"), and `Cell.set_vertical_alignment` (flat `cell-valign`). **Row / column
> styling** is solved by the anchor scheme: `table:N:row:R` → `RowAnchor` (a
> contiguous `Rows(R).Range`) and `table:N:col:C` → `ColumnAnchor` (fans across
> `Columns(C).Cells`), plus `Table.row(R)`/`Table.column(C)` — so the shipped
> `shading`/`borders`/`apply-style`/`format-run` verbs style a whole strip with
> **zero new styling surface**. **Live-probe findings baked in** (2026-06-20): cell
> vertical alignment maps 0/1/**3** (2 = `wdAlignVerticalJustify` is invalid for a
> cell); a table-style swap **overwrites direct cell shading** (restyle first, then
> overrides — documented); and a column has **no** Word per-column model on a merged /
> mixed-width table (`Column.Range` is absent and `Columns(C)` raises "mixed cell
> widths"), so a column op there raises a clean `OpError` pointing at per-cell
> `table:N:R:C` styling. See Part I's `table:N:row:R`/`col:C` taxonomy note and
> `CHANGELOG.md`.
>
> **Structural-polish strand — ✅ shipped (Unreleased, 2026-06-21).**
> `Table.add_column(values=None)` / `delete_column(index)` mirror the row ops
> (`add_column` fills top-to-bottom; `delete_column` raises a clean `OpError` on a
> merged / mixed-width table, pointing at per-cell deletion). `Cell.merge(other)` /
> `Cell.split(rows=1, cols=2)` are the **merged-cell addressing story**: either
> makes the table **non-uniform**, so `Table.is_uniform` reports `False`,
> `table:N:R:C` indexes *physical* cells, and `Table.read()`/`grid()` walk each
> row's physical cells (the read carries a `uniform` flag). Wired Python / CLI
> (`table add-column`/`delete-column`/`merge-cells`/`split-cell`) / `exec` ops /
> MCP; live-probed + smoke-validated 2026-06-21. See `CHANGELOG.md`. **Still
> deferred:** add/remove of an *interior* column at an index (only right-edge
> append today), and restyle nuances that preserve conditional formatting.

Two strands: a **table-styling surface** (the agent-publishing need — restyle
cells/rows/columns and whole tables with ease) and the **structural polish**
parked from earlier sweeps.

**What's shipped today (the baseline this builds on):**
- **Whole-table style at *creation* only** — `insert_table(style="…")` /
  `table create --style` / `create_table` op set `table_com.Style`; `style=None`
  defaults to the built-in `"Table Grid"`. Any built-in table style works by
  name (live-probed 2026-06-19: all **247** table-type styles are present in
  `doc.Styles` on a fresh doc — `"Plain Table 3"`, `"Grid Table 4 - Accent 1"`,
  etc. all resolve; the latent-style worry was unfounded on this build).
  `style list` filtered to `type=="table"` discovers them.
- **Cell-level styling** (a `Cell` *is* an `Anchor`): `apply_style`,
  `format_paragraph`, `format_run`, `set_shading(fill=…)`, `set_borders(…)`.
- **Structure:** `Table.autofit` (content/window/fixed), `set_heading_row`
  (repeating header + `AllowBreakAcrossPages`), `header=True` row-1 bolding.

**The gaps (this item):**
- **Restyle an *existing* table** — no verb takes a `table:N` and sets its table
  style; `table_com.Style = …` is reachable only at creation. The headline ask.
  → a `Table.set_style(name)` / `table set-style` CLI / `set_table_style` op.
- **Row / column styling in one call** — today you loop cells. Want
  `table.row(R).set_shading/​set_borders/​apply_style/​format_run` and a
  `table.column(C)` peer (a column is `Table.Columns(C).Cells` — its own
  range), so "shade the header row", "right-align the totals column",
  "bold column 1" are single ops. Decide the addressing: extend the anchor
  scheme (`table:N:row:R` / `table:N:col:C`) vs. methods on the `Table` wrapper.
- **Banding** — `Table.set_banding(rows=True, columns=False)` toggling
  `Table.Style*` band conditions (`wdTableStyleApply*`) / the table-style options
  (first-row / last-row / first-col / last-col / banded-row / banded-col), the
  knobs Word's "Table Style Options" ribbon group exposes. These ride on top of
  the applied table style.
- **Whole-table alignment & borders** — table alignment on the page
  (`Table.Rows.Alignment` left/center/right), and **table-wide** borders
  (`Table.Borders`) — currently explicitly out of scope in `set_borders`, which
  is per-range/per-cell only. A `Table.set_borders(…)` for the whole grid.
- **Cell vertical alignment** — `Cell.VerticalAlignment` (top/center/bottom);
  pairs naturally with the cell-styling surface.

**Structural polish (parked from the 2026-06-01 sweep):**
- Merged / split cells — the addressing model assumes a rectangular grid; needs
  a story for how a merged cell reports its `table:N:R:C` id.
- `add_column` / `delete_column` (mirrors `add_row` / `delete_row`).

Surfaces, as ever: Python + CLI (`table` group) + `exec` ops + MCP. **Probe
before shipping:** band toggles and table-wide borders only behave sensibly once
a real table style is applied — validate the interaction live (a styleless table
ignores band conditions). Restyle must preserve cell-level direct overrides the
user set (Word reapplies the style's conditional formatting — confirm it doesn't
clobber explicit shading/borders, or document that it does).

### 6. List polish

> **✅ Shipped (Unreleased, 2026-06-21).** `anchor.apply_list_format(levels)`
> **authors a custom multi-level list template** (`Document.ListTemplates.Add` +
> per-`ListLevel` mutation) and applies it — each per-level spec sets marker
> `format` / number `style` / `bullet` glyph + `font` / indentation (`number_position`
> / `text_position`) / `trailing` / `alignment` / marker `bold`/`italic`/`color`;
> >1 level mints an outline template. `anchor.read_list_levels()` is the read
> mirror. Live-probed 2026-06-21 (all per-level props settable; the one trap —
> baked in — is that a bullet level is the glyph + a symbol font, **not**
> `NumberStyle=bullet`, which raises `0x800a1200`). Wired Python / CLI (`list
> format` / `list levels`) / `exec` op (`apply_list_format`) / MCP; smoke-validated.
> New constants `WdListNumberStyle` / `WdTrailingCharacter` / `WdListLevelAlignment`.
> See `CHANGELOG.md`. **Deferred:** the "multi-section `LinkToPrevious`" sub-item is
> really a **header/footer** capability (`HeaderFooter.linked_to_previous` exists
> read-only in `_sections.py`) — making it writable is a separate create-but-can't-edit
> fix, not list work, so it's parked there rather than here.

Per-level bullet / number format and custom list-template authoring — **shipped**
above. (The "multi-section `LinkToPrevious`" sub-item was reclassified as a
header/footer gap — see the ✅ note.)

> **✅ Shipped (Unreleased, 2026-06-21).** `anchor.apply_list_format(levels)`
> **authors a custom multi-level list template** (`Document.ListTemplates.Add` +
> per-`ListLevel` mutation) and applies it — each per-level spec sets marker
> `format` / number `style` / `bullet` glyph + `font` / indentation (`number_position`
> / `text_position`) / `trailing` / `alignment` / marker `bold`/`italic`/`color`;
> >1 level mints an outline template. `anchor.read_list_levels()` is the read
> mirror. Live-probed 2026-06-21 (all per-level props settable; the one trap —
> baked in — is that a bullet level is the glyph + a symbol font, **not**
> `NumberStyle=bullet`, which raises `0x800a1200`). Wired Python / CLI (`list
> format` / `list levels`) / `exec` op (`apply_list_format`) / MCP; smoke-validated.
> New constants `WdListNumberStyle` / `WdTrailingCharacter` / `WdListLevelAlignment`.
> See `CHANGELOG.md`. **Deferred:** the "multi-section `LinkToPrevious`" sub-item is
> really a **header/footer** capability (`HeaderFooter.linked_to_previous` exists
> read-only in `_sections.py`) — making it writable is a separate create-but-can't-edit
> fix, not list work, so it's parked there rather than here.

Per-level bullet / number format and custom list-template authoring — **shipped**
above. (The "multi-section `LinkToPrevious`" sub-item was reclassified as a
header/footer gap — see the ✅ note.)

### 7. Chart depth (post-insert, static — no Excel respin) — ✅ shipped (Unreleased)
Error bars (`Series.ErrorBar`), series/point formatting (`MarkerStyle`/`MarkerSize`,
line `Smooth`, bar `GapWidth`/`Overlap`, pie `Explosion`, per-element data-label
`.Font`), trendline `Order`/`Period` knobs, and `HasDataTable` — all **live-probed
settable on the BreakLink-static chart, no Excel respin** (2026-06-21). Shipped as
`ChartAnchor.format_series` / `add_error_bars` + extended `format` (gap/overlap/
data-table) / `add_trendline` (order/period); wired across Python / CLI / `exec` /
MCP. **Deferred:** `ApplyChartTemplate(.crtx)` — `SaveChartTemplate`/
`ApplyChartTemplate` **block/hang under headless DispatchEx**, too risky for the
surface (note in [[charts-com-gotchas]]); secondary axes / multi-series authoring /
3-D stay out (the charts philosophy — keep each extension narrow). `Axis.Visible`
stays out (not settable under late binding, confirmed 2026-06-17). Full COM surface:
Part III.

## Priority 4 — Vision ↔ anchor bridge

### 8. Anchor-overlay snapshots — `snapshot(..., overlay="anchors")`

Render a page with `para:N` / `heading:N` / `table:N` / `image:N` ids drawn as
labelled overlays on the rasterised image, so a **vision model can see the layout
and name the exact anchor to act on** — closing the loop between wordlive's two
addressing modes (pixels + anchor ids). Small effort over the shipped snapshot
pipeline (PyMuPDF draws the boxes/labels post-rasterise from each anchor's
`location()` page + bounding rect). Python `snapshot` / `snapshot_anchor`, the
`--overlay` CLI flag, the `word_snapshot` param. **Probe:** extracting per-anchor
pixel rects from Word (`Information(wdHorizontalPositionRelativeToPage)` + page
geometry) and mapping them onto the PyMuPDF page raster.

## Priority 5 — Comparison & generation (larger surfaces; each its own design pass)

### 9. Compare / merge — `word.compare(a, b)` → tracked revisions
`Application.CompareDocuments` renders the delta of two drafts as **tracked
revisions** in a new document — which wordlive already reads and accepts/rejects end
to end. Needs a **second-document handle** in the model (also unlocks
cross-document section copy). Pairs with checkpoint/diff (Priority 1) as the
in-session counterpart. *Medium-high; the second-doc handle is the new primitive.*

### 10. Templating + mail merge
Fill a content-control / `{{placeholder}}` template from a JSON record
(`doc.fill(record)`), and/or `Document.MailMerge` for "N letters from a table." The
enterprise document-generation use case; builds on shipped content-control creation
+ variables. *High value, high cost — its own multi-step design pass.*

## Priority 6 — Deliverable hand-off

### 11. Prepare-for-sharing — `doc.prepare_for_sharing(...)`
One-call hand-off bundling three parked items: inspect & strip metadata / hidden
text / resolved comments (`Document.RemoveDocumentInformation`), optional
`Document.Protect(...)` to lock editing (gate like persistence), and an
accessibility audit (alt-text coverage, heading structure, reading order — feeds
the linter). The "make this safe to send" workflow.

## Priority 7 — Event loop / co-editing (sequenced after polish)

### 12. `doc.watch()` — event sink + reactive mode
The biggest architectural addition and wordlive's headline differentiator: an agent
that works *alongside* the user in a live session. Feasibility **proven live**
(2026-06-17) via `WithEvents` + `PumpWaitingMessages`. Headline use: a
**review-on-save** hook (`doc.on_save(review_fn)`) that intercepts
`DocumentBeforeSave`, runs `doc.lint()`/`regularize`/comments, then re-saves.

**Load-bearing event facts (live-probed 2026-06-17, Word 16 — `MSWORD.OLB` 8.7):**
- **No content-change event exists.** `TypeText`/`InsertAfter` fire nothing;
  `DocumentChange` is active-doc-switch, not an edit. Reacting to user edits must
  go through **checkpoint + diff** (Priority 1) — that's *why* it's Priority 1.
  Content-control `OnExit`/`ContentUpdate` are the only "content changed" signals,
  and only inside a CC.
- **Sinks:** `ApplicationEvents4` (app-wide — `DocumentBeforeSave(Doc, SaveAsUI,
  Cancel)`, `DocumentBeforeClose`, `DocumentBeforePrint`, `WindowSelectionChange(Sel)`,
  `DocumentChange()`, `DocumentOpen`, `NewDocument`, `WindowActivate/Deactivate`,
  `MailMerge*`); `DocumentEvents2` (per-doc — `ContentControlOnExit/OnEnter/
  ContentUpdate/AfterAdd/BeforeDelete`, `XMLAfterInsert/BeforeDelete`,
  `BuildingBlockInsert`).
- **Requires a running message pump** (`PumpWaitingMessages` in the COM/STA
  thread) — a long-running pumped process, not the one-shot CLI; pairs with the
  deferred **`asyncio` wrapper**.
- **Event args arrive as raw late-bound `PyIDispatch`** — `.Name`/`.Start` raise
  `AttributeError` until re-wrapped with `win32com.client.Dispatch(arg)`; the sink
  must re-wrap each arg into a wordlive type.
- **`DocumentBeforeSave` Cancel is byref** (a pywin32 handler must
  `return (SaveAsUI, Cancel)` to write it back); calling `Save()` inside the
  handler **re-fires** the event (re-entrancy guard needed), and a `Save()` that
  needs UI (`SaveAsUI=True`) **blocks the automation thread** on the dialog.
- **`WindowSelectionChange` is the closest edit signal but unreliable** — it does
  **not** fire per keystroke. Live-probed 2026-06-17: during 25s of continuous
  manual typing (~90 char edits) it fired only **4×**, once missing an 11s / 58-char
  run entirely (Word coalesces it during fast typing); discrete caret moves
  (click / arrow / programmatic `.Select()`) do fire it, and programmatic
  `TypeText` fires it *not at all*. So it can't even serve as a dependable
  "something changed, go look" trigger — **timer-based polling / checkpoint-diff is
  the only reliable edit detector** (a 170ms poll caught all 85 changes the events
  missed). This is the empirical proof that Priority 1's checkpoint/diff is
  mandatory, not optional.

(Mirrors the project memory `word-com-events-gotchas`.)

---

# Part III — Deferred & declined

## Declined

- **Native SmartArt** — declined 2026-05-31. `Shapes.AddSmartArt(Layout, …)` +
  driving the `Nodes` tree is heavy and brittle (GUID-indexed layouts,
  floating-only, hard to read back, locale/version-sensitive) for an intent that
  **charts** (shipped) and **render-diagram-to-image** (mermaid/graphviz →
  `insert_image`) already serve. Revisit only on a concrete SmartArt need.

## Deferred (no concrete trigger yet)

- **Chart formatting & design — the curated slice shipped (Unreleased).** A first
  pass (`ChartAnchor.format`/`set_axis`/`add_trendline`/`set_series_color` +
  `chart_style`/`has_legend` reads) is live (see Part I); it operates on the
  post-insert **static** chart with no Excel respin (re-probed 2026-06-17: 0 orphan
  EXCEL.EXE). The `.com` escape hatch (`doc.charts[N]._shape().Chart`) still
  reaches everything; the one rule holds — don't read chart *data* back. The depth
  items below were **promoted to Part II Priority 3, item 7 (chart depth) and
  shipped (Unreleased)** — `format_series` / `add_error_bars` + `format` gap/
  overlap/data-table + `add_trendline` order/period (live-probed 2026-06-21). This
  entry remains the detailed COM catalogue:
  - **Series/point depth ✅ (the core shipped):** pie `Points(i).Explosion`,
    `MarkerStyle`/`MarkerSize`, line `Smooth`, bar `ChartGroups(1).GapWidth`/
    `Overlap`, per-element data-label `.Font` — all shipped. `Series.Format`
    `.Shadow`/`.Glow` still deferred (no concrete need).
  - **Scientific ✅:** `Series.ErrorBar(Direction, Include, Type, Amount)` (error
    bars — enum args probed: Type fixed=1/percent=2/stdev=3/sterror=4) and
    polynomial `Order` / moving-average `Period` trendline knobs — all shipped.
  - **Whole-chart:** `HasDataTable` ✅ shipped. Still deferred — secondary axes,
    multi-series authoring, `ThreeD` (lives on the series `Format`); and
    `ApplyChartTemplate(.crtx)` is **dropped, not just deferred**: `SaveChartTemplate`/
    `ApplyChartTemplate` **block/hang under headless DispatchEx** (probed
    2026-06-21). Axis show/hide stays out — `Axis.Visible` is **not settable**
    under pywin32 late binding (confirmed 2026-06-17, so it's intentionally absent).

  Keep any extension narrow (the charts philosophy); `ChartColor`/"Change Colors"
  isn't one COM property — it's `ChartStyle` + per-series/point fills (now
  exposed). Colours are Office `OLE_COLOR` BGR longs — reuse `_format.to_bgr`.
- **Events / sinks** — `WithEvents(word.com, Handler)` for `DocumentBeforeSave`,
  `WindowSelectionChange`. **→ promoted to Part II Priority 7 (`doc.watch()`)** —
  use case found (review-on-save) and feasibility live-probed 2026-06-17; the
  marshalling gotchas are now captured there.
- **`asyncio` wrapper** — natural once events land. Sync core stays. **→ pairs
  with Part II Priority 7** (the pumped event loop wants it).
- **Read-model caching** — premature; live reads are correct. Cache when events
  arrive to invalidate on `DocumentChange`.
- **Styles deep cuts** — character/list/linked styles, theme-aware fonts, style
  import from template, `UpdateStyles`, style-usage inventory ("used vs. defined",
  "style near this anchor"). (Basic style creation/modification shipped v0.12.0;
  document themes shipped v0.16.0.)
- **Table styling & polish** — *core slice shipped (Unreleased):* restyle
  (`Table.set_style`), row/column styling in one call (the `table:N:row:R` /
  `table:N:col:C` anchors), banding / table-style options (`set_banding`),
  whole-table alignment + borders (`set_alignment` / `set_borders`), and cell
  vertical alignment (`Cell.set_vertical_alignment`) all landed (see Part II item 5
  + Part I). *Structural strand also shipped (Unreleased):* `add_column` /
  `delete_column` (mirror the row ops) and `Cell.merge` / `split` + `is_uniform`
  (the merged/split-cell addressing story — `table:N:R:C` then indexes physical
  cells). *Still open:* interior-index column insert (only right-edge append today).
  (AutoFit shipped v0.15.0 as `Table.autofit`; whole-table style settable at
  *creation* via `insert_table(style=…)`.)
  **→ Part II Priority 3 (item 5, table styling & polish).**
- **List polish** — custom list-template authoring + per-level bullet/number
  format **✅ shipped (Unreleased)** as `apply_list_format` / `read_list_levels`
  (see Part II item 6 + Part I). *Still open:* the "multi-section `LinkToPrevious`"
  sub-item, reclassified as a **header/footer** gap (`HeaderFooter.linked_to_previous`
  is read-only — make it writable). **→ Part II Priority 3 (item 6).**
- **Comment/revision polish** — comment replies (`comment.reply`), author/date
  filtering on `list()`. (Per-revision accept/reject is active backlog — Part II.)
- **Image polish** — *mostly shipped* (Unreleased): the floating-shape model
  (`shape:N`) delivered re-wrap, absolute/relative positioning (`Left`/`Top` via
  `set_position`), resize, alt-text, and **replace-in-place** (`replace_image`)
  for **floating** images; Slice 2 added the **inline `image:N` restyle** subset
  (`set_alt_text` / `set_size` — non-wrap) plus **shape depth** (`set_rotation`,
  `set_z_order`, `set_text_frame`), **group/ungroup** (`doc.group_shapes` /
  `ShapeAnchor.ungroup`), and the **`textbox:N`** alias; Slice 3 added **wrap
  *side* + text distance** (`set_wrap(side=…, distance_*=…)`, `WrapFormat.Side` /
  `Distance*`), **cropping** (`ShapeAnchor.set_crop` for a floating picture +
  `ImageAnchor.set_crop` for an inline one, `PictureFormat.Crop*`), and the **MCP
  `ImageContent` block** — `word_read command=read_image` now hands a vision model
  the actual picture (an inline image content block, like `word_snapshot`) instead
  of base64 text. *Still open:* EMF/WMF; chart-image export; OLE extraction.
  *Declined:* text-box **autosize** ("resize-to-fit-text") — no clean COM path
  (`TextFrame.AutoSize` no-ops; `TextFrame2.AutoSize` rejects the value,
  live-probed 2026-06-19). Live-probed 2026-06-19 (Slice 3): `WrapFormat.Side`
  is silently coerced back to `both` for `top-bottom`/`front`/`behind`;
  `PictureFormat.Crop*` raises on a non-picture shape (hence the `set_crop`
  picture guard); inline `InlineShape` shares `PictureFormat`, so one helper
  covers both.
- **Cursor/inline polish** — raw inline `insert_before`/`insert_after` (no new
  paragraph) on the CLI; a `write_cursor` exec op (fights `EditScope`'s
  cursor-restore in a batch).
- **Reference-apparatus extensions** — custom TOC/footnote marks & separators,
  numbering format/restart, footnote↔endnote conversion, content-control
  placeholder text (`SetPlaceholderText`), custom-XML data binding (below).

## Unexplored COM control surfaces (catalogue for prioritization)

Sketched 2026-06-01 — the broader sweep of "what else could an agent usefully
drive through COM." **Nothing here is approved or designed**; it's a parking lot.
Each entry: the COM entry point, the agent use-case, a rough **weight**. The
application/window/view surfaces (group I) are in real tension with *politeness*
(they change global UI state, not document content) — louder opt-in, or reads.

### A. Proofing & language

- **Spelling / grammar / readability** — ✅ **shipped v0.15.0** as
  `doc.proofing()` (`Range.SpellingErrors` / `GrammaticalErrors` + readability
  scores, each flagged run carrying a `range:START-END` to pair with comments /
  track-changes).
- **AutoCorrect / AutoText entries** — `Application.AutoCorrect`. Niche; mostly a
  write to global app state. *Low priority.*

### B. Document metadata & properties

- **Built-in & custom properties** — ✅ **shipped v0.15.0** as `doc.properties`
  (`read()` → `{builtin, custom}`; `set`/`delete`).
- **Document variables** — ✅ **shipped v0.15.0** as `doc.variables`.
- **Statistics** — ✅ **shipped v0.14.0** as `doc.stats()`.

### C. Document lifecycle & safety

- **Protection / editing restrictions** — `Document.Protect(type, …)` /
  `Unprotect`, read-only enforcement, `Permission` (IRM). Hand back a locked
  deliverable. *Weight: medium; gate like persistence.* **→ folded into Part II
  Priority 6 (item 11, prepare-for-sharing).**
- **Encryption / passwords** — `Document.Password` / `WritePassword`.
  *Security-sensitive; gate hard or leave to the human.*
- **Compare / merge** — `Application.CompareDocuments`, `Document.Merge`. "Diff
  these two drafts" as tracked revisions — genuinely agent-shaped. *Weight:
  medium-high; needs a second-document handle in the model.* **→ promoted to Part
  II Priority 5 (item 9).**
- **Inspect / redact**, **digital signatures** — *niche / low priority.*

### D. Mail merge & data-driven generation

- **`Document.MailMerge`** — data source, merge fields, `Execute` to a new doc /
  printer / email. The canonical "generate N letters from a table" workflow —
  squarely in wordlive's document-generation wheelhouse, but a large surface
  (data-source binding, field mapping, output routing). *Weight: high value, high
  cost — its own multi-step design pass.* **→ promoted to Part II Priority 5 (item
  10, with content-control templating).**

### E. Structured & data-bound content

- **Content-control *creation*** — ✅ **shipped v0.16.0**
  (`anchor.insert_content_control(...)` / `doc.content_controls.add(...)`).
- **Custom XML parts / data binding** — `Document.CustomXMLParts`, binding CCs to
  XML nodes. Powerful for structured generation; heavier and niche. *Defer.*
- **Legacy form fields** — `Document.FormFields`. *Superseded by content
  controls; skip unless asked.*

### F. Reference apparatus

- **Index**, **table of figures** — ✅ **shipped v0.16.0**.
- **Table of authorities**, **citations & bibliography** — ✅ **shipped v0.16.0**.
- Footnotes / endnotes / TOC / captions — ✅ shipped v0.12.0.

### G. Drawing layer (floating shapes & embedded objects)

- **General `Document.Shapes`** — *grouping + z-order shipped (Unreleased)* via
  `doc.group_shapes` / `ShapeAnchor.ungroup` / `set_z_order`; *still open:* lines,
  connectors, freeform, fill/gradient, shadow/3D. The watermark / text-box
  flourishes (Part II) + the `shape:N` model are the thin edge; the rest of the
  drawing layer is large and floating-only (hard to read back, fights the anchor
  model). *Weight: low — cherry-pick specific shapes.*
- **Embedded / OLE objects** — `InlineShapes.AddOLEObject` (embed an Excel range,
  a PDF, another doc). *Medium-niche; pairs with charts.*

### H. Range / navigation read surface

- **`Range.Information(wd…)`** — ✅ **shipped v0.14.0** as `anchor.location()`.
- **Story ranges** — `Document.StoryRanges` / `StoryType` iteration (main text,
  footnotes, headers, text frames, comments as distinct stories). Makes
  find/read *story-aware* instead of main-text-only. *Weight: medium.*
- **Range unit navigation** — `Expand`/`Collapse`/`MoveStart/End`, sentence/word/
  paragraph units. Building blocks for finer anchors. *Low, incremental.*

### I. Application / window / view control (politeness-sensitive — gate or read-only)

- **`Application.Options`** — spelling/grammar toggles, display settings. *Global
  app state; write only behind a loud opt-in.*
- **View / window** — `ActiveWindow.View.Type`, zoom, split, `Windows.Arrange`.
  Could help a vision-model workflow (set print-layout before a snapshot) but
  mutates the user's UI. *Weight: low; let `snapshot` handle layout implicitly.*
- **`DisplayAlerts` / `ScreenUpdating`** — internal optimization, not an agent verb.
- **Printing** — `Document.PrintOut`. Real side-effect; gate like persistence.

### J. Read-only discovery collections

✅ Effectively complete: `doc.lists` / `tables` / `images` / `footnotes` /
`endnotes` / `revisions` / `bookmarks.list()` (v0.8–0.13), plus `doc.hyperlinks`
/ `doc.fields` / `doc.properties` (v0.15.0). No outstanding gaps.

---

# Cross-cutting work (any release)

- **Image-source hardening** — ✅ **shipped v0.13.0** with persistence (CLI/MCP
  reject non-local `--path` / op `path` before the `is_file()` probe — UNC/NTLM,
  `file://`/URL/SSRF, local-file-disclosure; optional `--image-dir` /
  `WORDLIVE_IMAGE_DIRS` allowlist; Python API ungated; base64/bytes unaffected).
- **HRESULT coverage** — `_BUSY_HRESULTS` is a starter set; widen as real
  `pywintypes.com_error` HRESULTs surface in smoke runs.
- **Smoke fixtures** — a real `.docx` checked in with known bookmarks / CCs /
  headings / tables, so smoke tests have a known target.
- **Docs** — `spec.md` is the design doc; a `cookbook.md` of end-to-end LLM-tool
  examples is probably more useful than API reference at this stage (the 2026-05
  agent-test build — blank doc → styled multi-page catalogue — is a ready-made
  entry).
- **CLI surface shape — flat-first (decided 2026-06-08).** Top-level verbs are
  the default, mirroring the flat `exec` op vocabulary and keeping `SKILL.md` a
  single-lookup list. Noun-groups are the exception, justified only by a stable
  object/collection with ≥3 verbs that don't read as standalone verbs
  (grandfathered: `table`, `list`, `comment`, `style`, `track`, plus the
  `read`/`write` dispatch groups). New single-verb capabilities stay top-level.
  Do **not** deepen the tree for tidiness — the LLM's primary surface is MCP's
  four dispatch tools, not `--help`. **Explicitly rejected:** folding the
  `insert-*` family into an `insert` group (breaks the op parallel).
- **Post-creation restyle parity — the "create-but-can't-edit" gaps (audited
  2026-06-19).** A recurring asymmetry: several objects accept styling/config at
  `insert_*` time but expose **no setter** to change it afterward, so an agent's
  only recourse is delete-and-reinsert (lossy — re-typing content) or dropping to
  `.com`. This fights the iterative edit-review-refine loop agents actually run.
  The fix is uniform: give each its mutator surface, wired Python + CLI + `exec`
  op + MCP, idempotent where it makes sense. **Charts are the template that got
  it right** (`format`/`set_axis`/`add_trendline`/`set_series_color` on the
  post-insert handle); **Styles** (writable), **Themes** (`set_colors`/
  `set_fonts`), and **PageSetup** (`set_page_setup`) already mutate cleanly too —
  so the pattern exists, coverage is just uneven. The gaps, ranked by agent
  leverage:
  - **Tables** — the flagship; restyle style / banding / cell-row-column. **→
    fully specced in Part II Priority 3, item 5 (table styling & polish).**
  - **Content controls** — ✅ **done.** `ContentControl.set_properties(...)` +
    `set_items(...)` (Python + CLI `set-cc-properties`/`set-cc-items` + exec ops
    `set_cc_properties`/`set_cc_items` + MCP). Form-building is now iterable in
    place (relabel a field, lock it once filled, edit dropdown choices).
  - **Hyperlinks** — ✅ **done.** `doc.hyperlinks` is writable:
    `Hyperlink.update(...)` + `set_address`/`set_sub_address`/`set_text`/
    `set_screen_tip` (CLI `set-hyperlink --index N`, exec op `set_hyperlink` by
    1-based index, MCP `set_hyperlink` with `url`→address/`bookmark`→sub_address).
    Creation stays via `link_to`; a first-class `insert_hyperlink` was not needed.
  - **Images (floating)** — ✅ **done** (Unreleased). The floating-shape model
    closed this: a floating `insert_image` now returns a `shape:N`
    [`ShapeAnchor`] with `set_wrap`/`set_position`/`set_size`/`set_alt_text`/
    `replace_image` (in-place picture swap) — see Part I's `shape:N` taxonomy +
    live-Word gotcha. *Still inline-only / open:* `image:N` (inline pictures)
    stays read-only for alt-text/resize, and cropping / EMF-WMF / OLE extraction
    remain in "Image polish" below.
  - **Text boxes** — ✅ **done** (Unreleased). `insert_text_box` now **returns a
    `shape:N` [`ShapeAnchor`]** (was `None`); restyle in place via
    `set_wrap`/`set_position`/`set_size`/`format`/`set_text`. Folded into the
    unified floating-shape model (no separate `textbox:N` — `shape:N` covers text
    boxes, floating images, and WordArt).
  - **Reference objects (TOC / index / TOF / TOA)** — only `.update()` /
    `update_page_numbers()`; changing levels/columns/format flags means
    delete+reinsert. Somewhat **inherent** (Word rebuilds these from field
    switches), but a `reconfigure(...)` that rewrites the field switches in place
    would be more polite than a teardown. *Medium; partly inherent.*
  - **Fields** — `Field.code`/`result` are read-only; no edit-field-code path.
    *Niche / low — most field edits are better done via the typed verb that
    created them.*

  **Not gaps** (already mutable, listed so the audit is closed): styles, themes,
  page setup, watermark (`set`/`remove`), lists (`apply_list_template` re-applies
  a whole list), revisions, track-changes, bibliography style, document
  properties/variables. **Probe per object:** confirm Word lets the property be
  re-set on an existing instance under late binding (some COM props are
  create-only or byref-only — the kind of thing the table/chart probes already
  surfaced).

## Open papercuts

- **Inline JSON vs. Windows paths** — `exec --ops '{…}'` with backslash paths
  mangles under PowerShell quoting + JSON escaping; path-bearing batches should
  use `--script FILE` or `--ops -` (guide note still worth adding).
- ~~`heading:N` mis-described in the agent guide~~ ✅ fixed 2026-06-01.
- ~~Relative image paths fail~~ ✅ fixed v0.10.2 (`Path(p).resolve()` before COM).
