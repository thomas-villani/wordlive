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
| Structural query helpers (`doc.between`, `doc.nearest_heading`, `doc.find_paragraphs`; content-under-heading already shipped) | Unreleased |

## Load-bearing reference facts

The addressing scheme and the constraints every change must respect, distilled
from the shipped clusters above.

### Anchor-id taxonomy (the stable, LLM-visible addressing scheme)

`heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`, `table:N:R:C`,
`range:START-END`, `header:S:WHICH` / `footer:S:WHICH`, `footnote:N`,
`endnote:N`, `image:N`, `equation:N`, `chart:N`, `pin:CODE`, `start`, `end`. One
resolver: `doc.anchor_by_id(id)`. A malformed scheme (`banana:7`) reports
"unknown anchor type".

- **`chart:N`** is positional over the document's chart inline shapes
  (`HasChart`), in document order — it renumbers when an earlier chart is
  inserted; re-list (`doc.charts`), don't cache. Metadata only: `chart_type` /
  `title` / `chart_style` / `has_legend`. It's also the **write** target for
  formatting (`format`/`set_axis`/`add_trendline`/`set_series_color`). Reading
  chart *series data* back is deferred (and the data link is broken at insert, so
  it's static anyway).

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
  + borders/shading); cells don't appear in `doc.outline()`.
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

### 3. Token-budgeted whole-doc read — `doc.read(budget=N)`

A structure-aware compressed representation of an **entire** document sized to a
token budget, so an agent can load an 80-page document into context cheaply with
every anchor still addressable. Headings verbatim; body summarised/elided by depth;
tables as shapes; anchors preserved. Pure read; Python + CLI + MCP `word_read`.
**Probe:** the eliding heuristic (how to spend the budget across a deep vs. flat
doc) wants a live tuning pass.

### 4. Markdown / HTML export — `doc.to_markdown(within=anchor)`

The **read** mirror of the shipped `insert_markdown` compose path — clean Markdown
(or HTML) *out* of a document or any anchor's range. Constant agent need ("give me
this section as markdown"). Port the docx→md mapping + learnings from `all2md`
(which already does this well over structured docx). Pure read; Python + CLI + MCP.
**Deferred:** full-fidelity HTML and round-trip guarantees (export is lossy by
design, like the constrained-subset import).

## Priority 3 — Polish (incremental; lands before the event loop)

Small, well-scoped cuts rounding out shipped clusters. COM detail for each is in
Part III's catalogue (promoted here, not re-derived).

### 5. Table polish
Merged / split cells (the addressing model assumes a rectangular grid — needs a
story for how a merged cell reports its `table:N:R:C` id), `add_column` /
`delete_column`. (AutoFit + repeating heading rows already shipped.)

### 6. List polish
Per-level bullet / number format, custom list-template authoring, multi-section
`LinkToPrevious` editing.

### 7. Chart depth (post-insert, static — no Excel respin)
Error bars (`Series.ErrorBar` — wants the right enum args), series/point formatting
(`MarkerStyle`/`MarkerSize`, line `Smooth`, bar `GapWidth`/`Overlap`, pie
`Explosion`, per-element `.Font`), trendline `Order`/`Period` knobs, secondary axes
/ multi-series authoring, `HasDataTable`, `ApplyChartTemplate(.crtx)`. Keep each
extension narrow (the charts philosophy). `Axis.Visible` stays out (not settable
under late binding, confirmed 2026-06-17). Full COM surface: Part III.

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
- **`WindowSelectionChange` floods** (one per caret move) — debounce.

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
  items below are **→ promoted to Part II Priority 3, item 7 (chart depth)**; this
  entry remains the detailed COM catalogue for them:
  - **Series/point depth:** `Series.Format` `.Shadow`/`.Glow`, pie
    `Points(i).Explosion`, `MarkerStyle`/`MarkerSize`, line `Smooth`, bar
    `ChartGroups(1).GapWidth`/`Overlap`, per-element `.Font`.
  - **Scientific:** `Series.ErrorBar(...)` (error bars — wants the right enum
    args); polynomial `Order` / moving-average `Period` knobs on trendlines.
  - **Whole-chart:** secondary axes, multi-series authoring, `HasDataTable`,
    `ThreeD` (lives on the series `Format`), `ApplyChartTemplate(.crtx)` (needs a
    real `.crtx` path), and axis show/hide — `Axis.Visible` is **not settable**
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
- **Table polish** — merged/split cells (addressing assumes rectangular),
  `add_column`/`delete_column`. (AutoFit shipped v0.15.0 as `Table.autofit`.)
  **→ promoted to Part II Priority 5.**
- **List polish** — custom list-template authoring, per-level bullet/number
  format, multi-section `LinkToPrevious` editing. **→ promoted to Part II Priority 6.**
- **Comment/revision polish** — comment replies (`comment.reply`), author/date
  filtering on `list()`. (Per-revision accept/reject is active backlog — Part II.)
- **Image polish** — wrap *side* + text distance; absolute/relative positioning
  (`Left`/`Top`); cropping; replace-in-place; EMF/WMF; floating-shape / chart-image
  export; OLE extraction; an MCP `ImageContent` block so the model *sees* the
  extracted original directly.
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

- **General `Document.Shapes`** — lines, connectors, freeform, grouping, z-order,
  fill/gradient, shadow/3D. The watermark / text-box flourishes (Part II) are the
  thin edge; the full drawing layer is large and floating-only (hard to read
  back, fights the anchor model). *Weight: low — cherry-pick specific shapes.*
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

## Open papercuts

- **Inline JSON vs. Windows paths** — `exec --ops '{…}'` with backslash paths
  mangles under PowerShell quoting + JSON escaping; path-bearing batches should
  use `--script FILE` or `--ops -` (guide note still worth adding).
- ~~`heading:N` mis-described in the agent guide~~ ✅ fixed 2026-06-01.
- ~~Relative image paths fail~~ ✅ fixed v0.10.2 (`Path(p).resolve()` before COM).
