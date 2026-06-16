# wordlive — feature roadmap

How to read this document (refreshed 2026-06-15, after the **v0.16.0** release):

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

## Load-bearing reference facts

The addressing scheme and the constraints every change must respect, distilled
from the shipped clusters above.

### Anchor-id taxonomy (the stable, LLM-visible addressing scheme)

`heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`, `table:N:R:C`,
`range:START-END`, `header:S:WHICH` / `footer:S:WHICH`, `footnote:N`,
`endnote:N`, `image:N`, `equation:N`, `pin:CODE`, `start`, `end`. One resolver:
`doc.anchor_by_id(id)`. A malformed scheme (`banana:7`) reports "unknown anchor
type".

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

### Exit codes (CLI)

`0` ok · `1` other/bad-input (incl. `ImageSourceError`, `SnapshotError`,
`PathNotAllowedError`, `DocumentNotFoundError`) · `2` anchor/style not found or
zero `find` matches · `3` Word busy (retryable) · `4` Word not running · `5`
ambiguous `find` match.

---

# Part II — Approved / next up (priority order)

Everything here is **specced but not yet implemented** (re-verified 2026-06-15).
Ordered by leverage.

> **Durable handles & stale-anchor diagnostics — ✅ shipped (Unreleased).** All
> five pieces landed: `doc.pin`/`stamp` + the `pin:` anchor, `doc.pin_outline` /
> `outline(pin=True)`, insert-op `bind:"name"`, `$ops[N].field` references, and
> stale-anchor recovery hints. See Part I's load-bearing facts (the `pin:`
> taxonomy entry and the `Bookmarks.ShowHidden` gotcha) and `CHANGELOG.md`. Native
> `w14:paraId` was rejected (live probe 2026-06-09): assigned lazily and
> COM-invisible (`Range.WordOpenXML` strips it), so minted bookmarks are the only
> mechanism that survives the edits we care about.

## 1. Revision write surface (read side already shipped)

`doc.revisions`, `snapshot(markup="all")`, and MCP track-status shipped in
v0.12.0. Still open:

- **Accept / reject individual revisions** — a typed `revision.accept()` /
  `.reject()` (and a doc-wide accept/reject-all). Mutating a single `Revision`
  stays on `.com` until then.
- **Revision-aware text reads** — a tracked `find_replace` on the *same*
  paragraph still drifts because both inserted and deleted runs are present
  (workaround: re-read between tracked edits, or take a `markup="all"` snapshot).
  A proper revision-aware read model is the real fix.

## 2. Publishing flourishes — the floating-shape remainder

The publishing-quality cluster (character/paragraph formatting, borders/shading/
tab-stops, style creation, PageSetup writes, fields/page numbers, pagination
controls, **drop cap** — shipped v0.15.0) is essentially done. What's left is
pure floating-shape work, individually small; bundle whichever land cheap:

- **Watermark** — `doc.set_watermark(text, …)` via the header-story
  `Shapes.AddTextEffect` (WordArt) — DRAFT/CONFIDENTIAL stamps.
- **Text box / pull quote** — `anchor.insert_text_box(text, …)` →
  `Shapes.AddTextbox`.

The shipped `insert_field` primitive already covers the general `Fields.Add` case
these once leaned on.

## 3. Charts (Excel-backed)

Approved 2026-05-31 as the SmartArt substitute. `Range.InlineShapes.AddChart2`
embeds a chart whose data lives in an embedded Excel workbook. Heavier than
images (a new transitive dependency + a much larger surface), hence below the
items above.

- **`anchor.insert_chart(kind, data, *, title=None)`** — `kind` → `XlChartType`
  (`bar`→`xlColumnClustered`, `pie`→`xlPie`, `scatter`→`xlXYScatter`,
  `line`→`xlLine`); `data` (flat label→value mapping) populates
  `ChartData.Workbook.Worksheets(1)`.
- **Transitive Excel dependency** — `AddChart2` spins up hidden Excel. Gate
  behind an "is Excel available?" probe + a typed error (clean exit, not exit 1).
- `XlChartType` subset in `constants.py` (internal). **Keep narrow:** common
  kinds + flat `data` only. **Deferred:** multi-series, secondary axes,
  axis/series formatting, `BreakLink` policy, reading existing charts back out.

## 4. Structural query helpers (new, lower priority)

From the gpt-5.4 review: content-under-heading, block-between-headings,
nearest-heading-before/after, find-paragraph-by-approx-text. Several reduce to
thin compositions over `outline` + `find`, so they sit below the introspection
reads (already shipped). Not yet ticketed.

---

# Part III — Deferred & declined

## Declined

- **Native SmartArt** — declined 2026-05-31. `Shapes.AddSmartArt(Layout, …)` +
  driving the `Nodes` tree is heavy and brittle (GUID-indexed layouts,
  floating-only, hard to read back, locale/version-sensitive) for an intent that
  **charts** (Part II) and **render-diagram-to-image** (mermaid/graphviz →
  `insert_image`) already serve. Revisit only on a concrete SmartArt need.

## Deferred (no concrete trigger yet)

- **Events / sinks** — `WithEvents(word.com, Handler)` for `DocumentBeforeSave`,
  `WindowSelectionChange`. Wait for a use case before designing the marshalling
  layer.
- **`asyncio` wrapper** — natural once events land. Sync core stays.
- **Read-model caching** — premature; live reads are correct. Cache when events
  arrive to invalidate on `DocumentChange`.
- **Styles deep cuts** — character/list/linked styles, theme-aware fonts, style
  import from template, `UpdateStyles`, style-usage inventory ("used vs. defined",
  "style near this anchor"). (Basic style creation/modification shipped v0.12.0;
  document themes shipped v0.16.0.)
- **Table polish** — merged/split cells (addressing assumes rectangular),
  `add_column`/`delete_column`. (AutoFit shipped v0.15.0 as `Table.autofit`.)
- **List polish** — custom list-template authoring, per-level bullet/number
  format, multi-section `LinkToPrevious` editing.
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
  deliverable. *Weight: medium; gate like persistence.*
- **Encryption / passwords** — `Document.Password` / `WritePassword`.
  *Security-sensitive; gate hard or leave to the human.*
- **Compare / merge** — `Application.CompareDocuments`, `Document.Merge`. "Diff
  these two drafts" as tracked revisions — genuinely agent-shaped. *Weight:
  medium-high; needs a second-document handle in the model.*
- **Inspect / redact**, **digital signatures** — *niche / low priority.*

### D. Mail merge & data-driven generation

- **`Document.MailMerge`** — data source, merge fields, `Execute` to a new doc /
  printer / email. The canonical "generate N letters from a table" workflow —
  squarely in wordlive's document-generation wheelhouse, but a large surface
  (data-source binding, field mapping, output routing). *Weight: high value, high
  cost — its own multi-step design pass.*

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
