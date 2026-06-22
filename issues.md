# Review issues — work since v0.17.0

Findings from a fan-out review of the ~14k-line / 16-commit surface added since
`v0.17.0` (charts, checkpoint/diff, export, linting, shapes, tables, lists,
content-control/hyperlink restyle, cross-surface wiring). Each finding is graded
against the four core invariants (politeness, semantic anchors, atomic undo,
structured I/O) and the `.com` escape hatch.

**Tally:** 47 findings — **5 High**, **16 Medium**, **26 Low**.

Status legend: `[ ]` open · `[~]` in progress · `[x]` fixed/closed · `[wontfix]`.

---

> **Status — High + Medium addressed on branch `fix/post-v0.17-review`.** 19 of 21
> fixed with tests; 2 (EXPORT-3, LINT-1) determined to be **false positives** on
> deeper analysis (reasoning inline). All 1530 unit tests + the touched smoke
> tests pass; ruff + mypy clean. The 26 Low findings remain open for a later pass.

## High severity (correctness / data-loss / invariant break)

- [x] **EXPORT-1** — Underline never detected. *Fixed:* added `_font_underline` (enum-aware), live-validated `<u>`.
- [x] **EXPORT-2** — `read(budget=N)` not bounded. *Fixed:* lead snippets draw from a shared budget pool; spine documented as fixed backbone.
- [x] **CHECKPOINT-1** — Blank/duplicate paragraphs share one alignment key. *Addressed:* text-only key is correct by design (folding style breaks restyle detection — `track=True` is the real fix); documented limitation + regression test that real edits still classify correctly.
- [x] **CHECKPOINT-6** — Table-only edits returned `[]`. *Fixed:* added table-diff pass (`table_change`/`table_insert`/`table_delete`).
- [wontfix] **LINT-1** — *False positive.* The per-field fix only fires when the field is **uniform** across runs (`value is not None`) and writes only that one field, so it cannot clobber a deliberately different sub-run face (which would make the field `mixed` → skipped). The whole-range write is the regularizer's intended behavior.

## Medium severity (brittle / likely to break / wrong ergonomics)

- [x] **SHAPES-1** — `replace_image` dropped crop/rotation/side/standoffs. *Fixed + live-validated* (full layout preserved, empty diff).
- [x] **TABLES-2** — `ColumnAnchor._cells()` re-resolved via `cell()`. *Fixed:* styles already-read column cells (cached `_com_cell`); fake made faithful.
- [x] **TABLES-3** — `Cell.merge` docstring wrong. *Fixed* docstring (upper-left) + live-validated smoke test.
- [x] **LISTS-4** — `apply_list_format` orphaned a `ListTemplate`. *Fixed:* validate all levels before `Add`.
- [x] **CHARTS-1** — Error-bar amount check missed aliases. *Fixed:* keys off the resolved enum.
- [wontfix] **EXPORT-3** — *False positive.* The skipped-level HTML is valid, balanced (empty placeholder `<li>`s properly contain the nested lists); verified by running `_render_list_html` on a 1→3→1 sequence.
- [x] **EXPORT-4** — Multi-word links split. *Fixed:* coalesce adjacent same-href spans + range-overlap tagging; live-validated.
- [x] **EXPORT-5** — In-cell inline images vanished. *Fixed:* append `![alt](image:N)` per cell.
- [x] **CHECKPOINT-2** — `text+format` was O(n²). *Fixed:* fingerprint reads the already-iterated range.
- [x] **CHECKPOINT-3** — Field-code/view state leaked. *Fixed:* pinned `TextRetrievalMode`; live-validated.
- [x] **CHECKPOINT-4** — `reformat` misclassified. *Fixed:* gated on `include == "text+format"` + actual fmt diff.
- [x] **CHECKPOINT-5** — `range:` scope round-trip. *Fixed:* `changes_since` rejects offset scopes with a clear error.
- [x] **LINT-2** — Missed number-only/mixed lists. *Fixed:* broadened the numbered-type set.
- [x] **LINT-3** — `regularize` dropped failure detail. *Fixed:* failure rides on the raised error.
- [x] **SURFACE-1** — `word_exec` crop docstring. *Fixed:* docstring shows canonical names + the `crop_*` alias now works.
- [x] **SURFACE-2** — Crop param drift. *Fixed:* `crop_*` accepted as aliases in the op layer (all surfaces agree).

## Low severity (polish / consistency / test-gap)

- [ ] **SHAPES-2** — `apply_shape_size` leaves `LockAspectRatio` permanently FALSE when both dims passed w/o `lock_aspect`
- [ ] **SHAPES-3** — `insert_*`/`group_shapes` leave the `_wl_shape_*` probe name on shapes whose original name was empty
- [ ] **SHAPES-4** — Inherited revision/text-history reads on `ShapeAnchor` report the anchoring paragraph, not the shape
- [ ] **SHAPES-5** — `set_position(relative_to=…)` resets both axes' frames even when only one offset changes
- [ ] **SHAPES-6** — No test for the politeness/cursor-restore invariant on shape mutators
- [ ] **TABLES-1** — `add_column`/`delete_column` lack the index-drift warning the other structural ops carry
- [ ] **LISTS-5** — Bullet/number read-back misclassifies a numbered level whose format has no `%N` placeholder
- [ ] **LISTS-6** — `_configure_level` flattens outline numbering when `style` set but `format` omitted
- [ ] **TABLES-7** — No test that `delete_column`/`ColumnAnchor` surface `OpError` on a merged table
- [ ] **CHARTS-2** — `data_table` toggle unguarded → raw COM error on pie/scatter instead of a friendly message
- [ ] **CHARTS-3** — `format_series(data_labels=False, data_label_size=…)` re-enables labels
- [ ] **CHARTS-4** — No validation guidance on out-of-range `order` / `marker_size` / `explosion`
- [ ] **EXPORT-6** — Lead-snippet word accounting approximate; can emit spurious "N more words" markers
- [ ] **EXPORT-7** — Markdown link/image targets emitted unescaped (URLs with `)`/spaces break)
- [ ] **EXPORT-8** — `walk_blocks` reads `Style.NameLocal` → locale-dependent heading fallback
- [ ] **CHECKPOINT-7** — Table fingerprint drops un-readable cells without a sentinel → masks structural change
- [ ] **CHECKPOINT-8** — Test gaps: blank-line mis-align, table diff, scoped round-trip, `_SIM_PAIR_CAP` fallback
- [ ] **CHECKPOINT-9** — Politeness honored but unguarded by any test
- [ ] **LINT-4** — `lint` table-repeat-header rule double-repaginates per table and swallows all exceptions
- [ ] **LINT-5** — Float-rounding used as epsilon; fix writes rounded value back as a new direct override
- [ ] **LINT-6** — Test gaps: `within`-scoped consistency rules, multi-rule regularize, `_overlaps`/`_in_span` duplication
- [ ] **SURFACE-3** — Border line-style param name differs three ways (`--style` / `style`+`line_style` / `line_style`)
- [ ] **SURFACE-4** — `XlErrorBarInclude` docstring names a non-existent `direction=` param (it's `include=`)
- [ ] **SURFACE-5** — Error-bars `axis` choices differ between CLI (`y|value|x|category`) and MCP docs (`y|x`)

---

# Full detail by domain

## Shapes & image polish

### [SHAPES-1] `replace_image` silently drops crop, rotation, wrap-side, and text-distance standoffs
- **Severity:** Medium
- **Location:** `src/wordlive/_shapes.py:429-468` (`replace_shape_image`)
- **Problem:** The delete+reinsert swap restores only `pos`, `WrapFormat.Type`, `Left/Top`, `Width/Height`, `RelativeHorizontal/VerticalPosition`, `LockAspectRatio`, `AlternativeText`, `Name`. It does **not** preserve `PictureFormat.Crop*`, `Rotation`, `WrapFormat.Side`, or the `WrapFormat.Distance*` standoffs — all first-class facets the same module lets a user set. The docstring claims it preserves "wrap / position / size / lock-aspect / alt-text / name", quietly omitting these.
- **Why it matters:** Data/intent loss on a common edit (swap the logo, keep its layout). Crop loss is most surprising — the new full image overflows the intended box.
- **Suggested fix:** Capture `Rotation`, `WrapFormat.Side`, the four `Distance*`, and the four `Crop*` before `Delete()` and re-apply on `new_shape` (guard each in try/except). At minimum correct the docstring.

### [SHAPES-2] `apply_shape_size` leaves `LockAspectRatio` permanently FALSE when both dims passed without `lock_aspect`
- **Severity:** Low
- **Location:** `src/wordlive/_shapes.py:297-316`
- **Problem:** When both `width`/`height` are given and `lock_aspect` is omitted (`None`), the code drops `LockAspectRatio = FALSE` but never restores the prior lock state (the restore only runs in the `lock_aspect is not None` branch). `replace_shape_image:458-463` deliberately captures/restores `lock`, so the paths disagree. A later single-dimension `set_size` then scales freely instead of proportionally.
- **Suggested fix:** Capture `prior = int(shape.LockAspectRatio)` before dropping it and restore when `lock_aspect is None`.

### [SHAPES-3] `insert_*`/`group_shapes` leave the probe name on shapes whose original name was empty
- **Severity:** Low
- **Location:** `src/wordlive/_anchors.py:1502-1507`, `:1603-1608`; `src/wordlive/_document.py:490-495`
- **Problem:** The locate-by-unique-name dance does `if orig_name: shape.Name = orig_name`. If the new shape had an empty name, the `_wl_shape_<hex>` probe name is never cleared and surfaces in `doc.shapes.list()`. Rare (Word auto-names new Shapes) but possible.
- **Suggested fix:** Restore unconditionally, or generate a clean fallback name when `orig_name` is empty.

### [SHAPES-4] Inherited revision/text-history reads on `ShapeAnchor` report the anchoring paragraph, not the shape
- **Severity:** Low
- **Location:** `src/wordlive/_anchors.py:4331-4359`
- **Problem:** `ShapeAnchor` overrides `.text` (text-frame contents) but `_range()` returns the anchoring paragraph range, so inherited `text_final`/`text_original`/`revision_segments` return the *paragraph's* history — inconsistent with `.text`.
- **Suggested fix:** Override those on `ShapeAnchor` to mirror `.text` or raise a clear "not applicable" `OpError`.

### [SHAPES-5] `set_position(relative_to=…)` resets both axes' frames even when only one offset changes
- **Severity:** Low
- **Location:** `src/wordlive/_shapes.py:273-294`
- **Problem:** A single `relative_to` keyword always writes both `RelativeHorizontalPosition` and `RelativeVerticalPosition`, so a horizontal reposition can silently shift the shape vertically (margin↔page frames differ).
- **Suggested fix:** Only write the axis whose offset is set, or accept separate `relative_h`/`relative_v`.

### [SHAPES-6] No test for politeness/cursor-restore on shape mutators
- **Severity:** Low
- **Location:** `tests/test_shapes.py`
- **Problem:** Mutators are polite today (none touch `Selection`) but there's no test guarding it; a future `.Select()` would slip through the fake fixture.
- **Suggested fix:** Add a test asserting Selection is untouched and an UndoRecord wraps a representative shape op.

## Tables & lists

### [TABLES-1] `add_column`/`delete_column` lack the index-drift warning the other structural ops carry
- **Severity:** Low
- **Location:** `src/wordlive/_tables.py:572-589`, `:591-612`
- **Problem:** After a column add/delete the physical `table:N:col:C` numbering of columns to the right shifts, but unlike `Table.delete`/`Cell.merge` the column-op docstrings don't warn about it.
- **Suggested fix:** Add a one-line index-drift note matching the existing structural-op docstrings.

### [TABLES-2] `ColumnAnchor._cells()` re-resolves cells through `cell()` (logical bounds + redundant COM)
- **Severity:** Medium
- **Location:** `src/wordlive/_tables.py:283-300`
- **Problem:** `_cells()` reads live column cells correctly, then discards them and rebuilds via `self._table.cell(RowIndex, col)`, which bounds-checks against *logical* `column_count`/`Rows.Count` and does a COM round-trip per cell. On irregular grids this can raise `AnchorNotFoundError` (uncaught — the `except` only catches `ComError`).
- **Suggested fix:** Style each in-hand `com_cells` entry's `Range` directly, or wrap the `cell()` rebuild so an out-of-range physical index degrades to the per-cell hint.

### [TABLES-3] `Cell.merge` docstring wrong: merged cell keeps the upper-left id
- **Severity:** Medium
- **Location:** `src/wordlive/_tables.py:173-191`
- **Problem:** Docstring says the merged cell "keeps *this* cell's id". Word's `Cell.Merge(MergeTo)` collapses to the **upper-left** of the rectangle regardless of receiver, so `cell(2,2).merge(cell(1,1))` is addressed `table:N:1:1`. The fake-COM test only merges adjacent same-row cells where the two coincide.
- **Suggested fix:** Reword to "addressed by the upper-left coordinate"; add a live-validated test merging a non-top-left receiver.

### [LISTS-4] `apply_list_format` orphans a minted `ListTemplate` when a later level spec is invalid
- **Severity:** Medium
- **Location:** `src/wordlive/_lists.py:241-258`
- **Problem:** `ListTemplates.Add(outline)` runs *first*, then each level is configured; `_configure_level` raises `OpError` on a bad spec. A valid level 1 + malformed level 3 leaves an added-but-unapplied template lingering in the document. Validation is interleaved with mutation instead of up-front.
- **Suggested fix:** Validate every level spec into a normalized form before any COM write (mirror `insert_block`'s "resolve all up front").

### [LISTS-5] Bullet/number read-back misclassifies a numbered level with no `%N` placeholder
- **Severity:** Low
- **Location:** `src/wordlive/_lists.py:281`
- **Problem:** `is_bullet = style_int == BULLET or (bool(fmt) and "%" not in fmt)`. A numbered level with a static literal marker (no `%`) reads back as a bullet; the placeholder heuristic is the only discriminator.
- **Suggested fix:** Require a known symbol font as well, or surface the raw `NumberStyle` int.

### [LISTS-6] `_configure_level` flattens outline numbering when `style` set but `format` omitted
- **Severity:** Low
- **Location:** `src/wordlive/_lists.py:185-199`
- **Problem:** Setting `style="arabic"` on level *i* without `format` writes `NumberFormat = f"%{index}."`, so level 3 becomes `"%3."` instead of `"%1.%2.%3."` — silently flattening hierarchical numbering. The synthesized default only makes sense for single-level lists.
- **Suggested fix:** Only synthesize the default for single-level templates; document that multi-level numbering needs explicit `format`.

### [TABLES-7] No test that `delete_column`/`ColumnAnchor` surface `OpError` on a merged table
- **Severity:** Low
- **Location:** `tests/test_tables.py`
- **Problem:** The merged-cell `ComError`→`OpError` re-raise (the central COM gotcha) at `_tables.py:604-612` and `:291-299` is untested; the fake fixture doesn't replicate Word's "mixed cell widths" raise.
- **Suggested fix:** Monkeypatch the fake `Columns(...).Delete`/`.Cells` to raise `ComError`; assert `OpError` with the merged-table hint. Optionally a `@smoke` test.

## Charts

### [CHARTS-1] Error-bar amount-required check misses aliases
- **Severity:** Medium
- **Location:** `src/wordlive/_charts.py:543` (`_ERRORBAR_NEEDS_AMOUNT` `:141`, `ERRORBAR_KINDS` `:122`)
- **Problem:** `needs_amount` is tested against the raw caller string, not the resolved enum. `kind="percentage"`/`"standard_deviation"` (aliases) skip the no-amount validation and fall through to `amt = 1.0`, drawing bogus error bars; canonical `"percent"`/`"stdev"` correctly raise. Behavior diverges by spelling.
- **Suggested fix:** Key off the resolved enum: `needs_amount = etype != XlErrorBarType.STANDARD_ERROR`; drop `_ERRORBAR_NEEDS_AMOUNT`.

### [CHARTS-2] `data_table` toggle unguarded → raw COM error on pie/scatter
- **Severity:** Low
- **Location:** `src/wordlive/_charts.py:410-411`
- **Problem:** `chart_com.HasDataTable = bool(data_table)` has no try/except, unlike the adjacent `gap_width`/`overlap` block. Invalid for pie/scatter → generic HRESULT instead of a friendly message (it is still mapped to `OpError` upstream).
- **Suggested fix:** Wrap and raise `OpError("data_table applies only to category-axis charts")`.

### [CHARTS-3] `format_series(data_labels=False, data_label_size=…)` re-enables labels
- **Severity:** Low
- **Location:** `src/wordlive/_charts.py:515-519`
- **Problem:** `data_labels=False` sets `HasDataLabels=False`, then the font block unconditionally sets `HasDataLabels=True` to access `DataLabels()`, reversing the explicit off toggle.
- **Suggested fix:** If `data_labels is False` and a font field is passed, raise `OpError` or skip the font block.

### [CHARTS-4] No validation guidance on out-of-range `order` / `marker_size` / `explosion`
- **Severity:** Low
- **Location:** `src/wordlive/_charts.py:467-470`, `:509`, `:513-514`
- **Problem:** Docstrings advertise ranges (order 2–6, marker_size 2–72, explosion 0–400) but values pass straight to COM; out-of-range yields a generic HRESULT (still mapped to `OpError`). Polish only — consistent with the module's defer-to-COM stance.
- **Suggested fix:** Optional range checks raising `OpError` with the documented bounds.

> _Confirmed clean:_ All new formatting verbs operate on the static inserted chart only — no `AddChart2`, no workbook access, no series read-back, no orphaned-Excel risk. 1-based indexing and `Xl*` constants verified correct.

## Export (Markdown/HTML + token-budgeted read)

### [EXPORT-1] Underline never detected — `_font_bool` wrong for `Font.Underline`
- **Severity:** High
- **Location:** `src/wordlive/_export.py:405` (call site `:471`)
- **Problem:** `_font_bool(value)` returns `value == -1` (correct for Bold/Italic) but `Font.Underline` is a `WdUnderline` enum (`NONE=0`, `SINGLE=1`). A single-underlined word reports `1`, so `1 == -1` → `False`. Underline is **never** captured; the HTML `<u>` path is dead. The unit test passes only because it builds `Span(underline=True)` directly, bypassing the walk.
- **Why it matters:** `to_html` advertises it "keeps underline"; no underlined text ever renders `<u>` from a live doc.
- **Suggested fix:** Read separately: `underline = int(font.Underline) not in (0, _WD_UNDEFINED)`.

### [EXPORT-2] `read(budget=N)` not actually bounded
- **Severity:** High
- **Location:** `src/wordlive/_export.py:705-765` (`build_digest`), `:684-698` (`_render_body_segment`)
- **Problem:** `budget` only governs the body share. (1) Every heading is emitted verbatim with no cap — `fixed` is subtracted but never used to *limit*; a 500-heading doc blows the budget on the spine alone. (2) The `kept == 0` lead-snippet branch emits a snippet for *every* overflowing body segment regardless of remaining budget, so output grows with section count even at `budget=1`.
- **Why it matters:** The headline contract of `doc.read(budget=N)` is bounding output to avoid context blowout — the exact failure this prevents.
- **Suggested fix:** Either document `budget` as a body-only floor (and update the `_document.py:read` docstring) or enforce a real ceiling across headings + stubs + snippets, collapsing the spine once the hard budget is hit.

### [EXPORT-3] `_render_list_html` emits malformed nesting on skipped levels
- **Severity:** Medium
- **Location:** `src/wordlive/_export.py:321-326`
- **Problem:** A level jump >1 (e.g. 1→3, which Word can produce) opens `<{tag}>` but never an `<li>` to contain the inner `<ul>`, yielding `<ul><ul>…` (invalid); the unwind pairing (`:329-338`) won't symmetrically match. No test covers a >1-level jump.
- **Suggested fix:** Emit a placeholder `<li>` before opening a nested list on skipped levels and mirror on unwind; add a 1→3→1 test.

### [EXPORT-4] Hyperlink word-tagging mis-detects multi-word links
- **Severity:** Medium
- **Location:** `src/wordlive/_export.py:451-475`
- **Problem:** Links are tagged with a point test `s <= word.Start < e`. Word's hyperlink `Range` and per-word `Start` offsets don't always align (field codes, trailing-mark trimming), so multi-word links can split across spans or drop the last word. No smoke test asserts link boundaries.
- **Suggested fix:** Tag by overlap (`word.Start < e and word.End > s`); add a multi-word-link smoke test.

### [EXPORT-5] Table cells lose inline formatting and inline images silently
- **Severity:** Medium
- **Location:** `src/wordlive/_export.py:491-510` (`_table_block`), `:78-91`
- **Problem:** Cell text comes from `table.grid()` as plain strings — no spans, and inline *images* in a cell vanish entirely (not in `image_by_start` lookups, not in grid text), breaking the "every `image:N` stays addressable" invariant. Merged cells can also shift columns with no marker.
- **Suggested fix:** Scan cell ranges against `image_by_start` and append `![alt](image:N)` to cell text; document/marker the merged-cell shift.

### [EXPORT-6] Lead-snippet word accounting approximate
- **Severity:** Low
- **Location:** `src/wordlive/_export.py:622-628`, `:696`
- **Problem:** `rest = b.word_count - snip_words` mixes a raw-text word count with a rendered-Markdown word count (escapes/emphasis markers), so `rest` can be off and emit a spurious `…(para:N, K more words)…` on a fully-shown paragraph.
- **Suggested fix:** Count words off the same plain-text source.

### [EXPORT-7] Markdown link/image targets emitted unescaped
- **Severity:** Low
- **Location:** `src/wordlive/_export.py:165`, `:179-181`, `:226`
- **Problem:** URL/target goes into `](…)` verbatim. A URL with `)` or spaces (SharePoint URLs, `#bookmark` subaddresses) produces a broken Markdown link. HTML side is fine.
- **Suggested fix:** Angle-bracket-wrap or percent-encode targets containing `)`/space.

### [EXPORT-8] `walk_blocks` reads `Style.NameLocal` → locale-dependent headings
- **Severity:** Low
- **Location:** `src/wordlive/_export.py:571`, `:486-488`
- **Problem:** Heading fallback matches the localized `NameLocal` ("Überschrift 1") against `Heading (\d+)`, dead on non-English Word. `OutlineLevel` usually saves it, but the fallback exists for renamed-outline paragraphs.
- **Suggested fix:** Also match the built-in style id (`Style.BuiltIn` + `wdStyleHeading1..9`) or the English `Style.Name`.

> _Test-gap note:_ `tests/test_export.py` is pure-renderer coverage; the risky COM-walk paths (underline, link boundaries, in-cell images, localized headings, deep list jumps) have no unit or smoke coverage — which is why EXPORT-1 ships green.

## Checkpoint & diff

### [CHECKPOINT-1] Blank/duplicate paragraphs share one alignment key → systematic mis-alignment
- **Severity:** High
- **Location:** `src/wordlive/_checkpoint.py:229` (`"key": _sha1(norm)`), align at `:352-354`
- **Problem:** The alignment key is SHA-1 of normalized text only — no positional/structural component. Every blank line (and repeated boilerplate) collides, so `SequenceMatcher` mis-pairs around real edits, emitting spurious insert/delete pairs or attaching a `replace` to the wrong `para:N`. The spec (§5) flags this and offers `track=True`, but that's deferred — so the shipped default mis-aligns on the most common document feature. The docstring (`:67`) and spec also disagree on whether `style` is in the key.
- **Why it matters:** Wrong `anchor_id` / wrong op classification directly defeats the "verify my edits landed" use case.
- **Suggested fix:** Fold `style`/`level` into the key (as the spec originally specified), or loudly document the limitation and add a blank-line-separated test pinning current behavior.

### [CHECKPOINT-6] Table-only edits silently return an empty diff
- **Severity:** High
- **Location:** `src/wordlive/_checkpoint.py:236-242` (tables in `doc_hash`) vs `:351-382` (paragraphs-only diff)
- **Problem:** Table `cells_hash` is folded into `doc_hash`, so a table-only edit breaks the fast path — but `diff_checkpoints` aligns/classifies **paragraphs only** and never compares `a.tables`/`b.tables`. Net: fast path skipped → paragraph diff finds nothing → returns `[]`, positively asserting "no changes" for a doc that changed. Worse than the fast path.
- **Why it matters:** Silent missed change — the highest-severity class for this feature.
- **Suggested fix:** Add a table-comparison pass emitting a coarse `table_change` record when `cells_hash`/`shape` differs. Interim: either don't fold tables into `doc_hash` (consistent `[]` via fast path) or flag that table changes are undetected — the current half-state is the worst option.

### [CHECKPOINT-2] `text+format` checkpoint is O(n²)
- **Severity:** Medium
- **Location:** `src/wordlive/_checkpoint.py:220` → `_format_fingerprint` `:122-136` → `_anchors.py:3468-3478`
- **Problem:** The walk holds the live `para` COM object but `_format_fingerprint` throws it away and calls `doc.anchor_by_id(f"para:{idx}").format_info()`, which re-enumerates `Paragraphs` from 1 to `idx` — quadratic COM traffic over the whole doc.
- **Suggested fix:** Pass the already-iterated `para`/`rng` into the format probe; add an internal `format_info` helper accepting a live Range.

### [CHECKPOINT-3] Field codes / hidden text leak into the fingerprint
- **Severity:** Medium
- **Location:** `src/wordlive/_checkpoint.py:210` → `_anchors.range_text` `:72-92`
- **Problem:** `paragraph_text` returns `rng.Text`, reflecting the live field-code display state. Toggling `ShowFieldCodes` between checkpoints, or an auto-updating `DATE`/`PAGEREF`/TOC field, yields different text for unchanged content → false positives/negatives. "Ignores cosmetic churn" is claimed but field display state isn't normalized.
- **Suggested fix:** Capture with a fixed `TextRetrievalMode` (`IncludeFieldCodes=False`, fixed hidden-text) regardless of live view; add a smoke test toggling field-code display.

### [CHECKPOINT-4] `restyle`/`reformat` misclassified
- **Severity:** Medium
- **Location:** `src/wordlive/_checkpoint.py:362-364`
- **Problem:** Within an `equal` block, `op = "restyle" if style differs else "reformat"`. In `text`/`text+style` mode no format fingerprint is captured, yet a hash diff with equal styles still emits `reformat` — a verdict with no `fmt` data behind it.
- **Suggested fix:** Gate on `include`: only emit `reformat` when `include == "text+format"` and `fmt` actually differs; otherwise `restyle`.

### [CHECKPOINT-5] `changes_since` can't round-trip a `range:` scope
- **Severity:** Medium
- **Location:** `src/wordlive/_checkpoint.py:194-201` (scope stored as `anchor_id` string), `:389` (`within=base.scope`)
- **Problem:** `range:START-END` is re-resolved against the *current* doc. If an upstream edit shifted offsets (exactly what a diff detects), the re-derived range clips a different window → mismatched comparison, no error. Fine for stable semantic anchors.
- **Suggested fix:** Reject offset-only `range:` scopes for `changes_since` (typed error), or pin the scope; document the limitation; add a `within="heading:N"` round-trip test.

### [CHECKPOINT-7] Table fingerprint drops un-readable cells without a sentinel
- **Severity:** Low
- **Location:** `src/wordlive/_checkpoint.py:166-172`
- **Problem:** A cell whose `.Cell(r,c)` raises (merged) is `continue`d with no positional padding, so a same-shape re-merge that leaves the surviving text sequence unchanged hashes identically — a latent false negative for the eventual table diff.
- **Suggested fix:** Append a sentinel (`"\x00skip\x00"`) instead of dropping the cell.

### [CHECKPOINT-8] Test gaps for documented failure modes
- **Severity:** Medium
- **Location:** `tests/test_checkpoint.py`
- **Problem:** No coverage for blank/duplicate-paragraph mis-alignment, table-only change, scoped `changes_since` round-trip, the `_SIM_PAIR_CAP` positional fallback (`:294-301`), or the negative "`reformat` not emitted in `text+style`".
- **Suggested fix:** Add fake-word tests for each; the most likely real-world failures are exactly the untested paths.

### [CHECKPOINT-9] Politeness honored but unguarded
- **Severity:** Low
- **Location:** `src/wordlive/_checkpoint.py:183-250`
- **Problem:** The walk is read-only (no `.Select()`, no scroll) so invariant #1 holds — but nothing tests it; a future Selection-based format probe could regress silently.
- **Suggested fix:** Smoke test asserting Selection/scroll/`Saved` unchanged across `checkpoint()` and `changes_since()`.

## Linter & format regularizer

### [LINT-1] Consistency fix `format_run` writes the whole paragraph range
- **Severity:** High
- **Location:** `src/wordlive/_linting_consistency.py:64`; `Paragraph._range` at `_anchors.py:3489` (whole `Paragraph.Range` incl. pilcrow)
- **Problem:** The font-consistency fix is `{"op":"format_run","anchor_id":"para:N","font|size|bold":…}`. `format_run` writes the full paragraph range, so a `font` (face) fix overwrites a deliberately different face on a sub-run that didn't register as "mixed" for the name field — the opposite of the stated contract ("never disturbs other intentional formatting in the paragraph", `:11-13`).
- **Why it matters:** Violates the targeted-fix safety/idempotency promise; can clobber deliberate per-run formatting.
- **Suggested fix:** Document the whole-paragraph behavior, or restrict the write to exclude the trailing mark and skip a `font` fix when any other character field reads mixed.

### [LINT-2] `list-numbering-continuity` misses number-only / mixed numbered lists
- **Severity:** Medium
- **Location:** `src/wordlive/_linting.py:184`
- **Problem:** Filter is `read_list_info(rng)["type"] in ("numbered","outline")`, but `_LIST_TYPE_NAMES` also has `"number-only"` and `"mixed"`, which suffer the same "N independent 1. lists" footgun and are excluded. The `ListType` Word returns for `apply_list("numbered")` isn't guaranteed to be just `SIMPLE_NUMBERING`.
- **Suggested fix:** Include `"number-only"` (and consider `"mixed"`), or invert to "any type that isn't none/bulleted".

### [LINT-3] `regularize` exec-op path discards `run_batch` failure detail
- **Severity:** Medium
- **Location:** `src/wordlive/_linting.py:351-363`
- **Problem:** The `own_undo=True` path uses `run_batch` (stops at first failure, reports a `failure` dict) but `regularize` copies only `ops_run`/`warnings` and re-raises, losing the partial-progress detail. Atomic undo is preserved; diagnosability regresses.
- **Suggested fix:** Surface `result.get("failure")` into the raised error/report before re-raising.

### [LINT-4] `lint` table-repeat-header rule double-repaginates and swallows all exceptions
- **Severity:** Low
- **Location:** `src/wordlive/_linting.py:133-153`; `location()` at `_anchors.py:3231-3233`
- **Problem:** Two `location()` calls per table → two `Repaginate()` calls (wasteful on large docs; politeness still holds since `Saved` is preserved). Bare `except Exception` around cell reads / `HeadingFormat` silently drops a real defect on any transient COM error.
- **Suggested fix:** Compute first/last page from one `location()` over the table's whole range; narrow the `except` to the typed COM error.

### [LINT-5] Float-rounding used as epsilon; fix writes rounded value as a new override
- **Severity:** Low
- **Location:** `_anchors.py:3044` (`override = eff != sty`), `:338/415` (`round(…, 2)`); fix at `_linting_consistency.py:64,89`
- **Problem:** Spec §7b asked for an epsilon compare; the code rounds to 2 dp instead, then writes the *rounded* `style_val` back as a literal direct override (e.g. 12.0pt when the true style is 11.995pt). Idempotent (read also rounds), harmless today.
- **Suggested fix:** Carry the unrounded style value into the fix, or document the 0.01pt snap.

### [LINT-6] Test gaps + `_overlaps`/`_in_span` duplication
- **Severity:** Low
- **Location:** `tests/test_linting.py`; helpers `_linting.py:90` (`_overlaps`) and `_linting_consistency.py:35` (`_in_span`)
- **Problem:** No test for `within`-scoped consistency rules (which use a *different*, duplicated span helper), multi-rule `regularize` (one undo / all applied), `body-font-consistent` positive, or the `mixed-run-format` skip-on-mixed branch.
- **Suggested fix:** Add the missing tests and unify `_overlaps`/`_in_span` into one helper.

> _Confirmed clean:_ The write path routes through `doc.edit()`/`run_batch`, mutates via anchors (not Selection), and the exec-op path correctly uses `own_undo=False` to stay atomic in a batch.

## Cross-surface consistency

### [SURFACE-1] `word_exec` docstring documents crop field names the op layer rejects
- **Severity:** Medium
- **Location:** `src/wordlive/mcp/server.py:2133`, `:2137` vs `_ops.py:299,320` and apply_op `:776-779`, `:752-753`
- **Problem:** The `word_exec` docstring documents `set_shape_crop`/`set_image_crop` with `crop_left/crop_top/crop_right/crop_bottom`, but the op fields are `left/top/right/bottom` (no `crop_*` alias in `_ops.py`). The `crop_* → *` mapping exists only in MCP `word_write`. A batch author following the docstring gets an "unexpected field" warning, then `set_crop()` with no edges raises.
- **Suggested fix:** Change the `word_exec` docstring crop fields to `left/top/right/bottom` (or add `crop_*` aliases in the op layer — but matching docstring to op is smaller).

### [SURFACE-2] Crop edge params drift: CLI/op vs MCP `word_write`
- **Severity:** Medium
- **Location:** CLI `commands.py:4306-4309`, `:4780-4783` (`--left/--top/--right/--bottom`); `_ops.py:299,320` (`left/top/right/bottom`); MCP `server.py:1030-1042` (`crop_*`)
- **Problem:** Same capability, two parameter vocabularies. MCP's rename is deliberate (avoid confusing crop edges with a shape's `left/top` position), but it's the one capability whose param names don't agree across the op-bearing surfaces, and it feeds SURFACE-1.
- **Suggested fix:** Either accept the divergence and make every schema explicit, or add `crop_*` aliases in the op layer so all surfaces accept the same names.

### [SURFACE-3] Border line-style param differs three ways
- **Severity:** Low
- **Location:** CLI `commands.py:5944`, `:6654` (`--style`); op `_ops.py:331,429` (`style`, alias `line_style`); MCP `server.py:444-455,910-915` (`line_style`)
- **Problem:** Naming drift for one capability, but handled gracefully (op accepts both). MCP rename avoids colliding with `apply_style`/`table create`'s `style`.
- **Suggested fix:** No code change needed; optionally note the alias in CLI `--style` help.

### [SURFACE-4] `XlErrorBarInclude` docstring names a non-existent `direction=` param
- **Severity:** Low
- **Location:** `src/wordlive/constants.py:862-863` vs `_anchors.py:4162-4188`
- **Problem:** Docstring says `add_error_bars(direction=...)`; the param is `include` (axis selector is `axis`). Constant used correctly; docstring only.
- **Suggested fix:** Change docstring to `add_error_bars(include=...)`.

### [SURFACE-5] Error-bars `axis` choices differ between CLI and MCP docs
- **Severity:** Low
- **Location:** CLI `commands.py:4151` (`y|value|x|category`); MCP/`word_exec` docstrings `server.py:1852,2130` (`y|x`)
- **Problem:** Functionally equivalent (op passes through; `_anchors.py:4177-4179` accepts synonyms) but MCP help hides the `value`/`category` synonyms.
- **Suggested fix:** Align MCP docstrings to `axis=y|value|x|category` (or trim CLI choices).

> _Confirmed clean:_ No missing surfaces, no broken anchor handling, no undo-unsafe exec ops, no exit-code misuse; `__all__` is complete for every new feature class.
