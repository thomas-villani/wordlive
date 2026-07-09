# Linter + formatting regularizer — design sketch

Status: **Batch 6 — the first `adds_content` opt-in fixes shipped (Unreleased,
2026-07-08)**; Batch 5 + the gate before it — design was a **sketch** (2026-06-17).
Roadmap home: `feature-plan.md` Part II, Priority 1, item 1. This is the detailed
design; the roadmap keeps the one-paragraph summary.

## Progress

**Shipped** (all Unreleased; per-batch detail lives in `CHANGELOG.md`, and each
§5b cluster below carries an inline ✅). Every slice was live-probed against Word 16.

| Slice | What landed |
|---|---|
| Foundation (§10 steps 1–4) | `anchor.format_info()` + direct-override detection (§7); the 3 structural rules (`heading-keep-with-next`, `table-repeat-header`, `list-numbering-continuity`) + heading/font/spacing consistency rules + `mixed-run-format`; `doc.lint()` / `doc.regularize()` with the targeted-idempotent default fix + idempotency test; wired Python / CLI / `regularize` exec op / MCP |
| Batch 1 — typography hygiene (§A) | 10 P2 text-scan rules (`_linting_typography.py`); enablers `find_replace` `literal`/`regex` modes + `required=False`, and `Rule.default_on` |
| Batch 2 — finalization (§G) | 6 review-state rules (`_linting_finalization.py`), all off-by-default `finalization` tag; `format_info` gained `font.hidden` / `font.highlight` |
| Batch 3 + 3b — field-code backbone (§C) | `Range.Fields` walk (P1): `broken-cross-reference` + `caption-manual-numbering` (on, `academia`), `page-numbers-present` (off, `layout`), `xref-as-literal-text` (3b, off, heuristic) |
| Batch 4a — profiles + policy (§6) | `profile=` + the `Profile` loader; `body-justified` / `body-line-spacing` / `table-numeric-right-align`; `Rule.check(doc, span, profile)` |
| Batch 4b — hyperlinks (§I) | 3 rules over `doc.hyperlinks` (`_linting_hyperlinks.py`); no new read surface |
| Batch 4c — layout / document-level (§H) | 5 rules (`_linting_layout.py`) + a new `doc.watermark()` read (`WatermarkInfo`) |
| Batch 5 — heading & document structure (§B) | 6 rules over `doc.outline()` (`_linting_headings.py`); `heading-trailing-period` fixable (in-place strip), the new `structure` tag; no new read surface |
| `adds_content` gate (§8) | `Finding.adds_content` field + `regularize(allow_content=…)` surface gating; content-changing fixes withheld into a new `deferred` report bucket by default; wired Python / CLI (`--allow-content`) / exec op / MCP. The infrastructure the content fixes plug into |
| Batch 6 — first `adds_content` opt-in fixes | 5 report-only rules wired fixable + 1 new rule, all `adds_content=True`, pure composition (no new COM write): `draft-watermark-present`→`remove_watermark`; `page-numbers-present`→`{ PAGE }` in `footer:1:primary`; `confidentiality-notice`/`copyright-notice`→notice into `footer:1:primary`; `hyperlink-bare-for-print`→fold URL into display text; **new** `stray-empty-paragraph` (off, `whitespace` tag)→`delete_paragraph`. `regularize` applies deletes last, descending, so a multi-blank pass keeps earlier anchors valid |

Two standing decisions: `stale-fields` stays a **report-only nudge** (Word exposes
no staleness flag, so a presence-based refresh can't be idempotent), and the
**content-adding / repair fixes are now shipping rule-by-rule** — Batch 6 wired the
first six into the `adds_content` gate (watermark strip, page number, both notices,
bare-hyperlink fold, stray-blank delete). The rest stay report-only for concrete
reasons (see Remaining item 1). `heading-trailing-period` (Batch 5) remains the only
*ungated* fix beyond the foundation/typography set, because stripping a period is a
pure in-place edit. So the report-only rules across Batches 2–5 are being closed out,
not forgotten.

**Remaining** (priority order; the full rule catalogue is §5b):

1. **Close the loop — make the report-only rules fixable.** The gate (§8) and the
   first six fixes have **shipped** (Batch 6): strip watermark, insert page number,
   insert confidentiality/copyright notice, fold `(url)` into a bare hyperlink, and a
   new `stray-empty-paragraph` delete — each reuses an existing write verb, is
   idempotent, and flags `adds_content=True`. **Still deferred, for cause:**
   `document-properties-filled` (the fix needs a *value* the linter doesn't have — a
   profile-schema change, see item 2's `house_style`), `caption-manual-numbering`
   (rebuilding the number as a `SEQ` field in place is a fragile edit, not one existing
   verb), `figure-caption-present` (rule not yet written — needs an image/table-adjacency
   detection walk), and `broken-cross-reference` / `xref-as-literal-text` (repairing a
   `REF` needs a human-chosen target). These stay report-only until their blocker clears.
2. **`house_style` half of §6** — pin consistency-rule targets to named style values
   and fix by updating the style (`set_style`) — the brand/template path.
3. **Batch 1b typography heuristics** — `straight-quotes`, `nbsp-missing`,
   `sentence-spacing-consistent` (deferred from Batch 1 as false-positive-prone).
4. **Detection remainders** — §D citations cluster (`citation-as-literal-text`,
   `footnote-numbering-manual`, `mixed-citation-styles`, `orphan-citation`); §E
   `table-empty` / `table-overflows-margin`; §F `justify-misapplied` /
   `paragraph-too-long`; §C `caption-label-consistent` / `caption-position-consistent`;
   and `header-footer-consistent`'s cross-section **format** comparison (text-only shipped).
5. **Advanced / later** — the opt-in aggressive `Font.Reset()` strip-to-style fix
   (§7c); the `docx-plus` cascade-provenance hybrid (§7c); accessibility rules + the
   `prepare-for-sharing` product (§9); a custom-rule plugin API (only on a concrete need).
   These three (accessibility, custom rules, and the new **exemplar-driven** "extract a
   profile from / match against a reference document") are fleshed out in **§11 — Future
   ideas (post-v1.0)**.

> Audit a document for publishing-quality defects (`doc.lint()`), then autofix the
> mechanical ones in one atomic-undo step (`doc.regularize()`). Pure composition
> over shipped write verbs — **no new COM write surface**; the new work is a richer
> *format-probe read* and the rule engine.

---

## 1. Why this, why now

Every document hand-off involves the same tail of tedious, mechanical fixes —
dangling headings, a table that breaks across a page with no repeating header,
one `Heading 1` that's 15pt instead of 16, a numbered list Word silently split
into five "1." lists, a numeric table column left-aligned. They are:

- **objective** (you can write the rule down),
- **mechanical** (the fix is deterministic), and
- **already expressible** in wordlive's verbs (`format_paragraph`,
  `set_heading_row`, `apply_list`, `apply_style`, table cell alignment, …).

That combination is exactly what an agent should own. The linter is the
highest-utility next feature precisely because it's *composition*, not new COM.

## 2. The core reframing — consistency = "no direct formatting fighting the style"

Word documents have two formatting layers: **styles** (named, inheritable) and
**direct formatting** (per-range overrides on top of the style). Professional
documents are style-driven; the defects we keep fixing are almost always *direct
overrides that drifted from the style*. So most "consistency" rules become one
detection: **does this paragraph/run carry a direct override that deviates from
its applied style?** — and the fix is to bring it back to the style.

This makes the rules objective and the fixes idempotent, and it splits the
catalogue into three **kinds**:

| Kind | Needs config? | Detection | Example |
|---|---|---|---|
| **consistency** | no | direct override ≠ applied style | a `Heading 1` at 15pt; mixed body fonts |
| **structural** | no | objective defect in layout/structure | split numbered list; dangling heading; table broken with no repeat header; missing caption |
| **policy** | yes (a profile) | value ≠ the profile's target | body must be justified; numeric columns right-aligned |

Consistency + structural rules ship with sensible defaults and need no
configuration. Policy rules are opt-in and read their target from a **profile**
(§6).

## 3. Surface

```python
findings = doc.lint(rules=None, within=None, profile=None)      # pure read
report   = doc.regularize(rules=None, within=None, profile=None, dry_run=False)
```

- **`doc.lint(...)`** → a list of findings (`Finding`, §4). Read-only: snapshots
  nothing, mutates nothing, leaves `Saved` untouched (it *does* repaginate for the
  page-layout rules, like `stats()`/`location()` already do — content-neutral).
- **`doc.regularize(...)`** → applies the **fixable** subset of the matched
  findings inside a single `doc.edit("Regularize formatting")` (one Ctrl-Z reverts
  the whole pass), and returns `{applied: [...], skipped: [...], findings: [...]}`.
  `dry_run=True` runs detection + plans fixes but writes nothing (equivalent to
  `lint` plus the planned fix for each).
- **`rules`** selects/deselects by id or tag (`["headings", "lists"]`,
  `{"exclude": ["body-justified"]}`); `None` = the default rule set (all
  consistency + structural; no policy rules unless a `profile` enables them).
- **`within=anchor`** scopes both to any anchor's range (a heading's
  `section_range()`, a `range:`, a table) — "regularize just this section."

### Surfaces (all four must agree)

- **Python:** `doc.lint` / `doc.regularize`.
- **CLI:** `wordlive lint [--rules …] [--profile …] [--within ID]` (JSON findings)
  and `wordlive regularize [--dry-run] …`.
- **exec op:** `regularize` is a **write** op (so it joins an atomic batch);
  `lint` stays a read (CLI/MCP only, no op — like `stats`/`proofing`).
- **MCP:** `word_read command=lint`, `word_write command=regularize`.

## 4. The `Finding` shape

```jsonc
{
  "rule": "heading-keep-with-next",   // stable id
  "kind": "structural",               // consistency | structural | policy
  "severity": "warning",              // error | warning | info
  "anchor_id": "heading:7",           // where (a real anchor id)
  "message": "Heading 'Methods' may dangle at a page foot (keep-with-next off).",
  "fixable": true,
  "fix": {                            // present iff fixable; what regularize will do
    "op": "format_paragraph",
    "args": {"anchor_id": "heading:7", "keep_with_next": true}
  },
  "observed": "keep_with_next=false", // optional, for the report
  "expected": "keep_with_next=true"
}
```

`fix.op`/`fix.args` are literally an **exec op** — so `regularize` is "lint, then
run each finding's `fix` op through the existing `apply_op` loop." That keeps the
fix path on the audited, warning-emitting op vocabulary instead of a parallel
write path, and a caller can inspect/serialize exactly what will change.

## 5. The rule catalogue (v1)

Mapped to the recurring hand-edits. Each row: how it's **detected** (the COM read)
and how it's **fixed** (the wordlive verb / exec op). All fixes are idempotent
(re-running is a no-op) unless noted.

### Headings & paragraph spacing

| id | kind | detect | fix |
|---|---|---|---|
| `heading-keep-with-next` | structural | a heading paragraph with `KeepWithNext` off | `format_paragraph(keep_with_next=True)` |
| `heading-widow-orphan` | structural | `WidowControl` off on a heading/body para | `format_paragraph(widow_control=True)` |
| `heading-spacing-consistent` | consistency | a heading's `SpaceBefore`/`SpaceAfter` ≠ its style's | clear the override → style value (§7) |
| `body-line-spacing` | policy | `LineSpacingRule`/`LineSpacing` ≠ profile target | `format_paragraph(line_spacing=…)` |
| `stray-empty-paragraph` | structural | an empty `Normal` paragraph between blocks | `delete_paragraph` (✅ **shipped Batch 6** — off by default, `whitespace` tag; the fix is `adds_content` so the gate withholds it; deletes ordered descending) |
| `double-space` | consistency | runs of 2+ spaces in body text | `find_replace` collapse (skip code/verbatim styles) |

### Font / character consistency

| id | kind | detect | fix |
|---|---|---|---|
| `heading-font-consistent` | consistency | same-style headings whose `Font.Name`/`Size`/`Bold` carry direct overrides deviating from the style | re-apply style value / clear override (§7) |
| `body-font-consistent` | consistency | body paragraphs with a directly-set font face differing from the `Normal`/body style | clear override |
| `mixed-run-format` | consistency | a paragraph whose `Font.Size`/`Name` reads `wdUndefined` (mixed runs) where uniformity is expected | report-only (which run is the outlier needs run-walk; fix is opt-in) |

### Alignment & justification (policy) — ✅ shipped Batch 4a (with `body-line-spacing`)

| id | kind | detect | fix |
|---|---|---|---|
| `body-justified` | policy | body paragraphs not `wdAlignParagraphJustify` | `format_paragraph(alignment="justify")`, scoped to body styles |
| `table-numeric-right-align` | policy/heuristic | a table column whose non-empty body cells nearly all parse as numbers (`$`, `%`, `,`, `(neg)`) but aren't right-aligned | per-cell `format_paragraph(alignment="right")` |

### Lists

| id | kind | detect | fix |
|---|---|---|---|
| `list-numbering-continuity` | structural | a **contiguous** run of numbered paragraphs Word split into independent lists (each restarts at 1 / distinct `ListFormat.List`) — the documented "N independent 1. lists" footgun | `remove_list` the span then `apply_list("numbered")` over the single `range:` (the documented repair) |
| `list-false-continue` | structural | two lists separated by non-list content that share numbering (should restart) | `restart_numbering` at the second list's head |
| `list-bullet-consistent` | consistency | sibling bullet items at one level using different bullet chars / `ListTemplate` | re-apply the level's template |
| `list-indent-consistent` | consistency | sibling items at one logical level with differing `LeftIndent` | normalize to the level indent |

### Tables & captions

| id | kind | detect | fix |
|---|---|---|---|
| `table-repeat-header` | structural | a table spanning >1 page (`location().page != end_page`) whose row 1 isn't a heading row | `set_heading_row(1)` |
| `figure-caption-present` | structural | an inline image / table with no adjacent `Caption`-styled paragraph or `SEQ` field | **opt-in fix:** `insert_caption(label=…)` with an empty title placeholder to fill (adds content — off by default, report-only) |
| `caption-style-consistent` | consistency | a caption paragraph not on the `Caption` style | `apply_style("Caption")` |

## 5b. Catalogue v2 — brainstormed backlog (2026-06-19)

The v1 catalogue (§5) only needed the `format_info()` override probe. The backlog
below — gathered from a publishing/academia pass — pushes into **four new
detection primitives**. Each primitive unlocks a *cluster*, so we **batch by
primitive**, not by category. Build the primitive once, light up its rules.

**Default stance (decided 2026-06-19):** new **policy / opinion** rules ship
**off unless tagged-in** — consistent with §2 (policy needs a profile). Even
opinion-flavored *consistency* rules (sentence-spacing, em-dash, justify-on-short)
default **off**; the user enables a cluster by **tag** (`--rules academia`) or a
profile. Structural rules that are unambiguous defects (broken field, leftover
comments) stay **on** by default like the v1 structural set. Anything that
**adds content** or is **loud/irreversible** is **report-only** with an opt-in
fix flag (new Finding field `adds_content: bool`, gated by surfaces — same
treatment v1 gives `stray-empty-paragraph` / `figure-caption-present`).

### Detection primitives (build order for the backlog)

| Primitive | COM surface | Unlocks |
|---|---|---|
| **P1 · Field-code walk** | `Range.Fields` (SEQ, REF/PAGEREF, PAGE, TOC, HYPERLINK, CITATION) | caption-as-reference, xref-as-text, broken-ref, page-numbers, stale-fields, citation rules |
| **P2 · Run-walk / text scan** | `Range.Words` + wildcard `find` | manual-heading, typed-manual-lists, space-before-punct, em-dash, en-dash ranges, curly quotes |
| **P3 · Revision / markup state** | `Document.Revisions`, `.Comments`, `TrackRevisions` | leftover comments, unaccepted changes, track-changes-on |
| **P4 · Section / header-footer walk** | `_sections.py` (have it) + `Range.Fields` | page-numbers, confidentiality / copyright notice, header-footer consistency |

### A. Whitespace & typography hygiene  *(P2; cheap, high-frequency)* — ✅ shipped (Unreleased)

Shipped as `_linting_typography.py`: `trailing-whitespace`, `leading-whitespace`
(strip-only), `space-before-punctuation`, `double-space` are **on**;
`sentence-spacing-consistent`, `tabs-for-layout`, `manual-line-break`,
`nbsp-missing`, `straight-quotes`, `hyphen-as-range`, `em-dash-usage` are **off**
(tag/profile) — of which `straight-quotes`, `nbsp-missing`,
`sentence-spacing-consistent` were **deferred to Batch 1b** and the rest landed.
Fixes are regex-mode `find_replace` (see the status header).

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `sentence-spacing-consistent` | consistency | dominant 1-vs-2 spaces after `.?!`; flag the minority | `find_replace` to dominant | off (tag) |
| `trailing-whitespace` | structural | para text ends in space/tab | trim | on |
| `leading-whitespace` | structural | para starts with spaces/tabs used as indent | clear → real indent | on |
| `space-before-punctuation` | consistency | ` ,` ` .` ` ;` ` :` ` )` | collapse | on |
| `tabs-for-layout` | consistency | 2+ consecutive tabs / tab runs mid-para | report-only | off (tag) |
| `manual-line-break` | structural | `Chr(11)` Shift-Enter where a paragraph break belongs | report-only | off (tag) |
| `nbsp-missing` | policy | space in `Figure 3`, `5 km`, ` %`, before units/refs | insert nbsp | off (tag) |
| `straight-quotes` | consistency | `'`/`"` where the doc is otherwise curly (skip code styles) | smart-quote replace | off (tag) |
| `hyphen-as-range` | consistency | `1990-1995`, `pp. 10-15` using hyphen not en-dash | replace en-dash | off (tag) |
| `em-dash-usage` | policy | `—` present (the "AI tell") | report / optional `--` | off (tag) |

(`double-space`, `stray-empty-paragraph` already in v1.)

### B. Heading & document structure  *(P2 + outline walk)* — ✅ shipped (Unreleased)

Shipped as `_linting_headings.py` in **Batch 5** — six rules over `doc.outline()`.
`heading-level-skip` and `empty-heading` ship **on** (unambiguous outline defects);
the other four ship **off** behind the `headings` / `structure` tags. Five are
report-only; `heading-trailing-period` is **fixable** (a paragraph-scoped
`find_replace` regex strips the period in place — idempotent, no `adds_content` gate,
and it must scope to `para:N` not `heading:N`, since `find_replace` expands a heading
scope to its body section). `toc-present-and-current` is **presence-only** (Word
exposes no field-staleness flag — the same limit `stale-fields` hit — so the "current"
half stays a report). `manual-heading-formatting` (the first row) shipped earlier in
**Batch 1**.

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `manual-heading-formatting` | structural | a bold/large `Normal` para that looks like a heading but isn't styled | report → suggest `apply_style("Heading N")` | on (report) — Batch 1 |
| `heading-level-skip` | structural | outline jumps H1→H3 with no H2 | report-only | on (report) ✅ |
| `heading-numbering-manual` | consistency | heading text starts with literal `3.1` not list-numbered | report | off (tag) ✅ |
| `heading-trailing-period` | consistency | heading text ends with `.` | strip | off (tag) ✅ |
| `empty-heading` | structural | heading paragraph with no text | report | on (report) ✅ |
| `adjacent-headings` | structural | two headings, no body between | report | off (tag) ✅ |
| `toc-present-and-current` | structural | doc has Heading 1s but no TOC field / TOC stale | report (presence-only) | off (tag) ✅ |

### C. Captions & cross-references  *(P1 — the academia backbone)* — ⚙️ mostly shipped (Unreleased)

Shipped as `_linting_fields.py`: `broken-cross-reference` (on) and
`caption-manual-numbering` (on, report) landed in **Batch 3**, both report-only and
tagged `academia`. `xref-as-literal-text` landed in **Batch 3b** — report-only and
**off by default** (tags `crossref` / `academia`), deviating from the "on" column below
because it's heuristic/false-positive-prone. `caption-label-consistent` /
`caption-position-consistent` remain backlog.

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `caption-manual-numbering` | structural | a `Caption` para whose number is **literal text**, not a `SEQ` field | report → rebuild with SEQ (adds content, opt-in) | on (report) |
| `xref-as-literal-text` | structural | `see Figure 3` / `Table 2` typed as text, not a `REF` field | report (auto-fix needs target match) | off (tag; heuristic — shipped Batch 3b) |
| `caption-label-consistent` | consistency | mix of `Fig.`/`Figure`, `Table`/`Tbl`, `Eq.`/`Equation` | normalize label | off (tag) |
| `caption-position-consistent` | consistency | some figure captions above, some below the image | report | off (tag) |
| `broken-cross-reference` | structural | `REF`/`PAGEREF` rendering `Error! Reference source not found` | report | on |

### D. Citations & bibliography  *(P1; deep-academia, stageable)*

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `citation-as-literal-text` | structural | `(Smith 2020)` typed, no `CITATION` / reference-manager field | report | off (tag) |
| `footnote-numbering-manual` | structural | footnote refs typed as superscript text, not real footnotes | report | off (tag) |
| `mixed-citation-styles` | consistency | numeric `[1]` and author-date `(Smith, 2020)` both present | report | off (tag) |
| `orphan-citation` | structural | cited key absent from bibliography (needs parse) | report | off (later) |

### E. Tables  *(extends v1)*

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `table-style-consistent` | consistency | tables using different `Table.Style`; flag minority vs dominant | `apply` dominant style | on |
| `table-empty` | structural | table with all-empty cells | report | off (tag) |
| `table-overflows-margin` | structural | `PreferredWidth` / right edge > text width | report | off (tag) |

### F. Alignment & justification

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `justify-misapplied` | consistency | `wdAlignParagraphJustify` on a heading, list item, or short/one-line para (gappy last line) | clear → left / style default | off (tag) |
| `paragraph-too-long` | policy | single para spans > ½ page (`location()` page geometry or char threshold) | report-only | off (tag) |

(`body-justified`, `table-numeric-right-align` already in v1, policy.)

### G. Review-leftover & finalization hygiene  *(P3; "is this actually final?")* — ✅ shipped (Unreleased)

Shipped as `_linting_finalization.py`. **All six ship off by default** (the
`finalization` tag), not the per-rule defaults below — decided this pass: it's an
opt-in pre-send check, since a mid-authoring doc normally carries comments/
revisions. `stale-fields` is a **report-only nudge** (Word has no staleness flag,
so a presence-based `update_fields` fix can't be idempotent — the fixable version
waits for Batch 3's field-code backbone). `leftover-highlight` is the only fix.

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `unaccepted-revisions` | structural | `len(doc.revisions) > 0` | report (accept is loud) | off (tag) |
| `track-changes-on` | structural | `doc.track_changes == True` | report | off (tag) |
| `comments-present` | structural | `len(doc.comments) > 0` (+ `.done`) | report | off (tag) |
| `leftover-highlight` | consistency | `format_info().font.highlight` ≠ none | clear highlight | off (tag) |
| `hidden-text-present` | structural | `format_info().font.hidden` runs | report | off (tag) |
| `stale-fields` | structural | `doc.fields`, kind ∈ TOC/SEQ/REF/PAGE… | report nudge (fix → Batch 3) | off (tag) |

Two `format_info` fields landed with the batch — `font.hidden` (a Font property,
full override detection) and `font.highlight` (a Range property, so effective-only,
no style baseline) — the read mirror of `format_run`'s highlight/hidden writes.

These cluster as a coherent **`finalization`** tag — useful as a standalone
"is-this-ready-to-send?" check (and a building block for *prepare-for-sharing*).

### H. Page layout & document-level  *(P4)* — ✅ shipped Batch 4c (Unreleased)

Shipped as `_linting_layout.py` (`page-numbers-present` shipped earlier in Batch 3).
All **off by default**; the notice pair also carries the new `notices` tag. Detection
landed here; **Batch 6 then wired the opt-in fixes** (`adds_content=True`) for
`draft-watermark-present` (→`remove_watermark`), `page-numbers-present` (→`{ PAGE }` in
`footer:1:primary`), and both notices (→notice text into `footer:1:primary`).
`document-properties-filled` (needs a value) and `header-footer-consistent` stay
report-only — the latter scoped to **text** for v1 (cross-section *format* comparison
deferred — heuristic + not exercisable by the fake fixture). Detection reuses
`doc.properties` / `doc.sections` plus the new `doc.watermark()` read.

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `page-numbers-present` | policy | no `PAGE` field in any footer/header | insert (adds content, opt-in) | off (tag `layout`; shipped Batch 3) |
| `confidentiality-notice` | policy | profile-supplied text not found in H/F or body | report (insert opt-in) | off (profile; tags `layout`/`notices`) |
| `copyright-notice` | policy | profile `©` / text not present | report | off (profile; tags `layout`/`notices`) |
| `header-footer-consistent` | consistency | primary H/F **text** differs across the document's own (non-linked) sections | report | off (tag `layout`) |
| `document-properties-filled` | policy | a required built-in prop (default Title / Author) empty | report (set needs a value — opt-in) | off (tag `layout`) |
| `draft-watermark-present` | structural | a text watermark (DRAFT / CONFIDENTIAL) still present | report | off (tags `layout`/`finalization`) |

### I. Hyperlinks  *(print / sharing)* — ✅ shipped Batch 4b (Unreleased)

Shipped as `_linting_hyperlinks.py`, a thin walk over `doc.hyperlinks` (no new read
surface). `hyperlink-broken-internal` is **on**; the two print/sharing rules are
**off**, behind the `hyperlinks` / `print` tags. **Batch 6 wired** the opt-in fix
(`adds_content=True`) for `hyperlink-bare-for-print` — fold the URL into the display
text as `label (url)` via `set_hyperlink` (targeting the link by positional index);
the other two stay report-only (a broken jump / a bare-URL label need a human's call).

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `hyperlink-bare-for-print` | policy | hyperlink display text ≠ target URL (URL invisible on paper) | report / append `(url)` (opt-in) | off (tags `hyperlinks`/`print`) |
| `hyperlink-broken-internal` | structural | internal `HYPERLINK \l anchor` with no matching bookmark | report | on |
| `hyperlink-display-is-raw-url` | consistency | raw URL shown as the link's whole label where a label is wanted | report | off (tags `hyperlinks`/`print`) |

### Tag taxonomy

Rules carry **tags** so a user enables a *cluster* instead of naming ids. Proposed
top-level tags: `typography`, `headings`, `lists`, `tables`, `captions`,
`crossref`, `citations` (alias `academia` = captions + crossref + citations +
nbsp + en-dash), `finalization`, `layout`, `notices` (the confidentiality /
copyright notice pair, a sub-cluster of `layout`), `print`, `accessibility`.
`--rules academia` / `--rules finalization` become the headline ergonomics;
profiles (§6) toggle tags + supply policy targets.

### Suggested batch order (primitive-driven)

Batches 1–4c have shipped — see the **Progress** table up top for what each landed,
and the ✅ on each cluster below for which rules. The primitive-driven ordering that
guided them, and what it implies next:

1. **Typography (§A, P2)** → **Batch 1** ✅ — highest hit-rate, cheapest, no field plumbing.
2. **Finalization (§G, P3)** → **Batch 2** ✅ — reused the revision/comment/field wrappers.
3. **Field-code backbone (§C, P1)** → **Batch 3 / 3b** ✅ — the `Range.Fields` walk; the academia centerpiece.
4. **Profiles + §H/§I (P4 + §6)** → **Batch 4a** (profiles/policy) ✅, **4b** (hyperlinks §I) ✅, **4c** (layout §H + `doc.watermark()`) ✅.
5. **Heading & document structure (§B, P2 + outline)** → **Batch 5** ✅ — six rules over `doc.outline()`; `heading-trailing-period` fixable, the new `structure` tag.
6. **First `adds_content` opt-in fixes** → **Batch 6** ✅ — five report-only rules wired fixable + the new `stray-empty-paragraph`, the first fixes to populate the gate's `deferred` bucket.
7. **Later** — citations (§D), then the accessibility sub-product with *prepare-for-sharing* (§9).

Cross-cutting (not a detection batch): the `adds_content` gate (✅ shipped) whose first
fixes landed in Batch 6 — the rest close out rule-by-rule (Remaining item 1) — and the
`house_style` half of §6, still in **Remaining**, up top.

## 6. Profiles (policy rules + house style)

**Status: the policy half shipped in Batch 4a** (`_lint_profile.py` — `Profile.load`
accepts a path / dict / `None`; `body-justified` / `body-line-spacing` /
`table-numeric-right-align` are its first consumers). The **`house_style`** half below
(pinning consistency targets + `set_style` fixes) is **deferred** to a later pass.

A **profile** is a small declarative config that (a) enables policy rules and (b)
supplies their targets — and optionally pins consistency targets to an explicit
house style instead of the document's own applied styles.

```jsonc
// wordlive.lint.json  (or passed inline)
{
  "extends": "default",
  "rules": {
    "body-justified":          { "enabled": true,  "severity": "warning" },
    "body-line-spacing":       { "enabled": true,  "target": "1.5" },
    "table-numeric-right-align": { "enabled": true, "threshold": 0.8 },
    "double-space":            { "enabled": false }
  },
  "house_style": {            // optional: the canonical values for consistency rules
    "Heading 1": { "font": "Calibri Light", "size": "16pt", "space_before": "12pt" },
    "Normal":    { "font": "Calibri", "size": "11pt" }
  }
}
```

Without a profile: consistency rules judge each paragraph against **its own applied
style** (internal consistency), structural rules run, policy rules stay off. With a
`house_style`, consistency rules judge against the named targets *and* can fix by
updating the style definition (`set_style`) so the whole document follows — the
brand/template path. CLI `--profile PATH`; discoverable default file name so a repo
can check one in.

## 7. The format-read control + direct-override detection

Two things land here: a **new public read control** (the read mirror of the
existing `format_paragraph`/`format_run` write verbs — wordlive has the write side
but no read side today), and the **direct-override detection** the consistency
rules are built on. The detection mechanic was **live-validated 2026-06-17**
(throwaway doc, closed unsaved) — see the confirmed results below.

### 7a. `anchor.format_info()` — the public read mirror (new surface)

A read that returns an anchor's *effective* paragraph + character formatting, each
field annotated with whether it's a **direct override** vs inherited from the
applied style, plus the style's baseline value. Useful well beyond the linter
("what formatting is actually on this paragraph, and what's overriding the
style?"), and the substrate every consistency rule consumes.

```jsonc
// doc.anchor_by_id("heading:7").format_info()
{
  "anchor_id": "heading:7",
  "style": "Heading 1",
  "paragraph": {
    "alignment":    {"value": "left",  "style": "left",  "override": false},
    "space_before": {"value": "0pt",   "style": "0pt",   "override": false},
    "space_after":  {"value": "6pt",   "style": "0pt",   "override": true},
    "line_spacing": {"value": "1.15",  "style": "1.15",  "override": false},
    "keep_with_next": {"value": false, "style": false,   "override": false}
  },
  "font": {
    "name": {"value": "Aptos",     "style": "Aptos", "override": false},
    "size": {"value": 15.0,         "style": 12.0,    "override": true},
    "bold": {"value": true,         "style": true,    "override": false},
    "mixed": ["size"]   // fields that read wdUndefined (vary across runs)
  }
}
```

- **Read-only.** Same vocabulary as the write verbs (alignment / indents / spacing
  / `keep_*` / widow; font name/size/bold/italic/…), so read and write mirror each
  other field-for-field.
- `value` = effective, `style` = the applied style's resolved baseline, `override`
  = `value ≠ style`. A field listed in `mixed` reads `wdUndefined` (varies across
  runs within the paragraph) — surfaced explicitly, never as a bogus number.
- Surfaces: `anchor.format_info()` Python; CLI `read format --anchor-id ID`; MCP
  `word_read command=format_info`. **No exec op** (pure read).
- Optionally extend `doc.paragraphs.list()` with a compact `overrides: [...]` field
  so a bulk audit needs one read, not one per paragraph.

The **write** side already exists and is unchanged: `format_paragraph` /
`format_run` / `Style.format_paragraph` / `Style.format_run`. This item only adds
the missing reads.

### 7b. Direct-override detection (live-validated)

For a paragraph `p`: compare effective vs style and flag where they differ.

- effective: `p.Range.Font.Size`, `…Font.Name`, `p.Range.ParagraphFormat.SpaceBefore`, …
- style: `p.Range.ParagraphStyle.Font.Size`, `….ParagraphFormat.SpaceBefore`, …
- a direct override exists where **effective ≠ style**.

**Confirmed live (2026-06-17, Word 16 / Aptos default template):**

1. ✅ **Comparison works.** Clean `Heading 1`: effective `Font.Size` 12.0 == style
   12.0, `Font.Name` Aptos == Aptos, `SpaceBefore` 0 == 0. After a direct
   `Font.Size = 15`: effective 15.0 vs style **still 12.0** → override correctly
   detected.
2. ✅ **`wdUndefined` (9999999)** is exactly what a mixed-run paragraph returns for
   `Font.Size` — special-cased into `mixed` (the `mixed-run-format` rule), never
   treated as a number.
3. ✅ **Built-in styles resolve concrete values** — `Normal.Font.Size` → 12.0,
   `Heading 1.BaseStyle` → `Normal`; no need to walk `BaseStyle` by hand for the
   common case (Word resolves it). Keep an epsilon compare for float points.
4. ✅ **`Font.Reset()` strips to style** (15.0 → 12.0) — backs the opt-in aggressive
   fix below.
5. **Scope:** character props (`Font.*`) vary *within* a paragraph → read at the
   paragraph range, accept `wdUndefined` as the mixed signal; paragraph props
   (`ParagraphFormat.*`) are whole-paragraph and simpler.

### 7c. Borrowing from `docx-plus` (cascade provenance)

The sibling library `../docx-plus` has a mature OOXML cascade resolver
(`styles/inspect.py: resolve_effective_formatting` → `ResolvedFormatting` +
per-field `FormattingSource` provenance). It operates on `.docx`-on-disk, **not**
COM, so it isn't reused for effective values (Word's COM already resolves the
8-layer cascade — `Range.Font.Size` *is* the resolved value). Three concrete
borrowings:

- **Field schema parity.** `format_info()` (§7a) should mirror
  `ResolvedFormatting`'s field names (alignment / indents / spacing /
  `line_spacing_rule` / `keep_with_next` / `keep_lines` / `page_break_before` /
  `outline_level` / font name+size / the twelve ECMA-376 toggles / underline /
  color / highlight / `vert_align` / `num_id`+`num_level`) so the two libraries
  agree field-for-field and report the same shape for the same document.
- **Provenance is the upgrade past the 2-layer compare.** COM gives effective +
  applied-paragraph-style, so §7b detects "direct override on top of the paragraph
  style" — but is blind to *which* layer actually set it (a linked/character style
  via `rStyle`, a numbering-level `rPr`, table-style conditional formatting,
  docDefaults). The **fix differs** by layer (a character-style override does *not*
  yield to `Font.Reset()`). `FormattingSource` (layer + style_id + chain_depth +
  toggle-resolved) is exactly the "what contributes" attribution a precise fix
  needs.
- **Hybrid, plumbing already present.** For deep attribution, feed a range's OOXML
  (`Range.WordOpenXML` — already used by `read_image`) to docx-plus's resolver:
  live effective values + writes via COM, cascade provenance via docx-plus, no
  second cascade engine in wordlive. **Probe:** `WordOpenXML` is Flat OPC and the
  *final* (accepted) view — needs a small adapter to the python-docx `Document`
  docx-plus expects (reconstruct parts, or round-trip a temp `.docx`).

v1 ships the COM 2-layer compare (enough for the consistency rules); the
docx-plus hybrid is the attribution upgrade when a rule needs to explain or
precisely target *why* a value deviates.

**Targeted fix (idempotent).** wordlive/Word has no per-property "reset to style."
Two strategies:

- **Targeted (default):** write the *style's* value back as a direct property
  (e.g. `format_run(size=stylesize)`). Visually correct and **idempotent** (re-run
  writes the same value → no-op), though it leaves a redundant-but-matching direct
  property. Safe; never touches intentional formatting elsewhere in the paragraph.
- **Strip (opt-in, aggressive):** `ParagraphFormat.Reset()` / `Font.Reset()` clears
  *all* direct formatting back to the style — cleanest, but nukes intentional
  overrides (a bold term in a heading). Behind a profile flag
  (`"strip_direct_formatting": true`), not default.

**Idempotency contract.** `regularize` run twice must apply 0 fixes the second
time. This is a test invariant (a smoke test: build a messy doc → regularize →
regularize again → assert the second pass's `applied` is empty), and the reason the
targeted strategy is the default.

## 8. Politeness & safety

- `lint` is a pure read (repaginates content-neutrally for the layout rules,
  restoring `Saved` like `stats()` does); never moves selection/scroll.
- `regularize` runs inside `doc.edit()` → snapshots/restores selection + scroll,
  one atomic undo for the whole pass.
- **The `adds_content` gate (shipped).** A fixable finding whose fix inserts or
  destroys content — deletions (`stray-empty-paragraph`), content inserts
  (`figure-caption-present`), a watermark strip — sets `Finding.adds_content =
  True`. `regularize` **withholds** those by default (a formatting pass shouldn't
  silently change what the document says) and reports them in a `deferred` bucket
  (`{applied, skipped, deferred, findings}`). The caller opts in with
  `allow_content=True` (Python), `--allow-content` (CLI), or `allow_content: true`
  (the exec op / MCP `word_write`), which applies them in the same atomic-undo
  pass. Pure in-place formatting/text fixes leave `adds_content = False` and apply
  by default. (The individual content fixes are wired rule-by-rule — remaining
  item 1; the gate itself is the infrastructure they plug into.)
- Track-changes aware: if the document has Track Changes on, `regularize`'s edits
  are tracked like any other (the user reviews them) — call it out in docs.

## 9. Deferred (v1 boundaries)

- Cross-reference/bookmark/field-integrity rules (broken `REF`, dangling
  cross-ref) — a different detection family (the **P1 field-code walk**, §5b);
  now scoped in the v2 backlog (§5b·C/D, Batch 3), not open-ended.
- Reading-order / accessibility rules — tagged `accessibility` in §5b; the cheap
  structural ones (`heading-level-skip`) ship early, the rest belong with
  **prepare-for-sharing** (Part II Priority 6), which can call the linter.
- A custom-rule plugin API — start with the built-in catalogue; add extensibility
  only on a concrete need.
- Spelling/grammar — already covered by `doc.proofing()`; the linter is structure
  & formatting, not prose.

## 10. Build order

1. The **`anchor.format_info()` read control** + direct-override detection (§7) —
   the live-validated foundation, and a useful public read in its own right (the
   missing mirror of `format_paragraph`/`format_run`); every rule consumes it.
2. Structural rules (no config, objective): `heading-keep-with-next`,
   `table-repeat-header`, `list-numbering-continuity`. Highest signal, simplest.
3. Consistency rules (heading/font/spacing) on top of the format probe.
4. `regularize` (the `apply_op`-over-findings loop) + the idempotency smoke test.
5. Policy rules + the profile loader (`body-justified`, `body-line-spacing`,
   `table-numeric-right-align`). ✅ shipped as **Batch 4a**.
6. Wire CLI / exec op / MCP; docs (`docs/cli.md`, `docs/mcp.md`, `SKILL.md`,
   `cookbook.md` entry: "hand-off a clean document").

Steps 1–4 + wiring shipped as the foundation slice; the **v2 backlog (§5b)** then
continued primitive-driven through Batch 5 (§B heading structure), followed by the
cross-cutting `adds_content` gate (§8) and Batch 6 (its first content fixes) — see the
**Progress** table up top for the full shipped/remaining status. With the gate now in
place, **closing out the remaining report-only rules** (remaining item 1) and the
`house_style` half of §6 are next.

## 11. Future ideas (post-v1.0)

Three directions raised 2026-07-08, all **post-v1.0** (none blocks the release). They
share the same engine — `Rule.check(doc, span, profile)` dispatch + fixes-as-exec-ops —
so each is an extension, not a new subsystem.

### 11a. Exemplar-driven rules — extract a profile from / match against a reference

"Make this document look like *that* one." Two composable modes over one new reader:

- **Extract (`doc.derive_profile()` → `wordlive.lint.json`).** Read a *good* reference
  document and infer an implicit house style from it: the dominant `Normal` / heading
  font · size · spacing · alignment, list templates, table styles, margins/section
  setup. Emits the §6 `house_style` block (plus policy targets — justification, line
  spacing) so the output is a ready-to-check profile.
- **Match (`doc.lint(profile=derived)`).** Lint document A against the profile derived
  from document B — the exemplar becomes the policy. Pure composition once Extract
  exists: `lint`/`regularize` already take a `profile`.

**Dependency: this is the payoff of `house_style` (§6).** Extract produces a
`house_style`; Match consumes it. So it lands *after* the pre-v1.0 `house_style` work,
and is largely a **profile-inference reader** (dominant-value statistics over
`format_info()` across the reference doc) rather than new rules. Surfaces: Python
`derive_profile`, CLI `derive-profile REF.docx [-o profile.json]`, MCP `word_read`.
**Probe:** how to pick the "dominant" value per field (mode vs. the applied-style
baseline) and how much section/page setup to capture without over-fitting.

### 11b. Accessibility cluster (`accessibility` tag)

A coherent tag feeding **prepare-for-sharing** (Part II Priority 6). Mostly composes
*existing* reads — no new COM surface for the cheap ones:

| id | kind | detect (existing read) | fix | default |
|---|---|---|---|---|
| `image-missing-alt-text` | structural | `image:N` / `shape:N` `AlternativeText` empty | report → `set_alt_text` (needs a value → `adds_content`) | on (report) |
| `table-missing-header-row` | structural | a table with no designated header row (`set_heading_row`) | `set_heading_row(1)` | on |
| `link-text-not-descriptive` | consistency | link label is `click here` / a bare URL (overlaps `hyperlink-display-is-raw-url`) | report | off (tag) |
| `complex-merged-table` | structural | non-uniform table (`Table.is_uniform == False`) — screen-reader hostile | report | off (tag) |
| `heading-level-skip` | structural | **already shipped** — retag `accessibility` | report | on |
| `default-language-set` | structural | `Document`/style language unset | report | off (tag) |

Deferred (need real work): colour-contrast (colour math on effective run colour vs.
shading) and reading-order (floating-shape z-order vs. logical flow). The alt-text fix
is the one that needs a *value* — same shape as `document-properties-filled`, so it
rides the `adds_content` gate and (ideally) an alt-text suggestion the caller supplies.

### 11c. Custom rule definitions — declarative first, plugin later

Two tiers; the first is the "relatively straightforward" one and needs **no plugin API**:

- **Declarative rules in the profile (ship first).** A profile block lets a user define
  a rule from data — a `find` pattern (`literal`/`regex`, reusing the shipped
  `find_replace` modes), `severity`, `message`, `tags`, and an optional `fix` (a
  `find_replace` replacement). The engine already dispatches `Rule.check(doc, span,
  profile)` and runs fixes through `apply_op`, so a declarative rule is a thin
  parameterized `Rule` — a text-pattern rule authored in JSON, no code. Covers the bulk
  of "flag our forbidden phrase / enforce our term" asks (`"utilize"→"use"`, banned
  words, required boilerplate) safely.

  ```jsonc
  // in wordlive.lint.json
  "custom_rules": [
    { "id": "no-utilize", "find": "\\butiliz(e|es|ing|ed)\\b", "mode": "regex",
      "severity": "info", "message": "Prefer 'use'.", "tags": ["house"],
      "fix": { "replace": "use" } }
  ]
  ```

- **Programmatic `Rule` plugins (defer to concrete need).** Register a Python callable
  (entry-point / import path) implementing the `Rule` protocol for rules a pattern can't
  express. This is a code-loading surface (trust/sandboxing questions), so it stays
  gated on a real need — the declarative tier likely absorbs most demand first.
