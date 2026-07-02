# Linter + formatting regularizer — design sketch

Status: **foundation slice shipped (Unreleased, 2026-06-19)** — design was a
**sketch** (2026-06-17). Roadmap home: `feature-plan.md` Part II, Priority 1,
item 1. This is the detailed design; the roadmap keeps the one-paragraph summary.

**Shipped (build order §10 steps 1–4 + full wiring):** `anchor.format_info()`
(the read mirror + direct-override detection, §7), the three structural rules
(`heading-keep-with-next`, `table-repeat-header`, `list-numbering-continuity`),
the heading/font/spacing consistency rules + `mixed-run-format` (report-only),
`doc.lint()` / `doc.regularize()` with the **targeted, idempotent** default fix
and the idempotency test, wired across Python / CLI (`lint`, `regularize`, `read
format`) / `regularize` exec op / MCP. Live-validated against Word 16 (multi-page
table, split list, heading override → fix → idempotent re-run).
**Deferred to a follow-up:** the **policy** rules (`body-justified`,
`table-numeric-right-align`) and the **profile / house-style** loader (§6) — the
rule registry already carries `kind`, so they slot in without rework; the
aggressive `Font.Reset()` strip-to-style fix (§7c); the content-changing fixes
(`stray-empty-paragraph` delete, `figure-caption-present` insert); and the
`docx-plus` cascade-provenance hybrid (§7c).
**Backlog (v2, brainstormed 2026-06-19):** a primitive-driven catalogue of ~40
more rules for publishing/academia — typography hygiene, finalization, captions /
cross-references / citations, layout & notices — see **§5b**.
**Batch 1 — typography hygiene ✅ shipped (Unreleased).** The §5b·A cluster plus
`manual-heading-formatting` / `table-style-consistent`: 10 rules in
`_linting_typography.py` (the P2 paragraph-text scan), 6 on by default + 4
opinionated off-by-default behind the `typography` tag. Two enablers landed with
it: `find_replace` gained **`literal` / `regex` modes** + a `required=False`
no-op-on-zero-match flag (the fix path — Word's existing `find_replace` is fuzzy,
which **cannot** express a literal whitespace/punctuation edit; a fuzzy find of
`"  "` matches every single space — so the typography fixes are regex-mode
`find_replace` ops scoped to the offending `para:N`, re-scanning live text with no
precomputed offsets to drift), and `Rule` gained a **`default_on`** flag for the
default-off opinion rules. **Decisions baked in:** no `adds_content` field this
batch (every fixable rule is a pure in-place text/style edit; defer it to the
caption/finalization batches that add content); `straight-quotes` / `nbsp-missing`
/ `sentence-spacing-consistent` and `leading-whitespace`'s "→ real indent" half
**deferred to Batch 1b** (heuristic-heavy). Live-probed 2026-06-24: trailing-mark
preserved (no paragraph fuse), adjacent inline bold survives a sub-span collapse,
`regularize` idempotent on the second pass.
**Batch 2 — finalization ✅ shipped (Unreleased).** The §5b·G cluster: 6 rules in
`_linting_finalization.py`, all tagged `finalization` and — by decision this pass —
**all off by default** (an opt-in "is-this-ready-to-send?" check, not a default-lint
defect, since mid-authoring docs normally carry comments/revisions; this deviates
from §G's per-rule default column). `comments-present`, `unaccepted-revisions`,
`track-changes-on`, `hidden-text-present`, `stale-fields` are report-only;
`leftover-highlight` is the one fix (clears the highlight, idempotent). Detection
reuses the shipped `doc.comments` / `doc.revisions` / `doc.fields` /
`doc.track_changes` wrappers plus two new `format_info` fields that landed with it
(`font.hidden`, `font.highlight` — the read mirror of `format_run`'s writes).
**Decision baked in:** `stale-fields` is a **report-only nudge** — Word exposes no
field-staleness flag, so a presence-based `update_fields` fix couldn't satisfy the
idempotency contract; the fixable version lands with Batch 3's field-code backbone.
Live-probed 2026-07-01 (Word 16): all six fire on a seeded doc, none in the default
set, highlight fix + idempotent re-run confirmed on real COM.

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
| `stray-empty-paragraph` | structural | an empty `Normal` paragraph between blocks | `delete_paragraph` (report-only by default; deletes are loud) |
| `double-space` | consistency | runs of 2+ spaces in body text | `find_replace` collapse (skip code/verbatim styles) |

### Font / character consistency

| id | kind | detect | fix |
|---|---|---|---|
| `heading-font-consistent` | consistency | same-style headings whose `Font.Name`/`Size`/`Bold` carry direct overrides deviating from the style | re-apply style value / clear override (§7) |
| `body-font-consistent` | consistency | body paragraphs with a directly-set font face differing from the `Normal`/body style | clear override |
| `mixed-run-format` | consistency | a paragraph whose `Font.Size`/`Name` reads `wdUndefined` (mixed runs) where uniformity is expected | report-only (which run is the outlier needs run-walk; fix is opt-in) |

### Alignment & justification (policy)

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

### B. Heading & document structure  *(P2 + outline walk)*

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `manual-heading-formatting` | structural | a bold/large `Normal` para that looks like a heading but isn't styled | report → suggest `apply_style("Heading N")` | on (report) |
| `heading-level-skip` | structural | outline jumps H1→H3 with no H2 | report-only | on (report) |
| `heading-numbering-manual` | consistency | heading text starts with literal `3.1` not list-numbered | report | off (tag) |
| `heading-trailing-period` | consistency | heading text ends with `.` | strip | off (tag) |
| `empty-heading` | structural | heading paragraph with no text | report | on (report) |
| `adjacent-headings` | structural | two headings, no body between | report | off (tag) |
| `toc-present-and-current` | structural | doc has Heading 1s but no TOC field / TOC stale | `update_fields` or report | off (tag) |

### C. Captions & cross-references  *(P1 — the academia backbone)*

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `caption-manual-numbering` | structural | a `Caption` para whose number is **literal text**, not a `SEQ` field | report → rebuild with SEQ (adds content, opt-in) | on (report) |
| `xref-as-literal-text` | structural | `see Figure 3` / `Table 2` typed as text, not a `REF` field | report (auto-fix needs target match) | on (report) |
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

### H. Page layout & document-level  *(P4)*

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `page-numbers-present` | policy | no `PAGE` field in any footer/header | insert (adds content, opt-in) | off (tag) |
| `confidentiality-notice` | policy | profile-supplied text not found in H/F or body | report (insert opt-in) | off (profile) |
| `copyright-notice` | policy | profile `©` / text not present | report | off (profile) |
| `header-footer-consistent` | consistency | H/F text/format differs across sections unexpectedly | report | off (tag) |
| `document-properties-filled` | policy | Title / Author core props empty | set | off (tag) |
| `draft-watermark-present` | structural | a "DRAFT" watermark / shape still in final | report | off (tag) |

### I. Hyperlinks  *(print / sharing)*

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `hyperlink-bare-for-print` | policy | hyperlink display text ≠ target URL (URL invisible on paper) | report / append `(url)` (opt-in) | off (profile: print) |
| `hyperlink-broken-internal` | structural | internal `HYPERLINK \l anchor` with no matching bookmark | report | on |
| `hyperlink-display-is-raw-url` | consistency | long raw URL shown inline where a label is wanted | report | off (tag) |

### Tag taxonomy

Rules carry **tags** so a user enables a *cluster* instead of naming ids. Proposed
top-level tags: `typography`, `headings`, `lists`, `tables`, `captions`,
`crossref`, `citations` (alias `academia` = captions + crossref + citations +
nbsp + en-dash), `finalization`, `layout`, `print`, `accessibility`. `--rules
academia` / `--rules finalization` become the headline ergonomics; profiles
(§6) toggle tags + supply policy targets.

### Suggested batch order (primitive-driven)

1. **Batch 1 — Typography hygiene (P2): ✅ shipped (Unreleased).** the §A rules +
   `manual-heading-formatting`, `table-style-consistent`. Highest hit-rate,
   cheapest, no field plumbing. (`straight-quotes` / `nbsp-missing` /
   `sentence-spacing-consistent` deferred to a 1b follow-up.)
2. **Batch 2 — Finalization (P3, §G): ✅ shipped (Unreleased).** 6 rules in
   `_linting_finalization.py`, all off-by-default behind the `finalization` tag,
   one report-only fix (`leftover-highlight`); reused the shipped revision/comment/
   field wrappers + `doc.track_changes`, and added `format_info`'s `hidden` /
   `highlight` fields. `stale-fields` ships as a report-only nudge (its
   `update_fields` fix waits for Batch 3 — no COM staleness flag).
3. **Batch 3 — Field-code backbone (P1, §C):** build `Range.Fields` walk, then
   caption-manual-numbering, xref-as-literal-text, broken-cross-reference,
   page-numbers-present. The academia centerpiece.
4. **Batch 4 — Layout / notices (§H), hyperlinks (§I) + the profile/house-style
   loader** (the already-deferred v1 step-5).
5. **Later — citations cluster (§D) + the accessibility sub-product** (with
   *prepare-for-sharing*, §9).

## 6. Profiles (policy rules + house style)

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
- Deletions (`stray-empty-paragraph`) and content-adding fixes
  (`figure-caption-present`) are **report-only by default** — they change content,
  not just formatting, so they're opt-in per profile. Formatting fixes are safe to
  apply by default.
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
5. Policy rules + the profile loader (`body-justified`, `table-numeric-right-align`).
6. Wire CLI / exec op / MCP; docs (`docs/cli.md`, `docs/mcp.md`, `SKILL.md`,
   `cookbook.md` entry: "hand-off a clean document").

Steps 1–4 + wiring shipped (foundation slice). The **v2 backlog (§5b)** continues
the build, primitive-driven: Batch 1 typography (P2) ✅ · Batch 2 finalization
(P3) ✅ · Batch 3 field-code backbone (P1) · Batch 4 layout/notices + profile
loader · later citations + accessibility.
