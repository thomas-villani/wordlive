# Changelog

All notable changes to **wordlive** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.19.0] - 2026-07-15

### Added
- **A dedicated MCP agent guide — the server now teaches its own tools.** `word_read(command=
  "guide")` and the `wordlive://guide` resource previously returned the **CLI** skill, so an agent
  driving the four `word_*` dispatch tools was taught `wordlive outline`, `--anchor-id`, and exit
  codes it can't use — the single biggest gap the Claude Desktop trial found. A new
  `wordlive-mcp` skill teaches the tool surface natively (dispatch on `command` / `op` / `action`,
  the anchor model, `word_exec` batch mechanics) and carries the two things the CLI guide never
  did: a **CLI→MCP name crosswalk** (`read format`→`format_info`, `apply_list`→`list`+`action`,
  `--within`→`within`, `--profile PATH`→`profile` (which over MCP also takes an inline object)),
  and the house-style **`profile` JSON schema**. Both MCP guide surfaces now serve it; the CLI /
  Python guides (`wordlive llm-help`, `install-skill`) are unchanged. Drift tests pin the new
  skill's op list, its write-dispatcher actions, and every crosswalk command name to the code.
- **Clearer lint cluster wording + two `word_exec` docstring nits.** The CLI skill's `layout` /
  `structure` clusters now say naming the tag *enables* their off-by-default members (matching
  `typography`'s parenthetical), instead of a bare "all off by default" that read as "stays off".
  And the `word_exec` tool docstring now spells out that "atomic" is one **undo entry**, not a
  transaction (a mid-batch failure leaves the successful prefix applied), and that each op resolves
  its `anchor_id` fresh — so positional ids shift with edits earlier **in the same batch**, not
  only between calls.
- **Docs — two new guides plus a prominent one-click agent install.** A new
  [Advanced](docs/advanced.md) page walks the power features in one continuous session
  (token-budgeted `read` digests + `to_markdown` drill, durable `pin`/`pin_outline` handles,
  Excel-backed `insert_chart`, `snapshot` vision rendering, and profile-driven `regularize`).
  A new [Agent patterns](docs/agent-patterns.md) page consolidates the LLM-driving patterns that
  were scattered across the cookbook (skim/drill, anchor-not-cursor, pin-vs-renumber, one-intent
  batches, suggest-don't-overwrite, checkpoint/diff verify, typed-failure self-correction, vision
  snapshots, fuzzy locate). `docs/agents.md` now leads with the Claude Desktop `.mcpb` one-drop
  path; `docs/concepts.md` gains a "Durable handles (pins)" section; and Cookbook §22 adds a
  vision-snapshot recipe. Nav, landing-page cards, and cross-links updated; docs build clean under
  `mkdocs build --strict`.
- **Linter golden e2e — the published walkthrough is now pinned.** Two `e2e`/`smoke` tests in
  `tests/test_e2e_cli.py` drive the real CLI (`lint`, `regularize`, `regularize --dry-run`, and a
  `--profile` run) against the committed `examples/sample/messy-brief.docx` in live Word and assert
  the exact findings `docs/linting.md` publishes — the six default findings in order, the
  empty dry-run buckets, the 26-finding profile audit, and the fix → clean → idempotent-second-pass
  sequence. A companion test rebuilds the sample from its committed `build_messy_brief.py` and
  re-lints it, guarding against script↔binary drift. Closes the e2e-CLI linter coverage gap (the
  prior lifecycle tests never touched `lint`/`regularize`).
- **Linter Batch 6 — the first `adds_content` opt-in fixes (5 rules wired fixable + 1 new rule).**
  (Priority 1, item 1 — `spec-linter.md` Remaining item 1.) The `adds_content` gate finally has fixes
  plugged into it: six report-only rules now carry an opt-in `fix` flagged `adds_content=True`, so
  `regularize` reports them in `deferred` by default and applies them under `allow_content=True`
  (`--allow-content`). All compose existing exec ops — **no new COM write surface** — and all are
  idempotent. `draft-watermark-present` → `remove_watermark`; `page-numbers-present` → a `{ PAGE }`
  field in `footer:1:primary`; `confidentiality-notice` / `copyright-notice` → the notice text
  dropped into `footer:1:primary`; `hyperlink-bare-for-print` → the URL folded into the display text
  as `label (url)` (targeting the link by positional index). A **new** `stray-empty-paragraph` rule
  (structural, **off** by default behind the new `whitespace` tag) flags an empty `Normal` paragraph
  between content blocks and fixes it with `delete_paragraph`; `regularize` applies paragraph
  deletions last and in descending document order, so a multi-blank pass never invalidates an earlier
  fix's anchor. `document-properties-filled` (needs a value), `caption-manual-numbering` (fragile
  in-place `SEQ` rebuild), `xref-as-literal-text` / `broken-cross-reference` (need a human-chosen
  target) stay report-only. Live-probed against Word 16 (a messy doc — watermark, no page number,
  bare hyperlink, two stray blanks — regularized to clean, second pass a no-op).
- **Linter `adds_content` gate — content-changing fixes are now opt-in.** (Priority 1, item 1 —
  `spec-linter.md` §8.) `Finding` gains an `adds_content: bool` field: a fixable finding whose fix
  **inserts or destroys content** (rather than just re-formatting existing content) sets it `True`.
  `regularize` now **withholds** those fixes by default — a formatting pass no longer silently
  rewrites what a document *says* — and reports them in a new `deferred` bucket
  (`{applied, skipped, deferred, findings}`). Pass `allow_content=True` (Python), `--allow-content`
  (CLI `regularize`), or `allow_content: true` (the `regularize` exec op / MCP `word_write`) to apply
  them in the same atomic-undo pass. Pure in-place formatting/text fixes leave `adds_content` `False`
  and still apply by default; the existing idempotency invariant is unchanged. This is the
  cross-cutting infrastructure that unlocks the deferred content/repair fixes (strip watermark, fill
  a property, insert a notice, delete a stray paragraph, rebuild a caption — wired rule-by-rule next).
- **Linter Batch 5 — heading & document-structure rules (6 new rules).** (Priority 1, item 1 —
  `spec-linter.md` §5b·B, the P2 + outline-walk cluster.) `doc.lint` gains the §B cluster over
  `doc.outline()` (no new read surface): `heading-level-skip` (the outline jumps a level — an H1
  followed by an H3 with no H2) and `empty-heading` (a heading paragraph with no text) ship **on**
  by default like the v1 structural set; the opinionated four ship **off**, behind the new
  `structure` tag (and the existing `headings` tag) — `adjacent-headings` (two headings in a row
  with no body between), `heading-numbering-manual` (a heading numbered by hand, `3.1 Methods`,
  rather than by automatic numbering), `heading-trailing-period` (a heading whose text ends in a
  period), and `toc-present-and-current` (top-level headings but no table-of-contents field).
  `lint --rule structure` (or `--rule headings`) selects the cluster. Five are report-only;
  `heading-trailing-period` is **fixable** — the fix strips the trailing period in place via a
  paragraph-scoped `find_replace` regex (idempotent, no `adds_content` gate). `toc-present-and-current`
  is presence-only: Word exposes no field-staleness flag (the same limit `stale-fields` hit), so the
  "current" half stays a report. Rules auto-surface across Python / CLI / MCP. Live-probed against
  Word 16 (all six firing on a seeded outline, the two on-by-default in the default set, the
  trailing-period fix stripping + idempotent, and the TOC rule clearing once a TOC is inserted).
- **Linter Batch 4c — page-layout / document-level rules (5 new rules) + `doc.watermark()`.**
  (Priority 1, item 1 — `spec-linter.md` §5b·H, the P4 section/header-footer + document-level
  cluster.) `doc.lint` gains the §H cluster, all **off by default** (a §H issue is rarely a
  defect mid-authoring — an opt-in "getting it ready to hand off" check) and report-only:
  `header-footer-consistent` (the primary header/footer text disagrees across the document's
  own non-linked sections) and `draft-watermark-present` (a leftover DRAFT / CONFIDENTIAL
  watermark, also tagged `finalization`); plus three **policy** rules —
  `document-properties-filled` (a required built-in property left empty; `required` defaults
  to `["Title", "Author"]`), and `confidentiality-notice` / `copyright-notice` (a
  profile-supplied notice string — `copyright-notice` defaults to `"©"` — missing from every
  header/footer and the body). `lint --rule layout` selects the whole cluster; `--rule notices`
  selects just the two notice rules. Detection reuses the shipped `doc.properties` /
  `doc.sections` read wrappers plus a **new `doc.watermark()` read** — the mirror of
  `set_watermark` / `remove_watermark`, returning a `WatermarkInfo(text, sections)` or `None`,
  wired across Python, CLI (`read watermark`), and MCP (`word_read command=watermark`). All
  fixes (fill a property, insert a notice, strip a watermark) add or destroy content, so
  they're deferred behind the `adds_content` gate — no new COM write surface (the write verbs
  `doc.properties.set` / `remove_watermark` already exist). Rules auto-surface across Python /
  CLI / MCP. Live-probed against Word 16 (watermark round-trip through `set_watermark` →
  `doc.watermark()` → `remove_watermark`, and the cluster firing on a seeded doc).
- **Linter Batch 4b — hyperlink rules (3 new rules).** (Priority 1, item 1 —
  `spec-linter.md` §5b·I, the §I "are the links sound and print-ready?" cluster.)
  `doc.lint` gains a hyperlink cluster built on the existing `doc.hyperlinks` read
  wrapper (no new read surface): `hyperlink-broken-internal` (an in-document jump —
  `HYPERLINK \l` with an empty address — whose target bookmark no longer exists, a
  dead link) ships **on** by default like the v1 structural set; `hyperlink-bare-for-print`
  (an external link whose visible text doesn't contain its URL, so the destination is
  invisible on paper) and `hyperlink-display-is-raw-url` (a link whose whole label is a
  bare URL where a readable one was wanted) ship **off**, behind the `hyperlinks` / `print`
  tags (`lint --rule hyperlinks`, or `--rule print` for the two print/sharing rules). All
  three are report-only — the repairs either need a human to pick the intended target or
  add content (append `(url)`, invent a label), so they're deferred (the fixes' verb,
  `Hyperlink.update`, already exists). Broken-internal reuses `name in doc.bookmarks`
  (Word's `Bookmarks.Exists`), which also sees the hidden `_Toc…`/`_Ref…` bookmarks a live
  cross-reference targets, so a healthy jump isn't flagged. No new COM write surface; rules
  auto-surface across Python / CLI / MCP. Live-probed against Word 16 (internal vs external
  `Address`/`SubAddress`/`TextToDisplay`, and `Exists` over the target bookmark). The §H
  layout/notices detection rules become **Batch 4c**.
- **Linter Batch 4a — the house-style profile loader + 3 policy rules.** (Priority 1,
  item 1 — `spec-linter.md` §6, the profile/policy half of Batch 4.) `doc.lint` /
  `doc.regularize` gain a `profile=` argument (a path to a `wordlive.lint.json` file, an
  inline dict, or `None`) that enables **policy** rules — off in the default set until a
  profile opts them in — supplies their targets/thresholds, and can override a rule's
  severity or disable a default rule. Three policy rules land in `_linting_policy.py`, all
  fixing idempotently through the existing `format_paragraph` op: `body-justified` (body
  paragraphs not justified), `body-line-spacing` (line spacing ≠ the profile's `target`,
  e.g. `"1.5"`), and `table-numeric-right-align` (a table column that's mostly numbers,
  above a `threshold`, but not right-aligned). The profile threads through every surface —
  Python, CLI (`--profile PATH`), the `regularize` exec op, and MCP (`word_read
  command=lint` / `word_write command=regularize`). No new COM write surface (the fixes
  reuse `format_paragraph`); the `house_style` half of §6 (pinning consistency targets +
  `set_style` fixes) is deferred to a later pass. Live-probed against Word 16 (fires,
  fixes, and a second `regularize` is a no-op — the idempotency contract).
- **Linter Batch 3b — `xref-as-literal-text` (1 new rule).** (Priority 1, item 1 —
  `spec-linter.md` §5b·C.) The heuristic cross-reference rule deferred from Batch 3:
  a body paragraph that mentions a figure/table by literal number ("As shown in
  Figure 3, …") with no `REF`/`PAGEREF` field covering it, so it won't retarget when
  figures are renumbered or moved. Report-only (an auto-fix would have to guess the
  intended target). Because a bare "Table 2" in prose is often legitimate, it ships
  **off by default**, behind the `crossref` / `academia` tags (`lint --rule academia`) —
  deviating from §C's "on" column, matching how Batch 1b deferred its heuristic rules.
  Caption paragraphs (that's `caption-manual-numbering`'s job) and headings are
  skipped. No new COM write surface. Live-probed against Word 16.
- **Linter Batch 3 — field-code backbone (3 new rules).** (Priority 1, item 1 —
  `spec-linter.md` §5b·C, the P1 cross-reference/caption cluster.) `doc.lint` gains
  a field-integrity cluster built on a `Range.Fields` walk: `broken-cross-reference`
  (a `REF`/`PAGEREF` field rendering Word's "reference source not found" error) and
  `caption-manual-numbering` (a `Caption` paragraph whose figure/table number is
  literal text rather than a `SEQ` field) ship **on** by default (both also tagged
  `academia`, so `lint --rule academia` selects them); `page-numbers-present` (no
  `PAGE` field in any header/footer) is **off**, behind the `layout` tag. All three
  are report-only this batch — the fixes they imply either add content (rebuild a
  caption around a `SEQ` field, insert a page number) or need target matching (repair
  a broken reference), so they're deferred; the heuristic `xref-as-literal-text`
  ("see Figure 3" typed as text) is deferred to a Batch 3b. No new COM write surface
  (the fixes' future verbs — `insert_caption`, `insert_cross_reference`,
  `update_fields`, `insert_page_number` — already exist). Live-probed against Word 16.
  Rules auto-surface across Python / CLI / MCP. (This also settles the Batch 2
  `stale-fields` note: an idempotent auto-refresh is inherently infeasible without a
  Word staleness flag, so it stays a report-only nudge — not an IOU.)
- **Linter Batch 2 — finalization (6 new rules).** (Priority 1, item 1 —
  `spec-linter.md` §5b·G, the P3 "is this actually final?" cluster.) `doc.lint` /
  `doc.regularize` gain a leftover-review / markup-state cluster, all tagged
  `finalization`: `comments-present`, `unaccepted-revisions`, `track-changes-on`,
  `hidden-text-present`, `stale-fields` (updatable TOC/SEQ/REF/PAGE fields present
  — a "refresh before finalizing" nudge; Word exposes no staleness flag, so an
  idempotent auto-refresh fix is infeasible and it stays report-only), and
  `leftover-highlight` (the one fixable rule — clears the highlight, idempotent).
  The whole cluster ships **off by default**, behind the `finalization` tag: a
  mid-authoring document normally carries comments and revisions, so this is an
  opt-in pre-send check (`lint --rule finalization`), not a default-lint defect.
  Detection reuses the shipped `doc.comments` / `doc.revisions` / `doc.fields` /
  `doc.track_changes` wrappers plus the two new `format_info` fields (below).
  Live-probed against Word 16. Rules auto-surface across Python / CLI / exec op /
  MCP.
- **`format_info` gains `hidden` + `highlight`.** The `anchor.format_info()` font
  section (CLI `read format`, MCP `word_read command=format_info`) now reports
  Word's hidden-text attribute (`font.hidden`, with the usual effective-vs-style
  override annotation) and the highlight colour (`font.highlight`, a keyword or
  `"none"`) — mirroring the `format_run` write side. Highlight lives on the range,
  not the style, so it's effective-only (`style` is `null`).
- **Linter Batch 1 — typography hygiene (10 new rules).** (Priority 1, item 1 —
  `spec-linter.md` §5b·A, the P2 "run-walk / text-scan" cluster.) `doc.lint` /
  `doc.regularize` gain a cheap, high-frequency text-defect cluster, all tagged
  `typography`: `trailing-whitespace`, `leading-whitespace`,
  `space-before-punctuation`, `double-space`, `manual-heading-formatting`
  (a short, bold/enlarged `Normal` paragraph that reads like a heading but was
  never styled — report-only), and `table-style-consistent` (restyle the minority
  tables onto the document's dominant table style) ship **on** by default; the
  opinionated `hyphen-as-range` (en-dash), `em-dash-usage`, `tabs-for-layout`, and
  `manual-line-break` ship **off** (named or via the tag — `spec-linter.md` §5b
  "default stance"). Detection is a pure paragraph-text scan; the whitespace /
  punctuation / range fixes compose the new `find_replace` regex mode (below),
  scoped to the offending `para:N`, so they never flatten inline formatting and a
  second `regularize` is a clean no-op (live-probed). Wired Python / CLI (`lint
  --rules typography`, `regularize`) / `regularize` exec op / MCP — rules
  auto-surface from the registry.
- **`find_replace` literal & regex modes + a new `default_on` rule flag.** The
  find/replace surface (`doc.find` / `doc.find_replace`, the `find_replace` exec
  op, CLI `find`/`replace`, MCP `find`/`replace`) gains a `mode` ∈
  `fuzzy` (default, unchanged) / `literal` (exact, no folding) / `regex` (Python
  regex; the replacement may use `\1` backreferences, expanded per match — so a
  single call collapses runs or reorders groups). `find_replace` also gains
  `required=False` (zero matches → `[]` instead of raising), used by idempotent
  batch autofixes where an earlier fix already cleaned an overlapping match. The
  linter `Rule` registry gains a `default_on` flag so opinionated/policy rules can
  ship off-by-default yet stay reachable by id or tag.
- **Tutorial — a single guided editing session.** New `docs/tutorial.md` fills
  the learning-oriented gap between the quickstart and the random-access
  Cookbook: it drives one document end to end (attach → inspect → read a section
  → first polite edit → batch under one undo → suggest via comment + tracked
  change → verify with checkpoint/diff), dual-tabbed Python **and** CLI at every
  step, so the four invariants land by doing. Ships the sample it drives —
  `examples/sample/quarterly-report.docx` plus its committed, reviewable builder
  `build_quarterly_report.py` — and adds the page to the nav with cross-links
  from Getting started and Examples.
- **Inline code spans in the Markdown subset.** `` `code` `` in `insert_markdown` /
  `insert_block` item text (and `code: true` on a structured run) now becomes a real
  monospace run — direct `Font.Name = "Consolas"` character formatting, the same choice
  `**bold**` makes with `Font.Bold`, so the document's style gallery is never touched.
  `to_markdown` detects a monospace run and emits the backticks back, sizing the fence to
  contain literal backticks. Previously backticks landed in the document as literal
  characters — a high-frequency papercut, since agents backtick every identifier, filename,
  and command they describe. Code spans bind tighter than emphasis, and their content is
  literal; because runs are flat, emphasis does not reach across a code span.

### Changed
- **`insert_paragraph_before` / `insert_paragraph_after` now default to `Normal` instead of
  inheriting the anchor's style.** *(Behaviour change.)* Given no `style`, a paragraph inserted
  after a `Heading 2` used to come out as another heading — corrupting the navigation outline and
  shifting every `heading:N` id. Both now pin `Normal` when no `style` is named, so the result no
  longer depends on where you anchor. This closes out the same style-inheritance trap already fixed
  for `insert_block` / `insert_section` / `insert_markdown` / `insert_break` / `insert_table`; the
  low-level paragraph inserters were the last holdout. Pass `style` explicitly to match the
  surroundings (e.g. to continue a list — a paragraph's current style is in
  `doc.paragraphs.list()[i]["style"]`).
- **`body-font-consistent` now audits size and weight, not just the font face.** It flagged only a
  hand-set font *name*, while `heading-font-consistent` has always checked name/size/bold — so a
  body paragraph bumped to 16pt or bolded drifted from its style unnoticed (the trial's `para:8`
  observation). It now audits all three, each with the same targeted style-value fix. One carve-out
  keeps it honest: a *short, emphasized, non-sentence* paragraph is what `manual-heading-formatting`
  flags as a missing heading style, so its size/bold are left alone here — otherwise `regularize`
  would strip the very emphasis that points at the real defect, and the paragraph would stop
  tripping either rule. The two rules now share one predicate so they can't disagree. The font name
  is always audited (never a heading signal). No change to the default finding set on documents
  without size/bold body overrides.
- **`insert_block` / `insert_section` body items now default to the `Normal` style instead of
  inheriting the insertion point's.** An item that named no `style` used to inherit whatever
  paragraph it was inserted after. Because `insert_section` writes its heading *first*, and its
  natural anchor is a heading, the body of a section reliably came out as a **heading** —
  silently corrupting the outline and shifting every `heading:N` id. Anchoring an `insert_block`
  after a `Heading 1` had the same effect. Both now pin `Normal`, matching `insert_markdown`
  (which already did, deliberately) and `insert_break`. The result no longer depends on where
  you anchor. To match the surroundings instead, pass `style` explicitly — a paragraph's current
  style is in `doc.paragraphs.list()[i]["style"]`. (The low-level `insert_paragraph_before/after`
  followed in the Unreleased release above.)
- **`body-font-consistent` no longer audits table cells.** A cell's font is the table's business
  (a table style sets it, and `table-style-consistent` polices that), so auditing cells here
  double-reported and buried the real prose findings under one finding per cell — a wide table
  could emit hundreds. This also matches how the policy rules already define "body prose", and
  it is what kept a lint on a large table slow even after the walk was made linear.
- **Docs — to start a document from blank, lead with `insert_*`, not `append`.** The
  structural inserts (`insert_markdown` / `insert_block` / `insert_section`) reuse a blank
  document's lone empty paragraph; `append` promises a *new* final paragraph and so leaves
  that empty one stranded above your content. Working as designed — an `append` that
  silently reused `para:1` would depend on invisible document state — but it surprised an
  agent building a document top-down, so the CLI/MCP/Python guides now say which op to open
  with. Noted in both bundled skills, `docs/cli.md`, `docs/mcp.md`, and `docs/python-api.md`.
- **Internal — the four mega-modules are now packages (behavior-preserving).** No public
  API change: `wordlive.__all__`, every `Document`/`Anchor` member and signature, the
  anchor-id scheme, the CLI verbs and the MCP `word_*` tools are all byte-for-byte the
  same. `cli/commands.py` (8.2k lines) → a `cli/commands/` package aggregated by
  `register()`; `mcp/server.py` → `build_server`/`main` plus sibling dispatch modules;
  `_document.py` → a `_document/` package (`DocumentCore` + editing/reading/structure/
  persistence mixins); and `_anchors.py` (5.3k lines) → an `_anchors/` package of 20
  modules, with the 2,232-line `Anchor` ABC split into `AnchorCore` + six feature mixins.
  Only consumers reaching into *private* module paths are affected — e.g. the helper
  once at `wordlive._document._new_pin_code` now lives in `wordlive._document._persistence`
  (if you monkeypatch wordlive internals, patch the binding that uses it, not a re-export).
- **New top-level docs page: [Linting & regularizing](docs/linting.md).** The linter
  gets its own guide — motivation (formatting normalization is slow, error-prone
  handwork), the two-verb mental model (`lint` reads → findings → `regularize` writes),
  the anatomy of a finding, a guided walkthrough on a deliberately-messy sample
  (`examples/sample/messy-brief.docx` + `build_messy_brief.py` + a `wordlive.lint.json`
  house-style profile), the full 45-rule catalog (one canonical table: kind · default ·
  fixable · tags), rule selection, and house-style profiles. Linked from the home page,
  the nav (after Tutorial), and the Python/CLI reference sections. Corrected the
  `regularize` JSON example in `docs/cli.md` (`applied`/`skipped` are lists of findings,
  not counts — the counts are the `--text` summary). Docs/sample only; no code change.
- **Docs reorganised into browsable sections.** The CLI reference
  (`docs/cli.md`) was a flat list of 130 `##` command headings; it's now grouped
  into 23 thematic `##` sections (Inspecting · Reading · Editing · Formatting ·
  Tables · Citations · …) with each command demoted to `###` under its category.
  The Python API reference (`docs/python-api.md`) gained an 8-Part `##` layer over
  its 25 sections (now `###`), with mkdocstrings `heading_level` bumped 2→3 so the
  generated class docs nest one level deeper to match. Pure reorganisation — no
  command, op, or API surface changed, and all heading anchors are preserved
  (only heading *levels* changed), so existing cross-references still resolve.

### Fixed
- **A table dropped under a heading no longer strands empty heading paragraphs.** `insert_table`
  reset the new *cells* to `Normal` but not the separator/trailing paragraph marks it injects (the
  one it opens so a table can't end the document, and the ones that keep abutting tables from
  merging). Anchored under a `Heading 2` at the end of a document those marks inherited the
  heading style, leaving one or two empty `Heading 2` paragraphs after the table that showed up as
  phantom entries in the navigation outline. They're now reset to `Normal` — addressed relative to
  the finished table and walked as a run, since the terminal case leaves *two* stray marks (the
  injected one plus the document's original final mark) — while the anchor heading itself and any
  real content around the table are left untouched. A live-Word smoke test pins the behaviour.
- **Dispatch errors now suggest instead of reciting the vocabulary.** A wrong `command`
  used to fail pydantic's `Literal` check and dump all 45 (`word_read`) or 93 (`word_write`)
  alternatives, suggesting none — the miss never even reached wordlive's code. `command` is
  now typed `str` (the enum is still published in each tool's JSON schema, so nothing is lost
  on the happy path) and the dispatchers answer with `did you mean 'format_info'?`, validated
  *before* attaching to Word. The same treatment covers unknown `word_exec` ops and
  `list`/`comment`/`revision`/`table` actions. Matching is deliberately not plain `difflib`,
  which ranks `read_format` nearest `read_image` on the shared `read_` prefix: a new
  `_suggest` module blends character similarity with IDF-weighted token overlap, so the rare
  `format` token outweighs the common `read` one. A name with no near neighbour gets **no**
  suggestion — a confident wrong answer is worse than none.
- **An `exec` op name passed as a `word_write` command is answered exactly.** `add_row`,
  `apply_list`, `add_comment`, `accept_all_revisions` and 26 others are reached through a
  sub-dispatcher, and no string metric recovers `table` from `delete_row`. The error now
  names the pair outright: *"it is the `add_row` action of command `table` — call
  `word_write(command='table', action='add_row')`"*.
- **A missing or misspelled `action` now names the valid ones.** `command 'list' requires
  'action'` was true but unhelpful; it now lists `apply`, `remove`, `restart`, `indent`,
  `outdent`, `format` (and likewise for `comment` / `revision` / `table`).
- **`table:2` explains itself.** Addressing a whole table raised `table cell not found:
  'table:2'` while the guidance sat unread in `anchor_by_id`'s docstring. The error now
  carries a hint naming `table:2:R:C`, `table:2:row:R`, `table:2:col:C`, and `table_read`.
  A misspelled anchor *kind* (`tabel:1:1:1`) likewise gets a `did you mean 'table'?`.
- **`lint` no longer hangs on a document with a large table (a hard 4-minute lock).** Every
  table cell is a paragraph, so a big table inflated the paragraph count — and two things then
  went badly wrong. First, eight per-paragraph rules read formatting via
  `anchor_by_id(row["anchor_id"]).format_info()`; resolving a `para:N` re-walks the whole
  `Paragraphs` collection over COM, so doing it once per row was **quadratic** (against live
  Word, a 1.98× bigger document made `body-font-consistent` 3.53× slower — a clean quadratic
  is 3.93×). Rows already carry the paragraph's offsets, so the range is now rebuilt with one
  `Range(start, end)` call. Second, each rule rebuilt `paragraphs.list()` / `outline()` from
  scratch: **15 full document enumerations per default pass**, now memoised to one apiece for
  the duration of the pass (scoped to the pass, so any edit invalidates it). Two COM-level
  wastes went with them: `paragraphs.list()` fetched `para.Range` three times per paragraph
  (each fetch mints an object pywin32 must wrap — the profile's dominant cost), and it probed a
  per-paragraph `InlineShapes` collection just to learn the document had none. A default `lint`
  on a 61-paragraph table document went **110.7s → 5.2s**, and now scales linearly
  (~0.08s/paragraph) instead of quadratically. The `table-repeat-header` rule — the suspected
  culprit — was innocent all along, at 0.23s.
- **`regularize` no longer warns that `allow_content` "was ignored" while honouring it.** The op
  read the field but never declared it in the batch's optional-field registry, so every
  successful `allow_content` run also emitted `op 'regularize' does not use field
  'allow_content'; it was ignored`. A phantom warning on a field that demonstrably works
  discredits the whole `warnings` channel. An AST audit of `apply_op` against the registry
  confirms this was the only field affected.
- **`manual-heading-formatting` no longer fires on every cell of a bold table header row.** The
  cells of a `heading_row` are short, emphasized and `Normal`-styled, so they matched the
  faux-heading heuristic exactly — a wide header row produced one spurious finding per cell. The
  rule now skips paragraphs inside tables; repeating table headers are `table-repeat-header`'s
  domain.
- **Markdown round-trips no longer accrete backslashes (data corruption).** `to_markdown`
  escaped ``\`` ` ``*{}[]#_`` on the way out, but `insert_markdown` only ever unescaped
  `\*` and `\\`. The other five characters came back with their backslash attached, and
  because `\` is itself escaped, every read-modify-write cycle added one more: ``use `x` ``
  → ``use \`x\` `` → ``use \\`x\\` ``, without bound. `_runs.parse_markup` now unescapes
  exactly the set `_export._escape_inline` emits (asserted by a test, so the two can't
  drift), making the round-trip a fixed point.
- **`word_read(command="find")` no longer fails when `mode` is omitted.** The MCP tool builds
  its `params` dict with every key present, so an omitted `mode` arrived as an explicit
  `None` and `.get("mode", "fuzzy")` never returned its default — the documented default was
  never wired, and the very first `find` an agent made raised
  `unknown find/replace mode None`. `read_text` had the identical bug with `view`. Both now
  coalesce. (The CLI and `exec` paths were unaffected.)
- **`word_exec` gives an actionable error when `ops` is a JSON-encoded string.** Previously
  pydantic rejected it before the tool body ran, so the client got a raw `list_type` error
  and an `errors.pydantic.dev` URL. `ops` is now typed `list[dict] | str`: a string that
  decodes to an array of op objects is accepted outright, and anything else raises an
  `OpError` naming the field, the actual problem, and the shape wanted.

## [0.18.0] - 2026-06-23

### Added
- **Chart depth — error bars, series/point formatting & trendline knobs.**
  (Priority 3, item 7) Deepens the post-insert chart surface, all on the
  BreakLink-**static** chart with **no Excel** respin (live-probed 2026-06-21).
  `chart.format_series(series=1, point=None, ...)` styles a series or one point —
  markers (`marker` glyph name/int, `marker_size`), line `smooth`, pie
  `explosion`, and per-element data-label font (`data_labels`, `data_label_size`,
  `data_label_color`). `chart.add_error_bars(kind, amount, include, axis)` draws
  `fixed`/`percent`/`stdev`/`sterror` bars. `chart.format()` gains `gap_width` /
  `overlap` (bar spacing) and `data_table`; `chart.add_trendline()` gains `order`
  (polynomial degree) and `period` (moving-average window). Wired Python / CLI
  (`format-series`, `add-error-bars`, extended `format-chart` / `add-trendline`) /
  `exec` ops (`format_series`, `add_error_bars`) / MCP. New constants
  `XlMarkerStyle` / `XlErrorBarType` / `XlErrorBarInclude` / `XlErrorBarDirection`.
  (`SaveChartTemplate`/`ApplyChartTemplate` deferred — it blocks under headless
  COM.)
- **List polish — custom list-template authoring + per-level format.** Where
  `apply_list` only applies a gallery default, `anchor.apply_list_format(levels)`
  (Priority 3, item 6) **authors a custom multi-level list template** and applies
  it: each per-level spec sets the marker `format` (`"%1."`, `"%1)"`, `"%1.%2"`),
  number `style` (`arabic`/`upper-roman`/`lower-letter`/…) or `bullet` glyph +
  `font`, plus `start_at`/`number_position`/`text_position`/`trailing`/
  `alignment`/`bold`/`italic`/`color`. More than one level mints an outline
  template. `anchor.read_list_levels()` is the read mirror (one `{level, kind,
  format, number_style, style, trailing, number_position, text_position, font}`
  per template level).
  Composes over `Document.ListTemplates.Add` + per-`ListLevel` mutation, all
  settable under late binding (live-probed 2026-06-21 — the one trap, baked in:
  a bullet level is the glyph + a symbol font, never `NumberStyle=bullet`, which
  raises). Wired Python / CLI (`list format` / `list levels`) / `exec` op
  (`apply_list_format`) / MCP (`word_write` list action `format`, `word_read`
  command `list_levels`). New constants `WdListNumberStyle` / `WdTrailingCharacter`
  / `WdListLevelAlignment`.
- **Table structural polish — add/delete columns, merge/split cells.** Completes
  the table editing surface (Priority 3): `Table.add_column(values=None)` and
  `Table.delete_column(index)` mirror the existing `add_row`/`delete_row`
  (`add_column` fills its cells top-to-bottom). `Cell.merge(other)` joins two
  cells (and the rectangle they span) into one, and `Cell.split(rows=1, cols=2)`
  is its inverse — the merged-cell "addressing story": either makes the table
  **non-uniform**, so `Table.is_uniform` reports `False`, `table:N:R:C` indexes
  *physical* cells (a merged row has fewer than `column_count`), and `Table.read()`
  / `grid()` now walk each row's physical cells (robust to irregular tables; the
  read carries a `uniform` flag). `delete_column` on a merged / mixed-width table
  raises a clean `OpError` pointing at per-cell deletion (Word has no per-column
  model there — the same contract as a column-anchor style op). No new COM surface
  beyond `Columns.Add`/`Delete` + `Cell.Merge`/`Split` (live-probed 2026-06-21).
  Wired Python / CLI (`table add-column` / `delete-column` / `merge-cells` /
  `split-cell`) / `exec` ops (`add_column`/`delete_column`/`merge_cells`/
  `split_cell`; cell coords as `[row,col]` or `"R:C"`) / MCP.
- **Markdown / HTML export — `doc.to_markdown()` / `doc.to_html()`.** The read
  mirror of `insert_markdown`: serialise the whole document, or any anchor's
  range (`within=`), to clean Markdown or an HTML fragment. Both render from one
  COM document-walk so they agree on structure — headings, nested bullet/numbered
  lists, `**bold**`/`*italic*` (HTML keeps underline), GFM pipe tables (with
  pipe-escaping + alignment), inline images as `![alt](image:N)`, and hyperlinks
  as `[text](url)`. Export is **lossy by design** (like the constrained-subset
  import): it round-trips the dialect import speaks and reads the rest richer.
  Pure reads — no new COM surface, composed over `Range.Words` (per-word
  emphasis), `ListFormat`, the table range-interval walk, and `Range.Hyperlinks`
  (all live-probed). Wired Python / CLI (`read markdown` / `read html`, with
  `--within`; `--text` pipes the raw markup) / MCP (`word_read command=to_markdown`
  / `to_html`). Not `exec` ops (reads). New module `_export.py` holds the node
  model + the two pure emitters.
- **Token-budgeted whole-document read — `doc.read(budget=N, depth=None)`.** A
  structure-aware compressed read of the entire document sized to a token budget,
  so an agent can load a large doc into context cheaply while **every anchor
  stays addressable**. Headings are emitted verbatim (each tagged with its
  `<!-- heading:N -->` anchor — the navigation spine), tables become one-line
  shape stubs, and body text is sampled to fit `budget` (~4 chars/token),
  weighted so shallower sections keep more than deep ones; overflow elides to
  markers that name the `para:` range so an agent can drill in via
  `to_markdown(within=…)`. `depth` caps how deep a section keeps body. Reuses the
  `to_markdown` document walk; the eliding heuristic is tunable. Wired Python /
  CLI (`read digest [--budget] [--depth]`) / MCP (`word_read command=digest`).
- **Table styling & polish — restyle existing tables + row/column anchors.**
  Closes the "create-but-can't-edit" gap for tables: a table's style was settable
  only at `insert_table(style=…)` time, and styling a row/column/whole-table meant
  looping cells or dropping to `.com`. New surface, wired Python / CLI / `exec` op
  / MCP:
  - **`Table.set_style(name)`** — restyle an existing table (CLI `table set-style`,
    op `set_table_style`, MCP `table action=set_style`). Note: applying a table
    style reapplies its conditional formatting, so it **overwrites direct cell
    shading** — restyle *first*, then layer cell-level overrides.
  - **`Table.set_alignment(left|center|right)`** — the whole table across the page
    width (`Table.Rows.Alignment`; CLI `table set-alignment`, op
    `set_table_alignment`).
  - **`Table.set_borders(sides, style, weight, color)`** — borders across the
    **whole grid** in one call, the table-wide counterpart of the per-cell
    `set_borders` (interior gridlines via `horizontal`/`vertical`; CLI `table
    set-borders`, op `set_table_borders`).
  - **`Table.set_banding(first_row, last_row, first_column, last_column,
    banded_rows, banded_columns)`** — toggle Word's "Table Style Options"
    (tri-state; needs a real table style applied to show; CLI `table set-banding`,
    op `set_table_banding`).
  - **`Cell.set_vertical_alignment(top|center|bottom)`** — cell vertical alignment
    (flat CLI `cell-valign --anchor-id table:N:R:C`, op
    `set_cell_vertical_alignment`, MCP `cell_valign`). Maps onto 0/1/3 — the gap at
    2 (`wdAlignVerticalJustify`) is invalid for a cell.
  - **Row / column anchors — `table:N:row:R` (`RowAnchor`) and `table:N:col:C`
    (`ColumnAnchor`).** A row/column is now an addressable anchor, so the shipped
    `shading` / `borders` / `apply-style` / `format-run` verbs (and `set_shading` /
    `set_borders` ops) style a whole strip in one call — "shade the header row",
    "right-align the totals column", "bold column 1" — with zero new styling verbs.
    A row is a contiguous range; a **column** is not (Word exposes no `Column.Range`
    and rejects the whole column collection on a merged / mixed-width table), so a
    column op fans out across the column's cells and raises a clear `OpError` on a
    mixed-width table, pointing at per-cell `table:N:R:C` styling. `Table.row(R)` /
    `Table.column(C)` return the same anchors. Live-Word validated.
    (`RowAnchor` / `ColumnAnchor` exported from the package; `WdRowAlignment` /
    `WdCellVerticalAlignment` added to `constants`.)
  **Deferred:** merged/split-cell addressing and `add_column`/`delete_column`.
- **Checkpoint + diff — `doc.checkpoint()` / `doc.changes_since()` / `doc.diff()`.**
  Fingerprint the document's structure at one moment (`checkpoint`, a pure read),
  then produce a structured, content-aligned change list against a later moment —
  the only reliable way to answer "what changed in session" (Word emits no
  content-change event), and how a multi-step agent verifies its own edits landed
  without re-reading the whole document. `checkpoint(include=…, within=…)` returns
  an opaque, serialisable `Checkpoint` (`.to_json()` / `Checkpoint.from_json()`);
  `include` sets the fingerprint depth (`text` < `text+style` (default, a restyle
  surfaces) < `text+format` (a direct-format edit surfaces as a `reformat`)).
  `changes_since(cp)` diffs a stored checkpoint against the document now;
  `diff(cp_a, cp_b)` diffs two stored checkpoints. Each change is one of `replace`
  / `insert` / `delete` / `restyle` / `reformat`, carrying the **current** `para:N`
  (so the caller can act on it immediately) plus `index_before/after` and
  `text_before/after` / `style_before/after` as applicable. Alignment is by
  paragraph **content** (`difflib.SequenceMatcher`), not index — `para:N`
  renumbers under inserts/deletes — and an unchanged document short-circuits to
  `[]` via a whole-document `doc_hash`. **No new COM surface** — pure composition
  over `_findreplace._normalize` (the shared normalisation `find`/`find_paragraphs`
  use), `paragraph_text`, and `format_info`. Wired across Python / CLI
  (`checkpoint [--include] [--within] [--out FILE]`, `diff --since FILE` or
  `diff --from A --to B`) / MCP (`word_read command=checkpoint` / `command=diff`).
  Pure reads — not `exec` ops (the token round-trips through the caller, not Word).
  **Deferred:** pin-backed exact identity (`track=True`), move detection
  (`moves=True`), per-cell table diffing, and an in-document checkpoint store.
- **Floating-shape anchor model — `shape:N`.** A new positional anchor over the
  document's body-story floating shapes (text boxes, floating images, WordArt),
  in document order — the restyle handle the deferred "image polish" cluster was
  blocked on. Resolves via `doc.anchor_by_id("shape:N")`, `doc.shapes` (all body
  shapes), and `doc.text_boxes` (the text-box subset, a discovery filter that
  keeps each box's canonical `shape:N` id). `ShapeAnchor` carries `shape_type`
  (`text_box`/`picture`/`wordart`/…) and the in-place mutators `set_wrap`,
  `set_position(left/top/relative_to)`, `set_size(width/height/lock_aspect)`,
  `format(fill/border/border_weight)`, `set_alt_text`, `set_text` (text boxes),
  `replace_image` (floating pictures), and `delete` — wired across Python / CLI
  (`shapes`, `set-shape-wrap`, `set-shape-position`, `set-shape-size`,
  `format-shape`, `set-shape-alt-text`, `set-shape-text`, `replace-shape-image`,
  `delete-shape`) / `exec` ops (same names with `set_shape_*` / `format_shape` /
  `replace_shape_image` / `delete_shape`) / MCP (`word_read command=shapes`,
  `word_write` commands of the same names).
  - **`insert_text_box` now returns its `ShapeAnchor`** (was `None`), and a
    **floating `insert_image` returns the picture's `ShapeAnchor`** (an `inline`
    image still returns `None` and stays `image:N`) — so a just-placed shape can
    be restyled without re-discovering it.
  - `replace_image` swaps a floating picture's bits **in place** by delete +
    reinsert at the same anchor, preserving wrap / position / size / alt text
    (live-probed: `Shape.Fill.UserPicture` only *overlays* a second picture-fill
    on a picture shape, so it's not a true replace). Header-story watermarks are
    excluded from `doc.shapes`. `shape:N` is positional — it renumbers when an
    earlier shape is added/removed, so re-list rather than cache (the `image:N` /
    `chart:N` rule). (`MsoShapeType` / `WdStoryType` and `WdRelative*Position.PAGE`
    added to `constants`.)
- **Shape depth, inline-image restyle & a `textbox:N` alias** (building on
  `shape:N`):
  - **Shape depth.** `ShapeAnchor` gains `set_rotation(degrees)`,
    `set_z_order("front"|"back"|"forward"|"backward")` (restack within the float
    layer — distinct from wrap's in-front-of/behind-text), and
    `set_text_frame(margin_*/word_wrap)` for a text box's internal insets;
    `doc.shapes.list()` now reports `rotation` and `z_order`. **Group / ungroup:**
    `doc.group_shapes(*shape_ids)` collapses two or more floats into one group
    `shape:N` (it enables `AllowOverlap` on the members first — the live-probed
    prerequisite), and `ShapeAnchor.ungroup()` dissolves a group back into its
    members' `ShapeAnchor`s. (No autosize: Word's "resize-to-fit-text" doesn't
    set cleanly over COM — `TextFrame.AutoSize` no-ops and `TextFrame2.AutoSize`
    rejects the value — so it's intentionally omitted.)
  - **Inline-image restyle.** `ImageAnchor` (`image:N`) gains `set_alt_text` and
    `set_size(width/height/lock_aspect)` — alt text and resize for an *inline*
    picture without floating it (re-wrapping an image is still `insert_image`,
    which crosses it into `shape:N`).
  - **`textbox:N` addressing alias.** `doc.anchor_by_id("textbox:N")` resolves to
    the Nth text box, returning a `ShapeAnchor` that reports its canonical
    `shape:M` id (a thin alias, not a second id space).
  - Wired across Python / CLI (`set-shape-rotation`, `set-shape-z-order`,
    `set-shape-text-frame`, `group-shapes`, `ungroup-shape`, `set-image-alt-text`,
    `set-image-size`) / `exec` ops (same names) / MCP (`word_write` commands of
    the same names). (`MsoZOrderCmd` added to `constants`.)
- **Image polish — wrap side + text distance, cropping, and a vision-first
  `read_image`** (rounding out the "image polish" cluster on the `shape:N` model):
  - **Wrap side + standoff distance.** `ShapeAnchor.set_wrap` takes optional
    `side` (`both`/`left`/`right`/`largest` — which sides body text flows past;
    only `square`/`tight`/`through` honour it, the rest coerce back to `both`) and
    `distance_top`/`distance_bottom`/`distance_left`/`distance_right` (the standoff
    gaps between text and the shape). `wrap` is now optional (pass any one of the
    three). `doc.shapes.list()` reports `wrap_side`. (`WdWrapSideType` widened with
    `LEFT`/`RIGHT`/`LARGEST` in `constants`.)
  - **Cropping.** `ShapeAnchor.set_crop(left/top/right/bottom)` trims a floating
    *picture* shape in from its edges (guarded to picture shapes — cropping a text
    box raises a clean `OpError`); `ImageAnchor.set_crop(...)` does the same for an
    *inline* picture. `doc.shapes.list()` / `doc.images.list()` report a `crop`
    inset dict (or `None`).
  - **MCP `read_image` now hands back a real image block.** `word_read
    command=read_image` returns the extracted picture as an inline image content
    block (like `word_snapshot`) plus a compact `{anchor_id,mime,bytes}` label, so
    a vision model *sees* the original directly instead of base64 text it can't
    decode. (The CLI `read-image` and the Python `read_image()` are unchanged —
    they still return bytes/base64.)
  - Wired across Python / CLI (`set-shape-wrap` gains `--side` / `--distance-*`;
    new `set-shape-crop`, `set-image-crop`) / `exec` ops (`set_shape_crop`,
    `set_image_crop`; `set_shape_wrap` gains `side`/`distance_*`) / MCP
    (`word_write` `set_shape_crop` / `set_image_crop` with `crop_*` params,
    `set_shape_wrap` with `side`/`distance_*`).
- **Post-creation restyle parity — content controls & hyperlinks.** Two objects
  that accepted styling/config at insert time but had no way to change it
  afterward now have in-place mutators (the iterate-without-delete-and-reinsert
  pattern Charts already set):
  - **Content controls** — `ContentControl.set_properties(title=…, tag=…,
    lock_contents=…, lock_control=…)` re-sets a control's metadata (tri-state;
    a rename changes its `cc:NAME` id) and `ContentControl.set_items([...])`
    replaces a combo_box/dropdown's choice list. CLI `set-cc-properties` /
    `set-cc-items`; exec ops `set_cc_properties` / `set_cc_items`; MCP
    `word_write` commands of the same names.
  - **Hyperlinks** — `doc.hyperlinks` is no longer read-only:
    `Hyperlink.update(address=…, sub_address=…, text=…, screen_tip=…)` (and the
    individual `set_address` / `set_sub_address` / `set_text` / `set_screen_tip`)
    retarget or relabel a link in place. CLI `set-hyperlink --index N …`; exec op
    `set_hyperlink` (addressed by 1-based `index`); MCP `word_write` command
    `set_hyperlink` (`url`→address, `bookmark`→sub_address, matching
    `add_hyperlink`'s vocabulary).
- **Format read mirror — `anchor.format_info()`.** The missing read counterpart
  of `format_paragraph` / `format_run`: returns an anchor's *effective* paragraph
  and character formatting over the same field vocabulary the write verbs accept,
  each field annotated `{value, style, override}` — the effective value, the
  applied style's baseline, and whether a **direct override** sits on top.
  `font.mixed` lists fields that read `wdUndefined` because they vary across the
  range's runs (so they're never reported as a bogus number). Lengths in points,
  colour as `#RRGGBB` / `"auto"`, alignment/line-spacing as the same keywords the
  write side takes. Pure read. CLI `wordlive read format --anchor-id ID`; MCP
  `word_read` command `format_info`. (`WD_UNDEFINED` added to `constants`.)
- **Linter + formatting regularizer — `doc.lint()` / `doc.regularize()`.** A
  declarative rule set that audits a document for publishing-quality defects and
  autofixes the mechanical ones. Pure composition over shipped write primitives —
  no new COM write surface.
  - `doc.lint(rules=None, within=None)` → a severity-ranked list of findings
    `{rule, kind, severity, anchor_id, message, fixable, fix, observed,
    expected}`. `kind` is **consistency** (a direct override fighting the applied
    style — a `Heading 1` at 15pt), **structural** (an objective layout defect),
    or **policy** (a house-style target — none ship yet). A `fixable` finding
    carries an op-shaped `fix` (literally an `exec` op) describing exactly what
    `regularize` will change. Pure read.
  - `doc.regularize(rules=None, within=None, dry_run=False)` → applies the fixable
    subset in one `doc.edit("Regularize formatting")` (one Ctrl-Z reverts the
    whole pass). Returns `{applied, skipped, findings}`. The default fixes are
    **targeted and idempotent** — they write the style's own value back as a
    direct property, so a second `regularize` is a no-op (a tested invariant).
  - **Rules (v1):** structural — `heading-keep-with-next` (a heading that may
    dangle at a page foot), `table-repeat-header` (a multi-page table with no
    repeating header row → `set_heading_row(1)`), `list-numbering-continuity` (a
    numbered list Word split into independent "1." runs → remove + reapply one
    list); consistency — `heading-font-consistent` / `heading-spacing-consistent`
    / `body-font-consistent` (direct overrides drifted from the style → write the
    style value back), `mixed-run-format` (a heading with mixed runs, report-only).
  - `rules` selects by id / tag (`["headings", "lists"]`) or `{"exclude": [...]}`;
    `within=anchor` scopes the audit to a heading section / `range:` / table.
  - Surfaces: Python `doc.lint` / `doc.regularize`; CLI `wordlive lint` /
    `wordlive regularize [--dry-run]` (`--rule` / `--exclude` / `--within`);
    `regularize` **exec op** (a write, for the atomic-undo batch); MCP
    `word_read` command `lint`, `word_write` command `regularize`. `Finding` is
    exported from the package. Detailed design: `spec-linter.md`.

### Fixed
- **Export: underline never reached HTML output.** `to_html` advertised it kept
  underline, but the COM walk routed `Font.Underline` (a `WdUnderline` enum:
  `1`=single, `3`=double, …) through the `-1`/`0` bold/italic coercion, so every
  underline read as off. Underlined text now renders `<u>…</u>`.
- **Export: `doc.read(budget=N)` could exceed the budget on many-section docs.**
  Per-section lead snippets were emitted regardless of remaining budget, so output
  grew with section count. Snippets now draw from one shared budget pool (the
  heading spine remains the fixed navigation backbone, as documented).
- **Export: multi-word hyperlinks split into one link per word** (e.g.
  `[one](url) [two](url)`); adjacent spans of the same hyperlink now coalesce into
  a single `[one two three](url)`. Link word-tagging also switched from a
  point-test to range-overlap so a trailing word can't be dropped.
- **Export: inline images inside table cells became unaddressable** — they're now
  appended as `![alt](image:N)` so every `image:N` stays referenceable.
- **Checkpoint/diff: a table-only edit reported no changes.** The table hash
  broke the fast path but the diff only walked paragraphs, so `diff` returned `[]`
  for a changed document. Tables are now compared, emitting
  `table_change` / `table_insert` / `table_delete` records.
- **Checkpoint: the `text+format` fingerprint was O(n²)** (it re-resolved each
  `para:N`, re-enumerating paragraphs from the top); it now reads formatting off
  the already-iterated range.
- **Checkpoint: field-code/view state could cause phantom changes.** Paragraph
  text is now read with a pinned `TextRetrievalMode`, so toggling ShowFieldCodes
  between checkpoints no longer surfaces a spurious `replace`.
- **Checkpoint: `reformat` could be reported with no format data** (in
  `text`/`text+style` mode); a same-text change is now `reformat` only in
  `text+format` mode when the format fingerprint actually moved, else `restyle`.
- **Checkpoint: `changes_since` now rejects an offset-based `range:` scope** with
  a clear error (offsets shift under edits; use a stable anchor or `diff()` of two
  stored checkpoints).
- **Linter: `list-numbering-continuity` missed number-only / mixed numbered
  lists** (only simple/outline were recognised) — all numbered list types are now
  checked.
- **Linter: `regularize` dropped `run_batch`'s failure detail** on a failed fix;
  the structured failure now rides on the raised error.
- **Lists: `apply_list_format` could orphan a `ListTemplate`** when a later level
  spec was invalid — every level is now validated before the template is minted.
- **Charts: `add_error_bars` alias spellings bypassed amount validation**
  (`percentage` / `standard_deviation` silently drew amount-1.0 bars); the
  amount-required check now keys off the resolved error-bar type.
- **Tables: `Cell.merge` docstring corrected** — the merged cell is addressed by
  the rectangle's **upper-left** coordinate, not "this cell".
- **Tables: column styling re-resolved each cell through `Table.cell()`** (an
  extra COM round-trip + a logical-bounds check); `ColumnAnchor` now styles the
  already-read column cells directly.
- **Shapes: `replace_image` dropped crop, rotation, wrap side, and text-distance
  standoffs** — the delete+reinsert swap now preserves the full layout.

- **Low-severity review polish (correctness/robustness across modules).**
  - *Charts:* `format_chart(data_table=…)` on a pie/scatter now raises a clear
    `OpError` instead of a raw HRESULT; `format_series(data_labels=False, …font…)`
    no longer silently re-enables labels (it errors); `add_trendline(order)` /
    `format_series(marker_size, explosion)` validate their documented ranges.
  - *Shapes:* `set_size(width, height)` (both dims, no `lock_aspect`) restores the
    prior aspect-lock instead of leaving it permanently off; `set_position(left=…,
    relative_to=…)` only re-frames the axis being moved (a horizontal move no
    longer shifts the shape vertically); inserting/grouping a shape no longer
    leaves the internal `_wl_shape_*` probe name when the original name was empty;
    a `ShapeAnchor`'s `text_final`/`text_original`/`revision_segments` now mirror
    the shape's own text rather than the anchoring paragraph's history.
  - *Lists:* a multi-level number level with no explicit `format` keeps Word's
    built-in outline default instead of being overridden to a flat `%N.`;
    `read_list_levels()` now also surfaces the raw `number_style` int.
  - *Export:* Markdown link/image targets with spaces or parentheses are
    angle-bracket-wrapped (`<url>`) so they don't break; a fully-shown lead snippet
    in `read(budget=N)` no longer emits a spurious "N more words" marker; the
    heading-style fallback resolves localized built-in names (works on non-English
    Word), not just the English `Heading N`.
  - *Checkpoint:* an unreadable (merged) table cell now gets a positional sentinel
    in the fingerprint rather than being silently dropped.
  - *Linter:* the table-repeat-header rule computes pages from one `location()`
    over the table range (one repaginate, merged-grid-safe) and narrows its broad
    excepts to the typed COM error.
  - *Docstrings:* `XlErrorBarInclude` (`include=`, not `direction=`), MCP
    error-bar `axis` choices (`y|value|x|category`), and the CLI border `--style`
    ↔ exec/MCP `line_style` alias are now documented accurately.

### Changed
- **Crop edge fields accept `crop_*` aliases everywhere.** The `set_shape_crop` /
  `set_image_crop` exec ops (and CLI) accept `crop_left`/`crop_top`/… in addition
  to `left`/`top`/… — matching the MCP `word_write` names — so a batch authored
  against either vocabulary works.

## [0.17.0] - 2026-06-17

### Added
- **Structural query helpers.** Three pure document reads that navigate and
  locate by structure, composing over the existing outline/find primitives:
  - `doc.between(start, end, *, inclusive=False)` — a `RangeAnchor` spanning the
    gap between two anchors (the headline use is two `heading:N` ids: the block
    between two headings). Default excludes both bounding paragraphs;
    `inclusive=True` covers them. CLI `wordlive read between --start ID --end ID
    [--inclusive]`; MCP `word_read` command `between`.
  - `doc.nearest_heading(where, *, direction="before")` — the heading nearest a
    position (`anchor` id / `Anchor` / char offset). `before` = the enclosing /
    preceding heading, `after` = the next one. Returns an `outline()`-shaped row
    or `None`. CLI `wordlive read nearest-heading --anchor-id ID [--direction
    before|after]`; MCP `word_read` command `nearest_heading`.
  - `doc.find_paragraphs(text, *, limit=5, min_score=0.6)` — **fuzzy** paragraph
    search: scores every paragraph against `text` with `difflib.SequenceMatcher`
    over the same normalization `find()` uses (NFKC, smart quotes, dashes,
    whitespace), so a typo'd or paraphrased query still locates its `para:N`.
    Returns ranked `{anchor_id, index, score, text, level, is_heading}` rows.
    Unlike `find()` (exact substring → `range:START-END`), this is similarity
    ranked → `para:N`. CLI `wordlive find-paragraph --text T [--limit N]
    [--min-score F]`; MCP `word_read` command `find_paragraphs`.

  (Content-under-heading was already covered by `Heading.section_range()` /
  `read section`.) All three are pure reads — no `exec` ops, no `__init__`
  exports; they leave selection/scroll/`Saved` untouched.
- **Charts (Excel-backed).** `anchor.insert_chart(kind, data, *, title=None)`
  embeds a chart via `InlineShapes.AddChart2` — `kind` ∈
  `bar`/`pie`/`line`/`scatter`. `data` is a `{label: value}` mapping (bar/pie/
  line) or an array of `[x, y]` pairs (scatter — numeric axes, duplicate x kept;
  line accepts either). Returns a `ChartAnchor` (`chart:N`); discover charts via
  `doc.charts` (metadata only: kind, title, para). CLI `wordlive insert-chart
  --anchor-id ID --kind K --data JSON` / `wordlive charts`; exec op
  `insert_chart` (outputs `chart:N`); MCP `word_write`/`word_exec` command
  `insert_chart`, `word_read` command `charts`.
  - **Charts are Excel-backed**, so they need Excel installed: a non-invasive
    registry probe gates the insert and raises the new `ExcelNotAvailableError`
    (**CLI exit code 6**, parallel to "Word not running"'s 4) before touching the
    document. Several hard-won live-Word mechanics are encoded in `_charts.py`:
    `AddChart2` only works off the `Selection` (a `Range` raises "Requested
    object is not available"); data is written into the embedded workbook's cells
    and bound with a `=SERIES(...)` formula (the `Series.XValues`/`.Values` array
    setters are unreliable under pywin32 late binding, and a literal x-array
    stores text — breaking a scatter's numeric axis); and `ChartData.BreakLink()`
    runs before closing the workbook so the chart's data goes **static** and the
    hidden Excel terminates instead of orphaning (an orphaned data grid otherwise
    locks all later inserts with "the chart data grid is already open").
- **Chart formatting & design.** A curated formatting surface on `ChartAnchor`
  (`chart:N`) — Word's "Design"/"Format" tabs — operating on the **post-insert,
  static** chart, so it needs **no Excel** (live-probed: zero embedded-Excel
  respin / orphans). All fields are tri-state and the methods chain:
  - **`chart.format(...)`** — title, legend (+`legend_position`), `chart_style`
    (design-gallery int), chart/plot background fills, whole-chart font,
    `data_labels` (+number format), and `chart_type` to re-type in place.
  - **`chart.set_axis(which, ...)`** — `which` = `value`/`y` or `category`/`x`;
    title, min/max, `scale` (`linear`/`log`), number format, gridlines.
  - **`chart.add_trendline(...)`** — linear/exponential/logarithmic/
    moving_average/polynomial/power on a series, with `display_equation` /
    `display_r_squared` and forward/backward forecast (a power fit + equation
    draws the law of best fit).
  - **`chart.set_series_color(color, *, series=1, point=None)`** — recolour a
    whole series or one 1-based point/slice.
  - Read side gains `chart.chart_style` / `chart.has_legend` (and the same two
    fields in `doc.charts.list()`). CLI `format-chart` / `format-axis` /
    `add-trendline` / `set-series-color`; exec ops + MCP `word_write`/`word_exec`
    commands `format_chart` / `format_axis` / `add_trendline` / `set_series_color`.
- **Revision write surface — accept / reject tracked changes.** The read side
  (`doc.revisions`, `snapshot(markup="all")`) shipped in v0.12.0; mutating a
  `Revision` no longer needs the `.com` escape hatch:
  - **`doc.revisions[N].accept()` / `.reject()`** resolve a single tracked change
    (accepting consumes it and renumbers the rest). CLI `wordlive revision accept
    --index N` / `reject`; exec ops `accept_revision` / `reject_revision`; MCP
    `word_write` command `revision` (`action=accept|reject`).
  - **`doc.revisions.accept_all(within=anchor)` / `.reject_all(within=anchor)`**
    resolve every tracked change at once — whole-document by default, or scoped to
    any anchor's range (`within=heading` / `range:` / cell / …) so an agent can
    "accept all my edits in this section". Returns the count resolved. CLI
    `wordlive revision accept-all [--anchor-id …]` / `reject-all`; exec ops
    `accept_all_revisions` / `reject_all_revisions`; MCP `revision`
    (`action=accept_all|reject_all`). The top-level `revisions` command stays as
    the alias for the new `revision list`.
- **Revision-aware reads — `Anchor.text_final` / `text_original` /
  `revision_segments()`.** A tracked edit's two sides live in different places:
  Word's `Range.Text` returns the **final** view (inserted runs present, deleted
  runs gone), while the deleted text survives only on the delete `Revision`.
  These reconstruct both — `text_final` (as if accepted), `text_original` (as if
  rejected), and `revision_segments()` (the ordered `{text, change}` breakdown,
  `change` ∈ insert/delete/None). CLI `wordlive read text --anchor-id ID --view
  raw|final|original|segments`; MCP `word_read` command `read_text` (`view=…`).
- **Watermark — `doc.set_watermark(text, …)` / `doc.remove_watermark()`.** Stamps
  a text watermark (DRAFT / CONFIDENTIAL) behind every page via WordArt in each
  section's header story (the same shape name as Word's *Design → Watermark*, so
  it replaces a ribbon-added one). `layout="diagonal"|"horizontal"`, `color`,
  `font`, `semitransparent`; setting twice replaces rather than stacks; removal
  is idempotent. CLI `wordlive watermark --text … [--layout …]` / `--remove`;
  exec ops `set_watermark` / `remove_watermark`; MCP `word_write` `watermark`.
- **Text box / pull quote — `anchor.insert_text_box(text, …)`.** A floating
  `Shapes.AddTextbox` anchored to any anchor's paragraph, with `width` / `height`
  (points or unit strings), `wrap` (the `insert_image` vocabulary minus inline),
  `where`, text formatting (`font` / `size` / `bold` / `italic` / `alignment`),
  and `fill` / `border`. CLI `wordlive insert-text-box --anchor-id ID --text …`;
  exec op `insert_text_box`; MCP `word_write` `text_box`.
- **Durable handles (`pin:`) & stale-anchor diagnostics.** The fix for fragile
  positional `para:N` / `heading:N` ids that renumber under later inserts:
  - **`doc.pin(anchor, name=None)`** (alias `doc.stamp`) plants a Word-hidden
    bookmark (`_wl_<code>`) over an anchor's range and returns a `pin:<code>`
    anchor id — a random hex code, or a readable slug via `name="budget-intro"`.
    Word maintains the range↔bookmark association across inserts / deletes /
    edits natively, so the handle keeps pointing at the same content; a deleted
    paragraph's pin correctly vanishes (resolving raises `AnchorNotFoundError`).
    `pin:CODE` resolves through `doc.anchor_by_id` like any anchor. CLI:
    `wordlive pin ANCHOR_ID [--name SLUG]`; exec op `pin`; MCP `word_write`
    command `pin`.
  - **`doc.pin_outline(levels=…)`** (and `outline(pin=True)`) pins every heading
    in one call, returning the `{anchor_id: pin}` map — a durable navigation
    scaffold. Idempotent (reuses a heading's existing handle, keyed by range
    start). CLI `wordlive pin-outline [--levels LO HI]`; exec op `pin_outline`.
  - **`bind: "name"`** on an insert op (`insert` / `insert_block` /
    `insert_section` / `insert_markdown` / `create_table`) mints a pin on the
    freshly-inserted content and returns it in that op's `outputs` entry.
  - **Intra-batch output references.** Any exec-op field of the exact form
    `$ops[N].field` is replaced with an earlier op's recorded output before the
    op runs — e.g. create a table at op 0, then `set_cell` with
    `"table": "$ops[0].table"`.
  - **Stale-anchor recovery hints.** A missed positional `para:N` / `heading:N`
    now raises `AnchorNotFoundError` whose message explains *why* (out-of-range
    vs body-text-not-a-heading, the paragraph count, the nearest heading) and
    recommends pinning, instead of a bare "not found".
- **`python -m wordlive`.** The CLI is now runnable as a module (a thin
  `__main__` aliasing the `wordlive` console script), so tooling can drive it
  through the current interpreter without depending on the script being on PATH.
- **End-to-end CLI test suite (`tests/test_e2e_cli.py`, marker `e2e`).** Shells
  out to `python -m wordlive` against a live Word instance and walks a full
  document lifecycle — build via `exec` + verbs, read back, save/export (gated),
  then close, reopen from disk, and verify. Excluded from the default run and CI
  (needs Word); run with `uv run pytest -m e2e`.
- **`wordlive --version`/`-v` and `wordlive --about`/`-A`.** `--version` prints
  `wordlive <version>` (sourced from the package metadata via the new
  `wordlive.__version__`); `--about` renders a colourful banner with the version,
  author (Tom Villani, Ph.D.), license (MIT), and repo URL — the "word" half in
  blue and the "live" half in cyan on a terminal, clean ASCII when piped. Both
  are eager top-level flags: no
  subcommand needed and Word is never touched.

## [0.16.1] - 2026-06-16

### Fixed
- **`find_replace` no longer eats a trailing paragraph/cell mark at a segment
  boundary.** Replacing a *whole paragraph* that sat immediately before a table
  (or any segment edge) matched the trailing `\r` too, so the replacement
  deleted the paragraph break and fused the paragraph into the following table's
  first cell (e.g. a header cell read back as `"Costs decreased.Item"`). The
  normalization sentinel now maps to the offset one past the last *contributing*
  character rather than `len(s)`, so a folded-away trailing mark (`\r`, the
  `\x07` cell marker, a stripped space) is left intact. The earlier terminal-mark
  clamp only guarded the document's final mark; this fixes interior boundaries.


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


[Unreleased]: https://github.com/thomas-villani/wordlive/compare/v0.19.0...HEAD
[0.19.0]: https://github.com/thomas-villani/wordlive/compare/v0.18.0...v0.19.0
[0.18.0]: https://github.com/thomas-villani/wordlive/compare/v0.17.0...v0.18.0
[0.17.0]: https://github.com/thomas-villani/wordlive/compare/v0.16.1...v0.17.0
[0.16.1]: https://github.com/thomas-villani/wordlive/compare/v0.16.0...v0.16.1
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
