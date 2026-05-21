# wordlive — feature roadmap

Post-v0 sketch. v0 (the initial commit) covered `attach`/`connect`, anchors
(bookmark / content control / heading) with read + `set_text`, `doc.edit()`
with `UndoRecord` + Selection preservation, the LLM-first CLI (`status`,
`outline`, `read`, `write`, `insert`), typed exceptions, and magic-constant
enums. This document stages everything else, ordered by **LLM-agent
leverage** rather than spec order.

> Two big surfaces are missing from `spec.md` entirely: **tables** and
> **comments**. Both rank above styles and sections for the actual
> document-editing workflows wordlive targets. Called out below.

---

## v0.1 — close the LLM-tool loop

Small, no new COM territory. Makes the existing CLI actually usable as a
deterministic tool-call surface. All four items wire existing library
primitives into the JSON-in / JSON-out shell.

- **`replace --anchor-id <id>`** — `outline()` already emits IDs like
  `heading:3`; nothing consumes them yet. Add a unified
  `doc.anchor_by_id(id)` resolver (`heading:N`, `bookmark:NAME`, `cc:NAME`)
  and a `replace` subcommand that calls `set_text` on the resolved anchor.
- **`go_to --anchor-id <id>`** — `Document.go_to()` already exists in the
  library; expose via CLI for the same ID scheme as `replace`.
- **`exec --script ops.json`** — batch N anchor ops in one `doc.edit()` /
  `UndoRecord`, so a multi-step LLM edit reverts with one Ctrl-Z. Script
  format:
  ```json
  {
    "label": "Update report",
    "ops": [
      {"op": "write_bookmark", "name": "Address", "text": "..."},
      {"op": "write_cc", "name": "Signatory", "text": "..."},
      {"op": "insert_after_heading", "heading": "Risks", "text": "..."},
      {"op": "replace", "anchor_id": "heading:3", "text": "..."}
    ]
  }
  ```
  Failure semantics: if op K fails, ops 1..K-1 are already applied; we
  close the UndoRecord and report failure. User's Ctrl-Z reverts the
  partial work in one keystroke.
- **`NamedRange` anchor — deferred.** Spec lists it as an anchor type but
  Word has no separate "named range" concept (bookmarks *are* named
  ranges). A real `RangeAnchor` over `doc.range(start, end)` is a useful
  abstraction but needs its own design pass — push to v0.2 alongside the
  other Range work.

Estimate: one sitting. Tests build on existing `fake_word` fixture.

---

## v0.2 — styles + tables

Two parallel tracks; each is its own PR.

### Styles — ✅ shipped in v0.3

- ~~`doc.styles["Body Text"]`, `doc.styles.list()`, `style.exists`.~~ ✅
- ~~`anchor.apply_style(name)` for all anchor types.~~ ✅
- ~~Wire the existing `--style` arg on `insert` (currently stubbed).~~ ✅ —
  now validates via `doc.styles[]` before any mutation, exit code 2 on miss.
- ~~Paragraph formatting that ships alongside: alignment, indent, spacing.~~ ✅
  via `anchor.format_paragraph(**kwargs)` + `wordlive format-paragraph` CLI.
- ~~New typed error: `StyleNotFoundError`.~~ ✅ — subclass of
  `AnchorNotFoundError` so it reuses exit code 2 and `except AnchorNotFoundError`
  still catches it.
- Line spacing, character-style modelling, theme-aware fonts, and style
  creation/modification stay deferred to v0.6+.

### Tables — ✅ shipped in v0.4

- ~~`doc.tables` collection — `__getitem__` by index or by table caption.~~ ✅
  index by 1-based position or `Title`.
- ~~`Table` wrapper: `row_count`, `column_count`, `cell(row, col)`, iteration.~~ ✅
  plus `read()` / `grid()` / `to_dict()`.
- ~~`Cell.text` (read/write), `Cell.anchor` (so bookmarks/CCs inside cells
  resolve uniformly).~~ ✅ — `Cell` *is* an `Anchor`, so it inherits
  `apply_style` / `format_paragraph` / `set_text` directly.
- ~~`Table.add_row(values=None)` / `Table.delete_row(index)`.~~ ✅
- **Resolved:** anchor-id scheme is `table:N:R:C` for cells; the bare `table:N`
  is *not* an anchor (a whole table is a collection) and is addressed via
  `doc.tables[N]` / the `table` CLI group. Cells don't appear in
  `doc.outline()` (that stays heading-only).
- **Resolved:** bookmarks inside cells round-trip through `set_text` — covered
  by an E2E test (`t_bookmark_in_cell_roundtrip`).
- Deferred: merged/split-cell grids (cell addressing assumes rectangular),
  cell-level `add_column`/`delete_column`, table creation/deletion.

---

## v0.5 — collaboration features — ✅ shipped in v0.5

Genuinely LLM-shaped operations: "leave a comment on the Risks section" is
exactly the kind of polite, side-channel edit agents should prefer over
direct text mutation.

> Note: find/replace already shipped in v0.2 (commits `f90e0a9`, `cbe89cc`),
> styles + paragraph formatting in v0.3 (commit `c03b7d1`), and tables in
> v0.4. v0.5 closes out the genuinely-collaborative surface below.

- ~~**Comments** — `doc.comments.add(anchor, text, author=...)`,
  `doc.comments.list()`, `comment.resolve()`.~~ ✅ — plus `doc.comments[N]`,
  `comment.delete()` / `reopen()`, and `comment.scope_text`. Comments attach to
  any anchor's range without mutating the text.
- ~~**Track changes** — `with doc.tracked_changes(): ...` flips
  `TrackRevisions` on for the scope. Pairs with `doc.edit()`.~~ ✅ — plus the
  `doc.track_changes` read/write property, the `wordlive track on|off|status`
  CLI toggle, and a `"tracked": true` key on `exec` scripts.
- ~~**`RangeAnchor`** — `doc.range(start, end)` returns an `Anchor`-shaped
  wrapper.~~ ✅ — addressed as `range:START-END`, which is now what `find()`
  emits *and* what `anchor_by_id` resolves, so a find hit feeds straight back
  into `replace` / `comments.add`. This is what "NamedRange" was reaching for.
- **Resolved:** comments are addressed by 1-based index (`doc.comments[N]`),
  matching Word's `Comments(n)`; `comment.resolve()` uses the `Done` flag
  (Word 2013+). Track-changes exposes both a persistent CLI toggle and a
  self-restoring `tracked_changes()` scope.
- Deferred: comment replies (`comment.reply(...)`), per-revision
  accept/reject (`doc.revisions`), and author/date filtering on `list()`.

---

## v0.6 — lists & document structure — ✅ shipped in v0.6

Two tracks bundled into one release. Headers/footers turned out to be the
*easy* lift (a `HeaderFooter` is just a range, so it slots into the `Anchor`
pattern like `Cell` did); lists/numbering was the fiddly half the note below
warned about — `ListGalleries` / `ApplyListTemplate` / restart-vs-continue.

- ~~**Numbering / list management** — Word's `ListTemplates` /
  `ListGalleries` are genuinely painful; isolate this work into its own
  release rather than bundling.~~ ✅ — list verbs live on the base `Anchor`:
  `apply_list("bulleted"|"numbered"|"outline", continue_previous=...)`,
  `remove_list()`, `list_info()`, `restart_numbering()`, `indent_list()` /
  `outdent_list()`. `doc.lists` is a read-only discovery collection yielding a
  `RangeAnchor` per list.
- ~~**Sections / headers / footers** — useful, but mostly for
  template-generation workflows.~~ ✅ — `doc.sections` collection;
  `HeaderFooter` *is* an `Anchor` addressed `header:S:WHICH` / `footer:S:WHICH`
  (WHICH = primary/first/even), so `set_text` / `apply_style` /
  `format_paragraph` work on it; plus `Section.page_setup()` reads.
- **Resolved:** list ops are anchor-scoped (they act on a range's paragraphs),
  so they're methods on `Anchor` rather than a separate object — same shape as
  `apply_style`. Level control is `indent_list` / `outdent_list` (Word's own
  promote/demote), avoiding the unreliable direct `ListLevelNumber` write.
- Deferred: custom list-template authoring, per-level bullet/number-format
  control, multi-section `LinkToPrevious` editing, and `PageSetup` *writes*
  (margins/orientation are read-only for now).

---

## v0.7 — paragraph addressing & cursor surface — ✅ shipped in v0.7

Closes the "I can see headings but can't address the body" gap, unifies the odd
one-off `insert --after-heading` ergonomics, and adds the explicitly-opt-in
cursor surface that was the one piece deliberately missing from the
anchors-over-`Selection` model.

- ~~**`para:N` anchors** — every paragraph (not just headings) is addressable.~~
  ✅ — `Paragraph(Anchor)` + `doc.paragraphs` collection; `para:N` shares its
  index space with `heading:N` (a heading is `para:N` *and* `heading:N`).
  Inherits every anchor verb. New `wordlive paragraphs` listing (emits offsets);
  `outline --all` is an alias.
- ~~**`insert` ergonomics** — match every other command's `--anchor-id`.~~ ✅ —
  `insert --anchor-id ID --text … [--before|--after] [--style …]` works on any
  anchor; the old `--after-heading` flag is gone. `insert_paragraph_before/after`
  lifted to the base `Anchor`. Exec op renamed `insert_after_heading` →
  `insert_paragraph`.
- ~~**Cursor surface** — make writing at the live cursor *possible* but clearly
  non-default.~~ ✅ — `cursor read` / `cursor write` (and `Selection.write`),
  deliberately *not* an `anchor_by_id` scheme. `cursor write` opts into
  `EditScope.allow_cursor_move()` so it's the one op that intentionally moves
  the cursor; `cursor read` reports the containing `para:N`.
- **Resolved:** offset-precise, mid-paragraph insertion stays a collapsed
  `range:START-END` target (now discoverable because `paragraphs` emits offsets)
  rather than a new `insert --inline` verb.
- Deferred: raw inline `insert_before`/`insert_after` (no new paragraph) on the
  CLI, and a `write_cursor` exec op (a cursor write fights `EditScope`'s
  cursor-restore inside a batch, so it stays CLI-only for now).

---

## v0.8+ — defer

- **Events / sinks** — `WithEvents(word.com, Handler)` for
  `DocumentBeforeSave`, `WindowSelectionChange`. Wait for a concrete use
  case before designing the marshalling layer.
- **`asyncio` wrapper** — natural once events land. Sync core stays.
- **Read-model caching** — premature. Live reads are correct; cache when
  events arrive to invalidate on `DocumentChange`.
- **Styles deep cuts** — character styles, list styles, theme-aware fonts.
  Cover paragraph styles in v0.2 first, see what's actually missing.
- **PageSetup writes** — margins, orientation, page size. Reads shipped in
  v0.6; writes want their own small pass (units, section-vs-document scope).

---

## Cross-cutting work (any release)

- **HRESULT coverage** — `_BUSY_HRESULTS` is a starter set; widen as we
  hit real `pywintypes.com_error` HRESULTs in smoke runs.
- **Smoke fixtures** — a real `.docx` checked into the repo with known
  bookmarks / CCs / headings / tables, so smoke tests have a known target.
- **Docs** — `spec.md` is the design doc; a separate `cookbook.md` of
  end-to-end LLM-tool examples is probably more useful than API reference
  docs at this stage.
