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
  creation/modification stay deferred to v0.5+.

### Tables

*Not in spec, but most Word docs are mostly tables.* Worth its own design
pass before coding — open questions below.

- `doc.tables` collection — `__getitem__` by index or by table caption.
- `Table` wrapper: `row_count`, `column_count`, `cell(row, col)`, iteration.
- `Cell.text` (read/write), `Cell.anchor` (so bookmarks/CCs inside cells
  resolve uniformly).
- `Table.add_row(values=None)` / `Table.delete_row(index)`.
- Open: how do tables interact with `doc.outline()` / anchor IDs? Probably
  `table:N` and `table:N:R:C` for cell-level anchors.
- Open: bookmarks-inside-cells — confirm they round-trip through `set_text`
  (Word's bookmark-deletion-on-replace quirk gets weirder inside tables).

---

## v0.4 — collaboration features

Genuinely LLM-shaped operations: "leave a comment on the Risks section" is
exactly the kind of polite, side-channel edit agents should prefer over
direct text mutation.

> Note: find/replace already shipped in v0.2 (commits `f90e0a9`, `cbe89cc`)
> and styles + paragraph formatting in v0.3 (commit `c03b7d1`). What's left
> for v0.4 is the genuinely-collaborative surface below.

- **Comments** — `doc.comments.add(anchor, text, author=...)`,
  `doc.comments.list()`, `comment.resolve()`. Word's `Comments` collection
  is straightforward; main concern is anchor-range fidelity.
- **Track changes** — `with doc.tracked_changes(): ...` flips
  `TrackRevisions` on for the scope. Pairs with `doc.edit()` —
  "make this edit visibly", so the human can accept/reject.
- **`RangeAnchor`** — `doc.range(start, end)` returns an `Anchor`-shaped
  wrapper. Lets the same ops target arbitrary ranges, not just named ones.
  This is what "NamedRange" in the spec was reaching for.

---

## v0.5+ — defer

- **Sections / headers / footers** — useful, but mostly for
  template-generation workflows; lower priority than tables/comments for
  live-editing agents.
- **Numbering / list management** — Word's `ListTemplates` /
  `ListGalleries` are genuinely painful; isolate this work into its own
  release rather than bundling.
- **Events / sinks** — `WithEvents(word.com, Handler)` for
  `DocumentBeforeSave`, `WindowSelectionChange`. Wait for a concrete use
  case before designing the marshalling layer.
- **`asyncio` wrapper** — natural once events land. Sync core stays.
- **Read-model caching** — premature. Live reads are correct; cache when
  events arrive to invalidate on `DocumentChange`.
- **Styles deep cuts** — character styles, list styles, theme-aware fonts.
  Cover paragraph styles in v0.2 first, see what's actually missing.

---

## Cross-cutting work (any release)

- **HRESULT coverage** — `_BUSY_HRESULTS` is a starter set; widen as we
  hit real `pywintypes.com_error` HRESULTs in smoke runs.
- **Smoke fixtures** — a real `.docx` checked into the repo with known
  bookmarks / CCs / headings / tables, so smoke tests have a known target.
- **CLI `--text` mode** — currently every command emits JSON; the
  `--text` flag exists in spec but is unimplemented. Low priority.
- **Docs** — `spec.md` is the design doc; a separate `cookbook.md` of
  end-to-end LLM-tool examples is probably more useful than API reference
  docs at this stage.
