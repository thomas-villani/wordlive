# Checkpoint + diff — design sketch

Status: **sketch** (2026-06-17). Roadmap home: `feature-plan.md` Part II, Priority 1,
item 2. This is the detailed design; the roadmap keeps the one-paragraph summary.

> `doc.checkpoint()` captures an opaque, serialisable structural snapshot of the
> document right now; `doc.diff(a, b)` / `doc.changes_since(cp)` return a structured,
> content-aligned change list. Pure reads — no Word state touched, no exec ops.

---

## 1. Why this, and why it's load-bearing

Three forces converge on the same primitive:

1. **Word has no content-change event** (live-probed 2026-06-17 — see
   `feature-plan.md` Priority 7). The *only* way to answer "what did the user/agent
   change" is to fingerprint the document at two points and diff them.
2. **An agent needs to verify its own edits landed** — "did my 5 `replace` ops hit
   exactly the 5 paragraphs I meant, and nothing else?" — without re-reading the
   whole document into context.
3. **It's the substrate for two later features:** the save-hook review (Priority 7
   diffs `watch`-start vs. on-save to lint *only the changed regions*), and the
   in-session counterpart to compare/merge (Priority 5).

Distinct from existing surfaces: **Track Changes** is Word's own per-keystroke
redline (and only if the user has it on); **`doc.revisions`** reads those marks;
**compare/merge** (Priority 5) redlines two *saved files*. Checkpoint/diff is the
**in-session, track-changes-free, structural** diff between two arbitrary moments —
it works with TC off and catches changes TC never recorded.

## 2. Surface

```python
cp   = doc.checkpoint(include="text+style")     # pure read → an opaque snapshot
...                                             # edits happen (agent or user)
chg  = doc.changes_since(cp)                    # diff cp → now
chg2 = doc.diff(cp_a, cp_b)                     # diff two stored checkpoints
```

- **`doc.checkpoint(include=…, within=None)`** → a `Checkpoint` (serialisable;
  `.to_json()` / `Checkpoint.from_json()`). Pure read: walks paragraphs (like
  `outline`), touches no selection/scroll, leaves `Saved` untouched. `within=anchor`
  fingerprints just one section/range.
- **`doc.changes_since(cp)`** → diff the checkpoint against the live document now.
- **`doc.diff(cp_a, cp_b)`** → diff two checkpoints the caller is holding.
- **`include`** sets fingerprint depth (§3): `"text"` (cheapest) · `"text+style"`
  (default) · `"text+format"` (adds a paragraph-format fingerprint via `format_info`).

### Surfaces (all four agree)

- **Python:** `doc.checkpoint` / `changes_since` / `diff`.
- **CLI:** `wordlive checkpoint [--include …] [--within ID] [--out FILE]` emits the
  checkpoint JSON (stdout or file); `wordlive diff --since FILE` (cp → now) and
  `wordlive diff --from FILE --to FILE`.
- **exec op:** **none** — pure reads, and the checkpoint token round-trips through
  the *caller*, not Word (an exec batch is one Word undo step; a diff produces no
  document change to batch).
- **MCP:** `word_read command=checkpoint` / `command=diff`.

## 3. What a checkpoint contains

A structural fingerprint — enough to align and classify changes, not a full copy.

```jsonc
{
  "version": 1,
  "include": "text+style",
  "scope": null,                      // or an anchor id when `within` was set
  "paragraphs": [
    // per paragraph, in document order:
    {"i": 0, "text": "Introduction", "style": "Heading 1", "level": 1,
     "list": null, "fmt": null, "key": "h1:9f3a…", "hash": "c21b…"}
    // key  = alignment hash for move/identity (normalized text + style)
    // hash = content hash (what changed-detection compares)
    // fmt  = paragraph-format fingerprint, only when include="text+format"
  ],
  "tables": [ {"index": 1, "shape": [3, 4], "cells_hash": "…"} ],
  "doc_hash": "…"                     // whole-doc fast-path: equal ⇒ no changes
}
```

- **Text** is **normalised** with the *same* pipeline `find()` / `find_paragraphs`
  use (NFKC, smart quotes, dashes, whitespace) so the diff agrees with the rest of
  wordlive and ignores cosmetic-only churn.
- **`key`** (normalised text + style) is the alignment key for §4; **`hash`** is the
  change key. Separating them lets us classify *restyle* (same text, key differs
  only by style) vs *replace* (text differs).
- **`fmt`** is omitted unless `include="text+format"` — a fingerprint of the
  paragraph's `format_info()` (§ linter spec 7a), so a pure-formatting edit (a
  paragraph re-justified, spacing changed) shows up as a `restyle`/`reformat`
  without bloating the default checkpoint.
- Checkpoints are **not stored in the document** by default (a pure read). Caller
  holds the token. *(Deferred: an in-doc store via `doc.variables` so a checkpoint
  survives a close/reopen.)*

## 4. The diff — content alignment, not index alignment

`para:N` is **positional** and renumbers under inserts/deletes, so the diff must
**not** compare by index. Align the two paragraph sequences by content key with
`difflib.SequenceMatcher` (the same engine `find_paragraphs` already uses), then
classify each opcode:

| SequenceMatcher opcode | classified as | emitted |
|---|---|---|
| `equal`, keys identical | (no change) | — |
| `equal` text, style/fmt differ | **restyle** / **reformat** | `{op, anchor_id, style_before, style_after}` |
| `replace` | **replace** (text edit) | `{op, anchor_id, text_before, text_after}` |
| `insert` | **insert** | `{op, anchor_id, text_after}` |
| `delete` | **delete** | `{op, index_before, text_before}` (no live anchor) |
| matched block moved | **move** (opt-in) | `{op, anchor_id, index_before, index_after}` |

Change record shape:

```jsonc
{"op": "replace", "anchor_id": "para:14", "index_before": 12, "index_after": 14,
 "text_before": "Costs fell 4%.", "text_after": "Costs fell 9%.",
 "style_before": "Normal", "style_after": "Normal"}
```

- **Inserts/restyles/replaces** carry a live `anchor_id` (the *current* `para:N`)
  so the caller can act on the change immediately (re-lint it, comment on it).
- **Deletes** reference the *old* index + text (no live anchor — it's gone).
- **Move detection** is opt-in (`moves=True`): it's the noisy part (a cut-paste vs
  delete+insert is ambiguous); default off, surfaced as delete+insert.
- A fast path: equal `doc_hash` ⇒ return `[]` without walking (cheap "nothing
  changed" check).

## 5. Tracked checkpoints (opt-in, exact identity)

The §4 content alignment is heuristic — two identical paragraphs (e.g. blank lines,
repeated headers) can mis-align. For exact paragraph identity across edits, an
opt-in **tracked checkpoint** pins each paragraph first:

- `doc.checkpoint(track=True)` plants a hidden `pin:` over every paragraph (reusing
  the shipped pin machinery — Word maintains the range↔bookmark association
  natively across edits), so the diff matches by pin identity, not content
  guesswork — and a moved paragraph is *unambiguous*.
- **Cost:** `track=True` **mutates** the document (pins are hidden bookmarks), so it
  is **not** a pure read — it's gated like a write (its own `doc.edit()`), opt-in,
  and the inverse of the cheap default. Use it when exactness matters (legal
  redlines, move tracking); use the default content diff for everything else.
- Symmetry with the linter's politeness split: cheap pure-read default, explicit
  opt-in for the heavier/mutating mode.

## 6. Worked uses

- **Verify my edits:** `cp = doc.checkpoint(); …agent edits…;
  assert {c.anchor_id for c in doc.changes_since(cp)} == expected` — confirm the
  agent touched exactly what it meant to.
- **What did the user change:** checkpoint before handing control back; on return,
  `changes_since(cp)` → the user's edits as a structured list (no TC required).
- **Lint only the delta** (Priority 7): on `DocumentBeforeSave`, diff against the
  watch-start checkpoint and run `doc.lint(within=…)` over just the changed
  paragraphs — cheap, targeted review.
- **Two-draft diff in session:** two checkpoints of the same doc at different times,
  `doc.diff(a, b)` (the file-level redline is compare/merge, Priority 5).

## 7. Open questions / probes

- **Hash stability:** the content hash must be deterministic across a session
  (stable normalisation, no `Math.random`/locale drift). Pick a fixed algorithm
  (e.g. SHA-1 of the normalised string); document it so checkpoints are comparable
  across runs/versions (carry `version` for forward-compat).
- **Granularity:** paragraph-level v1. Sub-paragraph (run-level) diffing is
  deferred — a `replace` reports the whole paragraph's before/after text, and the
  caller can sub-diff the strings itself.
- **Tables:** v1 fingerprints a table as `{shape, cells_hash}` (detects *a* cell
  changed); per-cell diffing is a follow-up once the paragraph diff is solid.
- **Performance:** the checkpoint read enumerates every paragraph (≈ `outline`
  cost). Fine for normal docs; note the cost for very large ones. `doc_hash`
  short-circuits the common "unchanged" case.
- **Story scope:** v1 is the main story (like most reads). Headers/footers/footnotes
  are out until story-aware reads land (Part III).

## 8. Build order

1. `Checkpoint` dataclass + `doc.checkpoint(include="text+style")` (paragraph walk,
   normalised keys + hashes, `doc_hash`). Pure read, serialisable.
2. `doc.diff` / `changes_since` — the `SequenceMatcher` alignment + opcode
   classification (replace / insert / delete / restyle). The core.
3. CLI (`checkpoint`, `diff --since` / `--from/--to`) + MCP `word_read`.
4. `include="text+format"` (reuse the linter's `format_info` fingerprint) →
   `reformat` changes.
5. Opt-in `track=True` (pin-backed exact identity) + `moves=True`.
6. Table per-cell diffing; later, the in-doc checkpoint store (`doc.variables`).
