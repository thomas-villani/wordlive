# wordlive — feedback from the LLM that uses it

Notes from a live probe session driving an open Word doc end to end. The framing throughout: **the consumer is an LLM agent**, so the costs that matter are the ones that mislead a model or force brittle workarounds, not the ones a human would shrug off. Ordered by leverage.

Repro snippets use the MCP op shapes I actually called.

---

## 1. Tracked changes make the agent blind — both channels fail (highest leverage)

When `track on` is set, the agent loses the ability to perceive what it just did, in *both* of the ways it has to "see":

- **`snapshot` renders Final / No-Markup.** Revision marks (insert underline, delete strikethrough, change bars) and comment balloons do not appear — even though they are present and correct in live Word. The agent's only visual channel shows an accepted-looking document.
- **Text reads concatenate revision runs with no delimiter.** After a tracked `find_replace` of "First-level item one" → "First-level item one (revised)", `paragraphs` returned:

  ```
  "First-level item one (revised)First-level item one"
  ```

  Inserted text and the tracked deletion are glued together with no marker. To a model this reads as duplicated/corrupt content, and any subsequent `find`/offset reasoning on that paragraph is now wrong.
- **No `revisions` reader.** `comments` has a reader (good — it returned author/scope/done correctly). Tracked changes have none, so there's no structured fallback once the two channels above fail.
- **`track` can't report status over MCP.** The verb requires `on`; calling it bare errors. The CLI exposes `track status` but the MCP surface doesn't, so an agent can't even check whether it's currently in tracked mode.

**Requests**
- `snapshot` markup mode: `markup: "all" | "none"` (default could stay "none", but the agent needs to opt into seeing marks).
- A `revisions` read command returning structured entries: `{index, type: insert|delete|format, author, range, text}`.
- Make text reads revision-aware: either return "final" text by default with revision runs available on request, or tag runs with `revision_type`. The current silent concatenation is the actively-harmful part.
- Add `track status` (read) to the MCP surface to match the CLI.

This blocks a whole task class: "make tracked edits for review," "summarize/accept/reject the changes in this doc." Right now the agent can write revisions it can neither see nor enumerate.

---

## 2. In-cell `replace --find` overruns the cell boundary, and the documented fix doesn't work

The verification **guard** is great — it refused to overwrite rather than corrupt the cell, which looks like the right fix for the old occurrence-scoped table corruption. The problem is the search range.

Repro: a table with cells `Opus` (2,1) and `200K` (2,2).

```jsonc
// unscoped — expected to fail, did:
{"op":"replace","find":"Opus","text":"Claude Opus"}
// error: target resolved to 'Opus\r\x072', expected 'Opus\r\x07'
```

The match swallowed the cell/row end marker (`\r\x07`) **and a digit from the next cell** (`2` of `200K`). So far so good. But the documented remedy — scope to the cell — fails identically:

```jsonc
{"op":"replace","find":"Opus","in_anchor":"table:1:2:1","text":"Claude Opus"}
// SAME error: 'Opus\r\x072'
```

So cell-scoped `find` is **not** constraining the search to the cell's text range — it still sees across the boundary, which means the guard trips on essentially any in-cell `find` and the path the guide recommends is unusable.

Clean paths that *do* work (worth keeping as the recommended ones):
```jsonc
{"op":"replace","anchor_id":"table:1:2:1","text":"Claude Opus"}  // overwrite cell
{"op":"set_cell","table":1,"row":2,"col":1,"text":"Claude Opus"} // intended cell edit
```

**Requests**
- Fix `find` range-scoping so `in: table:N:R:C` searches only within the cell's content range (exclude the trailing cell/row markers).
- Until then, update the guide: stop recommending "re-scope to the cell anchor" for `find`-based in-cell edits; point to `set_cell` / cell-anchor `replace` instead.

---

## 3. Multi-paragraph insert + run-level formatting (the most common authoring friction)

Two gaps that together turned "add a styled bulleted section" into a heading insert + a reverse-order body insert + a second formatting pass.

- **No block insert.** There's no way to drop a contiguous run of styled paragraphs at one anchor. I had to insert paragraphs one anchor-relative op at a time and order them in reverse to dodge positional-anchor renumbering.
- **Inserted text can't carry inline runs.** A bullet like **`Bold lead`** — rest can't be created in one op; the lead-in bold required a second `find` → `apply_style "Strong"` pass per bullet.

**Requests**
- `insert_block` / `insert_paragraphs {anchor_id, before?, items:[{text, style}]}` that places a contiguous run atomically and returns the spanning range.
- Inline run support in insert ops: either lightweight markup (`**bold**`) or an explicit `runs:[{text, bold?, italic?, style?}]`. The bold-lead-in bullet is frequent enough (it's the standard "feature list" pattern) to deserve a one-shot path.

---

## 4. Numbered-list continuity is a silent trap

Applying `apply_list type:numbered` to N paragraphs individually produces N independent single-item lists — rendered **1. 1. 1. 1.** And the natural repair fails while reporting success:

```jsonc
// these return ok:true but the list still renders all "1."
{"op":"apply_list","anchor_id":"para:6","type":"numbered","continue_previous":true}
```

`continue_previous` only works on a **clean** apply. To fix an already-split list you must `remove_list` everything first, then re-apply chained in one batch — which *does* give 1, 2, 3.

**Requests**
- Apply-to-range list op (`apply_list` over a paragraph span, or accept a list of anchors) that numbers correctly by default — this is what an agent reaches for.
- Either auto-continue consecutive numbered applies, or have `continue_previous` *repair* an existing adjacent list instead of no-op'ing. At minimum, surface a warning when an apply creates a fresh list directly adjacent to another numbered list.

---

## 5. Intra-batch output references + durable handles on insert

`create_table` already returns its new `index` in `outputs` — but I can't *use* it later in the same batch (I build the whole `ops` array before seeing any output). Same for "the paragraph I just inserted."

**Requests**
- Let an op reference a prior op's output within the batch: `anchor_id: "$ops[0].table"` / `"$ops[1].anchor"`. Unlocks create-table-then-fill-cells, or chained inserts, in a single undo.
- Optional `bind: "name"` on insert ops that mints a bookmark on the new content and returns it — a *durable* handle that survives renumbering, vs. the fragile positional `para:N`. (This is the honest version of "return the anchor_id after insert," which we discussed: positional ids are positions, not identities.)

---

## 6. Smaller items

- **No `delete_paragraph`.** Stray paragraphs (e.g. the blank doc's leading empty `para:1` that sits above an appended title) can only be replaced or restyled, not removed. Add `delete_paragraph {anchor_id}` (and maybe `delete_range`).
- **Error messages leak raw control chars.** `'Opus\r\x072'` is fine for debugging but reads as noise to a model; a normalized message ("match extended past the cell boundary into the adjacent cell") would be more actionable. Low priority.
- **Document the batch anchor-resolution model.** Whether positional anchors re-resolve per-op against live state or against a pre-batch snapshot determines the correctness of every multi-insert pattern. Right now I have to design around *both* interpretations (reverse-order-after-a-fixed-anchor, which is safe either way). Stating the contract — and ideally guaranteeing per-op live resolution — removes the guesswork.

---

## What's already good (keep)

- The verification guard that refuses ambiguous/boundary-crossing replaces instead of corrupting (§2) — exactly right, just over-triggering.
- `outputs` reporting new table indices, and `warnings` for unused fields — the right instinct for agent feedback; extend it (§5).
- `comments` structured reader with scope/author/done.
- `snapshot` as a true WYSIWYG channel — it just needs a markup mode (§1).
- Name-based `bookmark:` / `cc:` anchors surviving edits — the durable-handle primitive already exists; §5 just asks to mint them on insert.
