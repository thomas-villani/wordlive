# Agent patterns

wordlive was built to be driven by an LLM, not just a human — that's the whole
premise. [Agents & LLM tools](agents.md) covers *connecting* an agent (skills,
MCP, the one-click `.mcpb`). **This page is the other half: how to drive it well
once it's wired up.** Every pattern below is distilled from the design invariants
and links down to the [Cookbook](cookbook.md) recipe that demonstrates it end to
end.

The patterns compose into one loop — **discover → decide → apply → verify →
recover** — and each addresses a specific way naïve automation goes wrong.

## 1. Skim cheap, drill precisely

Don't feed the whole document into the model. Start with a **structure-aware,
token-budgeted digest**, let the model pick the region it needs, then pull *just
that region* in full.

=== "Python"

    ```python
    digest = doc.read(budget=4000)              # whole doc, elided, every anchor kept
    # … model reads the digest, asks for heading:7 …
    section = doc.to_markdown(within="heading:7")
    ```

=== "MCP"

    ```text
    word_read command=digest budget=4000
    word_read command=to_markdown within=heading:7
    ```

Every elided span in the digest still names its `para:` range, so the model can
always drill deeper. This is the single biggest lever on context cost. See
[Cookbook §20](cookbook.md#20-load-a-big-document-into-context-cheaply).

## 2. Address anchors, never the cursor

The live `Selection` is invisible state and shifts under the user's hands — it's
the wrong target for a model. Operate on **stable anchor ids** (`heading:3`,
`bookmark:Address`, `table:1:2:2`, `range:412-429`) that round-trip through JSON.
When you must act on where the user is, read the caret and **map it to the anchor
it sits in**, then edit *there*:

=== "Python"

    ```python
    sel = doc.selection.info()                  # read-only; doesn't move them
    para = doc.paragraphs.at(sel["start"])      # → a para:N anchor
    with doc.edit("Annotate current paragraph"):
        para.insert_paragraph_after("…")        # edit the anchor, cursor stays put
    ```

=== "MCP"

    ```text
    word_read  command=status            # or read_text / find to locate
    word_write command=insert anchor_id=para:7 text="…"
    ```

See [Cookbook §13](cookbook.md#13-act-on-whatever-the-user-is-pointing-at). The
deliberate exception — "type at my cursor" — is the separately-labelled cursor
surface ([§3c](cookbook.md#3c-at-the-users-cursor-explicit-moves-them)).

## 3. Positional ids renumber — re-read or pin

`heading:N` / `para:N` are paragraph *indices*: an insert earlier in the document
shifts every id after it. Two disciplines keep an agent correct:

- **Single pass:** re-read `outline` / `paragraphs` after any structural edit
  before reusing an id. A stale positional id raises with a recovery hint.
- **Multi-pass:** [**pin**](advanced.md#step-2-pin-what-youre-about-to-move) the
  content first. `word_write command=pin` mints a durable `pin:CODE` Word keeps
  attached across edits — so the agent addresses the same block through several
  passes without re-reading. In a batch, `bind: "slug"` pins a creating op's new
  content inline.

Name-based anchors (`bookmark:NAME`, `cc:NAME`, `pin:CODE`) never renumber —
prefer them for anything the agent will revisit.

## 4. One intent = one batch = one undo

When several edits form a single user-visible intent, send them as **one batch**
(`exec` / `word_exec`), not N round-trips. The whole batch is a single Ctrl-Z the
user reverts the way they *think* about the change, it's one COM round-trip, and
it either fully applies or reports the exact failing op and rolls the prefix back.

=== "CLI"

    ```bash
    wordlive exec --script - <<'JSON'
    {"label": "Revise budget", "ops": [
      {"op": "set_cell", "table": 1, "row": 4, "col": 3, "text": "$5,000"},
      {"op": "add_row",  "table": 1, "values": ["Contingency", "$0", "$3,000"]},
      {"op": "append_paragraph", "text": "Report generated for Q2 review."}
    ]}
    JSON
    ```

=== "MCP"

    ```text
    word_exec ops=[ {op:set_cell,…}, {op:add_row,…}, {op:append_paragraph,…} ]
    ```

A batch can also **create then target without a round-trip**: any op field of the
form `$ops[N].field` is replaced with an earlier op's output before it runs (e.g.
`create_table` at op 0, then `set_cell` with `"table": "$ops[0].table"`). See
[Cookbook §5](cookbook.md#5-llm-tool-use-loop) and
[MCP → Batches](mcp.md#batches).

## 5. Suggest, don't overwrite

The most agent-shaped edits don't rewrite the user's text — they **propose**.
Flag a span with a comment (changes not a character), or make edits *visibly* as
tracked changes the human accepts or rejects. Both leave the user in control.

=== "Python"

    ```python
    hit = doc.find("as soon as possible")[0]
    with doc.edit("Flag vague deadline"):
        doc.comments.add(doc.anchor_by_id(hit["anchor_id"]),
                         "Commit to a concrete date?", author="ReviewBot")

    with doc.tracked_changes(), doc.edit("Plainer wording"):
        doc.find_replace("utilise", "use", all=True)
    ```

=== "MCP"

    ```text
    word_write command=comment  action=add anchor_id=range:413-432 text="…"
    word_exec  tracked=true ops=[ {op:find_replace, find:"utilise", text:"use", all:true} ]
    ```

`doc.tracked_changes()` (CLI/MCP `tracked: true`) turns Track Changes on for just
that scope and restores the prior state on exit. See
[Cookbook §10](cookbook.md#10-suggest-dont-overwrite-comments-tracked-changes).

## 6. Verify with checkpoint → diff

Word emits **no content-change event**, so an agent can't assume its edits landed
— and can't cheaply answer "what changed?" The reliable way is to fingerprint
before, then diff: `checkpoint()` returns an opaque, serialisable token;
`changes_since(token)` returns a **structured** change list aligned by paragraph
*content* (so a `para:N` that renumbered still lines up).

=== "Python"

    ```python
    token = doc.checkpoint().to_json()          # pure read — stash it
    # … the agent's edits …
    for c in doc.changes_since(token):
        print(c["op"], c.get("anchor_id"), c.get("text_after"))
    ```

=== "MCP"

    ```text
    word_read command=checkpoint          # → token; store it
    word_read command=diff checkpoint=<token>
    ```

This closes the agent's own loop: apply a batch, diff to confirm the intended ops
are present, and feed any mismatch back into the next turn. See
[Cookbook §19](cookbook.md#19-what-changed-this-session).

## 7. Self-correct on typed failures

Every failure is **typed** and carries recovery context — a stable `code` /
exit code plus the specific anchor or op that failed — so the agent branches on
the failure mode instead of looping blindly:

| Failure | `code` / exit | The agent's move |
| --- | --- | --- |
| stale/missing anchor, or `find` matched zero | `anchor_not_found` / `2` | re-read `outline`, retry |
| `find` matched several | `ambiguous_match` / `5` | re-issue with `occurrence`/`all` |
| Word busy (modal dialog) | `word_busy` / `3` | **back off and retry** |
| style not defined | `style_not_found` / `2` | read `styles` first |
| Word not running | `word_not_running` / `4` | stop until Word is up |

Because the payload names *what* failed (the anchor id, the op index, the match
list), the next turn's prompt can include it as feedback and the model fixes
itself. See [Cookbook §5](cookbook.md#5-llm-tool-use-loop) and the full
[error taxonomy](errors.md).

## 8. When it's visual, let the model *see* it

Structured reads answer "what does it say"; they can't answer "does the layout
look right." For anything visual — a heading stranded at a page foot, a chart's
placement, "does this styling land" — render the page and hand the model the
**image**. `word_snapshot` returns native image content; `max_dim` caps the
per-page token cost for a cheap whole-document glance; `markup="all"` shows
tracked changes and comments.

=== "MCP"

    ```text
    word_snapshot max_dim=1000                 # whole doc, cheap, model sees it
    word_snapshot page=2 markup=all            # one page, revision marks visible
    word_read command=read_image anchor_id=image:1   # SEE one embedded picture
    ```

=== "CLI"

    ```bash
    wordlive snapshot --max-dim 1000 --out doc.png
    wordlive snapshot --anchor-id heading:7 --out section.png
    ```

Pair it with the edit loop: **edit → snapshot → look → adjust.** See
[Advanced §4](advanced.md#step-4-see-what-you-just-did) and
[Cookbook §22](cookbook.md#22-let-a-vision-model-see-the-page).

## 9. Locate half-remembered text with fuzzy find

A model rarely quotes the document verbatim — it paraphrases, straightens quotes,
misremembers a word. Exact `find` returns nothing on drifted text; `find_paragraphs`
**scores every paragraph** for similarity and ranks the real one first. Use it to
home in, then a scoped exact `find_replace` for the surgical change.

=== "Python"

    ```python
    ranked = doc.find_paragraphs("the fast brown fox leaps over a lazy dog")
    target = ranked[0]["anchor_id"]             # "para:12", score 0.86
    with doc.edit("Fix the pangram"):
        doc.find_replace("jumps", "leaps", scope=doc.anchor_by_id(target))
    ```

=== "MCP"

    ```text
    word_read command=find_paragraphs text="the fast brown fox …" limit=5
    ```

See [Cookbook §15](cookbook.md#15-locate-a-paragraph-exact-find-vs-fuzzy-find-paragraph).

## Putting it together

A robust turn threads these into one loop:

1. **Discover** — `digest` (§1), drill with `to_markdown` where needed.
2. **Stabilise** — `pin` the blocks you'll revisit (§3); prefer named anchors.
3. **Decide** — the model picks anchors and values; fuzzy-locate if unsure (§9).
4. **Apply** — one `exec` batch per intent (§4); suggest rather than overwrite
   where appropriate (§5).
5. **Verify** — `diff` against the pre-edit checkpoint (§6); `snapshot` if the
   change was visual (§8).
6. **Recover** — branch on the typed failure and retry with feedback (§7).

## Where to next

- [Agents & LLM tools](agents.md) — connect the agent (skills, MCP, `.mcpb`).
- [Advanced session](advanced.md) — the same power features, walked end to end.
- [Cookbook §5](cookbook.md#5-llm-tool-use-loop) — a complete tool-use driver loop.
- [MCP server](mcp.md) — the four `word_*` tools and the full op vocabulary.
