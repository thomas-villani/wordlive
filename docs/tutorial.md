# Tutorial: a guided editing session

This tutorial walks **one document, end to end**. You'll attach to a live Word
instance, inspect it, read a section, make a polite edit, batch several edits
into a single undo, suggest changes instead of overwriting them, and finally
verify exactly what you changed — each step building on the last.

It's deliberately different from the other pages:

- [Getting started](getting-started.md) is the 60-second quickstart — the
  *smallest* possible read and write, each standing alone.
- The [Cookbook](cookbook.md) is random-access how-to — "I have problem X, here's
  the recipe."
- **This tutorial is a single continuous session.** By the end you'll have *felt*
  the four ideas that shape every wordlive API — politeness, semantic anchors,
  atomic undo, and structured I/O — because you used them in order, not because
  you read a list. Budget about 15 minutes.

Every step is reversible with Ctrl-Z, and nothing moves your cursor unless it
says so.

## Before you start

You need Windows, Microsoft Word, and wordlive installed — see
[Getting started → Install](getting-started.md#install) if you haven't yet.

### Get the sample document

The tutorial drives a short status report. Get it whichever way suits you:

=== "Download it"

    Grab
    [`quarterly-report.docx`](https://github.com/thomas-villani/wordlive/blob/main/examples/sample/quarterly-report.docx)
    from `examples/sample/` and open it in Word.

=== "Build it"

    Regenerate it from the committed script (so the binary stays reviewable):

    ```bash
    uv run --with python-docx python examples/sample/build_quarterly_report.py
    ```

=== "Make your own"

    Any document works if it has four Heading-1 sections named **Introduction**,
    **Risks**, **Budget**, and **Next Steps**, with a three-column table under
    *Budget*. The IDs printed below will differ, but every step still applies.

**Open the sample in Word now.** Everything below attaches to that
already-running instance — wordlive never launches Word behind your back when you
use `attach()`.

## Step 1 — Attach and see the structure

Start by asking the document what's in it. `outline()` returns every heading,
each tagged with the **anchor id** wordlive (and an LLM) uses to address it.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active
        for entry in doc.outline():
            indent = "  " * (entry["level"] - 1)
            print(f"{indent}{entry['text']}  [{entry['anchor_id']}]")
    ```

=== "CLI"

    ```bash
    wordlive outline
    ```

Either way you get the four sections:

```text
Introduction  [heading:3]
Risks  [heading:6]
Budget  [heading:10]
Next Steps  [heading:28]
```

!!! note "Why `heading:3`, not `heading:1`?"
    Heading ids share one index space with **every** paragraph: the title is
    `para:1`, the subtitle `para:2`, so the *Introduction* heading is `para:3` —
    and therefore `heading:3`. The numbers are paragraph positions, not "the Nth
    heading." That's the whole addressing scheme: anchor ids are stable, visible
    handles you pass back in. See [Anchor IDs](concepts.md#anchor-ids).

From here on, address sections by name (`doc.heading("Risks")`) or by the id you
just saw (`heading:6`) — they're interchangeable.

## Step 2 — Take a checkpoint

Before touching anything, fingerprint the document. This is a **pure read** — it
changes nothing — but it lets you prove, at the end, exactly what your session
did. Stash the token somewhere (a variable, a file, an agent's state).

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active
        cp = doc.checkpoint()          # pure read — fingerprint "now"
        token = cp.to_json()           # serialisable; keep it for Step 7
    ```

=== "CLI"

    ```bash
    wordlive checkpoint --out before.json
    ```

We'll come back to this in [Step 7](#step-7-verify-what-you-changed).

## Step 3 — Read one section

You rarely want the whole document — you want *one part*. Address the *Risks*
heading and pull its body (everything down to the next same-or-higher heading).

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active
        print(doc.heading("Risks").section_text())
    ```

=== "CLI"

    ```bash
    wordlive --text read section "Risks"
    ```

```text
The single biggest risk this quarter is schedule slip on the integration milestone.
The vendor has promised the updated components as soon as possible.
Mitigation owners are assigned but not yet confirmed.
```

That's the *skim cheap, drill precisely* loop: `outline` to find the anchor, then
read just that anchor. For a whole-document, token-budgeted digest instead, see
[Cookbook §20](cookbook.md#20-load-a-big-document-into-context-cheaply).

## Step 4 — Your first polite edit

Now change something — and watch wordlive keep your place. Add a scope note right
after the *Introduction* heading.

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active

        before = doc.selection.info()          # where the user's cursor is
        with doc.edit("Add scope note to introduction"):
            doc.heading("Introduction").insert_paragraph_after(
                "Scope: integration milestone and Q2 budget.",
                style="Body Text",
            )
        after = doc.selection.info()
        assert before["start"] == after["start"]   # cursor never moved
    ```

=== "CLI"

    ```bash
    wordlive insert --anchor-id heading:3 \
        --text "Scope: integration milestone and Q2 budget." \
        --after --style "Body Text"
    ```

Two things happened that the code doesn't spell out:

- **`doc.edit("…")` opened a Word `UndoRecord`.** One Ctrl-Z removes the whole
  edit, labelled *"Add scope note to introduction"* in Word's undo dropdown.
- **Your cursor and scroll position were snapshotted and restored.** The
  `assert` proves it: `info()` reports the same start offset before and after.
  The script edited the document without stealing your place in it — that's the
  [politeness contract](concepts.md#politeness-model), and it's the default for
  every write.

## Step 5 — Batch several edits under one Ctrl-Z

Real edits come in groups. Revise the budget: correct a figure, add a contingency
row, and stamp a generated-on line — as **one** user-visible intent, so one
Ctrl-Z reverts all three.

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active

        budget = doc.tables[1]
        with doc.edit("Revise Q2 budget"):
            budget.cell(4, 3).set_text("$5,000")          # fix the Travel figure
            budget.add_row(["Contingency", "$0", "$3,000"])
            doc.append_paragraph("Report generated for Q2 review.")
    ```

=== "CLI"

    ```bash
    wordlive exec --script - <<'JSON'
    {
      "label": "Revise Q2 budget",
      "ops": [
        {"op": "set_cell", "table": 1, "row": 4, "col": 3, "text": "$5,000"},
        {"op": "add_row", "table": 1, "values": ["Contingency", "$0", "$3,000"]},
        {"op": "append_paragraph", "text": "Report generated for Q2 review."}
      ]
    }
    JSON
    ```

Three mutations, **one** undo step. That's the third invariant —
[atomic undo](concepts.md#editscope-and-atomic-undo): a batch maps to a single intent, so the
user reverts it the way they think about it, not op by op. The `exec` form is the
same batch the [MCP](mcp.md) and LLM tool-use loops send.

## Step 6 — Suggest, don't overwrite

The most agent-shaped edits don't rewrite the user's text — they *propose*.
Flag a vague deadline with a comment, and soften some wording as a **tracked
change** the human can accept or reject. Neither destroys anything.

First locate the phrase to comment on (read-only), then attach the comment to
exactly that span:

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active

        hit = doc.find("as soon as possible")[0]        # {'anchor_id': 'range:413-432', …}
        with doc.edit("Flag vague deadline"):
            doc.comments.add(
                doc.anchor_by_id(hit["anchor_id"]),
                "Can we commit to a concrete date?",
                author="ReviewBot",
            )

        # And soften wording *visibly*, as a revision:
        with doc.tracked_changes(), doc.edit("Plainer wording"):
            doc.find_replace("utilise", "use", all=True)
    ```

=== "CLI"

    ```bash
    # Locate the span (read-only), then comment on it.
    wordlive find --text "as soon as possible"
    # → [{"anchor_id": "range:413-432", "start": 413, "end": 432, "text": "as soon as possible"}]

    wordlive comment add --anchor-id range:413-432 \
        --text "Can we commit to a concrete date?" --author ReviewBot

    # Soften wording as a tracked change (prior Track-Changes state is restored).
    wordlive exec --script - <<'JSON'
    {
      "label": "Plainer wording",
      "tracked": true,
      "ops": [{"op": "find_replace", "find": "utilise", "text": "use", "all": true}]
    }
    JSON
    ```

The comment hangs off the exact `range:` span `find` returned — no character of
the document changed. The `utilise → use` swap lands in Word's review pane as a
revision: the human accepts or rejects it. `doc.tracked_changes()` turns Track
Changes on for just that scope and restores the prior setting on exit.

## Step 7 — Verify what you changed

Word emits no "content changed" event, so the reliable way to answer *"what did I
just do?"* — and for an agent to confirm its edits actually landed — is to diff
against the checkpoint from [Step 2](#step-2-take-a-checkpoint).

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active
        for change in doc.changes_since(token):     # token from Step 2
            print(change["op"], change.get("anchor_id"),
                  repr((change.get("text_after") or "")[:48]))
    ```

=== "CLI"

    ```bash
    wordlive diff --since before.json
    ```

```text
insert   para:4   'Scope: integration milestone and Q2 budget.'
replace  para:6   'We use a weekly status cadence to keep stakehol…'
replace  para:27  '$5,000'
insert   para:37  'Report generated for Q2 review.'
…                 (table-row inserts omitted)
```

Every change is **structured data**, not a string you scrape — `op`,
`anchor_id`, the new text. That's the fourth invariant,
[structured I/O](design.md#design-principles): reads come back as objects with
deterministic shapes, so a script (or an LLM) can branch on them. Alignment is by
paragraph *content*, so a `para:N` that renumbered when you inserted text still
lines up. (Comments and tracked changes are tracked separately — this diff covers
content edits.)

## What you just learned

You drove a real document through a complete session — and in doing so used each
of wordlive's four invariants exactly once:

| Step | What you did | The idea behind it |
| --- | --- | --- |
| 1, 3 | Addressed sections by `heading:N` / name | [Semantic anchors & anchor IDs](concepts.md#semantic-anchors-over-selection) |
| 4 | Edited without moving the cursor | [Politeness](concepts.md#politeness-model) |
| 5 | Three writes, one Ctrl-Z | [Atomic undo](concepts.md#editscope-and-atomic-undo) |
| 2, 7 | Checkpointed and diffed as objects | [Structured I/O](design.md#design-principles) |
| 6 | Suggested via comment + tracked change | Politeness, taken to non-destructive edits |

## Where to next

- [Advanced](advanced.md) — the sequel session: the power features (digest reads,
  durable pins, charts, vision snapshots, house-style linting) on this same document.
- [Concepts](concepts.md) — the same four ideas, explained rather than walked.
- [Cookbook](cookbook.md) — 22 random-access recipes, including the full LLM
  tool-use loop ([§5](cookbook.md#5-llm-tool-use-loop)).
- [Examples](examples.md) — runnable Python and PowerShell scripts.
- [Python API](python-api.md) · [CLI](cli.md) — the complete reference for every
  verb you used above.
- **Wiring up an agent?** [Agents & LLM tools](agents.md) connects it (skills,
  MCP, the one-drop `.mcpb`); [Agent patterns](agent-patterns.md) is how to drive
  it well.
