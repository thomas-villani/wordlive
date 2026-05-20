# Cookbook

End-to-end recipes for the workflows wordlive was actually built for.

## 1. Update a contract template

You have a Word template open with three bookmarks (`Address`, `Date`,
`Signatory`) and want to populate them from a Python dict. The user is
mid-review and shouldn't notice the script ran.

=== "Python"

    ```python
    import wordlive as wl

    values = {
        "Address":   "123 Main St, Anytown",
        "Date":      "2026-05-19",
        "Signatory": "Jane Doe",
    }

    with wl.attach() as word:
        doc = word.documents.active

        # Sanity-check that every target bookmark actually exists *before*
        # opening the edit scope. Cheaper than failing partway through.
        missing = [name for name in values if name not in doc.bookmarks]
        if missing:
            raise SystemExit(f"missing bookmarks: {missing}")

        with doc.edit("Populate contract template"):
            for name, text in values.items():
                doc.bookmarks[name].set_text(text)
    ```

    All three writes collapse to a single Ctrl-Z labelled
    *"Populate contract template"*. The user's cursor and scroll position are
    restored on exit.

=== "CLI (single shot)"

    ```bash
    cat > ops.json <<'JSON'
    {
      "label": "Populate contract template",
      "ops": [
        {"op": "write_bookmark", "name": "Address",   "text": "123 Main St, Anytown"},
        {"op": "write_bookmark", "name": "Date",      "text": "2026-05-19"},
        {"op": "write_bookmark", "name": "Signatory", "text": "Jane Doe"}
      ]
    }
    JSON

    wordlive exec --script ops.json
    ```

=== "CLI (one command per write)"

    Works, but you get three separate Ctrl-Z steps and three round-trips:

    ```bash
    wordlive write bookmark Address   --text "123 Main St, Anytown"
    wordlive write bookmark Date      --text "2026-05-19"
    wordlive write bookmark Signatory --text "Jane Doe"
    ```

    Prefer `exec` whenever the writes belong to a single user-visible intent.

## 2. Read what the user is looking at

A very common case: the user has selected (or just clicked into) a passage,
and your script needs to act on whatever they're focused on — without
moving them.

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active            # whichever doc is in focus
    sel = doc.selection.info()             # {"start": int, "end": int, "text": str}

    print(f"Active document: {doc.name}")
    print(f"Cursor at offset {sel['start']}–{sel['end']}")
    if sel["start"] == sel["end"]:
        print("(no selection — cursor is collapsed)")
    else:
        print(f"Selected text: {sel['text']!r}")
```

`info()` is read-only and never moves the user. It's the right primitive when
you want to:

- Capture what the user just highlighted, send it to an LLM as context, and
  feed the response back as an edit at a *named anchor* (not back at the
  selection — keep the cursor still).
- Detect "no current selection" (`start == end`) before deciding whether to
  show a "nothing to act on" message.
- Drive a hotkey-style workflow: user highlights a phrase, presses a key,
  your script reads the selection and reacts.

If you need offsets beyond `start` / `end` / `text` (e.g. page number, line
number), drop to `doc.selection.com` — that's the raw
[`Application.Selection`](concepts.md#the-com-escape-hatch) and has the full
Word object model under it.

The CLI does not currently surface live selection reads; for now this is a
Python-only flow.

### Variant: read a whole section by heading

When the user's request is "summarize the Risks section", you don't want the
selection — you want every paragraph under a heading. Use
`Heading.section_text()` (or `wordlive read section`):

```python
with wl.attach() as word:
    doc = word.documents.active
    section = doc.heading("Risks").section_text()
    # section is the body from after the Risks heading up to the next heading
    # at level ≤ Risks's level (or end of document).
```

```bash
$ wordlive --text read section "Risks"
The operational risks identified this quarter are …
```

`--text` mode prints just the body — perfect for piping into a prompt:

```bash
prompt=$(wordlive --text read section "Risks")
claude -p "Summarize these risks in two bullets: $prompt"
```

## 3. Add text to a document

Three patterns, picked by *where* you want the text to land. All three use
[`doc.edit()`](python-api.md#wordlive.Document) so the insert collapses into a
single Ctrl-Z.

### 3a. After a named anchor (recommended)

This is the polite default: target a stable anchor, don't touch the cursor.

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    with doc.edit("Append note to introduction"):
        # After a heading — gets its own paragraph.
        doc.heading("Introduction").insert_paragraph_after(
            "(Added 2026-05-19: see attached appendix.)"
        )

        # After a bookmark — inline, no new paragraph.
        doc.bookmarks["Address"].insert_after(" (verified)")

        # Before a content control — inline, on the left.
        doc.content_controls["Signatory"].insert_before("Dr. ")
```

Use `insert_paragraph_after` when you want a new paragraph (with optional
`style="Body Text"` etc.); use `insert_before` / `insert_after` for inline
inserts that don't break the surrounding paragraph.

!!! note
    `insert_before` and `insert_after` leave the bookmark's stored range
    *unchanged* — the new text lands outside the bookmark's span. Only
    `set_text` re-creates the bookmark to cover its new content. If you
    want the bookmark to grow with the appended text, use
    `bm.set_text(bm.text + " (verified)")` instead.

### 3b. Append to the end of the document

There's no high-level helper for "end of doc" (it isn't a named anchor) —
drop to `.com`:

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    with doc.edit("Append closing note"):
        # Content is the full document range; InsertAfter appends past it.
        doc.com.Content.InsertAfter("\n\nClosing note added by automation.\n")
```

Politeness still holds — `doc.edit()` snapshots and restores the user's
selection and scroll position even when the underlying mutation is raw COM.

### 3c. At the user's cursor (explicit, moves them)

For "type at the cursor" semantics, opt out of the selection restore:

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    with doc.edit("Insert at cursor") as scope:
        scope.allow_cursor_move()           # do NOT restore selection
        word.com.Selection.TypeText("This text lands at the cursor.")
```

Without `allow_cursor_move()`, wordlive would snap the cursor back to where
the user had it — collapsing the just-typed text would still be visible, but
the user would be confused by the cursor jump. Always pair cursor-moving
edits with `allow_cursor_move()`.

For "replace the user's current selection with X" (e.g. an LLM-rewrite of
highlighted text), `Selection.TypeText` already does the right thing: if a
range is selected it replaces it, otherwise it inserts at the caret.

## 4. Fuzzy find + replace (LLM-friendly)

The classic LLM editing flow: the model sees the document, decides "replace
*this* sentence with *that*", and emits a `(find, replace)` pair. Naïve
substring matching breaks the moment the model normalizes the source text
through its tokenizer — smart quotes get straightened, NBSPs become spaces,
em-dashes turn into hyphens. `wordlive.find_replace()` normalizes both sides
the same way so cosmetic drift doesn't blow up the match:

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    with doc.edit("Apply LLM-suggested rewrite"):
        # The LLM said: replace "Q1 2025" with "Q2 2025".
        # The doc actually contains "Q1 2025" with a NBSP between Q1 and 2025.
        applied = doc.find_replace("Q1 2025", "Q2 2025")

    print(f"replaced {len(applied)} occurrence(s)")
```

What's preserved automatically:

- **Character formatting.** Word's range-replace inherits the formatting of
  the first character of the matched span, so a bold "Q1 2025" becomes a bold
  "Q2 2025". You don't need to teach the LLM to write markdown.
- **The user's cursor.** `doc.edit()` snapshots and restores selection +
  scroll position, just like every other polite write.

### Disambiguating multiple matches

If `find` matches more than one occurrence and you didn't pass `all=True` or
`occurrence=N`, you get an [`AmbiguousMatchError`](errors.md#ambiguousmatcherror)
carrying every match's offsets. The CLI version returns the same payload
on stdout with exit code `5`:

```bash
$ wordlive replace --find "Q1" --text "Q2"
{"ok": false, "error": "ambiguous_match", "find": "Q1",
 "matches": [{"anchor_id": "range:412-414", "start": 412, "end": 414, "text": "Q1"},
             {"anchor_id": "range:887-889", "start": 887, "end": 889, "text": "Q1"}]}
$ echo $?
5
```

The agent's recovery is a fresh call with `--occurrence N` (or `--all`):

```bash
$ wordlive replace --find "Q1" --text "Q2" --occurrence 2
{"ok": true, "replacements": [{"anchor_id": "range:887-889", "start": 887, "end": 889, "text": "Q1"}]}
```

### Scoping the search

For "replace this phrase, but only inside the Risks section":

```python
with doc.edit("Targeted rewrite"):
    doc.find_replace(
        "needs review",
        "approved",
        scope=doc.heading("Risks"),
        all=True,
    )
```

When `scope` is a `Heading`, wordlive expands it to the heading's section
(the body up to the next same-or-higher heading) — so the replacement won't
accidentally touch identical phrasing in unrelated parts of the document.

CLI equivalent: `wordlive replace --find "..." --text "..." --in heading:N --all`.

### Read-only locate first

If the agent isn't sure whether its match will be unique, use `wordlive find`
to peek without writing:

```bash
$ wordlive find --text "the risk register" --in heading:3
[{"anchor_id": "range:412-429", "start": 412, "end": 429, "text": "the risk register"}]
```

This is the same matcher as `replace --find`, but read-only — useful as a
pre-flight check or to enumerate candidates for an `--occurrence` pick.

## 5. LLM tool-use loop

The CLI's JSON-in / JSON-out shape is designed to drop straight into a
tool-use loop. The pattern is:

1. **Discover** with `wordlive outline` — gives the model addressable anchors.
2. **Decide** — model picks anchors and new values.
3. **Apply** with `wordlive exec --script ops.json` — single round-trip, one
   Ctrl-Z.
4. **Branch on the exit code** — `2` means a stale anchor (re-fetch outline),
   `3` means Word is busy (retry).

### Tool schema

A minimal tool definition the agent sees:

```json
{
  "name": "wordlive_apply",
  "description": "Apply a batch of edits to the user's open Word document under one Ctrl-Z. All ops succeed or the failure point is reported.",
  "input_schema": {
    "type": "object",
    "required": ["label", "ops"],
    "properties": {
      "label": {"type": "string"},
      "ops":   {"type": "array", "items": {"type": "object"}}
    }
  }
}
```

The supported op shapes are documented on the [CLI page](cli.md#exec-script-opsjson).

### Driver loop

```python
import json, subprocess

def wordlive(*args: str) -> tuple[int, dict | list]:
    """Run a wordlive subcommand and return (exit_code, parsed_stdout)."""
    proc = subprocess.run(["wordlive", *args], capture_output=True, text=True)
    try:
        payload = json.loads(proc.stdout) if proc.stdout.strip() else {}
    except json.JSONDecodeError:
        payload = {"raw": proc.stdout}
    return proc.returncode, payload


def agent_turn(claude, outline):
    """One round trip: ask the model what to change, return its plan."""
    response = claude.messages.create(
        model="claude-opus-4-7",
        max_tokens=1024,
        tools=[WORDLIVE_APPLY_TOOL],
        messages=[
            {"role": "user", "content": [
                {"type": "text", "text":
                 "Here is the document outline. Update section 3 to reflect "
                 "the new risk register."},
                {"type": "text", "text": json.dumps(outline)},
            ]},
        ],
    )
    for block in response.content:
        if block.type == "tool_use" and block.name == "wordlive_apply":
            return block.input
    return None


def main():
    code, outline = wordlive("outline")
    if code == 4:
        raise SystemExit("Word is not running")

    for attempt in range(3):
        plan = agent_turn(claude, outline)
        if plan is None:
            return

        with open("ops.json", "w") as f:
            json.dump(plan, f)

        code, result = wordlive("exec", "--script", "ops.json")
        if code == 0:
            print(f"applied {result['ops_run']} ops")
            return
        if code == 2:                       # stale anchor — refresh and retry
            code, outline = wordlive("outline")
            continue
        if code == 3:                       # Word busy — back off
            time.sleep(2 ** attempt)
            continue
        raise SystemExit(f"wordlive failed (exit {code}): {result}")
```

The key property: every failure is *labelled* (anchor name, op index) so the
next iteration's prompt can include `result["failure"]` as feedback. The
model corrects itself instead of looping blindly.

## 6. Insert a section without disturbing the user

You want to add a new "Action items" paragraph after the *Risks* heading
without moving the user's cursor or scroll position — even if they're
typing on the same page.

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    # Snapshot the user's position so we can prove we didn't move them.
    before = wl._selection.snapshot(word)        # noqa: SLF001 — diagnostic only

    with doc.edit("Add action items"):
        doc.heading("Risks").insert_paragraph_after(
            "Action items: follow up with risk owners by Friday.",
            style="Body Text",
        )

    after = wl._selection.snapshot(word)
    assert before.start == after.start, "user's cursor moved!"
    assert before.vertical_percent == after.vertical_percent, "scroll moved!"
```

The `_selection.snapshot` call is private (and only used here for the
assertion). In real code you'd just trust the politeness contract — but the
assertion is useful proof when you're verifying the behaviour in a smoke
test.

If you *do* want to move the user — say, jump them to the freshly inserted
paragraph — opt out of restoration:

```python
with doc.edit("Add and jump") as scope:
    risks = doc.heading("Risks")
    risks.insert_paragraph_after("Action items: …")
    scope.allow_cursor_move()
    doc.go_to(risks)
```

## 7. Restyle and format a paragraph politely

You want to promote the *Risks* heading from H3 to H2 and tighten its
spacing, without touching the user's cursor or scrolling the view.

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    # 1. Sanity-check the style exists before we mutate anything.
    if "Heading 2" not in doc.styles:
        raise SystemExit("doc is missing the Heading 2 style")

    risks = doc.heading("Risks")
    with doc.edit("Restyle Risks heading"):
        risks.apply_style("Heading 2")
        risks.format_paragraph(space_before=12, space_after=4, alignment="left")
```

Both calls go through `doc.edit("…")`, so a single Ctrl-Z reverts the whole
change. `apply_style` raises [`StyleNotFoundError`](errors.md#stylenotfounderror)
(exit code `2`) if the style doesn't exist — discover real names with
`doc.styles.list()` or `wordlive style list`.

The same intent from the CLI, in one atomic batch:

```bash
$ wordlive exec --script - <<'JSON'
{
  "label": "Restyle Risks",
  "ops": [
    {"op": "apply_style",      "anchor_id": "heading:3", "name": "Heading 2"},
    {"op": "format_paragraph", "anchor_id": "heading:3",
      "space_before": 12, "space_after": 4, "alignment": "left"}
  ]
}
JSON
```

Indent and spacing values are in **points** — the same unit Word uses in its
paragraph dialog. If the anchor spans a partial paragraph (e.g., a bookmark
covering five words inside a longer paragraph), `format_paragraph` applies
to the *enclosing* paragraph, mirroring how Word's own UI behaves.

## 8. Read and edit a table

A cell is just another anchor (`table:N:R:C`), so the same polite, atomic-undo
patterns apply. Discover the grid first, then address cells by id.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        budget = doc.tables[1]            # by 1-based position
        # …or doc.tables["Budget"] by Title.

        # Read the whole grid as plain text.
        for row in budget.grid():
            print(row)

        with doc.edit("Update budget"):
            budget.cell(2, 2).set_text("$450")          # bump a figure
            budget.add_row(["Lodging", "$600"])         # append a row
            budget.cell(1, 1).apply_style("Heading 4")  # restyle a header cell
    ```

    Cell text is returned clean — Word's internal end-of-cell markers are
    stripped. The whole block reverts with one Ctrl-Z.

=== "CLI"

    ```bash
    # Discover cells (each carries its anchor_id).
    wordlive table read 1

    # Write a single cell by its anchor id.
    wordlive replace --anchor-id table:1:2:2 --text "$450"

    # Append / drop rows.
    wordlive table add-row --table 1 --values '["Lodging", "$600"]'
    wordlive table delete-row --table 1 --row 5
    ```

=== "CLI (one atomic batch)"

    ```bash
    wordlive exec --script - <<'JSON'
    {
      "label": "Update budget",
      "ops": [
        {"op": "set_cell", "table": 1, "row": 2, "col": 2, "text": "$450"},
        {"op": "add_row",  "table": 1, "values": ["Lodging", "$600"]},
        {"op": "apply_style", "anchor_id": "table:1:1:1", "name": "Heading 4"}
      ]
    }
    JSON
    ```

!!! note
    Cell addressing assumes a rectangular grid. Tables with merged or split
    cells follow Word's own `Table.Cell(row, col)` indexing and may raise
    inside a merged region.

## 9. Multi-document workflows

When several documents are open, `--doc NAME` picks the target:

```bash
wordlive --doc Draft.docx outline
wordlive --doc Draft.docx write bookmark Address --text "456 Elm St"
```

In Python, `word.documents` is iterable and `word.documents[name]` does the
lookup:

```python
with wl.attach() as word:
    for doc in word.documents:
        if "draft" in doc.name.lower():
            with doc.edit("Touch all drafts"):
                doc.bookmarks["LastReviewed"].set_text("2026-05-19")
```

Document not found raises [`DocumentNotFoundError`](errors.md), which the CLI
maps to exit code `1`.

## 10. Suggest, don't overwrite: comments + tracked changes

The most agent-shaped edits are the *non-destructive* ones. Instead of
rewriting a passage, an agent can flag it with a comment, or make its edits
*visibly* as tracked changes the human accepts or rejects. Both leave the user
in control.

### Comment on what `find` located

A `find()` hit returns a `range:START-END` id, which resolves to a
[`RangeAnchor`](python-api.md#wordlive.RangeAnchor) — so you can attach a
comment to exactly the span you matched, without changing a character of it:

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    hits = doc.find("as soon as possible")
    if hits:
        target = doc.anchor_by_id(hits[0]["anchor_id"])   # a RangeAnchor
        with doc.edit("Flag vague deadline"):
            doc.comments.add(
                target,
                "Can we commit to a concrete date here?",
                author="ReviewBot",
            )
```

```bash
# CLI: discover the span, then comment on it.
$ wordlive find --text "as soon as possible"
[{"anchor_id": "range:512-531", "start": 512, "end": 531, "text": "as soon as possible"}]

$ wordlive comment add --anchor-id range:512-531 \
      --text "Can we commit to a concrete date here?" --author ReviewBot
{"ok": true, "anchor_id": "range:512-531", "comment": {"index": 1, "author": "ReviewBot"}}
```

List, resolve, and delete comments by their 1-based index:

```bash
$ wordlive comment list
$ wordlive comment resolve --index 1
$ wordlive comment delete --index 1
```

### Make edits visible as tracked changes

When you *do* want to change the text but let the human vet it, wrap the edit
in [`doc.tracked_changes()`](python-api.md#wordlive.Document). Track Changes is
turned on for the scope and restored to its prior state on exit:

```python
with wl.attach() as word:
    doc = word.documents.active

    with doc.tracked_changes(), doc.edit("Suggest plainer wording"):
        doc.find_replace("utilise", "use", all=True)
```

Every replacement lands as a revision in Word's review pane — one Ctrl-Z
removes the batch, or the user accepts/rejects each suggestion. From the CLI,
set `"tracked": true` on an `exec` script so the whole batch is recorded as
tracked changes and the prior setting is restored afterwards:

```bash
$ wordlive exec --script - <<'JSON'
{
  "label": "Suggest plainer wording",
  "tracked": true,
  "ops": [
    {"op": "find_replace", "find": "utilise", "text": "use", "all": true}
  ]
}
JSON
```

The standalone `wordlive track on` / `track off` toggle is *persistent* —
useful when a human will keep editing in tracked mode — but it doesn't
auto-restore, so prefer the scoped forms above for one-shot agent edits.

## 11. Number a procedure and stamp the header/footer

Template-generation work: take the paragraphs under a *Steps* heading, turn
them into a numbered list, and brand the page with a header and footer — all in
one atomic-undo batch. List verbs and header/footer writes are just anchor
operations, so they compose with everything else.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        steps = doc.heading("Steps")          # body under the heading
        with doc.edit("Number the procedure"):
            steps.apply_list("numbered")      # 1., 2., 3., …

            sec = doc.sections[1]
            sec.header().set_text("ACME Corporation — Internal")
            sec.footer().set_text("Confidential — do not distribute")
    ```

    `apply_list` accepts `"bulleted"`, `"numbered"`, or `"outline"`. To pick up
    numbering from a list just above, pass `continue_previous=True`; to force a
    fresh count on an existing list, call `steps.restart_numbering()`. Read the
    current state with `steps.list_info()` (`{type, level, number, string}`).

=== "CLI"

    ```bash
    # Discover lists already in the document (each carries a range anchor id).
    wordlive list show

    # Number a heading's paragraphs, then restart at 1 if needed.
    wordlive list apply --anchor-id heading:6 --type numbered
    wordlive list restart --anchor-id heading:6

    # Headers/footers by section (default section 1, which=primary).
    wordlive header write --section 1 --text "ACME Corporation — Internal"
    wordlive footer write --section 1 --text "Confidential — do not distribute"
    ```

=== "CLI (one atomic batch)"

    ```bash
    wordlive exec --script - <<'JSON'
    {
      "label": "Number the procedure",
      "ops": [
        {"op": "apply_list",   "anchor_id": "heading:6", "type": "numbered"},
        {"op": "write_header", "section": 1, "text": "ACME Corporation — Internal"},
        {"op": "write_footer", "section": 1, "text": "Confidential — do not distribute"}
      ]
    }
    JSON
    ```

A header/footer is just a range (`header:S:WHICH` / `footer:S:WHICH`, WHICH ∈
`primary`/`first`/`even`), so the same id also works with `replace`,
`style apply`, and `format-paragraph` when you need more than plain text.
