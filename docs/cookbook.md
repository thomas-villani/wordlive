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

From the CLI, `cursor read` is the same read — plus it resolves which
paragraph the cursor sits in, so you can pivot straight to an anchored edit:

```bash
$ wordlive cursor read
{"start": 142, "end": 142, "collapsed": true, "text": "", "paragraph": {"anchor_id": "para:7"}}
```

The `paragraph.anchor_id` is the bridge: read where the user is, then act on
`para:7` (or its `heading:N`) with the polite, cursor-preserving verbs instead
of writing back at the live caret.

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

The end of the document is the one position no content names, so it gets its
own helper — [`doc.append_paragraph(...)`](python-api.md#wordlive.Document)
adds a new final paragraph (no need to find the last one first):

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    with doc.edit("Append closing note"):
        doc.append_paragraph("Closing note added by automation.")
        # Optional style, and \r / \n to append several paragraphs at once:
        # doc.append_paragraph("Heading\rFirst line", style="Body Text")
```

Use [`doc.append(text)`](python-api.md#wordlive.Document) instead when you want
the text to continue the last paragraph inline rather than start a new one
(the direct, polite form of the old `doc.com.Content.InsertAfter(...)`).

Both also surface as an anchor — `doc.end` (id `end`) — so the end composes
with the usual verbs and the CLI's `--anchor-id`:

```python
doc.end.insert_paragraph_after("Closing note.")   # same as append_paragraph
doc.end.insert_image("logo.png", wrap="inline")   # drop an image at the end
```

```bash
$ wordlive append --text "Closing note added by automation."
$ wordlive insert --anchor-id end --text "Closing note."   # equivalent
```

The start of the document mirrors all of this:
[`doc.prepend_paragraph(...)`](python-api.md#wordlive.Document) /
[`doc.prepend(...)`](python-api.md#wordlive.Document), the `doc.start` anchor
(id `start`), and `wordlive prepend` — for a title or a "DRAFT" banner above
everything else.

### 3c. At the user's cursor (explicit, moves them)

The cursor is the deliberately *non-default* target. `doc.selection.write()` is
the first-class way to type at it; pair it with `allow_cursor_move()` so the
edit doesn't snap the cursor back afterwards:

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    with doc.edit("Insert at cursor") as scope:
        scope.allow_cursor_move()           # this edit is *allowed* to move them
        doc.selection.write("This text lands at the cursor.")
```

Without `allow_cursor_move()`, wordlive snaps the cursor back to where the
user had it — the typed text is still there, but the cursor jump confuses the
user. Always pair cursor-moving edits with `allow_cursor_move()`.

By default `write` replaces the current selection (like typing over
highlighted text); pass `replace=False` to insert at the selection start
without removing it. Either way the cursor is left after the inserted text.

The CLI mirrors this with the dedicated, intentionally-separate `cursor` group
— `cursor write` already opts into the cursor move for you:

```bash
$ wordlive cursor write --text "This text lands at the cursor."
{"ok": true, "replace": true}

# Insert without overwriting the user's selection:
$ wordlive cursor write --text "(draft) " --no-replace
```

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

`find` matches **exact** text (after cosmetic normalization). When the agent
only *approximately* remembers the wording — a paraphrase, a typo, a half-recalled
sentence — reach for `find-paragraph` instead; see
[recipe 15](#15-locate-a-paragraph-exact-find-vs-fuzzy-find-paragraph).

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
        model="<your-model>",
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

### Merged and split cells

Merged cells are a modelled, supported surface, not an edge to avoid.
`cell(r, c).merge(cell(r2, c2))` joins a rectangle into its upper-left cell;
`cell(r, c).split(rows=…, cols=…)` is the inverse. Either makes the table
**non-uniform** — `Table.is_uniform` flips to `False`, and `table:N:R:C` then
indexes *physical* cells (a merged row is short, so an index can shift or fall
off the row's end). Re-read with `budget.read()` (or `wordlive table read N`,
whose `uniform` flag reports this) after a merge to see the new shape.

```python
with doc.edit("Span the header"):
    budget.cell(1, 1).merge(budget.cell(1, 2))   # one header cell across two cols
    budget.add_column(["", "$0", "$0"])          # Columns.Add tolerates a merged table
```

`add_column` / `delete_column` mirror `add_row` / `delete_row`, and the whole
row/column anchors `table:N:row:R` / `table:N:col:C` style a strip in one call.
`delete_column` and column anchors raise `OpError` on a merged / mixed-width
table (Word has no per-column model there) — pointing you back at per-cell
`table:N:R:C` styling; rows are always contiguous, so they're unaffected.

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

## 12. Address and edit *any* paragraph, not just headings

`outline` only shows headings, so a document of plain prose looks unaddressable.
It isn't: every paragraph is a `para:N` anchor. Discover them with `paragraphs`
(or `outline --all`), then act on a body paragraph with the same verbs you'd use
on a heading.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        # Every paragraph, with offsets — headings AND body text AND list items.
        for p in doc.paragraphs.list():
            flag = f"H{p['level']}" if p["is_heading"] else "  "
            print(f"{flag} {p['anchor_id']:8} {p['text'][:50]!r}")

        with doc.edit("Tidy the opening"):
            # Rewrite the second paragraph in place (trailing ¶ preserved).
            doc.paragraphs[2].set_text("A clearer opening sentence.")

            # Drop a new paragraph *before* paragraph 2, styled as body text.
            doc.paragraphs[2].insert_paragraph_before(
                "Executive summary.", style="Body Text"
            )
    ```

    `doc.paragraphs[N]` returns a `Paragraph` anchor (`para:N`) that inherits
    every verb — `apply_style`, `format_paragraph`, `apply_list`, the insert
    pair. Because `para:N` and `heading:N` share an index space, a heading at
    `para:5` is also `heading:5`; use whichever reads better.

=== "CLI"

    ```bash
    # Discover every paragraph (these two are identical):
    wordlive paragraphs
    wordlive outline --all

    # Edit a body paragraph by its para:N id.
    wordlive replace --anchor-id para:2 --text "A clearer opening sentence."

    # Insert a new paragraph before / after any anchor.
    wordlive insert --anchor-id para:2 --text "Executive summary." --before --style "Body Text"
    wordlive insert --anchor-id heading:1 --text "Background follows." --after
    ```

### Inserting *inside* a paragraph at an offset

`insert` always makes a *new* paragraph. To splice text into the middle of an
existing one, target a **collapsed range** — `range:OFFSET-OFFSET` — and write
to it. The offsets come straight from `paragraphs` (or `find`):

```bash
# Paragraph 2 starts at offset 13; insert a marker 5 chars in (offset 18).
$ wordlive replace --anchor-id range:18-18 --text "[NOTE] "
```

Setting text on a zero-width range inserts without overwriting; a non-zero
`range:START-END` replaces that span. Range offsets are *live*, so compute and
use them in the same breath — an edit elsewhere shifts everything after it.

## 13. Act on whatever the user is pointing at

The hotkey workflow: the user clicks into (or selects) something, triggers your
script, and you decide whether to act *politely at an anchor* or *directly at
the cursor*. `cursor read` gives you both the raw position and the containing
`para:N`, so you can choose.

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active
    sel = doc.selection.info()              # {start, end, collapsed, text}

    # Map the caret to a stable anchor and edit *there* — cursor stays put.
    para = doc.paragraphs.at(sel["start"])
    if para is not None:
        with doc.edit("Annotate current paragraph"):
            para.insert_paragraph_after(f"(reviewed: {para.text[:30]}…)")
```

That's the polite path: read the cursor, but write at the anchor it resolves
to, leaving the user where they were. When the user genuinely wants text *at*
the caret — "insert my signature here" — reach for the explicit cursor write
from [recipe 3c](#3c-at-the-users-cursor-explicit-moves-them):

```bash
$ wordlive cursor read
{"start": 142, "end": 142, "collapsed": true, "text": "", "paragraph": {"anchor_id": "para:7"}}

# Polite: act on the resolved anchor instead of the caret.
$ wordlive insert --anchor-id para:7 --text "Reviewed by automation." --after

# Or explicit, when the caret is genuinely the target.
$ wordlive cursor write --text "— J. Doe"
```

The split is deliberate: anchors are addressable, stable, and visible to an LLM
as JSON; the cursor is none of those, so wordlive keeps it behind its own
clearly-labelled `cursor` surface rather than letting it leak into
`--anchor-id`.

## 14. Drop a figure into a document

`insert_image` works on any anchor and embeds the picture (it never links to a
path that could vanish). `wrap` is required, so layout intent is always
explicit; `"auto"` floats small images with Square wrap and large ones
top-and-bottom.

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    with doc.edit("Add diagram after Risks"):
        # From a file, letting the size heuristic pick the wrap.
        doc.heading("Risks").insert_image("diagram.png", wrap="auto")

        # Inline (in the text flow), with explicit size and alt text.
        doc.bookmarks["Logo"].insert_image(
            "logo.png", wrap="inline", width=96, alt_text="Company logo"
        )
```

An LLM usually holds image **data**, not a path — pass `bytes` or a base64
string and wordlive temp-files it, embeds it, and cleans up:

```python
import base64, wordlive as wl

png_b64 = "...base64 from a vision/diffusion model..."

with wl.attach() as word:
    doc = word.documents.active
    with doc.edit("Insert generated chart"):
        doc.heading("Results").insert_image(png_b64, wrap="square", width=240)
        # Equivalently: insert_image(base64.b64decode(png_b64), wrap="square")
```

From the CLI, use `--path` for files and `--base64` (or `--base64 -` from
stdin) for in-memory data:

```bash
$ wordlive insert-image --anchor-id heading:2 --path diagram.png --wrap auto
{"ok": true, "anchor_id": "heading:2", "anchor": {"kind": "heading", "name": "Risks"}, "wrap": "auto", "where": "after"}

$ base64 logo.png | wordlive insert-image --anchor-id bookmark:Logo --base64 - \
    --wrap inline --width 96 --alt-text "Company logo"
```

A missing file, malformed base64, or an unrecognised format raises
[`ImageSourceError`](errors.md#imagesourceerror) (exit code 1) before anything
is inserted — the batch never half-mutates the document.

## 15. Locate a paragraph: exact `find` vs fuzzy `find-paragraph`

Two locators, two jobs. **`find`** does exact substring matching (after
normalizing smart quotes / dashes / whitespace) and returns `range:START-END`
hits — the right tool when you know the literal text and want to *edit* it (it
feeds straight into `replace`; see [recipe 4](#4-fuzzy-find-replace-llm-friendly)).
**`find_paragraphs`** scores *every paragraph* for similarity to your query with
`difflib.SequenceMatcher` (over the same normalization) and returns ranked
`para:N` candidates — the right tool when the agent only **approximately**
remembers the wording and wants the paragraph it lives in.

A model holding a paraphrase would get **zero** hits from `find`, but
`find_paragraphs` still ranks the real paragraph first:

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        # The doc says: "The quick brown fox jumps over the lazy dog."
        # The agent only half-remembers it:
        hits = doc.find("the fast brown fox leaps over a lazy dog")   # exact → []
        ranked = doc.find_paragraphs("the fast brown fox leaps over a lazy dog")

    print(hits)                       # []  — no exact substring
    print(ranked[0]["anchor_id"],     # "para:12"
          round(ranked[0]["score"], 2))  # 0.86  — best fuzzy match
    ```

    `find_paragraphs` returns
    `[{anchor_id, index, score, text, level, is_heading}, …]` sorted by
    descending `score`, keeping only matches at or above `min_score` (default
    `0.6`), capped at `limit` (default `5`). Headings are included and flagged by
    `is_heading` / `level`, but everything is addressed by `para:N`.

=== "CLI"

    ```bash
    # Exact: nothing, because the wording drifted.
    $ wordlive find --text "the fast brown fox leaps over a lazy dog"
    []

    # Fuzzy: the real paragraph, ranked, with a score.
    $ wordlive find-paragraph --text "the fast brown fox leaps over a lazy dog"
    [{"anchor_id": "para:12", "index": 12, "score": 0.8605,
      "text": "The quick brown fox jumps over the lazy dog.",
      "level": 10, "is_heading": false}]
    ```

    Tune the breadth with `--limit N` and the strictness with `--min-score F`
    (0–1). An empty or whitespace-only query returns `[]`.

Then act on the winner by its `para:N` id — read it, edit it, or use it as a
`scope` for a precise `find_replace`:

```python
target = ranked[0]["anchor_id"]                       # "para:12"
with doc.edit("Fix the pangram"):
    doc.find_replace("jumps", "leaps", scope=doc.anchor_by_id(target))
```

Rule of thumb: **`find` to edit known text, `find_paragraphs` to locate
half-remembered text.** Pair them — `find_paragraphs` to home in on the
paragraph, then a scoped exact `find` / `find_replace` for the surgical change.

## 16. Insert a chart and dress it up

`insert_chart` embeds an Excel-backed chart at any anchor, then breaks the data
link so the chart ships static (no live workbook). The post-insert formatting
verbs — `add_trendline`, `format` (whole chart), `set_axis`, `add_error_bars`,
`format_series` — operate on the static chart and need no Excel. Charts are the
one feature that reaches a second Office app, so Excel must be installed:
`insert_chart` raises [`ExcelNotAvailableError`](errors.md) (CLI exit code `6`)
up front, leaving the document untouched.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        with doc.edit("Add revenue chart"):
            chart = doc.heading("Results").insert_chart(
                "scatter",
                [[1, 12], [2, 19], [3, 31], [4, 28]],   # [x, y] pairs
                title="Weekly revenue",
            )
            chart.add_trendline(kind="linear", display_equation=True)
            chart.add_error_bars(kind="percent", amount=5)
            chart.set_axis("y", title="$000s", minimum=0)
    ```

    `insert_chart` returns a `ChartAnchor` (`chart:N`); the formatting verbs
    chain (each returns `self`). `kind` is `"bar"`, `"pie"`, `"line"`, or
    `"scatter"`; bar/pie/line also accept a `{label: value}` mapping.

=== "CLI"

    ```bash
    wordlive insert-chart --anchor-id heading:2 --kind scatter \
        --data '[[1,12],[2,19],[3,31],[4,28]]' --title "Weekly revenue"

    wordlive add-trendline --anchor-id chart:1 --kind linear --display-equation
    wordlive add-error-bars --anchor-id chart:1 --kind percent --amount 5
    wordlive format-axis --anchor-id chart:1 --which y --title "\$000s" --minimum 0
    ```

Discover existing charts with `doc.charts` (or `wordlive charts`); each carries
its `chart:N` id. The series data isn't read back (the link is broken), so
chart reads report metadata only — `chart_type`, `title`, `chart_style`.

## 17. Float a pull-quote or stamp a watermark

Floating shapes (`shape:N`) sit in the drawing layer, not the text flow.
`insert_text_box` drops a pull-quote / call-out anchored to a paragraph and
hands back its `ShapeAnchor`, which restyles in place — `set_wrap`,
`set_position`, `set_size`, `set_crop` (pictures), `format`. A text watermark is
`doc.set_watermark(...)`, stamped into every section's header story behind the
body.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        with doc.edit("Add pull quote + DRAFT stamp"):
            quote = doc.heading("Summary").insert_text_box(
                "“The single biggest risk is schedule slip.”",
                width="2.5in", wrap="square", fill="#eef3ff", italic=True,
            )
            quote.set_position(left="center", relative_to="margin")
            quote.set_wrap("tight", side="largest")

            doc.set_watermark("DRAFT", layout="diagonal")
    ```

    A floating shape anchors to a *paragraph*, so `shape:N` renumbers as shapes
    come and go — re-list via `doc.shapes` (or just the text boxes via
    `doc.text_boxes`) rather than caching an id. `remove_watermark()` clears it.

=== "CLI"

    ```bash
    wordlive insert-text-box --anchor-id heading:1 \
        --text "“The single biggest risk is schedule slip.”" \
        --width 2.5in --fill "#eef3ff" --italic

    wordlive set-shape-position --anchor-id shape:1 --left center
    wordlive set-shape-wrap --anchor-id shape:1 --wrap tight --side largest

    wordlive set-watermark --text DRAFT --layout diagonal
    ```

## 18. Audit and autofix publishing defects

`lint` is a pure-read audit for publishing-quality defects — direct formatting
fighting the applied style, headings that may dangle at a page foot, multi-page
tables with no repeating header, numbered lists Word split into independent
runs. Each finding is severity-ranked and flags whether it's `fixable`.
`regularize` then applies the mechanical fixes in one atomic-undo step; the
default fixes are targeted and idempotent (a second pass is a no-op).

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        findings = doc.lint()                 # pure read — nothing changes
        fixable = [f for f in findings if f["fixable"]]
        print(f"{len(fixable)}/{len(findings)} findings are auto-fixable")

        plan = doc.regularize(dry_run=True)   # preview without writing
        report = doc.regularize()             # one Ctrl-Z reverts the whole pass
        print(f"applied {len(report['applied'])}, skipped {len(report['skipped'])}")
    ```

    Scope either with `within=doc.heading("Risks")`, and narrow the rules with
    `rules=["headings", "lists"]` or `rules={"exclude": [...]}`. Content-changing
    fixes are out of scope — `regularize` touches formatting / structure only.

=== "CLI"

    ```bash
    wordlive lint                              # severity-ranked findings
    wordlive lint --rule headings --within heading:3
    wordlive regularize --dry-run              # plan the fixes
    wordlive regularize                        # apply them (atomic-undo)
    ```

## 19. "What changed this session?"

Word emits no content-change event, so the reliable way to answer "what changed"
— or for an agent to verify its own edits landed without re-reading the whole
document — is to fingerprint, then diff. `doc.checkpoint()` returns an opaque,
serialisable token; `doc.changes_since(cp)` diffs it against the document *now*,
and `doc.diff(a, b)` diffs two stored checkpoints.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        cp = doc.checkpoint()                  # pure read — fingerprint now
        token = cp.to_json()                   # stash it (file, DB, the agent's state)

        # … the user (or the agent) edits the document …

        for change in doc.changes_since(token):
            # op ∈ replace | insert | delete | restyle | reformat
            print(change["op"], change.get("anchor_id"), change.get("text_after"))
    ```

    `include="text+style"` (default) surfaces restyles; `"text+format"` also
    catches pure direct-formatting edits. Alignment is by paragraph *content*,
    so a `para:N` that renumbered still aligns. An unchanged document returns
    `[]` via a fast-path hash.

=== "CLI"

    ```bash
    wordlive checkpoint --out before.json      # store the token
    # … edits happen …
    wordlive diff --since before.json          # changes vs the document now
    wordlive diff --from before.json --to after.json   # two stored checkpoints
    ```

## 20. Load a big document into context cheaply

Two reads sized for an agent's context window. `doc.read(budget=…)` (`read
digest`) is a token-budgeted, structure-aware digest of the **whole** document:
headings verbatim (each tagged with its anchor), tables as one-line stubs, body
text sampled to fit the budget, overflow elided to markers that still name the
`para:` range. `doc.to_markdown(within=…)` (`read markdown`) then serialises any
one region in full — so the loop is *skim cheap, drill precisely*.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        digest = doc.read(budget=4000)         # ~4 chars/token; whole doc, elided
        # … agent picks heading:7 as the section it needs in full …
        section = doc.to_markdown(within="heading:7")   # or any range:S-E from find
    ```

    `read` keeps every anchor addressable, so an agent can drill into any elided
    region with `to_markdown(within=…)`. Markdown export is lossy by design
    (underline, colours, merged cells don't survive); `to_html` keeps underline.

=== "CLI"

    ```bash
    wordlive --text read digest --budget 4000        # whole-doc digest
    wordlive --text read markdown --within heading:7 # one section, in full
    ```

## 21. Author a custom multi-level list template

`apply_list` applies a gallery default (`"bulleted"` / `"numbered"` /
`"outline"`). When you need a *specific* outline — say "1)" at level 1, lettered
sub-items at level 2 — author it with `apply_list_format`, a 1-based list of
per-level specs that defines each level's marker, numbering scheme, indentation,
and marker font.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        with doc.edit("Author custom outline"):
            doc.heading("Steps").apply_list_format([
                {"kind": "number", "format": "%1)", "style": "arabic", "bold": True},
                {"kind": "number", "format": "%2.", "style": "lower-letter"},
                {"kind": "bullet", "bullet": "–", "font": "Symbol"},
            ])
    ```

    Each spec's keys are optional except a bullet level's glyph; a number
    level's `format` uses `%N` to reference level N's number (`"%1.%2"`).
    `read_list_levels()` is the read mirror.

=== "CLI"

    ```bash
    wordlive list format --anchor-id heading:6 --levels '[
      {"kind":"number","format":"%1)","style":"arabic","bold":true},
      {"kind":"number","format":"%2.","style":"lower-letter"},
      {"kind":"bullet","bullet":"–","font":"Symbol"}
    ]'
    ```
