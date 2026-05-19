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

## 4. LLM tool-use loop

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

## 5. Insert a section without disturbing the user

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

## 6. Multi-document workflows

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
