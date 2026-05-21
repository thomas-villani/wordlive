# Getting started

This page takes you from zero to a first polite edit, both from Python and
from the CLI.

## Prerequisites

- **Windows.** wordlive talks to Word over COM (`pywin32`); there is no
  cross-platform path.
- **Microsoft Word**, installed and running. Anything from Word 2010 onward
  should work; older Word versions silently lose the atomic-undo feature but
  everything else still works.
- **Python 3.13+**.

## Install

```bash
pip install wordlive
```

`pywin32` is pulled in automatically on Windows. Click is the only other
runtime dependency.

## Hello, document

Open a Word document, then run:

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active
    print(doc.name)

    for entry in doc.outline():
        indent = "  " * (entry["level"] - 1)
        print(f"{indent}{entry['text']}  [{entry['anchor_id']}]")
```

`attach()` connects to the *already-running* Word instance — it won't launch
one. If Word isn't running you get a [`WordNotRunningError`](errors.md). Use
[`wl.connect()`](python-api.md#wordlive.connect) when you'd rather launch
Word if it isn't already up.

Every heading in the outline carries an `anchor_id` like `heading:3` — those
strings are how the CLI and LLM tool-use loops address ranges. See
[Anchor IDs](concepts.md#anchor-ids) for the scheme.

## Your first polite edit

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    with doc.edit("Update address block"):
        doc.bookmarks["Address"].set_text("123 Main St")
        doc.content_controls["Signatory"].set_text("Jane Doe")
        doc.heading("Introduction").insert_paragraph_after(
            "New context paragraph."
        )
```

A few things are happening that aren't obvious from the code:

- `doc.edit("…")` opens a Word `UndoRecord`. All three mutations are bundled
  into a single Ctrl-Z.
- Before the block runs, the user's cursor position and scroll offset are
  snapshotted. On exit they're restored — your script does not steal the
  user's place in the document. See [Politeness](concepts.md#politeness-model).
- If any bookmark / content control / heading is missing, you get a typed
  [`AnchorNotFoundError`](errors.md), not a raw COM error.

## Same task from the CLI

The CLI is intentionally thin over the Python API. Same atomic-undo, same
politeness:

```bash
# What's open?
wordlive status

# What's in the active doc?
wordlive outline          # headings (heading:N)
wordlive paragraphs       # every paragraph (para:N) with offsets

# Mutate.
wordlive write bookmark Address --text "123 Main St"
wordlive write cc Signatory --text "Jane Doe"
wordlive insert --anchor-id heading:1 --text "New context paragraph."

# Or batch all three under a single Ctrl-Z:
wordlive exec --script ops.json
```

Every command emits one JSON object on stdout (`--text` if you'd rather read
it) and uses deterministic exit codes:

| Exit | Meaning                |
| ---- | ---------------------- |
| `0`  | ok                     |
| `2`  | anchor or style not found |
| `3`  | Word busy / modal      |
| `4`  | Word not running       |
| `5`  | ambiguous match (`replace --find` matched >1 occurrence) |
| `1`  | other error            |

Full reference: [CLI](cli.md).

## Two more everyday flows

### Read what the user is looking at

Most useful when you want to feed the user's current focus to an LLM, or
trigger something off a hotkey-driven selection:

```python
with wl.attach() as word:
    doc = word.documents.active        # whichever doc is focused
    sel = doc.selection.info()         # {"start": int, "end": int, "text": str}

    if sel["start"] != sel["end"]:
        print(f"User has selected: {sel['text']!r}")
    else:
        print(f"Cursor sits at offset {sel['start']} (nothing selected).")
```

`info()` never moves the user. See
[Cookbook §2](cookbook.md#2-read-what-the-user-is-looking-at) for the full
pattern.

### Add a new paragraph to the document

The polite default: anchor to an existing heading and insert *after* it.

```python
with wl.attach() as word:
    doc = word.documents.active
    with doc.edit("Append note"):
        doc.heading("Introduction").insert_paragraph_after(
            "New note added by automation."
        )
```

For "append to the very end of the document" or "type at the user's cursor",
see [Cookbook §3](cookbook.md#3-add-text-to-a-document).

## Where to next

- [Concepts](concepts.md) — the four ideas that shape every wordlive API:
  politeness, semantic anchors, anchor IDs, and `EditScope`.
- [Cookbook](cookbook.md) — end-to-end recipes including an LLM tool-use loop.
- [Python API](python-api.md) — auto-generated reference for every public
  symbol.
