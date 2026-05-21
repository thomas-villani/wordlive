# Concepts

Four ideas drive almost every API decision in wordlive. If you understand
these, the rest of the surface follows.

## Politeness model

The user is editing the same document as your script. Naïve automation
clobbers their cursor, their scroll position, and their undo history. wordlive
refuses to do that.

Every [`Document.edit()`](python-api.md#wordlive.Document) scope:

1. Snapshots the user's `Selection` (start / end character offsets) and
   `ActiveWindow.VerticalPercentScrolled`.
2. Runs your mutations.
3. Restores the snapshot on the way out.

The captured snapshot is a plain dataclass:

```python
from wordlive import SelectionSnapshot

# Captured at the start of every doc.edit() block.
SelectionSnapshot(start=412, end=412, vertical_percent=37)
```

If you genuinely want to move the user — say, jumping their cursor to a new
section after inserting it — opt in explicitly:

```python
with doc.edit("Add risk paragraph") as scope:
    doc.heading("Risks").insert_paragraph_after("New risk identified.")
    scope.allow_cursor_move()       # don't restore on exit
    doc.go_to(doc.heading("Risks"))
```

Snapshots use best-effort restoration: if the document shrank and the old
offset is now invalid, wordlive collapses to the start of the snapshot range
rather than raising.

!!! info "Implementation"
    The snapshot dataclass, plus the `snapshot()` / `restore()` helpers, live in
    [`src/wordlive/_selection.py`](https://github.com/thomas-villani/wordlive/blob/main/src/wordlive/_selection.py).

## Semantic anchors over `Selection`

The Word object model encourages you to drive everything through the live
`Selection` — the user's cursor. That's hostile to both humans (your script
fights their typing) and LLM agents (the cursor is invisible state).

wordlive operates on **anchors** instead: named handles for ranges that don't
depend on the cursor.

| Anchor type        | What it names                              | Persistence                |
| ------------------ | ------------------------------------------ | -------------------------- |
| `Bookmark`         | A bookmark by name                         | Stored in the `.docx`      |
| `ContentControl`   | A structured field by Title (or Tag)       | Stored in the `.docx`      |
| `Heading`          | A heading paragraph by visible text        | Reads the doc structure    |
| `Cell`             | A table cell by (table, row, column)       | Reads the doc structure    |
| `HeaderFooter`     | A section's header/footer by (section, which) | Reads the doc structure |
| `RangeAnchor`      | An arbitrary character span by offsets     | Ephemeral (resolved live)  |

They all subclass [`Anchor`](python-api.md#wordlive.Anchor) and share the
same operations:

```python
addr = doc.bookmarks["Address"]
addr.text                           # read
addr.set_text("123 Main St")        # replace
addr.insert_before("Mailing: ")     # insert without replacing
addr.insert_after(" (verified)")
addr.delete()
addr.com                            # raw COM Range — escape hatch
```

Why not Selection-driven? Two reasons:

1. **Idempotent operations are easier to reason about.** "Set the Address
   bookmark to X" is repeatable; "type X at the cursor" is not.
2. **LLM tool use needs stable identifiers.** A bookmark name is stable; a
   character offset isn't.

## Anchor IDs

Each anchor kind has its own collection (`doc.bookmarks`,
`doc.content_controls`, `doc.tables`, `doc.sections`, `doc.paragraphs`,
`doc.heading(name)`, …). For programmatic addressing across all of them —
especially from JSON tool-use payloads — wordlive uses a single string scheme:

```
heading:3            # 1-based paragraph index of a heading
para:5               # 1-based index of any paragraph (same index space as heading:N)
bookmark:Address     # bookmark by name
cc:Signatory         # content control by Title (or Tag)
table:1:2:3          # cell at row 2, column 3 of the 1st table
range:412-429        # arbitrary character span (the form find() emits)
header:1:primary     # primary header of section 1
footer:2:first       # first-page footer of section 2
```

`para:N` and `heading:N` index the same paragraph stream, so a heading at
`para:5` is also `heading:5` — the difference is that `heading:N` refuses to
resolve a non-heading paragraph, while `para:N` resolves any paragraph.
[`doc.outline()`](python-api.md#wordlive.Document) emits the heading-only view;
[`doc.paragraphs.list()`](python-api.md#wordlive.ParagraphCollection) (and
`outline --all`) emits every paragraph with offsets.

The bare `table:N` form is deliberately *not* an anchor — a whole table is a
collection, not a single range — so it's addressed through `doc.tables[N]` and
the `table` CLI group instead. Only cells (`table:N:R:C`) resolve via
`anchor_by_id`. Header/footer ids take a section index `S` and a `WHICH` of
`primary` / `first` / `even`; the bare `section:N` is likewise a collection, not
an anchor (use `doc.sections[N]`).

The `range:START-END` form is what [`find()`](python-api.md#wordlive.Document)
emits for each hit, and it round-trips: feed it back into `replace --anchor-id`
or `comments.add` to act on exactly the span that was found. Range offsets are
*live* — they're resolved against the document on each use, so an edit that
shifts the text earlier can leave a stale range pointing at the wrong place.
Resolve, act, discard.

These IDs are emitted directly by [`doc.outline()`](python-api.md#wordlive.Document):

```python
doc.outline()
# [
#   {"level": 1, "text": "Introduction", "anchor_id": "heading:1"},
#   {"level": 2, "text": "Context",      "anchor_id": "heading:3"},
#   {"level": 1, "text": "Risks",        "anchor_id": "heading:8"},
# ]
```

And consumed by [`doc.anchor_by_id()`](python-api.md#wordlive.Document) and
every CLI command that takes `--anchor-id`:

```python
anchor = doc.anchor_by_id("heading:3")
anchor.set_text("Updated section heading")
```

Why a paragraph index for headings instead of the heading text? Two headings
can share the same text ("Background", "Background") and the index
disambiguates. The `heading:N` form always refers to the *Nth paragraph in
the document*, which is stable across the lifetime of a session.

!!! info "Implementation"
    Resolution is centralised in
    [`Document.anchor_by_id`](python-api.md#wordlive.Document); see
    [`src/wordlive/_document.py`](https://github.com/thomas-villani/wordlive/blob/main/src/wordlive/_document.py).

## `EditScope` and atomic undo

`doc.edit("label")` returns an [`EditScope`](python-api.md#wordlive.EditScope).
Inside the `with` block, wordlive opens `Application.UndoRecord` so every
mutation is bundled into a single Ctrl-Z step labelled with your string.

```python
with doc.edit("Replace boilerplate"):
    doc.bookmarks["Greeting"].set_text("Hello,")
    doc.bookmarks["Closing"].set_text("Best,")
    doc.heading("Footer").set_text("Signed electronically.")

# In Word's undo dropdown: a single entry, "Replace boilerplate".
```

Two responsibilities are bundled into the same context manager:

1. **`UndoRecord`** — start/end the recording. On Word versions that don't
   support `UndoRecord` (pre-2010), wordlive silently falls back to running
   the ops without atomic-undo; everything still works, you just get N undo
   entries instead of one.
2. **`SelectionSnapshot`** — see [Politeness](#politeness-model).

The scope object itself exposes one knob:

```python
with doc.edit("Insert and jump") as scope:
    doc.heading("Introduction").insert_paragraph_after("…")
    scope.allow_cursor_move()       # skip the snapshot restore
    doc.go_to(doc.heading("Introduction"))
```

Most code never touches the scope — just `with doc.edit("label"):` and write
your mutations.

!!! info "Implementation"
    [`EditScope`](python-api.md#wordlive.EditScope) lives in
    [`src/wordlive/_edit.py`](https://github.com/thomas-villani/wordlive/blob/main/src/wordlive/_edit.py).

## The `.com` escape hatch

wordlive deliberately covers a small surface. When you need something it
doesn't, every wrapper exposes the raw COM object via `.com`:

```python
with wl.attach() as word:
    doc = word.documents.active

    # Anything wordlive covers, use the wordlive API.
    with doc.edit("Bold the first ten characters"):
        # Anything it doesn't, drop to COM.
        doc.com.Range(0, 10).Font.Bold = True
```

`word.com`, `doc.com`, `anchor.com`, and `selection.com` all return the
underlying pywin32 dispatch object. Treat this as a forward-compatibility
seam: as wordlive grows, today's COM call may become tomorrow's high-level
helper, but the escape hatch is permanent.
