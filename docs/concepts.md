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
| `Pin`              | A durable handle (a wordlive-managed hidden bookmark) | Stored in the `.docx` |
| `Heading`          | A heading paragraph by visible text        | Reads the doc structure    |
| `Paragraph`        | Any paragraph by 1-based index             | Reads the doc structure    |
| `Cell` / `Row` / `Column` | A table cell, row, or column          | Reads the doc structure    |
| `Image` / `Shape`  | An inline picture / floating shape (text box, WordArt, float) | Reads the doc structure |
| `Equation` / `Chart` | A math zone / embedded chart by order    | Reads the doc structure    |
| `Footnote` / `Endnote` | A note by 1-based number               | Reads the doc structure    |
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
pin:7f3a2c           # a durable handle (doc.pin / bind:) that survives renumbering
cc:Signatory         # content control by Title (or Tag)
table:1:2:3          # cell at row 2, column 3 of the 1st table
table:1:row:2        # whole row 2 of the 1st table (table:1:col:3 for a column)
image:2              # the 2nd inline picture     (equation:N / chart:N likewise, by order)
shape:1              # the 1st floating shape     (textbox:N is an alias onto a text box's shape:N)
footnote:4           # the 4th footnote           (endnote:N likewise)
range:412-429        # arbitrary character span (the form find() emits)
header:1:primary     # primary header of section 1
footer:2:first       # first-page footer of section 2
start                # the position before the first paragraph (the prepend target)
end                  # the position past the last paragraph (the append target)
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

The bare `start` and `end` keywords are the two schemes without a `kind:value`
shape: they name the positions before the first and past the last paragraph —
the spots no content names — resolving to a
[`StartAnchor`](python-api.md#wordlive.StartAnchor) /
[`EndAnchor`](python-api.md#wordlive.EndAnchor) whose insert verbs all prepend /
append. They back [`doc.prepend_paragraph`](python-api.md#wordlive.Document) /
[`doc.append_paragraph`](python-api.md#wordlive.Document) (and the matching
inline [`doc.prepend`](python-api.md#wordlive.Document) /
[`doc.append`](python-api.md#wordlive.Document)) plus the `wordlive prepend` /
`append` commands, so building a document from either end needs no `.com` drop.

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

### Durable handles (pins)

`heading:N` / `para:N` are **positional** — they're paragraph indices, so an
insert or delete earlier in the document renumbers every id after it. That's fine
within a single read-decide-write, but a multi-step session (or an agent making
several passes) has to re-read `outline` to stay in sync. A **pin** removes that
churn: [`doc.pin("para:7")`](python-api.md#wordlive.Document) plants a
wordlive-managed hidden bookmark over that range and hands back a `pin:<code>` id
that Word keeps attached to the same *content* across later inserts, deletes, and
edits — the durability comes from Word maintaining the association natively.

```python
with doc.edit("Pin the budget block"):
    handle = doc.pin("heading:10", name="budget")   # → {"anchor_id": "pin:budget", …}

# …insert paragraphs above it; heading:10 has shifted, pin:budget has not…
doc.anchor_by_id("pin:budget").text                 # still the Budget heading
```

Resolve a pin like any other anchor (`doc.anchor_by_id("pin:budget")`), or feed it
straight into another op. Omit `name=` for a random code; reuse a slug to move the
handle. `doc.pin_outline()` pins every heading at once, and if pinned content is
later deleted the handle correctly vanishes (resolving it raises
`AnchorNotFoundError`). The rule of thumb: **positional ids for one pass, pins (or
name-based `bookmark:` / `cc:` anchors) for anything you'll revisit.** The
[Advanced session](advanced.md#step-2-pin-what-youre-about-to-move) walks a pin
end to end.

!!! info "Implementation"
    Resolution is centralised in
    [`Document.anchor_by_id`](python-api.md#wordlive.Document); see
    [`src/wordlive/_document/_core.py`](https://github.com/thomas-villani/wordlive/blob/main/src/wordlive/_document/_core.py).

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
