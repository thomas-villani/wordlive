# CLI

`wordlive` ships a Click-based CLI designed for LLM tool-use loops: JSON in,
JSON out, deterministic exit codes, one structured object per invocation on
stdout. The CLI is a thin wrapper over the [Python API](python-api.md) — same
politeness, same atomic-undo.

## Global flags

```
wordlive [--json|--text] [--doc DOC_NAME] <subcommand> [args]
```

| Flag             | Default     | Purpose                                                |
| ---------------- | ----------- | ------------------------------------------------------ |
| `--json/--text`  | `--json`    | Output format. `--text` prints a per-command human form (indented outline tree, bare text for reads, one-line acks for writes); JSON stays the LLM-friendly default. |
| `--doc DOC_NAME` | active doc  | Target a specific open document by name (e.g. `Report.docx`). |
| `-h`, `--help`   | —           | Show help for the command or subgroup.                  |

## Exit codes

The CLI's error boundary classifies every [`WordliveError`](errors.md) into a
deterministic exit code so an LLM tool-use loop can branch on the failure
mode without parsing strings:

| Code | Meaning                | Source exception           |
| ---- | ---------------------- | -------------------------- |
| `0`  | OK                     | —                          |
| `1`  | Other / unclassified   | `WordliveError` (default), `DocumentNotFoundError`, `ImageSourceError` |
| `2`  | Anchor or style missing | `AnchorNotFoundError` / `StyleNotFoundError` (also used for zero-match `find`/`replace --find`) |
| `3`  | Word busy / modal      | `WordBusyError` (retryable) |
| `4`  | Word not running       | `WordNotRunningError`      |
| `5`  | Ambiguous match        | `AmbiguousMatchError` (multiple `find` hits without `--all`/`--occurrence`) |

See the [Errors page](errors.md) for the full exception taxonomy and
retry guidance.

---

## `status`

```
wordlive status
```

List all open documents and mark which is active.

```bash
$ wordlive status
[{"name": "Report.docx", "path": "C:\\Users\\me\\Report.docx", "is_active": true},
 {"name": "Draft.docx",  "path": "C:\\Users\\me\\Draft.docx",  "is_active": false}]
```

Failures: `4` if Word isn't running (returns `[]` to stdout *and* the error
on stderr). Useful as a probe before issuing other commands.

## `outline`

```
wordlive outline [--all] [--doc DOC_NAME]
```

Heading outline of the target document, with addressable anchor IDs. Pass
`--all` to list **every** paragraph (headings *and* body text *and* list items)
as `para:N` — identical to [`paragraphs`](#paragraphs).

```bash
$ wordlive outline
[{"level": 1, "text": "Introduction", "anchor_id": "heading:1"},
 {"level": 2, "text": "Context",      "anchor_id": "heading:3"},
 {"level": 1, "text": "Risks",        "anchor_id": "heading:8"}]
```

This is the entry point for LLM workflows that need to discover what's
addressable in the document. The emitted `anchor_id` strings are exactly
what `replace`, `go-to`, `insert`, and `exec` consume.

## `paragraphs`

```
wordlive paragraphs [--doc DOC_NAME]
```

List **every** paragraph in document order — headings, body text, and list
items alike — each with a `para:N` anchor, its outline `level`, an
`is_heading` flag, character `start`/`end` offsets, and its text. `outline
--all` is an alias.

```bash
$ wordlive paragraphs
[{"index": 1, "anchor_id": "para:1", "level": 1,  "is_heading": true,  "start": 0,  "end": 13, "text": "Introduction"},
 {"index": 2, "anchor_id": "para:2", "level": 10, "is_heading": false, "start": 13, "end": 29, "text": "Body text here."},
 {"index": 3, "anchor_id": "para:3", "level": 2,  "is_heading": true,  "start": 29, "end": 35, "text": "Risks"}]
```

`para:N` shares its index space with `heading:N` — paragraph 1 is both
`para:1` and (because it's a heading) `heading:1`. The emitted offsets feed a
[`range:START-END`](concepts.md) target for an offset-precise, mid-paragraph
insertion via `replace`.

## `read bookmark NAME`

```
wordlive read bookmark NAME [--doc DOC_NAME]
```

Read the text of a bookmark.

```bash
$ wordlive read bookmark Address
{"text": "123 Main St"}
```

Failures: `2` if the bookmark doesn't exist.

## `read cc NAME`

```
wordlive read cc NAME [--doc DOC_NAME]
```

Read the text of a content control. `NAME` matches the control's **Title**
first, then **Tag**.

```bash
$ wordlive read cc Signatory
{"text": "Jane Doe"}
```

Failures: `2` if no content control with that Title/Tag exists.

## `read section HEADING`

```
wordlive read section HEADING                [--doc DOC_NAME]
wordlive read section --anchor-id heading:N  [--doc DOC_NAME]
```

Read the body text under a heading — from the end of the heading paragraph up
to the next heading whose level is **less than or equal to** this one's (or to
the end of the document if there's no such boundary).

```bash
$ wordlive read section "Introduction"
{"heading": "Introduction",
 "anchor_id": "heading:1",
 "level": 1,
 "text": "This document covers the Q2 risk register …"}

$ wordlive --text read section "Introduction"
This document covers the Q2 risk register …
```

`--text` mode emits only the section body — handy for piping into an LLM
prompt without ceremony. Use `--anchor-id heading:N` to disambiguate when the
same visible heading text appears more than once.

Failures: `2` heading not found.

## `write bookmark NAME --text "…"`

```
wordlive write bookmark NAME --text "..." [--doc DOC_NAME]
```

Replace a bookmark's text inside a single atomic-undo scope.

```bash
$ wordlive write bookmark Address --text "456 Elm St"
{"ok": true, "anchor": {"kind": "bookmark", "name": "Address"}}
```

The bookmark is preserved — wordlive re-adds it covering the new content
after the Word `Range.Text` assignment (which would otherwise delete it).

Failures: `2` anchor not found, `3` Word busy.

## `write cc NAME --text "…"`

```
wordlive write cc NAME --text "..." [--doc DOC_NAME]
```

Replace a content control's text inside a single atomic-undo scope.

```bash
$ wordlive write cc Signatory --text "Jane Doe"
{"ok": true, "anchor": {"kind": "content_control", "name": "Signatory"}}
```

Failures: `2` anchor not found, `3` Word busy.

## `insert --anchor-id ID --text "…"`

```
wordlive insert --anchor-id ID --text "..." [--before | --after] [--style "Body Text"] [--doc DOC_NAME]
```

Insert a new paragraph relative to **any** anchor — addressed the same way
every other command addresses things, with `--anchor-id` (`heading:N`,
`para:N`, `bookmark:NAME`, a cell, a range). `--after` (the default) lands the
new paragraph just below the anchor; `--before` lands it just above. `--after`
works even when the anchor is the document's last paragraph — the new paragraph
is appended before the final mark — so you can build a document top-down from a
single empty paragraph.

```bash
$ wordlive insert --anchor-id heading:8 --text "New risk identified."
{"ok": true, "anchor_id": "heading:8", "where": "after", "style": null}

$ wordlive insert --anchor-id para:3 --text "Section preamble." --before
{"ok": true, "anchor_id": "para:3", "where": "before", "style": null}
```

`--style` is optional; if given it must be a Word style name that exists in
the document — the style is validated before the paragraph is inserted, so a
typo never partially mutates the document. Use `wordlive style list` to see
the available names. Failures: `2` anchor not found or style not found, `3`
Word busy.

To insert text *inside* a paragraph at a precise offset rather than as a new
paragraph, target a collapsed range instead — `replace --anchor-id
range:120-120 --text "…"` — using offsets from `paragraphs` or `find`.

## `insert-break --anchor-id ID [--kind …] [--before | --after]`

```
wordlive insert-break --anchor-id ID
    [--kind page|column|section_next|section_continuous]
    [--before | --after] [--doc DOC_NAME]
```

Insert an explicit page, column, or section break at any anchor — the clean,
discoverable replacement for appending a paragraph whose text is a literal
form-feed. `--kind` defaults to `page` (the common case); `column` breaks a
multi-column layout, and the two `section_*` kinds start a new document section
(which can carry its own headers/footers and page setup — see `section`).
`--after` (default) drops the break just past the anchor; `--before`, just
before it.

```bash
$ wordlive insert-break --anchor-id para:12
{"ok": true, "anchor_id": "para:12", "kind": "page", "where": "after"}

$ wordlive insert-break --anchor-id heading:3 --kind section_next --before
{"ok": true, "anchor_id": "heading:3", "kind": "section_next", "where": "before"}
```

To make a *style* (e.g. every `Heading 1`) open a new page without a stray
break character that drifts on reflow, prefer `format-paragraph --anchor-id ID
--page-break-before` instead — it's a paragraph property, not an inserted mark.
Failures: `1` unknown `--kind` (usage error), `2` anchor not found, `3` Word
busy.

## `prepend --text "…"` / `append --text "…"`

```
wordlive prepend --text "..." [--paragraph | --inline] [--style "Body Text"] [--doc DOC_NAME]
wordlive append  --text "..." [--paragraph | --inline] [--style "Body Text"] [--doc DOC_NAME]
```

`prepend` is the mirror of `append`: it adds to the very **start** of the
document (a new first paragraph by default, or `--inline` to join the opening
paragraph) — equivalent to `insert --anchor-id start --text "…"`. Everything
below applies to both; just swap "end" for "start".

```bash
$ wordlive prepend --text "DRAFT — not for distribution"
{"ok": true, "mode": "paragraph", "style": null}
```

Append text to the very end of the document — the high-level "end of doc"
helper, no anchor needed. `--paragraph` (the default) makes `text` a new final
paragraph; `--inline` continues the document's last paragraph instead. This is
exactly `insert --anchor-id end --text "…"` (the `end` anchor names the
position past the last paragraph), spelled as its own verb.

```bash
$ wordlive append --text "Closing note added by automation."
{"ok": true, "mode": "paragraph", "style": null}

$ wordlive append --text " (verified)" --inline
{"ok": true, "mode": "inline", "style": null}
```

`--style` is optional, paragraph-mode only, and must name a style that exists
in the document — it's validated before anything is written, so a typo never
partially mutates the document (`wordlive style list` shows the names).
Failures: `2` style not found, `3` Word busy.

## `insert-image --anchor-id ID (--path FILE | --base64 VALUE) --wrap WRAP`

```
wordlive insert-image --anchor-id ID (--path FILE | --base64 VALUE) --wrap WRAP \
    [--before | --after] [--width N] [--height N] [--alt-text "…"] \
    [--lock-aspect | --no-lock-aspect] [--doc DOC_NAME]
```

Embed an image at **any** anchor. Exactly one image source is required:
`--path` reads a file from disk (best for large images); `--base64` takes
base64 data inline, or `--base64 -` reads base64 from **stdin** (handy when an
LLM holds image data in memory). The picture is embedded in the document, not
linked, so the source file can move or vanish afterwards. Word auto-detects the
image's natural size; `--width`/`--height` (points) override it and
`--lock-aspect` (the default) keeps the aspect ratio.

`--wrap` is **required** so layout intent is always explicit:

| `--wrap`                                            | Effect                                              |
| --------------------------------------------------- | --------------------------------------------------- |
| `inline`                                            | Stays in the text flow, like a character.           |
| `auto`                                              | Floats: Square if ≤ half the page's usable width, else top-and-bottom. |
| `square` `tight` `through` `top-bottom` `front` `behind` | Floats with that wrap type.                    |

`--after` (default) places the image just below the anchor; `--before` above.

```bash
$ wordlive insert-image --anchor-id heading:2 --path diagram.png --wrap auto
{"ok": true, "anchor_id": "heading:2", "anchor": {"kind": "heading", "name": "Risks"}, "wrap": "auto", "where": "after"}

$ base64 logo.png | wordlive insert-image --anchor-id bookmark:Header --base64 - --wrap square --width 96
{"ok": true, "anchor_id": "bookmark:Header", "anchor": {"kind": "bookmark", "name": "Header"}, "wrap": "square", "where": "after"}
```

Failures: `1` the image is missing, unreadable, or not a recognised raster
format (PNG/JPEG/GIF/BMP/TIFF) — an `ImageSourceError`; `2` anchor not found
or an invalid `--wrap` value; `3` Word busy.

## `snapshot [--anchor-id ID | --page N | --pages A-B]`

```
wordlive snapshot [--anchor-id ID | --page N | --pages A-B] \
    [--out FILE] [--dpi 150] [--doc DOC_NAME]
```

Render document page(s) to PNG so a **vision model can see the layout** — real
fonts, spacing, and page geometry, not just the text. Word exports a
pixel-faithful PDF of the document it has open and wordlive rasterises the
requested pages. Read-only: the document and the user's cursor are untouched.

Pick **at most one** target; with none, the whole document is rendered:

| Target          | Renders |
| --------------- | ------- |
| `--anchor-id ID` | the page(s) the anchor occupies — a `heading:` expands to its **whole section** (heading + body) |
| `--page N`       | a single 1-based page |
| `--pages A-B`    | an inclusive page span, e.g. `2-4` |

Output: with `--out FILE` the image is written to disk — a single page to `FILE`,
multiple pages alongside it as `<stem>-p<N><suffix>`. **Without `--out`, base64
PNG data is returned inline** in the JSON (`images[].base64`), which suits an LLM
that wants to look at the page directly. `--dpi` (default `150`) sets resolution.

This needs the optional **`snapshot` extra** (PyMuPDF):
`pip install "wordlive[snapshot]"` (or `uv add "wordlive[snapshot]"`).

```bash
$ wordlive snapshot --anchor-id heading:3 --out section.png
{"ok": true, "selector": "heading:3", "dpi": 150, "count": 1, "images": [{"page": 4, "bytes": 81234, "path": "section.png"}]}

$ wordlive snapshot --page 1
{"ok": true, "selector": 1, "dpi": 150, "count": 1, "images": [{"page": 1, "bytes": 64210, "base64": "iVBORw0KGgo…"}]}
```

Failures: `1` PyMuPDF isn't installed, or rasterising the PDF failed — a
`SnapshotError`; `2` `--anchor-id` not found; `3` Word busy.

## `cursor read` / `cursor write --text "…"`

```
wordlive cursor read [--doc DOC_NAME]
wordlive cursor write --text "..." [--replace | --no-replace] [--doc DOC_NAME]
```

The **explicit cursor surface**. Every other command targets a semantic anchor
and preserves the user's cursor; `cursor` is the deliberate exception, for when
the user genuinely wants to read or write at their current position. It is *not*
addressable by `--anchor-id` — that separation is intentional, signalling it's
the non-preferred mode.

`cursor read` reports the selection's `start`/`end`, whether it's `collapsed`
(an insertion point with no selected text), the selected `text`, and the
containing `para:N` so you can pivot back to anchored edits:

```bash
$ wordlive cursor read
{"start": 142, "end": 142, "collapsed": true, "text": "", "paragraph": {"anchor_id": "para:7"}}
```

`cursor write` types at the cursor and — unlike anchor writes — deliberately
leaves the cursor after the inserted text. With a spanning selection, the
default `--replace` overwrites it (like typing); `--no-replace` inserts at the
selection start without removing it.

```bash
$ wordlive cursor write --text "inserted at cursor"
{"ok": true, "replace": true}
```

Failures: `3` Word busy, `4` Word not running.

## `find --text "…"`

```
wordlive find --text "..." [--in ANCHOR_ID] [--doc DOC_NAME]
```

Locate every fuzzy occurrence of the given text in the document (read-only).
Matching is forgiving of cosmetic differences that show up when LLMs re-emit
text — whitespace runs collapse, smart quotes/dashes fold to ASCII, NBSPs
become spaces, and the strings are NFKC-normalized before comparison.

```bash
$ wordlive find --text "the risk register"
[{"anchor_id": "range:412-429",
  "start": 412,
  "end": 429,
  "text": "the risk register"}]
```

Use `--in ANCHOR_ID` to restrict the search. Headings expand to their
*section* (the body under the heading), which is the common case for
"replace this phrase, but only inside the Risks section":

```bash
$ wordlive find --text "Q1" --in heading:8
```

`text` in each match is the actual original substring (with smart quotes,
NBSP, etc.) — that's the form Word will preserve when you replace it.

Failures: returns `[]` with exit `0` for no matches; `2` if `--in ANCHOR_ID`
refers to a missing anchor.

## `replace`

```
wordlive replace --anchor-id ID --text "..."                            [--doc DOC_NAME]   # anchor mode
wordlive replace --find OLD --text NEW [--in ID] [--all|--occurrence N] [--doc DOC_NAME]   # fuzzy mode
```

Two modes share the verb:

### Anchor mode — replace an entire range

Replace the text at an anchor identified by [anchor ID](concepts.md#anchor-ids).
Works across all three anchor kinds.

```bash
$ wordlive replace --anchor-id heading:3 --text "Updated section text"
{"ok": true,
 "anchor_id": "heading:3",
 "anchor": {"kind": "heading", "name": "Context"}}
```

The response's `anchor.name` resolves the ID back to a human-readable name.
Failures: `2` anchor not found, `3` Word busy.

### Fuzzy mode — find + replace a substring

Locate `--find OLD` (same fuzzy match as `find`) and replace with `--text NEW`.
Word's native range replacement preserves the *character formatting* of the
matched range — bold stays bold, italics stay italic — so you don't need to
re-state formatting on the replacement.

```bash
$ wordlive replace --find "Q1 2025" --text "Q2 2025"
{"ok": true,
 "replacements": [{"anchor_id": "range:412-419",
                   "start": 412, "end": 419, "text": "Q1 2025"}]}
```

| Flag           | Meaning                                                                                       |
| -------------- | --------------------------------------------------------------------------------------------- |
| `--in ID`      | Restrict search to the given anchor's range (headings expand to their section).               |
| `--all`        | Replace every match. Mutually exclusive with `--occurrence`.                                  |
| `--occurrence N` | Replace only the Nth match (1-based). Mutually exclusive with `--all`.                      |

Failures:

- **Exit `2`** — zero matches. Same code as anchor-not-found because the
  agent's recovery is identical: re-fetch state and retry.
- **Exit `5`** — multiple matches and neither `--all` nor `--occurrence` was
  given. Stdout still emits a JSON payload listing all matches so the agent
  can pick an occurrence and retry:

  ```json
  {"ok": false, "error": "ambiguous_match", "find": "Q1",
   "matches": [{"start": 412, "end": 414, "text": "Q1"},
               {"start": 887, "end": 889, "text": "Q1"}]}
  ```

- **Exit `3`** — Word busy.

## `go-to --anchor-id ID`

```
wordlive go-to --anchor-id ID [--no-scroll] [--doc DOC_NAME]
```

Move the user's cursor to an anchor. **This is the one CLI command that
deliberately disturbs the user's selection** — every other write preserves
it.

```bash
$ wordlive go-to --anchor-id bookmark:Address
{"ok": true,
 "anchor_id": "bookmark:Address",
 "anchor": {"kind": "bookmark", "name": "Address"}}
```

`--no-scroll` collapses the selection at the anchor without scrolling the
view to it. Failures: `2` anchor not found, `3` Word busy.

## `style list`

```
wordlive style list [--doc DOC_NAME]
```

Enumerate every style defined in the document.

```bash
$ wordlive style list
[{"name": "Normal",    "type": "paragraph", "builtin": true, "in_use": true},
 {"name": "Body Text", "type": "paragraph", "builtin": true, "in_use": true},
 {"name": "Heading 1", "type": "paragraph", "builtin": true, "in_use": true}]
```

`type` is one of `"paragraph"`, `"character"`, `"table"`, `"list"`. Built-in
Word styles set `builtin: true`; user-defined styles set `false`. Failures:
`3` Word busy, `4` Word not running.

## `style apply --anchor-id ID --name NAME`

```
wordlive style apply --anchor-id ID --name "Heading 2" [--doc DOC_NAME]
```

Apply a style to the anchor's range. Atomic-undo. The style must already
exist in the document — wordlive does not create styles on demand.

```bash
$ wordlive style apply --anchor-id heading:3 --name "Heading 2"
{"ok": true,
 "anchor_id": "heading:3",
 "anchor": {"kind": "heading", "name": "Risks"},
 "style": "Heading 2"}
```

Word picks paragraph- vs. character-style behaviour from the style's own
`Type`; you don't need to model that distinction. Failures: `2` anchor or
style not found, `3` Word busy.

## `format-paragraph --anchor-id ID [...]`

```
wordlive format-paragraph --anchor-id ID
    [--alignment left|center|centre|right|justify]
    [--left-indent POINTS] [--right-indent POINTS] [--first-line-indent POINTS]
    [--space-before POINTS] [--space-after POINTS]
    [--page-break-before | --no-page-break-before]
    [--doc DOC_NAME]
```

`centre` is accepted as a synonym for `center` (UK spelling).

Set paragraph-formatting properties on the anchor's range. At least one
formatting flag is required. Indent and spacing values are in **points** —
the unit Word's COM API uses natively for these fields.
`--page-break-before` forces the paragraph to begin on a new page (and
`--no-page-break-before` clears it) — the *clean*, reflow-safe way to
page-break, leaving no stray break character (contrast `insert-break`, which
inserts an explicit one-off break). Atomic-undo.

```bash
$ wordlive format-paragraph --anchor-id heading:3 \
      --alignment center --space-before 6
{"ok": true,
 "anchor_id": "heading:3",
 "anchor": {"kind": "heading", "name": "Risks"},
 "applied": {"alignment": "center", "space_before": 6.0}}
```

Only the flags you pass are written; everything else on the paragraph is
left alone. If the anchor spans a partial paragraph (e.g., a bookmark
covering five words inside a longer paragraph), Word applies the formatting
to the *enclosing* paragraph — that's the COM behaviour, not a wordlive
quirk. Failures: `2` anchor not found, `3` Word busy.

## `table list`

```
wordlive table list [--doc DOC_NAME]
```

Enumerate every table in the document, in top-to-bottom order.

```bash
$ wordlive table list
[{"index": 1, "title": "Budget", "rows": 4, "columns": 3},
 {"index": 2, "title": "",       "rows": 2, "columns": 2}]
```

`index` is the 1-based position used to address the table and its cells.
`title` is the table's Title (empty string if unset). Failures: `3` Word busy,
`4` Word not running.

## `table read INDEX`

```
wordlive table read INDEX [--doc DOC_NAME]
```

Read table `INDEX` (1-based) as a grid. Each cell carries its **`anchor_id`**
(`table:N:R:C`) so you can feed it straight into `replace`, `style apply`, or
`format-paragraph`.

```bash
$ wordlive table read 1
{"index": 1, "title": "Budget", "rows": 2, "columns": 2,
 "cells": [[{"row": 1, "col": 1, "text": "Item",  "anchor_id": "table:1:1:1"},
            {"row": 1, "col": 2, "text": "Cost",  "anchor_id": "table:1:1:2"}],
           [{"row": 2, "col": 1, "text": "Travel","anchor_id": "table:1:2:1"},
            {"row": 2, "col": 2, "text": "$400",  "anchor_id": "table:1:2:2"}]]}
```

Cell text is stripped of Word's internal end-of-cell markers. To **write** a
cell, use its anchor id with `replace`:

```bash
$ wordlive replace --anchor-id table:1:2:2 --text "$450"
{"ok": true, "anchor_id": "table:1:2:2", "anchor": {"kind": "cell", "name": "table:1:2:2"}}
```

Failures: `2` table index out of range, `3` Word busy.

## `table create`

```
wordlive table create --anchor-id ID --rows R --cols C
                      [--style NAME] [--header] [--before|--after]
                      [--data '[["…"],…]' | --data -] [--doc DOC_NAME]
```

Create a new `R`×`C` table at a **position anchor** (`heading:`, `para:`,
`start`, `end`, `range:` — *not* a bare `table:N`, which addresses an existing
table). Every other verb edits existing structure; this is how you build a table
from nothing. Atomic-undo. Reports the new table's 1-based `index` for an
immediate follow-up `set-cell` / `add-row`.

`--data` populates the cells at creation from a **row-major** JSON 2-D array
(`[[r1c1, r1c2], …]`), validated against `R`×`C` up front — a short/partial
array leaves trailing cells empty; an array that *overflows* the grid is a clean
error (exit 1). Pass `--data -` to read the JSON from stdin, which sidesteps
Windows quoting/backslash fights (mirrors `exec --ops -`).

`--style` names a table style defined in the document; it defaults to the
built-in **`Table Grid`** so a new table has visible borders rather than only
faint gridlines. A style name not in the document fails (exit 2). `--header`
bolds the first row.

```bash
$ wordlive table create --anchor-id end --rows 3 --cols 3 --header \
    --data '[["Tier","Monthly","SLA"],["Wobble","$9","best effort"],["Finch","$99","99.9%"]]'
{"ok": true, "table": 2, "rows": 3, "columns": 3}
```

A table appended where another already sits flush against it (e.g. two tables in
a row at the end of the document) is kept distinct: Word would otherwise merge
adjacent tables, so a separating paragraph is inserted automatically.

Failures: `1` bad dimensions / `--data` shape, `2` anchor or style not found,
`3` Word busy, `4` Word not running.

## `table add-row`

```
wordlive table add-row --table INDEX [--values '["a","b"]'] [--doc DOC_NAME]
```

Append a row at the end of the table. `--values` is an optional JSON array of
cell values, matched to columns left-to-right (extras ignored, short lists
leave trailing cells empty). Atomic-undo.

```bash
$ wordlive table add-row --table 1 --values '["Lodging", "$600"]'
{"ok": true, "table": 1, "rows": 3}
```

Failures: `2` table index out of range, `3` Word busy.

## `table delete-row`

```
wordlive table delete-row --table INDEX --row R [--doc DOC_NAME]
```

Delete the 1-based row `R` from the table. Atomic-undo.

```bash
$ wordlive table delete-row --table 1 --row 3
{"ok": true, "table": 1, "rows": 2}
```

Failures: `2` table index or row out of range, `3` Word busy.

## `table delete INDEX`

```
wordlive table delete INDEX [--doc DOC_NAME]
```

Delete table `INDEX` (1-based) and all its cells — the structural mirror of
`table create` / `delete-row`. Atomic-undo. The indices of any tables below it
shift down by one afterwards.

```bash
$ wordlive table delete 2
{"ok": true, "deleted": 2}
```

Failures: `2` table index out of range, `3` Word busy.

## `comment list`

```
wordlive comment list [--doc DOC_NAME]
```

Enumerate every review comment, in document order.

```bash
$ wordlive comment list
[{"index": 1, "author": "ReviewBot", "text": "Please verify this figure.",
  "scope": "$400", "done": false}]
```

`index` is the 1-based handle used by `resolve` and `delete`. `scope` is the
document text the comment is attached to; `done` is the resolved flag (always
`false` on Word builds older than 2013). Failures: `3` Word busy, `4` Word not
running.

## `comment add --anchor-id ID --text "…"`

```
wordlive comment add --anchor-id ID --text "..." [--author NAME] [--doc DOC_NAME]
```

Attach a comment to the anchor's range. **The document text is untouched** —
this is the polite, side-channel alternative to rewriting a passage. Atomic-undo.
The anchor id is any of the [recognised forms](concepts.md#anchor-ids), including
a `range:START-END` from `find`.

```bash
$ wordlive comment add --anchor-id heading:3 --text "Please expand this." --author "ReviewBot"
{"ok": true, "anchor_id": "heading:3", "comment": {"index": 1, "author": "ReviewBot"}}
```

`--author` is optional; without it Word uses the running app's user name.
Failures: `2` anchor not found, `3` Word busy.

## `comment resolve --index N`

```
wordlive comment resolve --index N [--doc DOC_NAME]
```

Mark comment `N` (from `comment list`) as resolved/done. Requires Word 2013+.
Atomic-undo.

```bash
$ wordlive comment resolve --index 1
{"ok": true, "index": 1, "done": true}
```

Failures: `2` index out of range, `3` Word busy.

## `comment delete --index N`

```
wordlive comment delete --index N [--doc DOC_NAME]
```

Delete comment `N`. Remaining comments re-index, so re-list before deleting
another by index. Atomic-undo.

```bash
$ wordlive comment delete --index 1
{"ok": true, "index": 1, "deleted": true}
```

Failures: `2` index out of range, `3` Word busy.

## `track status | on | off`

```
wordlive track status [--doc DOC_NAME]
wordlive track on      [--doc DOC_NAME]
wordlive track off     [--doc DOC_NAME]
```

Inspect or toggle the document's **Track Changes** setting. While on, every
edit (yours or the user's) is recorded as a revision the user can accept or
reject.

```bash
$ wordlive track status
{"tracked": false}

$ wordlive track on
{"ok": true, "tracked": true}
```

The toggle is **persistent** — `track on` leaves Word recording revisions until
`track off`. For a self-restoring scope, prefer the library's
`doc.tracked_changes()` context manager, or set `"tracked": true` on an
[`exec` script](#exec-script-opsjson) to record a single batch as tracked
changes and restore the prior setting afterwards. Failures: `3` Word busy, `4`
Word not running.

## `list show`

```
wordlive list show [--doc DOC_NAME]
```

Enumerate every bullet/numbered list in the document, top to bottom. Each row
carries a `range:START-END` **`anchor_id`** covering the whole list, so you can
feed it straight into `list restart`, `replace`, or `comment add`.

```bash
$ wordlive list show
[{"index": 1, "type": "numbered", "count": 4, "anchor_id": "range:512-690"}]
```

`type` is `bulleted` / `numbered` / `outline` / `number-only` / `mixed`.
Failures: `3` Word busy, `4` Word not running.

## `list apply --anchor-id ID --type …`

```
wordlive list apply --anchor-id ID [--type bulleted|numbered|outline] [--continue] [--doc DOC_NAME]
```

Turn the anchor's paragraphs into a list. `--type` defaults to `bulleted`.
Numbering starts fresh at 1 unless `--continue` is given, which continues from
a list immediately above. Atomic-undo.

```bash
$ wordlive list apply --anchor-id heading:6 --type numbered
{"ok": true, "anchor_id": "heading:6",
 "anchor": {"kind": "heading", "name": "Steps"},
 "type": "numbered", "continue_previous": false}
```

Failures: `2` anchor not found, `3` Word busy.

## `list info --anchor-id ID`

```
wordlive list info --anchor-id ID [--doc DOC_NAME]
```

Report the list state at an anchor (read-only): `{type, level, number, string}`,
where `string` is the rendered marker (`"1."`, `"a)"`, `"•"`). `type` is
`"none"` when the anchor isn't in a list.

```bash
$ wordlive list info --anchor-id range:512-540
{"type": "numbered", "level": 1, "number": 3, "string": "3."}
```

Failures: `2` anchor not found, `3` Word busy.

## `list remove | restart | indent | outdent --anchor-id ID`

```
wordlive list remove  --anchor-id ID [--doc DOC_NAME]   # strip list formatting
wordlive list restart --anchor-id ID [--doc DOC_NAME]   # restart numbering at 1
wordlive list indent  --anchor-id ID [--doc DOC_NAME]   # demote one level (1 -> 2)
wordlive list outdent --anchor-id ID [--doc DOC_NAME]   # promote one level (2 -> 1)
```

All four are atomic-undo. `restart` re-applies the list's own template starting
at 1; it errors if the anchor isn't part of a list.

```bash
$ wordlive list restart --anchor-id range:512-540
{"ok": true, "anchor_id": "range:512-540", "anchor": {"kind": "range", "name": "range:512-540"}}
```

Failures: `2` anchor not found, `3` Word busy.

## `section list`

```
wordlive section list [--doc DOC_NAME]
```

List the document's sections with each one's page setup.

```bash
$ wordlive section list
[{"index": 1,
  "page_setup": {"orientation": "portrait",
                 "top_margin": 72.0, "bottom_margin": 72.0,
                 "left_margin": 72.0, "right_margin": 72.0,
                 "page_width": 612.0, "page_height": 792.0}}]
```

Margins and page dimensions are in **points**. Headers and footers live in the
`header` / `footer` commands. Failures: `3` Word busy, `4` Word not running.

## `header read | write` · `footer read | write`

```
wordlive header read  [--section N] [--which primary|first|even] [--doc DOC_NAME]
wordlive header write [--section N] [--which primary|first|even] --text "..." [--doc DOC_NAME]
wordlive footer read  [--section N] [--which primary|first|even] [--doc DOC_NAME]
wordlive footer write [--section N] [--which primary|first|even] --text "..." [--doc DOC_NAME]
```

Read or set a section's header/footer. `--section` defaults to `1` and
`--which` to `primary` (the other options are `first` for the first-page
header/footer and `even` for even pages). `write` is atomic-undo. A header/footer
is just a range, so its id (`header:S:WHICH` / `footer:S:WHICH`) also works with
`replace`, `style apply`, and `format-paragraph`.

```bash
$ wordlive header read --section 1
{"anchor_id": "header:1:primary", "section": 1, "which": "primary", "text": "Confidential"}

$ wordlive header write --section 1 --text "ACME Corporation — Q2 Report"
{"ok": true, "anchor_id": "header:1:primary", "section": 1, "which": "primary"}
```

`--text` mode (`wordlive --text header read`) emits just the header text.
Failures: `2` section out of range, `3` Word busy.

## `exec` — `--script ops.json` or `--ops '{…}'` { #exec-script-opsjson }

```
wordlive exec (--script ops.json | --ops JSON | --ops -) [--doc DOC_NAME]
```

Apply a batch of operations under a single atomic-undo scope. This is the
most useful command for LLM tool-use: one round-trip per *intent*, not one
per *operation*.

Provide the batch one of three ways (exactly one of `--script` / `--ops` is
required):

- `--script ops.json` — read it from a file.
- `--ops '{"ops": [...]}'` — pass the JSON inline on the command line.
- `--ops -` — read the JSON from stdin (e.g. `… | wordlive exec --ops -`),
  which sidesteps the shell's command-line length limit and is the right
  choice for large payloads such as inline base64 images.

In every form a bare `[...]` array is accepted as shorthand for `{"ops": [...]}`,
and malformed JSON returns a clean error (exit `1`) rather than a traceback.

Script shape:

```json
{
  "label": "Update report",
  "ops": [
    {"op": "write_bookmark",    "name": "Address",        "text": "123 Main St"},
    {"op": "write_cc",          "name": "Signatory",      "text": "Jane Doe"},
    {"op": "insert_paragraph",  "anchor_id": "heading:8",  "text": "New risk paragraph.",
                                "where": "after",          "style": "Body Text"},
    {"op": "replace",           "anchor_id": "heading:3",  "text": "Updated section text"}
  ]
}
```

### Supported ops

| `op`                   | Required fields                            | Optional                          |
| ---------------------- | ------------------------------------------ | --------------------------------- |
| `write_bookmark`       | `name`, `text`                             | —                                 |
| `write_cc`             | `name`, `text`                             | —                                 |
| `insert_paragraph`     | `anchor_id`, `text`                        | `where` (`after`/`before`) or `before: true`, `style` |
| `append_paragraph`     | `text`                                     | `style`                           |
| `append`               | `text`                                     | —                                 |
| `prepend_paragraph`    | `text`                                     | `style`                           |
| `prepend`              | `text`                                     | —                                 |
| `insert_image`         | `anchor_id`, `wrap`, and one of `path` / `base64` | `where` or `before: true`, `width`, `height`, `alt_text`, `lock_aspect` |
| `replace`              | `anchor_id`, `text`                        | —                                 |
| `find_replace`         | `find`, `text`                             | `in`, `all`, `occurrence`         |
| `apply_style`          | `anchor_id`, `name`                        | —                                 |
| `format_paragraph`     | `anchor_id`                                | `alignment`, `left_indent`, `right_indent`, `first_line_indent`, `space_before`, `space_after`, `page_break_before` |
| `set_cell`             | `table`, `row`, `col`, `text`              | —                                 |
| `add_row`              | `table`                                    | `values`                          |
| `delete_row`           | `table`, `row`                             | —                                 |
| `create_table`         | `anchor_id`, `rows`, `cols`                | `style`, `data` (row-major 2-D), `header`, `where` or `before: true` |
| `delete_table`         | `table`                                    | —                                 |
| `insert_break`         | `anchor_id`                                | `kind` (`page`/`column`/`section_next`/`section_continuous`), `where` or `before: true` |
| `add_comment`          | `anchor_id`, `text`                        | `author`                          |
| `resolve_comment`      | `index`                                    | —                                 |
| `delete_comment`       | `index`                                    | —                                 |
| `apply_list`           | `anchor_id`                                | `type` (`bulleted`/`numbered`/`outline`), `continue` |
| `remove_list`          | `anchor_id`                                | —                                 |
| `restart_numbering`    | `anchor_id`                                | —                                 |
| `indent_list`          | `anchor_id`                                | —                                 |
| `outdent_list`         | `anchor_id`                                | —                                 |
| `write_header`         | `section`, `text`                          | `which` (`primary`/`first`/`even`) |
| `write_footer`         | `section`, `text`                          | `which`                           |

The `find_replace` op mirrors `wordlive replace --find …` — fuzzy whitespace
+ smart-quote match, optional `in` anchor to scope it, and either `all` or
`occurrence` to handle multi-match. Ambiguous-match failures surface in the
batch response's `failure.matches` so the LLM can rewrite the op and retry.

`insert_paragraph` mirrors the `insert` command: a new paragraph relative to
any anchor, with placement defaulting to `after` and an optional `style` that's
validated before the batch mutates anything. Placement accepts either the
verbose `"where": "before"|"after"` or the boolean `"before": true` — the latter
mirrors the command's `--before`/`--after` flags, so the same intent encodes the
same way whether you type it or batch it. (`insert_image` accepts both forms too.)

`append_paragraph` and `append` mirror the `append` command — they add a new
final paragraph (optional `style`, validated first) or inline text at the very
end of the document, with no anchor to resolve. Equivalent to an
`insert_paragraph` op targeting the `end` anchor. `prepend_paragraph` and
`prepend` are their start-of-document mirrors (the `start` anchor).

`insert_image` mirrors `insert-image`. Supply the image with either a `path`
(read from disk) or `base64` (inline data — the natural choice in a JSON op,
with no command-line length limit). `wrap` is required; the optional fields
match the command's flags. A bad image source surfaces as the batch's
`failure` with `type: "ImageSourceError"`.

`apply_style` and `format_paragraph` are the same as their dedicated CLI
verbs — the style must already exist in the document, indent and spacing
values are in points, alignment is one of `left`/`center`/`right`/`justify`.
`format_paragraph`'s `page_break_before` (a bool) forces or clears a
reflow-safe page break before the paragraph — the clean way to page-break a
style without a stray break character.

`set_cell`, `add_row`, and `delete_row` operate on tables by 1-based `table`
index. `set_cell` is shorthand for a `replace` on a `table:N:R:C` anchor;
`add_row`'s optional `values` is a JSON array matched to columns. All three
table ops join the same atomic-undo scope as the rest of the batch.

`create_table` builds a new table at a **position** `anchor_id` (`heading:`,
`para:`, `start`, `end`, `range:` — not a bare `table:N`); `delete_table`
removes one by 1-based `table` index. `create_table`'s `data` is a row-major 2-D
array validated against `rows`×`cols` before the batch mutates anything,
`style` defaults to `Table Grid`, and `header` bolds the first row. Because a
successful batch reports structure it created, the response carries an
`outputs` array — `[{"index": <op index>, "op": "create_table", "table": N,
"rows": R, "columns": C}]` — so a later op (or a follow-up call) can address the
new table by its reported index. Filling the whole grid through `data` in the
create op keeps it one atomic undo and avoids a `set_cell` storm.

`insert_break` mirrors the `insert-break` command — an explicit page, column,
or section break at any `anchor_id`. `kind` defaults to `page`; placement
accepts the same `where` / `before` forms as the other insert ops. For a
reflow-safe page break tied to a paragraph (rather than a one-off mark), use a
`format_paragraph` op with `page_break_before` instead.

`add_comment`, `resolve_comment`, and `delete_comment` mirror the `comment`
verbs — `add_comment` attaches a side-channel annotation to an `anchor_id`
without touching the text, while `resolve_comment` / `delete_comment` take a
1-based `index`. Since deletes re-index, ordering matters within a batch.

`apply_list`, `remove_list`, `restart_numbering`, `indent_list`, and
`outdent_list` mirror the `list` verbs — all take an `anchor_id`, and
`apply_list`'s optional `type` defaults to `bulleted`. `write_header` /
`write_footer` set a section's header/footer by 1-based `section` index, with
an optional `which` (`primary` / `first` / `even`, default `primary`) — handy
for stamping a client name or page footer across a generated document in the
same atomic-undo batch as the body edits.

### Recording the batch as tracked changes

Set `"tracked": true` at the top level of the script to flip Word's Track
Changes on for the whole batch and restore the prior setting when it finishes —
so the user sees every op as an accept/reject-able revision under one Ctrl-Z:

```json
{
  "label": "Suggest rewordings",
  "tracked": true,
  "ops": [
    {"op": "find_replace", "find": "utilise", "text": "use", "all": true}
  ]
}
```

### Behaviour on partial failure

If any op fails, the entire scope's `UndoRecord` still closes cleanly — but
operations *before* the failure have already been applied. The user can
roll the whole batch back with one Ctrl-Z. The response reports the failure
point precisely so an LLM can retry with a corrected payload:

```json
{
  "ok": false,
  "ops_run": 2,
  "label": "Update report",
  "failure": {
    "index": 2,
    "op": {"op": "insert_paragraph", "anchor_id": "heading:99", "text": "…"},
    "error": "heading not found: 'heading:99'",
    "type": "AnchorNotFoundError"
  }
}
```

The exit code reflects the **first** failed op: `2` for anchor-not-found,
`3` for Word-busy, etc. — so an LLM's retry policy can branch on it.

### Example invocation

```bash
$ wordlive exec --script ops.json
{"ok": true, "ops_run": 4, "label": "Update report"}
```

## `llm-help`

```
wordlive llm-help
```

Print the full **agent guide** — the bundled `SKILL.md` — to stdout: the anchor
model, every read/write verb, image insertion, the `exec` batch format, and the
exit-code taxonomy. `wordlive --help` points an agent straight here, so a model
can get everything it needs in one call without an install step.

Unlike every other command, the output is raw Markdown rather than JSON (and is
unaffected by `--json/--text`) — it's documentation, exactly like `--help`
itself, meant to read cleanly into a model's context. The YAML frontmatter that
fronts the installed skill is stripped. Offline: it never touches Word.

```bash
$ wordlive llm-help
# wordlive

`wordlive` drives a **running** Microsoft Word instance over COM (Windows only).
...
```

This is the same content [`install-skill`](#install-skill) writes to disk; reach
for `llm-help` when you just want it in context now, and `install-skill` when you
want coding tools to discover it on their own.

## `install-skill`

```
wordlive install-skill [--system] [--force]
```

Install the bundled **agent skill** (`SKILL.md`) so LLM coding tools can pick up
how to drive wordlive. By default it writes to the current project at
`./.agents/skills/wordlive/SKILL.md`; `--system` installs it for your user at
`~/.agents/skills/wordlive/SKILL.md`. This command is offline — it never touches
Word. It refuses to clobber an existing file unless you pass `--force`.

```bash
$ wordlive install-skill
{"ok": true, "scope": "local", "path": ".../.agents/skills/wordlive/SKILL.md", "bytes": 6172}

$ wordlive install-skill --system --force
{"ok": true, "scope": "system", "path": "/home/you/.agents/skills/wordlive/SKILL.md", "bytes": 6172}
```

The skill is a concise CLI reference — anchors, the read/write verbs, image
insertion, the `exec` batch format, and the exit-code contract — written for an
agent to load into context. Failures: `1` if the target can't be written.

## LLM tool-use example

A typical agent loop looks like:

```python
# 1. Discover what's addressable.
outline = json.loads(run(["wordlive", "outline"]))

# 2. Model picks an anchor and a new value, returns:
#    {"anchor_id": "heading:3", "text": "Revised context section"}

# 3. Apply.
result = run(["wordlive", "replace",
              "--anchor-id", anchor_id,
              "--text",      text])

# 4. Branch on exit code.
if result.returncode == 2:        # anchor not found — re-fetch outline
    ...
elif result.returncode == 3:      # Word busy — back off and retry
    ...
```

For multi-step intents, batch into one `exec --script ops.json` call instead.
See the [Cookbook](cookbook.md) for full worked examples.
