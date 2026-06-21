# CLI

`wordlive` ships a Click-based CLI designed for LLM tool-use loops: JSON in,
JSON out, deterministic exit codes, one structured object per invocation on
stdout. The CLI is a thin wrapper over the [Python API](python-api.md) — same
politeness, same atomic-undo.

## Global flags

```
wordlive [--json|--text] [--doc DOC_NAME] [--save-dir DIR]... [--image-dir DIR]... <subcommand> [args]
```

| Flag             | Default     | Purpose                                                |
| ---------------- | ----------- | ------------------------------------------------------ |
| `--json/--text`  | `--json`    | Output format. `--text` prints a per-command human form (indented outline tree, bare text for reads, one-line acks for writes); JSON stays the LLM-friendly default. |
| `--doc DOC_NAME` | active doc  | Target a specific open document by name (e.g. `Report.docx`). |
| `--save-dir DIR` | none (deny) | Allow `save`/`save-as`/`export-pdf` to write under `DIR` (repeatable; merges with `WORDLIVE_SAVE_DIRS`). **Default-deny:** with no directory configured, saving is off. |
| `--image-dir DIR`| none        | Restrict `insert-image --path` to files under `DIR` (repeatable; merges with `WORDLIVE_IMAGE_DIRS`). Non-local paths (UNC, URLs) are *always* rejected regardless. |
| `-h`, `--help`   | —           | Show help for the command or subgroup.                  |
| `-v`, `--version`| —           | Print `wordlive <version>` and exit.                    |
| `-A`, `--about`  | —           | Print the [about screen](#about) — banner, version, author, license, repo — and exit. |

## Exit codes

The CLI's error boundary classifies every [`WordliveError`](errors.md) into a
deterministic exit code so an LLM tool-use loop can branch on the failure
mode without parsing strings:

| Code | Meaning                | Source exception           |
| ---- | ---------------------- | -------------------------- |
| `0`  | OK                     | —                          |
| `1`  | Other / unclassified   | `WordliveError` (default), `DocumentNotFoundError`, `ImageSourceError`, `PathNotAllowedError` (save/image policy denial) |
| `2`  | Anchor or style missing | `AnchorNotFoundError` / `StyleNotFoundError` (also used for zero-match `find`/`replace --find`) |
| `3`  | Word busy / modal      | `WordBusyError` (retryable) |
| `4`  | Word not running       | `WordNotRunningError`      |
| `5`  | Ambiguous match        | `AmbiguousMatchError` (multiple `find` hits without `--all`/`--occurrence`) |
| `6`  | Excel not available    | `ExcelNotAvailableError` (`insert-chart` needs Excel installed) |

See the [Errors page](errors.md) for the full exception taxonomy and
retry guidance.

---

## `status`

```
wordlive status
```

List all open documents and mark which is active. Each entry carries a `name`
(always non-empty — `Document1` for a document never saved), the on-disk `path`
(empty until the document is saved), a `saved` flag, and `is_active`.

```bash
$ wordlive status
[{"name": "Report.docx", "path": "C:\\Users\\me\\Report.docx", "saved": true,  "is_active": true},
 {"name": "Document2",   "path": "",                           "saved": false, "is_active": false}]
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
items alike — each with a `para:N` anchor, its outline `level`, the applied
Word `style` name, an `is_heading` flag, character `start`/`end` offsets, and
its text. `outline --all` is an alias. Use `style` to mirror an existing
document's formatting: `level` is `10` for *every* non-heading paragraph, so
only the style name tells a `List Number` item apart from `Normal` body text.

```bash
$ wordlive paragraphs
[{"index": 1, "anchor_id": "para:1", "level": 1,  "style": "Heading 1",   "is_heading": true,  "start": 0,  "end": 13, "text": "Introduction"},
 {"index": 2, "anchor_id": "para:2", "level": 10, "style": "List Number", "is_heading": false, "start": 13, "end": 29, "text": "First item."},
 {"index": 3, "anchor_id": "para:3", "level": 2,  "style": "Heading 2",   "is_heading": true,  "start": 29, "end": 35, "text": "Risks"}]
```

`para:N` shares its index space with `heading:N` — paragraph 1 is both
`para:1` and (because it's a heading) `heading:1`. The emitted offsets feed a
[`range:START-END`](concepts.md) target for an offset-precise, mid-paragraph
insertion via `replace`.

## `read bookmark NAME` / `read bookmark --list`

```
wordlive read bookmark NAME                    [--doc DOC_NAME]
wordlive read bookmark --list [--include-hidden] [--doc DOC_NAME]
```

Read the text of a bookmark, or with `--list` emit every bookmark **name** in
document order (`--include-hidden` also returns Word's internal bookmarks —
`_Toc…`, `_Ref…`).

```bash
$ wordlive read bookmark Address
{"text": "123 Main St"}

$ wordlive read bookmark --list
["Address", "Intro"]
```

Failures: `2` if the named bookmark doesn't exist. Pass exactly one of `NAME`
or `--list`.

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

## `read markdown` / `read html`

```
wordlive read markdown [--within ANCHOR]   [--doc DOC_NAME]
wordlive read html     [--within ANCHOR]   [--doc DOC_NAME]
```

Serialise the whole document — or one anchor's range — to clean **Markdown**
(or an **HTML** fragment). The read mirror of `insert-markdown`: headings,
bullet / numbered lists (nested), `**bold**` / `*italic*`, GFM pipe tables,
inline images as `![alt](image:N)`, and hyperlinks as `[text](url)`. Export is
**lossy by design** (underline, colours, and merged table cells don't survive),
so it round-trips the constrained subset import speaks and reads the rest richer.

`--within` scopes to an anchor's **literal range** — a `range:START-END` (e.g.
from `find`), or any anchor id. A `heading:N` covers only the heading line, not
its section body (use `read between` or a `range:` for "the section under X").

```bash
$ wordlive read markdown --within heading:3
{"markdown": "### Pricing\n\nOur tiers are **flexible** …"}

$ wordlive --text read markdown > document.md   # pipe the raw Markdown out
```

`--text` emits the raw Markdown / HTML (no JSON envelope) — handy for piping
into a file or an LLM prompt. Failures: `2` anchor not found.

## `read digest [--budget N] [--depth D]`

```
wordlive read digest [--budget 6000] [--depth D]   [--doc DOC_NAME]
```

A **token-budgeted, structure-aware** read of the whole document — load a large
document into context cheaply while **every anchor stays addressable**. Headings
are verbatim (each tagged with its `<!-- heading:N -->` anchor — the navigation
spine), tables become one-line shape stubs (`> table:N — R rows × C cols: …`),
and body text is sampled to fit `--budget` (an approximate token count,
~4 chars/token), weighted so shallower sections keep more than deep ones.
Overflow is elided to markers that still name the `para:` range, so you can drill
into any region with `read markdown --within …`. `--depth D` caps how deep a
section keeps body (deeper sections collapse to a marker).

```bash
$ wordlive --text read digest --budget 4000
# Q3 Report  <!-- heading:1 -->

The quarter closed ahead of plan …

…(para:5–para:18, 1240 words elided)…

> table:2 — 9 rows × 4 cols: Quarter, Region, Revenue, Growth
…
```

A pure read. For the full text of any region, use `read markdown --within`.

## `read between --start ID --end ID [--inclusive]`

```
wordlive read between --start ID --end ID [--inclusive]   [--doc DOC_NAME]
```

Read the content **between two anchors** — the headline use is two `heading:N`
ids (the block between two headings), but any anchors work. By default the span
is *strictly between* them (e.g. the body between two headings, excluding both
heading lines); `--inclusive` covers both bounding paragraphs. Returns the
spanning `range:START-END` id plus its text.

```bash
$ wordlive read between --start heading:1 --end heading:3
{"start": "heading:1", "end": "heading:3", "inclusive": false,
 "anchor_id": "range:13-352", "text": "Body of the first section …"}
```

The returned offsets are live — use them before further edits shift the
document. Failures: `1` if `end` begins before `start`; `2` anchor not found.

## `read nearest-heading --anchor-id ID [--direction before|after]`

```
wordlive read nearest-heading --anchor-id ID [--direction before|after]   [--doc DOC_NAME]
```

Find the heading nearest to a position. `--direction before` (default) returns
the enclosing / preceding heading — the section the anchor sits in; `after`
returns the next heading past it. Emits the heading row (`{level, text,
anchor_id}`) under `heading`, or `null` when there is none in that direction.

```bash
$ wordlive read nearest-heading --anchor-id para:42 --direction before
{"anchor_id": "para:42", "direction": "before",
 "heading": {"level": 2, "text": "Risks", "anchor_id": "heading:17"}}
```

Failures: `1` bad `--direction`; `2` anchor not found.

## `read format --anchor-id ID`

```
wordlive read format --anchor-id ID [--doc DOC_NAME]
```

Read the **effective formatting** at an anchor — the read mirror of
`format-paragraph` / `format-run`. Returns `{anchor_id, style, paragraph,
font}`, where `style` is the applied paragraph style's name and `paragraph` /
`font` each map a field name to `{value, style, override}`: `value` is the
effective value, `style` is the value the applied style contributes as its
baseline, and `override` is `true` when `value` differs from `style` (a direct
override Word would show with the style cleared). `font` also carries a `mixed`
key — the font fields that read `wdUndefined` because they vary across the
range's runs (their `value` is `null` and they're never flagged as overrides).

Lengths are in **points** (floats); `color` is `#RRGGBB` or `"auto"`;
`alignment` is `left`/`center`/`right`/`justify`; `line_spacing` is
`single`/`1.5`/`double`, `"1.15"` (a multiple), `"14pt"` (exactly), or
`"at_least:14pt"`. Paragraph fields: `alignment`, `left_indent`, `right_indent`,
`first_line_indent`, `space_before`, `space_after`, `line_spacing`,
`page_break_before`, `keep_together`, `keep_with_next`, `widow_control`. Font
fields: `name`, `size`, `bold`, `italic`, `underline`, `strikethrough`, `color`,
`subscript`, `superscript`, `small_caps`, `all_caps`, `spacing`. Non-mutating.

```bash
$ wordlive read format --anchor-id heading:3
{"anchor_id": "heading:3",
 "style": "Heading 2",
 "paragraph": {"alignment": {"value": "center", "style": "left", "override": true},
               "space_before": {"value": 6.0, "style": 12.0, "override": true}},
 "font": {"name": {"value": "Calibri Light", "style": "Calibri Light", "override": false},
          "bold": {"value": true, "style": true, "override": false},
          "size": {"value": null, "style": 13.0, "override": false}},
 "mixed": ["size"]}
```

Diff `override: true` to see what a paragraph carries beyond its style — the
input for `regularize`, which writes those back to the style's own value.

Failures: `2` anchor not found, `3` Word busy, `4` Word not running.

## `write bookmark NAME (--text "…" | --create --anchor-id ID)`

```
wordlive write bookmark NAME --text "..."             [--doc DOC_NAME]   # set existing
wordlive write bookmark NAME --create --anchor-id ID  [--doc DOC_NAME]   # create new
```

Two modes, both in a single atomic-undo scope:

- `--text "…"` replaces an **existing** bookmark's text. The bookmark is
  preserved — wordlive re-adds it covering the new content after the Word
  `Range.Text` assignment (which would otherwise delete it).
- `--create --anchor-id ID` **creates** a new bookmark `NAME` over an anchor's
  range (e.g. `heading:2`, `range:120-140`) — the prerequisite for internal
  links (`link --bookmark NAME`) and cross-references (`cross-ref --target
  bookmark:NAME`). `NAME` must start with a letter and contain only letters,
  digits, and underscores.

```bash
$ wordlive write bookmark Address --text "456 Elm St"
{"ok": true, "anchor": {"kind": "bookmark", "name": "Address"}}

$ wordlive write bookmark Intro --create --anchor-id heading:1
{"ok": true, "bookmark": "Intro", "anchor_id": "heading:1", "created": true}
```

(Creating a bookmark was previously `bookmark add NAME --anchor-id ID`, now a
hidden deprecated alias.)

Failures: `2` anchor not found, `3` Word busy.

## `pin ANCHOR_ID [--name SLUG]`

```
wordlive pin ANCHOR_ID [--name SLUG] [--doc DOC_NAME]
```

Plant a **durable handle** on any anchor and print its `pin:CODE` id. A pin is a
Word-hidden bookmark (`_wl_<code>`) over the anchor's range; Word keeps it pinned
to the same content across the inserts/deletes that renumber positional `para:N`
/ `heading:N` ids, so it is the escape hatch when you need a stable address.
Resolve it later with `--anchor-id pin:CODE`. Omit `--name` for a random code, or
pass a readable slug (`--name budget-intro`, lowercase words joined by hyphens).
If the pinned content is later deleted the handle vanishes (the next resolve is
exit `2`). One atomic-undo scope.

```bash
$ wordlive pin heading:4 --name methods
{"anchor_id": "pin:methods", "pin": "pin:methods", "target": "heading:4"}
```

Failures: `2` anchor not found, `3` Word busy.

## `pin-outline [--levels LO HI]`

```
wordlive pin-outline [--levels LO HI] [--doc DOC_NAME]
```

Pin every heading at once and print the `{heading:N → pin:CODE}` map — a durable
navigation scaffold to set up before a batch of structural edits. Idempotent: a
heading already carrying a handle reuses it (keyed by range start), so re-running
returns the same map. `--levels LO HI` restricts to an inclusive heading-level
band (e.g. `--levels 1 2`). One atomic-undo scope.

```bash
$ wordlive pin-outline
{"heading:1": "pin:a3f9c2", "heading:4": "pin:b6223a"}
```

Failures: `3` Word busy.

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

## `insert --anchor-id ID (--text "…" | --runs JSON)`

```
wordlive insert --anchor-id ID (--text "..." | --runs JSON) [--before | --after] [--style "Body Text"] [--doc DOC_NAME]
```

Insert a new paragraph relative to **any** anchor — addressed the same way
every other command addresses things, with `--anchor-id` (`heading:N`,
`para:N`, `bookmark:NAME`, a cell, a range). `--after` (the default) lands the
new paragraph just below the anchor; `--before` lands it just above. `--after`
works even when the anchor is the document's last paragraph — the new paragraph
is appended before the final mark — so you can build a document top-down from a
single empty paragraph.

Give exactly one of `--text` (a literal string — no markup) or `--runs` (a JSON
array of inline-formatted spans, or `-` to read it from stdin). Each run is
`{text, bold?, italic?, underline?, style?}`, so a bold lead-in is one op:

```bash
$ wordlive insert --anchor-id heading:8 --text "New risk identified."
{"ok": true, "anchor_id": "heading:8", "where": "after", "style": null}

$ wordlive insert --anchor-id para:3 --text "Section preamble." --before
{"ok": true, "anchor_id": "para:3", "where": "before", "style": null}

$ wordlive insert --anchor-id end --runs '[{"text":"Bold lead","bold":true},{"text":" — rest"}]'
{"ok": true, "anchor_id": "end", "where": "after", "style": null}
```

`--style` is optional; if given it must be a Word style name that exists in
the document — the style is validated before the paragraph is inserted, so a
typo never partially mutates the document. Use `wordlive style list` to see
the available names. Failures: `2` anchor not found or style not found, `3`
Word busy.

To insert text *inside* a paragraph at a precise offset rather than as a new
paragraph, target a collapsed range instead — `replace --anchor-id
range:120-120 --text "…"` — using offsets from `paragraphs` or `find`. To insert
several styled paragraphs at once, use `insert-block`.

## `insert-block --anchor-id ID --items JSON`

```
wordlive insert-block --anchor-id ID --items JSON [--before | --after] [--doc DOC_NAME]
```

Insert a **contiguous run of styled paragraphs** at an anchor in one op, in
natural reading order — the multi-paragraph counterpart to `insert`. Use it to
drop a whole styled section (a feature list, a heading plus its body) without a
reverse-ordered storm of `insert` calls dodging positional-anchor renumbering.

`--items` is a JSON array (or `-` for stdin). Each item is one paragraph, given
as either a plain string or an object `{text | runs, style?}`:

- `text` carries a tiny inline **markdown**: `**bold**`, `*italic*`,
  `***both***` (escape a literal asterisk as `\*`).
- `runs` is the structured form — `[{text, bold?, italic?, underline?,
  style?}]` — for unambiguous control or a per-run character style.
- `style` names the paragraph style for that item.

It reports the spanning `range:START-END` of the inserted block, so you can act
on the whole run next — e.g. bullet the section you just inserted:

```bash
$ wordlive insert-block --anchor-id heading:1 --items \
    '[{"text":"**Politeness** first.","style":"List Bullet"},
      {"runs":[{"text":"Atomic undo","bold":true},{"text":" — one Ctrl-Z."}],"style":"List Bullet"},
      "Plain third bullet."]'
{"ok": true, "anchor_id": "range:412-470", "paragraphs": 3, "where": "after"}

$ wordlive list apply --anchor-id range:412-470 --type bulleted
```

Styles are validated before anything is inserted, so a bad name fails the whole
block cleanly. Failures: `2` anchor not found or style not found, `3` Word busy.

## `insert-section --anchor-id ID --heading TEXT --body JSON`

```
wordlive insert-section --anchor-id ID --heading TEXT --body JSON [--level N] [--before | --after] [--doc DOC_NAME]
```

The opinionated common case over `insert-block`: a `Heading {level}` paragraph
followed by its body, in reading order and one op. `--heading` carries the same
inline markdown an item's `text` does; `--body` is the same JSON items shape
`insert-block` takes (or `-` for stdin). `--level` is 1–9 (default `1`). Reports
the section's spanning `range:START-END`.

```bash
$ wordlive insert-section --anchor-id end --heading "Results" --level 2 --body \
    '["We observed a **20%** lift.", {"text":"Caveats apply.","style":"Body Text"}]'
{"ok": true, "anchor_id": "range:512-560", "where": "after"}
```

## `insert-markdown --anchor-id ID --markdown TEXT`

```
wordlive insert-markdown --anchor-id ID --markdown TEXT [--before | --after] [--doc DOC_NAME]
```

Drop a chunk of **constrained Markdown** as real Word structure. The dialect is a
documented subset, not CommonMark:

- `#` / `##` / `###` → `Heading 1` / `Heading 2` / `Heading 3`.
- `-` / `*` → a bulleted list; `1.` → a numbered list (each list is numbered
  1..N over its own span).
- a blank line separates paragraphs; consecutive plain lines join into one
  `Normal` paragraph.
- inline `**bold**` / `*italic*` / `***both***` are honoured.

Out of scope in v1: code fences, nested lists, block quotes, tables — anything
unrecognised becomes literal paragraph text. Multi-line input is easiest from
stdin with `--markdown -`. Reports the spanning `range:START-END`.

```bash
$ printf '# Plan\n\nKick-off notes.\n\n- scope it\n- staff it\n' | \
    wordlive insert-markdown --anchor-id end --markdown -
{"ok": true, "anchor_id": "range:600-640", "where": "after"}
```

## `replace-section --anchor-id heading:N (--body JSON | --markdown TEXT)`

```
wordlive replace-section --anchor-id heading:N (--body JSON | --markdown TEXT) [--doc DOC_NAME]
```

Rewrite a heading's body — everything from after the heading up to the next
same-or-higher heading — while **keeping the heading paragraph**. The "rewrite
section X" workflow. Give exactly one of `--body` (the `insert-block` items shape)
or `--markdown` (the constrained subset above); either accepts `-` for stdin.
Reports the new body's spanning `range:START-END`. Needs a `heading:N` anchor
(exit 1 otherwise).

```bash
$ wordlive replace-section --anchor-id heading:3 --markdown "Updated findings.

- point one
- point two"
{"ok": true, "anchor_id": "range:300-340"}
```

## `delete-paragraph --anchor-id ID`

```
wordlive delete-paragraph --anchor-id ID [--doc DOC_NAME]
```

Delete the paragraph(s) at an anchor — **text and the trailing mark** — so the
surrounding text closes up with no empty line left behind (unlike
`replace --text ""`, which empties the paragraph but keeps it). Handy for a
stray leading empty `para:1`. Deleting the document's last paragraph clears it
but keeps Word's mandatory final mark.

```bash
$ wordlive delete-paragraph --anchor-id para:1
{"ok": true, "anchor_id": "para:1", "deleted": true}
```

Failures: `2` anchor not found, `3` Word busy.

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

## `insert-field --anchor-id ID --kind KIND [--text CODE] [--before | --after]`

```
wordlive insert-field --anchor-id ID
    --kind page|numpages|date|time|filename|author|title|field
    [--text "FIELD CODE"] [--before | --after] [--doc DOC_NAME]
```

Insert a self-updating Word field — a value Word keeps current rather than literal
text. `page` / `numpages` are the page number and total page count (combine for
"Page X of Y"); `date` / `time` stamp the current date/time; `filename` /
`author` / `title` pull document metadata. `--kind field` is the escape hatch:
pass the raw field code via `--text` (e.g. `--text "REF myBookmark \h"`).

Page numbers belong in a footer or header — the anchor id `footer:1:primary`
works here like any other:

```bash
$ wordlive insert-field --anchor-id footer:1:primary --kind page
{"ok": true, "anchor_id": "footer:1:primary",
 "anchor": {"kind": "footer", "name": "footer:1:primary"},
 "applied": {"kind": "page", "text": null, "where": "after"}}
```

Newly inserted fields render once; run `update-fields` (or `snapshot`, which
repaginates) to refresh them after later edits. Failures: `1` `--kind field`
without `--text`, or bad input; `2` anchor not found; `3` Word busy.

## `update-fields`

```
wordlive update-fields [--doc DOC_NAME]
```

Recompute the document's fields (page numbers, cross-references, dates, a TOC) —
the "make the numbers right again" verb after edits. Atomic-undo; scope is the
main text story.

```bash
$ wordlive update-fields
{"ok": true, "updated": true}
```

Failures: `3` Word busy, `4` Word not running.

## `insert-footnote --anchor-id ID --text "…"` / `insert-endnote …`

```
wordlive insert-footnote --anchor-id ID --text "NOTE BODY" [--before | --after] [--doc DOC_NAME]
wordlive insert-endnote  --anchor-id ID --text "NOTE BODY" [--before | --after] [--doc DOC_NAME]
```

Attach a footnote (bottom of the page) or endnote (end of the document) to the
anchor's range; Word auto-numbers the reference mark. Reports the new note's id
so you can edit or remove it later (`footnote:N` / `endnote:N` resolve through
`--anchor-id` everywhere).

```bash
$ wordlive insert-footnote --anchor-id range:120-140 --text "See appendix B."
{"ok": true, "anchor_id": "range:120-140", "footnote": 1, "note_id": "footnote:1"}
```

Failures: `1` bad input; `2` anchor not found; `3` Word busy.

## `insert-toc [--anchor-id ID] [--levels A-B] [--no-heading-styles] [--no-hyperlinks]`

```
wordlive insert-toc [--anchor-id ID] [--levels 1-3]
    [--heading-styles | --no-heading-styles] [--hyperlinks | --no-hyperlinks]
    [--before | --after] [--doc DOC_NAME]
```

Insert a table of contents built from the document's headings. `--anchor-id`
defaults to `start` (the top of the document); `--levels` is the inclusive
heading-level span (`1-3` = Heading 1–3). Entries link to their headings unless
`--no-hyperlinks`.

```bash
$ wordlive insert-toc --levels 1-2
{"ok": true, "anchor_id": "start",
 "applied": {"levels": [1, 2], "use_heading_styles": true, "hyperlinks": true, "where": "after"}}
```

Page numbers populate only after repagination — run `update-fields` (or
`snapshot`) before reading them. Failures: `1` bad input (e.g. malformed
`--levels`); `2` anchor not found; `3` Word busy.

## `table-of-figures [--anchor-id ID] [--label L] [--no-label] [--no-hyperlinks]`

```
wordlive table-of-figures [--anchor-id ID] [--label Figure] [--label | --no-label]
    [--hyperlinks | --no-hyperlinks] [--before | --after] [--doc DOC_NAME]
```

Insert a table of figures built from the document's captions — every caption of
one `--label` (`Figure`/`Table`/`Equation`/a custom string) with its page number.
`--anchor-id` defaults to `start`; `--no-label` drops the "Figure"/"Table" prefix
from each entry, and entries link to their captions unless `--no-hyperlinks`.

```bash
$ wordlive table-of-figures --label Figure
{"ok": true, "anchor_id": "start",
 "applied": {"label": "Figure", "include_label": true, "hyperlinks": true, "where": "after"}}
```

Page numbers populate only after repagination — run `update-fields` (or
`snapshot`) before reading them. Failures: `1` bad input; `2` anchor not found;
`3` Word busy.

## `mark-index-entry --anchor-id ID --entry 'topic'`

```
wordlive mark-index-entry --anchor-id ID --entry "topic"
    [--cross-reference X] [--bold] [--italic] [--doc DOC_NAME]
```

Mark the anchor's range as a back-of-book index entry (an `XE` field) — the first
of the index's two steps (build it with `insert-index`). `--entry` is the headword;
use `"main:sub"` to nest a subentry. `--cross-reference X` makes the entry a "see
X" pointer instead of a page number, and `--bold`/`--italic` style the page number.

```bash
$ wordlive mark-index-entry --anchor-id range:120-140 --entry "risk:market"
{"ok": true, "anchor_id": "range:120-140", "entry": "risk:market"}
```

Failures: `1` bad input; `2` anchor not found; `3` Word busy.

## `insert-index [--anchor-id ID] [--columns N] [--run-in] [--right-align-page-numbers]`

```
wordlive insert-index [--anchor-id ID] [--columns 2] [--run-in]
    [--right-align-page-numbers] [--before | --after] [--doc DOC_NAME]
```

Build a back-of-book index from the entries marked with `mark-index-entry`.
`--anchor-id` defaults to `end` (the back of the document); `--columns` is the
column count (default `2`), `--run-in` lays subentries inline rather than each on
its own line, and `--right-align-page-numbers` flushes the page numbers right with
a tab leader.

```bash
$ wordlive insert-index --columns 2
{"ok": true, "anchor_id": "end",
 "applied": {"columns": 2, "run_in": false, "right_align_page_numbers": false, "where": "after"}}
```

Page numbers populate only after repagination — run `update-fields` (or
`snapshot`) before reading them. Failures: `1` bad input; `2` anchor not found;
`3` Word busy.

## `bibliography-style --style STYLE`

```
wordlive bibliography-style --style STYLE [--doc DOC_NAME]
```

Set the document's bibliography style — the citation scheme Word renders citations
and the bibliography in (`APA`, `MLA`, `Chicago`, `IEEE`, `Turabian`, …). Which
names are accepted is **build-dependent** (Word ships a fixed set of style XSLTs);
an unsupported value fails. Atomic-undo.

```bash
$ wordlive bibliography-style --style APA
{"ok": true, "style": "APA"}
```

Failures: `1` unsupported style (an `OpError`), `3` Word busy.

## `add-source --type TYPE [...]`

```
wordlive add-source --type book
    [--tag T --author "Last, First" (repeatable) --title "..." --year YYYY
     --publisher "..." --city "..." --journal-name "..." --volume V --issue I
     --pages P --url URL --edition E --doi D | --xml RAW] [--doc DOC_NAME]
```

Register a citation **source** in the document's master/source list — the first of
the citation workflow's two steps (cite it with `insert-citation`, then build the
list with `insert-bibliography`). `--type` is one of `book`, `book_section`,
`journal_article`, `article_in_periodical`, `conference_proceedings`, `report`,
`web_site`, `document_from_site`, `electronic_source`, `art`, `sound_recording`,
`performance`, `film`, `interview`, `patent`, `case`, `misc` (plus aliases).
`--author` is `"Last, First"` (or `"First Last"`, and repeatable for several
authors); `--tag` is the short handle later passed to `insert-citation` and
**auto-derives** from the first author's surname + year when omitted. `--xml`
is the escape hatch: a raw `<b:Source>` OOXML element (which must carry its own
`<b:Tag>`), bypassing the typed flags.

```bash
$ wordlive add-source --type book --author "Smith, Jane" --title "On Risk" --year 2020
{"ok": true, "source": "Smith2020"}
```

Failures: `1` bad input (e.g. `--xml` without a `<b:Tag>`); `3` Word busy.

## `insert-citation --anchor-id ID --tag TAG [...]`

```
wordlive insert-citation --anchor-id ID --tag TAG
    [--pages P] [--prefix "..."] [--suffix "..."] [--volume V]
    [--suppress-author] [--suppress-year] [--suppress-title]
    [--locale 1033] [--before | --after] [--doc DOC_NAME]
```

Insert an in-text citation to a registered source at the anchor — a Word `CITATION`
field that renders per the active `bibliography-style` (e.g. `(Smith 2020, 15)`).
`--tag` is the source's tag (from `add-source`); `--pages` adds a page locator,
and the `--suppress-*` flags hide the author / year / title in this one citation.
A citation to a tag with no matching source still inserts, but renders
**"Invalid source specified."** — so register the source first.

```bash
$ wordlive insert-citation --anchor-id range:120-140 --tag Smith2020 --pages 15
{"ok": true, "citation": "Smith2020"}
```

Failures: `1` bad input; `2` anchor not found; `3` Word busy.

## `insert-bibliography [--anchor-id end] [--before | --after]`

```
wordlive insert-bibliography [--anchor-id end] [--before | --after] [--doc DOC_NAME]
```

Insert the bibliography / works-cited block — a Word `BIBLIOGRAPHY` field listing
every cited source, formatted per the active `bibliography-style`. `--anchor-id`
defaults to `end` (the back of the document). Its entries and page-dependent
formatting populate only after repagination — run `update-fields` (or `snapshot`)
before reading them.

```bash
$ wordlive insert-bibliography
{"ok": true, "bibliography": true}
```

Failures: `1` bad input; `2` anchor not found; `3` Word busy.

## `mark-citation --anchor-id ID --long "full citation" [...]`

```
wordlive mark-citation --anchor-id ID --long "FULL CITATION"
    [--short ABBREV] [--category cases] [--before | --after] [--doc DOC_NAME]
```

Mark the anchor's range as a **table-of-authorities** entry (a Word `TA` field) —
the first of the table-of-authorities' two steps (build it with
`table-of-authorities`). `--long` is the full citation as it appears in the table;
`--short` is the abbreviated form for later same-authority references (defaults to
`--long`). `--category` groups the entry — `cases` / `statutes` / `other` /
`rules` / `treatises` / `regulations` / `constitutional`, or an int `1`–`16`.

```bash
$ wordlive mark-citation --anchor-id range:120-140 \
    --long "Brown v. Board, 347 U.S. 483 (1954)" --short "Brown" --category cases
{"ok": true, "anchor_id": "range:120-140"}
```

Failures: `1` bad input; `2` anchor not found; `3` Word busy.

## `table-of-authorities [--anchor-id end] [--category C] [--no-passim] [--no-keep-formatting]`

```
wordlive table-of-authorities [--anchor-id end] [--category all]
    [--passim | --no-passim] [--keep-formatting | --no-keep-formatting]
    [--entry-separator "..."] [--page-range-separator "..."]
    [--before | --after] [--doc DOC_NAME]
```

Build a table of authorities from the entries marked with `mark-citation`.
`--anchor-id` defaults to `end`. `--category` is `all` (every category) or a single
named category / int (as in `mark-citation`). `--no-passim` lists every page even
when an authority is cited on five or more (the default `passim` collapses those to
"passim"); `--no-keep-formatting` drops the marked entries' own character
formatting. `--entry-separator` / `--page-range-separator` override the punctuation
between an entry and its pages, and within a page range.

```bash
$ wordlive table-of-authorities --category cases
{"ok": true, "table_of_authorities": true}
```

Page numbers populate only after repagination — run `update-fields` (or
`snapshot`) before reading them. Failures: `1` bad input; `2` anchor not found;
`3` Word busy.

## `theme` · `list-themes`

```
wordlive theme [--doc DOC_NAME]
wordlive list-themes [--doc DOC_NAME]
```

Read the document's theme — the document-wide brand primitive. `theme` reports the
12 theme colours (as `#RRGGBB`) and the major/minor fonts; `list-themes` reports the
built-in themes, colour schemes, and font schemes Office ships (the names the
`apply-theme` / `set-theme-*` commands accept). Both are non-mutating.

```bash
$ wordlive --json theme
{"colors": {"text1": "#000000", "accent1": "#156082", ...}, "major_font": "Aptos Display", "minor_font": "Aptos"}
```

## `apply-theme --theme NAME`

```
wordlive apply-theme --theme NAME [--doc DOC_NAME]
```

Apply a whole document theme — colours, fonts, and effects in one step. `--theme`
is a built-in name (e.g. `Facet`, `Ion` — see `list-themes`) or a `.thmx` file path
(a brand file). Wrap-free atomic undo. Override individual brand colours/fonts
afterwards with `set-theme-colors` / `set-theme-fonts`.

```bash
$ wordlive apply-theme --theme Facet
{"ok": true, "applied": {"theme": "Facet"}}
```

Failures: `1` bad input (incl. an unknown theme name); `3` Word busy.

## `set-theme-colors [--scheme S] [--accent1 C] ...`

```
wordlive set-theme-colors [--scheme NAME] [--text1 C] [--background1 C]
    [--text2 C] [--background2 C] [--accent1 C] ... [--accent6 C]
    [--hyperlink C] [--followed-hyperlink C] [--doc DOC_NAME]
```

Set the theme's colour scheme and/or individual brand colours. `--scheme` loads a
named built-in colour scheme (e.g. `Blue`) or a Theme-Colors `.xml` path; the
per-colour flags override individual slots. A colour value is a name (`navy`) or a
hex string (`#1A73E8`).

```bash
$ wordlive set-theme-colors --accent1 "#1A73E8" --accent2 "#34A853"
{"ok": true, "colors": {"accent1": "#1A73E8", "accent2": "#34A853", ...}, "applied": {...}}
```

## `set-theme-fonts [--scheme S] [--major F] [--minor F]`

```
wordlive set-theme-fonts [--scheme NAME] [--major FONT] [--minor FONT] [--doc DOC_NAME]
```

Set the theme's fonts. `--scheme` loads a named built-in font scheme (e.g.
`Garamond`) or a Theme-Fonts `.xml` path; `--major` / `--minor` override the heading
/ body font names.

```bash
$ wordlive set-theme-fonts --major Arial --minor Calibri
{"ok": true, "major_font": "Arial", "minor_font": "Calibri", "applied": {...}}
```

Failures for `set-theme-*`: `1` bad input (incl. an unknown scheme/colour); `3` Word busy.

## `create-content-control --anchor-id ID [--kind K] [--title T] [--tag T] [--item ...]`

```
wordlive create-content-control --anchor-id ID [--kind rich_text]
    [--title T] [--tag T] [--item 'Text' --item 'Label=Value' ...]
    [--where wrap|before|after] [--lock-contents] [--lock-control] [--doc DOC_NAME]
```

Create a content control — the structured-document fill-in field (the read/write
side is `read cc` / `write cc`). `--where wrap` (the default) surrounds the
anchor's existing range — e.g. a `range:START-END` from `find` — while `before` /
`after` insert a fresh empty control. `--kind` is one of `rich_text` (default),
`text`, `picture`, `combo_box`, `dropdown`, `date`, `checkbox`, `building_block`,
`group`, `repeating_section`. A `--title` (falling back to `--tag`) names the
control so it's addressable later as `cc:TITLE`. `--item` (repeatable, combo_box /
dropdown only) adds a choice — `'Text'` or `'Label=Value'`. `--lock-contents`
stops edits to the value; `--lock-control` stops deletion.

```bash
$ wordlive create-content-control --anchor-id range:120-140 --kind dropdown \
    --title Status --item "Open" --item "Done=closed"
{"ok": true, "content_control": "Status", "cc_anchor_id": "cc:Status"}
```

Failures: `1` bad input; `2` anchor not found; `3` Word busy.

## `set-cc-properties --anchor-id cc:NAME [--title T] [--tag T] [--lock-contents] [--lock-control]`

```
wordlive set-cc-properties --anchor-id cc:NAME [--title T] [--tag T]
    [--lock-contents | --no-lock-contents] [--lock-control | --no-lock-control] [--doc DOC_NAME]
```

Re-set an existing content control's metadata in place — no delete + reinsert.
Pass at least one option; `""` clears `--title` / `--tag` while omitting one
leaves it untouched. `--lock-contents` / `--no-lock-contents` toggle whether the
value is editable; `--lock-control` / `--no-lock-control` whether the control can
be deleted. A title (or tag, when untitled) rename changes the control's
`cc:NAME` anchor id.

```bash
$ wordlive set-cc-properties --anchor-id cc:Status --title Stage --lock-contents
{"ok": true, "anchor_id": "cc:Status", "applied": {"title": "Stage", "lock_contents": true}}
```

Failures: `1` bad input / wrong-kind anchor; `2` anchor not found; `3` Word busy.

## `set-cc-items --anchor-id cc:NAME --item ... [--item ...]`

```
wordlive set-cc-items --anchor-id cc:NAME --item 'Text' [--item 'Label=Value' ...] [--doc DOC_NAME]
```

Replace a `combo_box` / `dropdown` control's choice list in place — the new
`--item` set *replaces* the existing entries (it does not append). Each `--item`
is `'Text'` or `'Label=Value'`. Only valid on a combo_box / dropdown control.

```bash
$ wordlive set-cc-items --anchor-id cc:Status --item "Open" --item "Done=closed"
{"ok": true, "anchor_id": "cc:Status", "applied": {"items": ["Open", {"text": "Done", "value": "closed"}]}}
```

Failures: `1` bad input / not a list control; `2` anchor not found; `3` Word busy.

## `footnotes` / `endnotes`

```
wordlive footnotes [--doc DOC_NAME]
wordlive endnotes  [--doc DOC_NAME]
```

List the document's footnotes / endnotes — each note's `footnote:N` / `endnote:N`
id, number, body text, and the `para:N` its reference mark sits in.

```bash
$ wordlive footnotes
[{"index": 1, "anchor_id": "footnote:1", "marker": "1",
  "text": "See appendix B.", "para": "para:6"}]
```

Failures: `3` Word busy, `4` Word not running.

## `revisions`

```
wordlive revisions [--doc DOC_NAME]
```

List the document's tracked changes — the structured counterpart to
`snapshot --markup all`. Each revision reports its `type` (`insert` / `delete` /
`format` / …), `author`, the affected `text`, and a `range:START-END` id.
Reading is non-mutating. (Toggle Track Changes with `track on` / `off`; check it
with `track status`.) This is an alias for `revision list` — see the `revision`
group below to *resolve* changes.

```bash
$ wordlive revisions
[{"index": 1, "type": "insert", "author": "A. Reviewer", "text": "swift",
  "anchor_id": "range:5-10", "start": 5, "end": 10, "date": "2026-06-08T16:57:00"}]
```

Failures: `3` Word busy, `4` Word not running.

## `revision list | accept | reject | accept-all | reject-all`

```
wordlive revision list
wordlive revision accept  --index N
wordlive revision reject  --index N
wordlive revision accept-all [--anchor-id ID]
wordlive revision reject-all [--anchor-id ID]
```

Resolve tracked changes (the write side of `revisions`). `accept` makes the
change at 1-based `--index` permanent; `reject` undoes it — both renumber the
remaining revisions, so re-list between resolves. `accept-all` / `reject-all`
resolve every change at once and report the count; pass `--anchor-id` to scope
to a single anchor's range ("accept all my edits in this section" — note an
anchor's range is literal, so a `heading:N` covers only the heading line, not the
body below it; target a paragraph/range that spans the changes). All wrap one
atomic-undo step.

```bash
$ wordlive revision accept --index 1
{"ok": true, "index": 1, "accepted": true}
$ wordlive revision accept-all --anchor-id heading:3
{"ok": true, "accepted": 4, "anchor_id": "heading:3"}
```

Failures: `2` anchor not found / index out of range, `3` Word busy, `4` Word not
running.

## `read text --anchor-id ID [--view raw|final|original|segments]`

```
wordlive read text --anchor-id ID [--view raw|final|original|segments]
```

Read an anchor's text, optionally resolving tracked changes. `raw` (default) is
`Range.Text` as Word reports it — already the *final* view (inserted runs
present, deleted runs gone). `final` reconstructs the accept-all view explicitly;
`original` the reject-all view (the deleted wording, which lives only on the
delete revisions, spliced back in); `segments` returns the ordered
`{text, change}` breakdown (`change` ∈ `insert` / `delete` / `null`).

```bash
$ wordlive read text --anchor-id para:5 --view original
{"anchor_id": "para:5", "view": "original", "text": "the quick brown fox"}
```

Failures: `2` anchor not found, `3` Word busy, `4` Word not running.

## `locate --anchor-id ID`

```
wordlive locate --anchor-id ID [--doc DOC_NAME]
```

Report where an anchor sits in the **laid-out** document — a non-visual layout
read that answers "what page is this on" without a `snapshot` vision pass.
Returns the anchor's page span (`page` / `end_page` — the pages its first and
last characters fall on, equal for a single-line anchor), its first character's
`line` and `column`, and whether it's `in_table`.

Page numbers are only meaningful in print layout, so the document is
**repaginated first** — content-neutral, so the user's selection, scroll, and
view are left untouched (the same guarantee `snapshot` gives). Non-mutating.

```bash
$ wordlive locate --anchor-id heading:8
{"anchor_id": "heading:8", "page": 2, "end_page": 2, "line": 1, "column": 1, "in_table": false}
```

Scan `paragraphs` and watch `page` step up to find "which paragraph starts page
2"; compare a table/section anchor's `page` and `end_page` for its page span.

Failures: `2` anchor not found, `3` Word busy, `4` Word not running.

## `stats`

```
wordlive stats [--doc DOC_NAME]
```

A one-call document summary — the "what am I looking at before I act" read.
Returns `pages`, `words`, `characters`, `paragraphs`, and `lines` (Word's own
`ComputeStatistics`), plus the structural counts `sections`, `headings`,
`tables`, `images`, `comments`, `revisions` (from wordlive's discovery
collections, so they agree with `table list` / `images` / `outline`), and
`saved`. Page/line counts are print-layout truth, so the document is
repaginated first (selection/scroll/view untouched). Non-mutating.

```bash
$ wordlive stats
{"pages": 2, "words": 312, "characters": 1840, "paragraphs": 24, "lines": 48,
 "sections": 1, "headings": 5, "tables": 1, "images": 0, "comments": 1,
 "revisions": 0, "saved": false}
```

Composes with `locate`: `stats.pages` + an anchor's `locate` page answers "why
is this 2 pages" structurally, no image pass.

Failures: `3` Word busy, `4` Word not running.

## `proofing`

```
wordlive proofing [--doc DOC_NAME]
```

Run Word's proofing tools over the document. Returns `spelling` and `grammar`,
each `{count, errors}` — the exact error count plus a (capped) list of flagged
runs as `{text, anchor_id, para}`, so a `range:START-END` can be fed back into
`read` / `comment add` — and `readability`, Word's readability statistics
(`flesch_reading_ease`, `flesch_kincaid_grade_level`, `passive_sentences`,
`words_per_sentence`, …). Heavier than `stats` (it asks Word to (re)check the
document) but still a pure read; if proofing is disabled or the document is
protected, the affected section reports a `null` count / empty readability.

```bash
$ wordlive proofing
{"spelling": {"count": 1, "errors": [{"text": "teh", "anchor_id": "range:14-17", "para": "para:2"}]},
 "grammar": {"count": 0, "errors": []},
 "readability": {"flesch_reading_ease": 65.5, "flesch_kincaid_grade_level": 7.2}}
```

Failures: `3` Word busy, `4` Word not running.

## `lint`

```
wordlive lint [--rule ID|TAG ...] [--exclude ID|TAG ...] [--within ID] [--doc DOC_NAME]
```

Audit the document for formatting inconsistency, structural slips, and policy
breaches — the read that answers "what's off about this document before I hand
it over". Returns a severity-ranked list of findings, each
`{rule, kind, severity, anchor_id, message, fixable, fix, observed, expected}`:
`kind` is `consistency` / `structural` / `policy`, `severity` is `error` /
`warning` / `info`, and when `fixable` is `true` the `fix` is an op-shaped dict
(or list of them) — literally an `exec` op `regularize` would run.

`--rule` / `--exclude` (repeatable, mutually exclusive) narrow the rule set by
id or tag; with neither, the **default set** runs (every consistency +
structural rule; policy rules are off — none ship yet). `--within` scopes the
audit to one anchor (`heading:N` / `range:S-E` / `table:N:R:C`). It's a pure
read: layout rules repaginate content-neutrally, leaving selection, scroll, and
the document's `Saved` state untouched.

v1 rules — structural: `heading-keep-with-next`, `table-repeat-header`,
`list-numbering-continuity`; consistency: `heading-font-consistent`,
`heading-spacing-consistent`, `body-font-consistent`, and `mixed-run-format`
(report-only, not fixable).

```bash
$ wordlive lint --within heading:3
[{"rule": "heading-keep-with-next", "kind": "structural", "severity": "warning",
  "anchor_id": "heading:3", "message": "Heading does not keep with next paragraph.",
  "fixable": true, "observed": false, "expected": true,
  "fix": {"op": "format_paragraph", "anchor_id": "heading:3", "keep_with_next": true}}]
```

`fixable` findings feed straight into `regularize`; report-only ones
(`mixed-run-format`) are yours to resolve by hand.

Failures: `1` unknown rule id/tag, `2` `--within` anchor not found,
`3` Word busy, `4` Word not running.

## `regularize`

```
wordlive regularize [--rule ID|TAG ...] [--exclude ID|TAG ...] [--within ID] [--dry-run] [--doc DOC_NAME]
```

Apply the **fixable** `lint` findings in one atomic-undo edit (labelled
"Regularize formatting", so one Ctrl-Z reverts them all; selection and scroll
are preserved). Returns `{applied, skipped, findings}` plus `ops_run` (and
`dry_run` when set). The default fixes are **targeted and idempotent** — each
writes the style's own value back as a direct property, so a *second*
`regularize` applies nothing (a tested invariant). `--dry-run` plans without
writing. Same `--rule` / `--exclude` / `--within` selection as `lint`.

Content-changing fixes (deletes, caption inserts) are out of scope — this is a
formatting/structure regularizer only. It's Track-Changes-aware: with Track
Changes on, the edits are recorded as revisions.

```bash
$ wordlive regularize --within heading:3 --dry-run
{"applied": 0, "skipped": 0, "dry_run": true, "ops_run": 1,
 "findings": [{"rule": "heading-keep-with-next", "anchor_id": "heading:3", "fixable": true}]}

$ wordlive regularize --within heading:3
{"applied": 1, "skipped": 0, "ops_run": 1,
 "findings": [{"rule": "heading-keep-with-next", "anchor_id": "heading:3", "fixable": true}]}
```

Run `lint` first to preview; `regularize --dry-run` to see exactly which fixes
would fire. Atomic-undo.

Failures: `1` unknown rule id/tag, `2` `--within` anchor not found,
`3` Word busy, `4` Word not running.

## `checkpoint`

```
wordlive checkpoint [--include text|text+style|text+format] [--within ID] [--out FILE] [--doc DOC_NAME]
```

Fingerprint the document's structure now → a checkpoint token (a pure read).
Store the token, edit the document, then `diff --since FILE` for a structured
change list — the way an agent verifies its edits landed, or sees what the user
changed (Word emits no content-change event). `--include` sets the fingerprint
depth: `text` (a restyle is invisible) < `text+style` (default — restyle
surfaces) < `text+format` (a direct-format edit surfaces as a `reformat`).
`--within ID` fingerprints just one anchor's range. Without `--out` the token is
the JSON object on stdout; with `--out FILE` the token is written to the file and
a small summary is emitted.

```bash
$ wordlive checkpoint --out cp.json
{"out": "cp.json", "include": "text+style", "scope": null, "paragraphs": 42, "tables": 1}
```

## `diff`

```
wordlive diff --since FILE [--doc DOC_NAME]
wordlive diff --from FILE --to FILE
```

Diff a stored checkpoint against the document **now** (`--since`, needs live
Word), or diff two stored checkpoints (`--from` / `--to`, no live Word). Emits a
content-aligned change list: each change is `replace` / `insert` / `delete` /
`restyle` / `reformat`, carrying the current `para:N` so a follow-up op can act on
it immediately. An unchanged document returns `[]` (a `doc_hash` fast-path).

```bash
$ wordlive diff --since cp.json
[{"op": "replace", "anchor_id": "para:14", "index_before": 12, "index_after": 13,
  "text_before": "Costs fell 4%.", "text_after": "Costs fell 9%."},
 {"op": "insert", "anchor_id": "para:7", "index_after": 6, "text_after": "A new note."}]
```

Failures: `1` bad/missing file or argument combination, `2` `--within` anchor not
found, `3` Word busy, `4` Word not running.

## `hyperlinks`

```
wordlive hyperlinks [--doc DOC_NAME]
```

List the document's hyperlinks — the read mirror of `link`. Each reports its
visible `text`, external `address` or internal `sub_address` bookmark,
`screen_tip`, and a `range:START-END` / `para:N`. Non-mutating.

```bash
$ wordlive hyperlinks
[{"index": 1, "text": "Acme", "address": "https://acme.example", "sub_address": "",
  "screen_tip": "", "anchor_id": "range:15-19", "para": "para:2"}]
```

Failures: `3` Word busy, `4` Word not running.

## `set-hyperlink --index N [--address URL] [--sub-address BM] [--text T] [--screen-tip T]`

```
wordlive set-hyperlink --index N [--address URL] [--sub-address BOOKMARK]
    [--text T] [--screen-tip T] [--doc DOC_NAME]
```

Retarget or relabel an existing hyperlink in place — no delete + reinsert.
Address the link by its 1-based `--index` (from `hyperlinks`). `--address` is the
external destination (URL / mailto / file path); `--sub-address` is the
in-document bookmark target; `--text` is the visible link text; `--screen-tip`
the hover tooltip. Pass at least one; omitting one leaves it untouched.
`--address` and `--sub-address` are left orthogonal (setting one does not clear
the other). These *retarget* a link, they don't unlink it: `--sub-address` /
`--screen-tip` clear with `""`, but `--address` / `--text` cannot be emptied
(Word keeps every link pointing somewhere with visible text).

```bash
$ wordlive set-hyperlink --index 1 --address https://new.example --text "New site"
{"ok": true, "index": 1, "applied": {"address": "https://new.example", "text": "New site"}}
```

Failures: `1` bad input; `2` index out of range; `3` Word busy.

## `fields`

```
wordlive fields [--doc DOC_NAME]
```

List the document's fields — the read mirror of `insert-field`. Each reports its
`kind` (the code's leading keyword — `PAGE` / `REF` / `TOC` / …), raw `code`,
last-rendered `result`, `locked`, and a `range:START-END` / `para:N`. Run
`update-fields` first to refresh stale results. Non-mutating.

```bash
$ wordlive fields
[{"index": 1, "kind": "PAGE", "type": 33, "code": "PAGE", "result": "1",
  "locked": false, "anchor_id": "range:16-17", "para": "para:2"}]
```

Failures: `3` Word busy, `4` Word not running.

## `properties list | set | delete`

```
wordlive properties list [--doc DOC_NAME]
wordlive properties set --name NAME --value VALUE [--custom] [--doc DOC_NAME]
wordlive properties delete --name NAME [--doc DOC_NAME]
```

Read and edit the document's metadata. `list` returns `{builtin, custom}` —
the built-in bag (Title, Author, Subject, Keywords, …, plus read-only stats like
the word count) and any custom name/value pairs. `set` writes a built-in
property by name, or a custom one with `--custom` (created if absent); `delete`
removes a custom property (built-ins can't be removed). Writes are atomic-undo.

```bash
$ wordlive properties set --name Title --value "Q3 Report"
{"ok": true, "name": "Title", "value": "Q3 Report", "custom": false}
```

Failures: `1` read-only/unknown built-in (set) or missing custom (delete),
`3` Word busy, `4` Word not running.

## `variables list | set | delete`

```
wordlive variables list [--doc DOC_NAME]
wordlive variables set --name NAME --value VALUE [--doc DOC_NAME]
wordlive variables delete --name NAME [--doc DOC_NAME]
```

Read and edit the document's variables — invisible named string storage (the
backing store for `{ DOCVARIABLE name }` fields). `list` returns a `{name: value}`
map; `set` creates or updates a variable; `delete` removes one. Writes are
atomic-undo.

```bash
$ wordlive variables list
{"ClientName": "Acme"}
```

Failures: `2` missing variable (delete), `3` Word busy, `4` Word not running.

## `images` / `read-image --anchor-id ID [--out FILE]`

```
wordlive images [--doc DOC_NAME]
wordlive read-image --anchor-id ID [--out FILE] [--doc DOC_NAME]
```

The read side of images — pull an embedded picture's bytes back out (the write
side is `insert-image`). `images` lists every embedded picture: its `image:N` id,
MIME type, size (points), alt text, and the `para:N` it sits in.

```bash
$ wordlive images
[{"index": 1, "anchor_id": "image:1", "mime": "image/png",
  "width": 240.0, "height": 160.0, "alt_text": "Chart", "para": "para:6"}]
```

`read-image` extracts one picture, resolved by `--anchor-id image:N` (or any
anchor whose range contains exactly one image). With `--out` the raw bytes are
written to that file and the JSON reports `{path, mime, bytes}`; without it,
base64 data is returned inline (`{mime, bytes, base64}`).

```bash
$ wordlive read-image --anchor-id image:1 --out logo.png
{"ok": true, "anchor_id": "image:1", "mime": "image/png", "bytes": 8462, "path": "logo.png"}

$ wordlive read-image --anchor-id image:1
{"ok": true, "anchor_id": "image:1", "mime": "image/png", "bytes": 8462, "base64": "iVBORw0KG…"}
```

Reading is non-mutating. Failures: `1` bad input (a range with no image, or more
than one); `2` anchor not found; `3` Word busy, `4` Word not running.

## `bookmark add NAME --anchor-id ID` *(deprecated)*

Hidden, deprecated alias for **`write bookmark NAME --create --anchor-id ID`**
(see above). Kept for one release; prefer the `write bookmark --create` form.

## `link --anchor-id ID (--url U | --bookmark B) [--text T] [--screen-tip S]`

```
wordlive link --anchor-id ID (--url URL | --bookmark NAME) [--text "…"]
    [--screen-tip "…"] [--doc DOC_NAME]
```

Turn an anchor into a hyperlink. Pass exactly one destination: `--url` for an
external link (URL / `mailto:` / file path) or `--bookmark` for an internal jump
to a bookmark in this document. With no `--text` the anchor's existing range
becomes the link; `--text` instead **inserts** new linked text at the end of the
range (so linking a heading or phrase keeps its content).

```bash
$ wordlive link --anchor-id range:120-140 --url "https://example.com"
{"ok": true, "anchor_id": "range:120-140",
 "applied": {"url": "https://example.com", "bookmark": null, "text": null}}
```

Failures: `1` bad input (not exactly one of `--url`/`--bookmark`); `2` anchor not
found; `3` Word busy.

## `cross-ref --anchor-id ID --target TARGET [--kind …] [--no-hyperlink]`

```
wordlive cross-ref --anchor-id ID --target TARGET
    [--kind text|page|number|above_below] [--hyperlink | --no-hyperlink]
    [--before | --after] [--doc DOC_NAME]
```

Insert a cross-reference to another anchor. `--target` is a `bookmark:NAME`,
`heading:N`, `footnote:N`, or `endnote:N` id. `--kind` selects what shows: the
referenced `text` (default), its `page` number, its `number`, or its
`above_below` position.

```bash
$ wordlive cross-ref --anchor-id end --target bookmark:Intro --kind page
{"ok": true, "anchor_id": "end",
 "applied": {"target": "bookmark:Intro", "kind": "page", "hyperlink": true, "where": "after"}}
```

References go stale when the document shifts — run `update-fields` (or
`snapshot`) to refresh them. Failures: `1` bad input; `2` anchor / target not
found; `3` Word busy.

## `caption --anchor-id ID [--label Figure] [--text "…"]`

```
wordlive caption --anchor-id ID [--label Figure] [--text "…"]
    [--position above|below] [--doc DOC_NAME]
```

Insert an auto-numbered caption (Figure 1, Table 2, …) as its **own paragraph**
at an anchor. `--label` is a built-in (`Figure`/`Table`/`Equation`) or a custom
string; `--text` is the title after the label and number. Pairs with `cross-ref`
for "see Figure 2". The caption always becomes its own `Caption`-styled
paragraph (it never fuses into the target paragraph); on a table cell anchor it
captions the **whole table**. `--position` controls above/below — default is
**above for a `Table`, below otherwise** (the figure/table convention).

```bash
$ wordlive caption --anchor-id range:120-140 --label Figure --text "System overview"
{"ok": true, "anchor_id": "range:120-140",
 "applied": {"label": "Figure", "text": "System overview", "position": null}}
```

Failures: `1` bad input; `2` anchor not found; `3` Word busy.

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
    [--before | --after] [--block | --no-block] [--width N] [--height N] \
    [--alt-text "…"] [--lock-aspect | --no-lock-aspect] [--doc DOC_NAME]
```

Embed an image at **any** anchor. Exactly one image source is required:
`--path` reads a file from disk (best for large images); `--base64` takes
base64 data inline, or `--base64 -` reads base64 from **stdin** (handy when an
LLM holds image data in memory). The picture is embedded in the document, not
linked, so the source file can move or vanish afterwards. Word auto-detects the
image's natural size; `--width`/`--height` (points) override it and
`--lock-aspect` (the default) keeps the aspect ratio.

A `--path` source is **screened** before any filesystem access: a non-local
path — a UNC `\\host\share\…`, a `file://`, or any URL — is rejected (exit `1`),
and if `--image-dir` / `WORDLIVE_IMAGE_DIRS` is configured the path must resolve
inside it. Prefer `--base64`/stdin for untrusted (e.g. LLM-supplied) images.

`--wrap` is **required** so layout intent is always explicit:

| `--wrap`                                            | Effect                                              |
| --------------------------------------------------- | --------------------------------------------------- |
| `inline`                                            | Stays in the text flow, like a character.           |
| `auto`                                              | Floats: Square if ≤ half the page's usable width, else top-and-bottom. |
| `square` `tight` `through` `top-bottom` `front` `behind` | Floats with that wrap type.                    |

`--after` (default) places the image just below the anchor; `--before` above.
`--block` puts the image on its own new (`Normal`) line instead of in the
anchor's text run — use it with `--before` at a heading so the image lands on its
own line above the heading instead of joining the heading text.

A **floating** wrap (anything but `inline`) leaves the text flow, so the image is
no longer an `image:N`; the output then carries a `shape:N` handle (see
**Floating shapes**) to re-wrap / reposition / resize / replace it. An `inline`
image stays `image:N` and reports `"shape": null`.

```bash
$ wordlive insert-image --anchor-id heading:2 --path diagram.png --wrap auto
{"ok": true, "anchor_id": "heading:2", "anchor": {"kind": "heading", "name": "Risks"}, "shape": "shape:1", "wrap": "auto", "where": "after"}

$ base64 logo.png | wordlive insert-image --anchor-id bookmark:Header --base64 - --wrap square --width 96
{"ok": true, "anchor_id": "bookmark:Header", "anchor": {"kind": "bookmark", "name": "Header"}, "shape": "shape:1", "wrap": "square", "where": "after"}
```

Failures: `1` the image is missing, unreadable, or not a recognised raster
format (PNG/JPEG/GIF/BMP/TIFF) — an `ImageSourceError`; `2` anchor not found
or an invalid `--wrap` value; `3` Word busy.

## `watermark (--text "…" | --remove) [--layout …] [--color …] [--font …] [--solid]`

```
wordlive watermark --text "DRAFT" [--font NAME] [--color COLOR] \
    [--layout diagonal|horizontal] [--transparent | --solid] [--doc DOC_NAME]
wordlive watermark --remove [--doc DOC_NAME]
```

Stamp (or `--remove`) a text watermark behind every page — WordArt drawn into
each section's header story, named like Word's own *Design → Watermark* object
(so it replaces a ribbon-added one). `--layout` is `diagonal` (45°, default) or
`horizontal`; `--color` the fill (default `#C0C0C0`); `--solid` turns off the
default 50% wash. Setting twice replaces rather than stacks. One atomic-undo step.

```bash
$ wordlive watermark --text "CONFIDENTIAL" --color "#ff0000"
{"ok": true, "text": "CONFIDENTIAL", "sections": 1}
$ wordlive watermark --remove
{"ok": true, "removed": 1}
```

Failures: `1` bad `--layout`/`--color`; `3` Word busy, `4` Word not running.

## `insert-text-box --anchor-id ID --text "…" [--width …] [--height …] [--wrap …]`

```
wordlive insert-text-box --anchor-id ID --text "…" [--width W] [--height H] \
    [--wrap square|tight|through|top-bottom|front|behind] [--before | --after] \
    [--font F] [--size S] [--bold | --no-bold] [--italic | --no-italic] \
    [--align left|center|right|justify] [--fill COLOR] \
    [--border-color COLOR | --no-border] [--doc DOC_NAME]
```

Insert a floating text box / pull quote anchored to the anchor's paragraph.
`--width`/`--height` accept points or unit strings (`3in`, `8cm`); `--wrap` is
how body text flows around it (the `insert-image` vocabulary minus `inline`).
`--fill` sets a background colour and `--no-border` / `--border-color` control
the outline. One atomic-undo step. The output carries the new text box's
`shape:N` id (see **Floating shapes** below) for restyling it afterward.

```bash
$ wordlive insert-text-box --anchor-id heading:2 --text "Key takeaway" --width 2.5in --fill "#eeeeff"
{"ok": true, "anchor_id": "shape:1", "wrap": "square"}
```

Failures: `2` anchor not found or invalid `--wrap`; `3` Word busy.

## Floating shapes — `shapes` / `set-shape-*` / `format-shape` / `group-shapes` / `ungroup-shape` / `replace-shape-image` / `delete-shape` / `set-image-*`

```
wordlive shapes [--doc DOC_NAME]
wordlive set-shape-wrap     --anchor-id shape:N [--wrap square|tight|through|top-bottom|front|behind] [--side both|left|right|largest] [--distance-top D] [--distance-bottom D] [--distance-left D] [--distance-right D]
wordlive set-shape-crop     --anchor-id shape:N [--left L] [--top T] [--right R] [--bottom B]
wordlive set-shape-position --anchor-id shape:N [--left L] [--top T] [--relative-to margin|page]
wordlive set-shape-size     --anchor-id shape:N [--width W] [--height H] [--lock-aspect | --no-lock-aspect]
wordlive format-shape       --anchor-id shape:N [--fill C] [--border-color C | --no-border | --default-border] [--border-weight W]
wordlive set-shape-alt-text --anchor-id shape:N --text "…"
wordlive set-shape-text     --anchor-id shape:N --text "…"
wordlive set-shape-rotation --anchor-id shape:N --degrees 30
wordlive set-shape-z-order  --anchor-id shape:N --order front|back|forward|backward
wordlive set-shape-text-frame --anchor-id shape:N [--margin-left L] [--margin-right R] [--margin-top T] [--margin-bottom B] [--word-wrap | --no-word-wrap]
wordlive replace-shape-image --anchor-id shape:N (--path FILE | --base64 VALUE)
wordlive group-shapes       --anchor-id shape:N --anchor-id shape:M [...]
wordlive ungroup-shape      --anchor-id shape:N
wordlive delete-shape       --anchor-id shape:N
wordlive set-image-alt-text --anchor-id image:N --text "…"
wordlive set-image-size     --anchor-id image:N [--width W] [--height H] [--lock-aspect | --no-lock-aspect]
wordlive set-image-crop     --anchor-id image:N [--left L] [--top T] [--right R] [--bottom B]
```

`shape:N` addresses the document's **floating** shapes — a text box from
`insert-text-box`, a floating image from `insert-image` (any `--wrap` but
`inline`), or WordArt — in document order; header-story watermarks are excluded.
`shapes` lists them (id, kind, size, rotation, z-order, wrap, wrap-side, crop, the
`para:N` they're anchored in). `shape:N` is **positional**: inserting or deleting a shape renumbers
the rest, so re-list rather than caching an id (the `image:N` / `chart:N` rule).
`textbox:N` is an addressing alias onto a text box's canonical `shape:N`.

`--left`/`--top` are lengths (`2in`) or `center`. `set-shape-text` /
`set-shape-text-frame` need a text box; `set-shape-z-order` restacks within the
float layer (distinct from wrap's front/behind-text); `set-shape-rotation` is an
absolute angle. `group-shapes` collapses two or more floats into one group
`shape:N` (pass `--anchor-id` twice or more), and `ungroup-shape` dissolves a
group back into its members. `replace-shape-image` needs a picture shape and swaps
its image **in place** (delete + reinsert at the same anchor, preserving wrap /
position / size / alt text) — pass exactly one of `--path` / `--base64`, and
`--path` is screened like `insert-image`. `set-shape-wrap` takes any one of
`--wrap` / `--side` (which sides text flows past — `square`/`tight`/`through`
honour it) / `--distance-*` (standoff gaps). `set-shape-crop` / `set-image-crop`
trim a **picture** in from its edges (a non-picture shape is rejected).
`set-image-*` restyle an **inline** picture (`image:N`) without floating it
(re-wrapping is `insert-image --wrap`). There is no autosize knob — Word's
resize-to-fit-text doesn't expose over COM.
The exec ops are `set_shape_wrap`, `set_shape_crop`, `set_shape_position`,
`set_shape_size`, `format_shape`, `set_shape_alt_text`, `set_shape_text`,
`set_shape_rotation`, `set_shape_z_order`, `set_shape_text_frame`,
`replace_shape_image`, `delete_shape`, `group_shapes`, `ungroup_shape`,
`set_image_alt_text`, `set_image_size`, `set_image_crop`.

```bash
$ wordlive set-shape-size --anchor-id shape:1 --width 3in --no-lock-aspect
{"ok": true, "anchor_id": "shape:1", "applied": {"width": "3in", "lock_aspect": false}}
```

Failures: `2` anchor not found or not a `shape:N`; `3` Word busy.

## `insert-equation --anchor-id ID (--unicodemath … | --latex … | --mathml …)` / `equations`

```
wordlive insert-equation --anchor-id ID (--unicodemath "…" | --latex "…" | --mathml "…") \
    [--display | --inline] [--before | --after] [--doc DOC_NAME]
wordlive equations [--doc DOC_NAME]
```

Insert a mathematical equation at **any** anchor. Exactly one input dialect is
required:

| Input          | Needs            | Example |
| -------------- | ---------------- | ------- |
| `--unicodemath`| nothing (native) | `"x=(-b±√(b^2-4ac))/(2a)"`, `"a^2+b^2=c^2"` |
| `--latex`      | the `latex` extra | `"\frac{-b\pm\sqrt{b^2-4ac}}{2a}"` |
| `--mathml`     | nothing          | `"<math>…</math>"` (or `--mathml -` from **stdin**) |

UnicodeMath is typed into a math zone and *built up* by Word itself; LaTeX and
MathML travel LaTeX→MathML→OMML→Word through Office's own shipped transform, so
only the LaTeX→MathML hop needs a third party (`pip install "wordlive[latex]"`).
The equation lands on its own paragraph with a pinned style (so it never
inherits a neighbouring heading's style and pollutes the outline/TOC):
`--display` (default) gives it the centred `Equation` paragraph style
(auto-created, based on `Normal`); `--inline` makes it `Normal` and left-aligned
(still its own paragraph, not mid-sentence). `--after` (default) places it below
the anchor; `--before` above (including a clean prepend at the document start).

`equations` is the read side: every equation's `equation:N` id, `type`
(display/inline), a linear preview, and the `para:N` it sits in. `equation:N` is
positional (Word's `OMaths` order) — inserting another equation before an
existing one renumbers it, so re-list rather than caching an id across inserts.

```bash
$ wordlive insert-equation --anchor-id heading:2 --unicodemath "a^2+b^2=c^2"
{"ok": true, "anchor_id": "heading:2", "equation": 1, "equation_anchor_id": "equation:1", "display": true, "where": "after"}

$ wordlive equations
[{"index": 1, "anchor_id": "equation:1", "type": "display", "linear": "𝑎2+𝑏2=𝑐2", "para": "para:1"}]
```

Failures: `1` malformed input (no dialect, or more than one; unparseable
MathML/LaTeX; the `latex` extra not installed) — an `EquationError`; `2` anchor
not found; `3` Word busy.

## `insert-chart --anchor-id ID --kind bar|pie|line|scatter --data JSON` / `charts`

```
wordlive insert-chart --anchor-id ID --kind bar|pie|line|scatter --data JSON \
    [--title "…"] [--before | --after] [--doc DOC_NAME]
wordlive charts [--doc DOC_NAME]
```

Insert an **Excel-backed** chart at any anchor (atomic-undo). `--data` is JSON
(or `--data -` to read it from stdin): an object `{"label": value}` for
`bar`/`pie`/`line`, or an array of `[x, y]` pairs for `scatter` (both axes
numeric, duplicate x preserved); `line` accepts either. `--title` sets the chart
title and series name.

Charts embed a hidden Excel workbook (`InlineShapes.AddChart2`), so **Excel must
be installed** — if it isn't, the command exits **`6`** (`ExcelNotAvailableError`)
and the document is left untouched. wordlive then breaks the data link, so the
chart's data is **static**: no embedded workbook ships in the document, and the
series data isn't read back. `charts` is the read side — every chart's `chart:N`
id, `kind`, `title`, and `para:N`; `chart:N` is positional (document order), so
inserting another chart earlier renumbers it.

```bash
$ wordlive insert-chart --anchor-id end --kind bar --data '{"Q1": 10, "Q2": 25, "Q3": 18}' --title "Quarterly"
{"ok": true, "anchor_id": "end", "chart": 1, "chart_anchor_id": "chart:1", "kind": "bar", "where": "after"}

$ echo '[[1.2, 3.4], [1.2, 3.9], [2.5, 6.1]]' | wordlive insert-chart --anchor-id end --kind scatter --data -
{"ok": true, "anchor_id": "end", "chart": 2, "chart_anchor_id": "chart:2", "kind": "scatter", "where": "after"}

$ wordlive charts
[{"index": 1, "anchor_id": "chart:1", "kind": "bar", "title": "Quarterly", "chart_style": 201, "has_legend": false, "para": "para:1"}]
```

Failures: `1` malformed `--data` (bad JSON, empty, wrong shape, non-numeric
value — an `OpError`); `2` anchor not found; `3` Word busy; **`6`** Excel not
installed (`ExcelNotAvailableError`).

## `format-chart` / `format-axis` / `add-trendline` / `set-series-color`

```
wordlive format-chart --anchor-id chart:N [--title TXT] [--legend|--no-legend] \
    [--legend-position right|left|top|bottom|corner] [--chart-style INT] \
    [--background COLOR] [--plot-background COLOR] \
    [--font NAME] [--font-size SIZE] [--font-color COLOR] \
    [--data-labels|--no-data-labels] [--data-label-format FMT] \
    [--chart-type bar|pie|line|scatter]
wordlive format-axis --anchor-id chart:N --which value|y|category|x \
    [--title TXT] [--minimum N] [--maximum N] [--scale linear|log] \
    [--number-format FMT] [--gridlines|--no-gridlines]
wordlive add-trendline --anchor-id chart:N [--series N] \
    [--kind linear|exponential|logarithmic|moving_average|polynomial|power] \
    [--display-equation] [--display-r-squared] [--forward N] [--backward N]
wordlive set-series-color --anchor-id chart:N --color COLOR [--series N] [--point N]
```

Format and design an **existing** `chart:N` — Word's chart "Design"/"Format"
tabs. These operate on the post-insert, static chart, so **Excel is not needed**
(no exit 6). Every option is tri-state: only what you pass is written. Colours are
a name, hex (`#2E86C1`), or comma-separated `r,g,b`. `format-chart --chart-type`
re-types the chart in place; `--chart-style` is the built-in design-gallery id.
`format-axis --scale log` suits order-of-magnitude data; `--which` takes
`value`/`y` or `category`/`x`. `add-trendline --kind power --display-equation`
draws the law of best fit. `set-series-color --point N` recolours one bar / pie
slice / marker (1-based); omit it to colour the whole series.

```console
$ wordlive format-chart --anchor-id chart:1 --chart-style 240 --legend --title "Quarterly"
{"ok": true, "anchor_id": "chart:1", "applied": {"title": "Quarterly", "legend": true, "chart_style": 240}}

$ wordlive format-axis --anchor-id chart:1 --which value --scale log --title "USD (M)"
{"ok": true, "anchor_id": "chart:1", "which": "value", "applied": {"title": "USD (M)", "scale": "log"}}

$ wordlive add-trendline --anchor-id chart:2 --kind power --display-equation
{"ok": true, "anchor_id": "chart:2", "applied": {"series": 1, "kind": "power", "display_equation": true, "display_r_squared": false}}

$ wordlive set-series-color --anchor-id chart:1 --color "#2E86C1" --point 2
{"ok": true, "anchor_id": "chart:1", "series": 1, "point": 2, "color": "#2E86C1"}
```

Failures: `1` bad input (unknown colour / scale / trendline kind, or the anchor
isn't a `chart:N`); `2` chart not found; `3` Word busy.

## `snapshot [--anchor-id ID | --page N | --pages A-B]`

```
wordlive snapshot [--anchor-id ID | --page N | --pages A-B] \
    [--out FILE] [--dpi 150] [--max-dim N] [--markup none|all] [--doc DOC_NAME]
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

`--max-dim N` caps each page's **long edge** to `N` pixels (only ever lowering
resolution). A vision model is billed on an image's pixel area, so a long-edge
cap is a predictable per-page token budget regardless of paper size — the lever
for a cheap **whole-document** layout check (pair it with no page target; ~1000
stays legible for "did my styling land"). It composes with `--dpi` (the cap wins
when it implies a lower resolution); `--dpi 72` is a coarser alternative.

```bash
$ wordlive snapshot --max-dim 1000          # whole doc, every page capped to 1000px long edge
{"ok": true, "selector": "all", "dpi": 150, "max_dim": 1000, "count": 12, "images": [...]}
```

`--markup all` renders tracked changes and comments as visible revision marks
and balloons (default `none` renders the final document); the structured list is
the `revisions` command.

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

## `save` / `save-as PATH` / `export-pdf PATH` (gated)

```
wordlive save                                              [--doc DOC_NAME]
wordlive save-as PATH [--format docx] [--overwrite]        [--doc DOC_NAME]
wordlive export-pdf PATH [--pages A-B]                     [--doc DOC_NAME]
```

Persist the document or hand back a deliverable. **These three verbs are gated:**
they only write inside a directory whitelisted with the global `--save-dir`
(repeatable) or `WORDLIVE_SAVE_DIRS` (an `os.pathsep`-separated list). With no
whitelist configured, saving is **off** (exit `1`, `PathNotAllowedError`). The
target is resolved *first* (so `..`/symlinks can't escape) and must then sit
inside an allowed directory. The Python API (`doc.save()` etc.) is ungated.

- `save` writes to the document's existing file (fails if it was never saved —
  use `save-as` first). Its existing path must itself be whitelisted.
- `save-as PATH` writes a `.docx` (the only `--format`; PDF is `export-pdf`).
  Refuses to clobber an existing file unless `--overwrite` is given.
- `export-pdf PATH` exports a PDF — a pixel-faithful render via Word's PDF
  engine (the same one `snapshot` uses), the recommended way to hand back a
  deliverable. `--pages A-B` (or a single `--pages N`) limits the page range;
  the whole document by default. Overwrites an existing PDF.

```bash
$ wordlive --save-dir C:\out save-as C:\out\report.docx
{"ok": true, "path": "C:\\out\\report.docx", "format": "docx"}

$ wordlive --save-dir C:\out export-pdf C:\out\report.pdf
{"ok": true, "path": "C:\\out\\report.pdf"}

$ wordlive save-as C:\elsewhere\x.docx          # no whitelist → denied
error: saving is disabled: no save directories are configured …
# exit 1
```

Failures: `1` policy denial (no/wrong whitelist, refused overwrite) or bad
input; `3` Word busy; `4` Word not running.

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

## `find-paragraph --text "…" [--limit N] [--min-score F]`

```
wordlive find-paragraph --text "…" [--limit 5] [--min-score 0.6]   [--doc DOC_NAME]
```

**Fuzzy** paragraph search — the typo-/paraphrase-tolerant counterpart to
`find`. Where `find` does an *exact substring* match (on normalized text) and
returns `range:START-END` hits, `find-paragraph` scores **every paragraph**
against the query with a similarity ratio (`difflib.SequenceMatcher`, over the
same NFKC + smart-quote/dash/whitespace normalization) and returns the best
`para:N` candidates. Use it when you have approximately-remembered text and want
the paragraph it belongs to.

```bash
$ wordlive find-paragraph --text "the quick brown fox jumped over the dog"
[{"anchor_id": "para:12", "index": 12, "score": 0.9438,
  "text": "The quick brown fox jumps over the lazy dog.",
  "level": 10, "is_heading": false}]
```

Rows are sorted by descending `score`; only those with `score >= --min-score`
(default `0.6`) are kept, capped at `--limit` (default `5`). Headings are
included (flagged by `is_heading` / `level`), addressed by `para:N`. An empty or
whitespace-only query returns `[]`.

Failures: returns `[]` with exit `0` for no matches above the threshold; `1` for
`--limit < 1` or `--min-score` outside `[0, 1]`.

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
exist in the document — define one first with [`style add`](#style-add-name)
if needed.

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
    [--line-spacing MULTIPLE|single|1.5|double|LENGTH]
    [--page-break-before | --no-page-break-before]
    [--keep-together | --no-keep-together]
    [--keep-with-next | --no-keep-with-next]
    [--widow-control | --no-widow-control]
    [--doc DOC_NAME]
```

`centre` is accepted as a synonym for `center` (UK spelling).

Set paragraph-formatting properties on the anchor's range. At least one
formatting flag is required. Indent and spacing values are in **points** —
the unit Word's COM API uses natively for these fields. `--line-spacing` sets
the leading *within* the paragraph (distinct from `--space-before`/`--space-after`,
which space paragraphs apart): a **multiple** of single spacing (`1`, `1.5`,
`2`), the keywords `single`/`1.5`/`double`, or an **exact length** (`14pt`,
`1.5cm`) for a fixed line height.
`--page-break-before` forces the paragraph to begin on a new page (and
`--no-page-break-before` clears it) — the *clean*, reflow-safe way to
page-break, leaving no stray break character (contrast `insert-break`, which
inserts an explicit one-off break). The remaining flags are Word's
**pagination** controls for clean multi-page layout: `--keep-together` keeps all
lines of a paragraph on one page, `--keep-with-next` keeps it with the following
paragraph (e.g. a heading with its first body line), and `--widow-control`
prevents a lone first/last line stranded at a page boundary. Atomic-undo.

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

## `format-run --anchor-id ID [...]`

```
wordlive format-run --anchor-id ID
    [--bold | --no-bold] [--italic | --no-italic]
    [--underline | --no-underline] [--strikethrough | --no-strikethrough]
    [--font NAME] [--size POINTS|UNIT] [--color NAME|HEX|R,G,B]
    [--highlight NAME] [--subscript | --no-subscript]
    [--superscript | --no-superscript] [--small-caps | --no-small-caps]
    [--all-caps | --no-all-caps] [--spacing POINTS|UNIT]
    [--doc DOC_NAME]
```

Set **character-formatting** (run-level) properties on the anchor's range —
the *bold this phrase* layer, distinct from `style apply` (named styles) and
`format-paragraph` (paragraph scope). Pairs naturally with a `range:START-END`
anchor to style a sub-paragraph span. At least one flag is required; only the
flags you pass are written. Atomic-undo.

`--color` takes a colour name (`red`, `navy`, …), a hex string (`#FF0000` or
`FF0000`), or comma-separated `R,G,B`. `--highlight` is a named text-highlight
palette colour (`yellow`, `green`, …, or `none` to clear). `--size`/`--spacing`
accept a bare number (points) or a unit string (`12pt`, `1.5mm`).

```bash
$ wordlive format-run --anchor-id range:120-145 --bold --color FF0000 --highlight yellow
{"ok": true,
 "anchor_id": "range:120-145",
 "anchor": {"kind": "range", "name": "range:120-145"},
 "applied": {"bold": true, "color": "FF0000", "highlight": "yellow"}}
```

Failures: `2` anchor not found, `1` bad colour/size/highlight value, `3` Word busy.

## `shading --anchor-id ID --fill COLOR`

```
wordlive shading --anchor-id ID --fill NAME|HEX|R,G,B [--doc DOC_NAME]
```

Set the background-fill shading of the anchor's range. Because a table cell is
an anchor (`table:N:R:C`), this is also how you shade a cell. Atomic-undo.

## `borders --anchor-id ID [...]`

```
wordlive borders --anchor-id ID
    [--sides all|box|top|bottom|left|right|horizontal|vertical]
    [--style single|double|dot|dash|dash-dot|none]
    [--weight POINTS] [--color NAME|HEX|R,G,B] [--doc DOC_NAME]
```

Draw borders on the anchor's range or cell. `--sides` is comma-separated for
several edges (default `all` = the four outer edges). `--weight` is in points,
snapped to Word's discrete line-width set (0.25/0.5/0.75/1/1.5/2.25/3 pt).
Page-wide and table-wide borders are out of scope (this sets per-range/per-cell
borders). Atomic-undo.

## `drop-cap --anchor-id ID [...]`

```
wordlive drop-cap --anchor-id ID
    [--position dropped|margin|none] [--lines N]
    [--distance POINTS|UNIT] [--font NAME] [--doc DOC_NAME]
```

Turn the first letter of the anchor's **paragraph** into a drop cap — the
editorial oversized initial, a real Word `DropCap` so the body text wraps around
it natively (not a faked big-font run). `--position` is `dropped` (default — the
letter sits into the text), `margin` (it hangs in the left margin), or `none`
(remove an existing drop cap; the other flags are then ignored). `--lines` is
how many lines tall the letter is (default 3), `--distance` the gap from the body
text, and `--font` an optional font for the dropped letter. Word rejects a drop
cap on an **empty** paragraph (no letter to drop) — that surfaces as a `3`/COM
error. Atomic-undo. Failures: `2` anchor not found, `1` bad value.

## `tab-stop --anchor-id ID --position POS [...]`

```
wordlive tab-stop --anchor-id ID --position POINTS|UNIT
    [--align left|center|right|decimal|bar]
    [--leader none|dots|dashes|lines|heavy|middle-dot] [--doc DOC_NAME]
```

Add a tab stop to the anchor's paragraph(s) — `--leader dots` gives the dotted
fill of a price list or table-of-contents row without a table. `--position`
accepts points or a unit string (`3in`). Atomic-undo.

## `style add NAME [...]`

```
wordlive style add NAME
    [--type paragraph|character|table|list]
    [--based-on NAME] [--next-style NAME] [--doc DOC_NAME]
```

Define a new style. `--based-on` is the inheritance parent; `--next-style` is
the style applied to the paragraph after one in this style. Style its defaults
with `style set`, then `style apply` it. Atomic-undo. Failures: `1` bad
`--type`, `2` unknown `--based-on`/`--next-style`.

```bash
$ wordlive style add Brand --based-on Normal
{"ok": true, "style": "Brand", "type": "paragraph"}
```

## `style set NAME [...]`

```
wordlive style set NAME
    [--bold|--no-bold] [--italic|--no-italic] [--underline|--no-underline]
    [--font NAME] [--size POINTS|UNIT] [--color NAME|HEX|R,G,B]
    [--alignment left|center|right|justify]
    [--space-before POINTS|UNIT] [--space-after POINTS|UNIT]
    [--line-spacing MULTIPLE|single|1.5|double|LENGTH]
    [--based-on NAME] [--next-style NAME] [--doc DOC_NAME]
```

Set the font / paragraph defaults of an existing style (built-in or one you
created). At least one property is required. The brand/template workflow:
`style add` once, `style set` its look, then `style apply` it everywhere.
Atomic-undo. Failures: `2` style not found, `1` bad value, `3` Word busy.

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

## `table records INDEX`

```
wordlive table records INDEX [--doc DOC_NAME]
```

Read table `INDEX` as **records** — the read mirror of `table create --data`
from a list of objects. Row 1 is taken as the header; each row below becomes a
`{header: cell_text}` object. Non-mutating.

```bash
$ wordlive table records 1
[{"Item": "Travel", "Cost": "$400"}]
```

A duplicate header label collapses (rightmost column wins); a blank header cell
yields an empty-string key — both the caller's responsibility, matching the
write path. Failures: `2` table index out of range, `3` Word busy.

## `table create`

```
wordlive table create --anchor-id ID [--rows R] [--cols C]
                      [--style NAME] [--header] [--before|--after]
                      [--data JSON | --data -] [--doc DOC_NAME]
```

Create a new table at a **position anchor** (`heading:`, `para:`, `start`,
`end`, `range:` — *not* a bare `table:N`, which addresses an existing table).
Every other verb edits existing structure; this is how you build a table from
nothing. Atomic-undo. Reports the new table's 1-based `index` for an immediate
follow-up `set-cell` / `add-row`.

`--data` populates the cells at creation and accepts two shapes:

- a **row-major 2-D array** (`[[r1c1, r1c2], …]`); or
- **records** — a list of objects (`[{"Tier":"Wobble","Monthly":"$9"}, …]`),
  whose keys become a header row (each object is one body row); this implies
  `--header`.

When `--data` is given, **`--rows`/`--cols` are optional** — they're inferred
from the data's shape, so the common case is just `--data …`. Pass them
explicitly to pad the grid *larger* than the data (a short/partial array leaves
trailing cells empty; an array that *overflows* the grid is a clean error, exit
1). Without `--data`, both `--rows` and `--cols` are required. Pass `--data -`
to read the JSON from stdin, which sidesteps Windows quoting/backslash fights
(mirrors `exec --ops -`).

`--style` names a table style defined in the document; it defaults to the
built-in **`Table Grid`** so a new table has visible borders rather than only
faint gridlines. A style name not in the document fails (exit 2). `--header`
bolds the first row.

```bash
$ wordlive table create --anchor-id end --header \
    --data '[["Tier","Monthly","SLA"],["Wobble","$9","best effort"],["Finch","$99","99.9%"]]'
{"ok": true, "table": 2, "rows": 3, "columns": 3}

$ wordlive table create --anchor-id end \
    --data '[{"Tier":"Wobble","Monthly":"$9"},{"Tier":"Finch","Monthly":"$99"}]'
{"ok": true, "table": 3, "rows": 3, "columns": 2}
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

## `table append-record`

```
wordlive table append-record --table INDEX --record '{"Item":"Lodging","Cost":"$600"}' [--doc DOC_NAME]
```

Append a row from a JSON **object**, mapping its keys to the header columns
(row 1). A header with no matching key gets an empty cell and an extra key is
ignored — the same lenient mapping `table create --data` (records) uses. The new
row inherits the table's existing formatting/banding. Atomic-undo. The header-
name companion to `add-row` (which is positional).

```bash
$ wordlive table append-record --table 1 --record '{"Item":"Lodging","Cost":"$600"}'
{"ok": true, "table": 1, "rows": 3}
```

Failures: `1` `--record` not a JSON object, `2` table index out of range, `3` Word busy.

## `table update-row`

```
wordlive table update-row --table INDEX --key VALUE --values '{"Cost":"$450"}'
    [--column HEADER] [--doc DOC_NAME]
```

Update the **first** row whose key-column cell equals `--key`, setting cells by
header name (`--values` is a `{header: new_value}` object). The key column is the
first column by default, or the header named by `--column`. First match wins on
duplicate keys. Atomic-undo — addresses a row by *content* instead of a fragile
1-based index.

```bash
$ wordlive table update-row --table 1 --key Travel --values '{"Cost":"$450"}'
{"ok": true, "table": 1, "key": "Travel"}
```

Validation happens before any edit: an unknown `--column`, or a `--values` key
that isn't a header, is a clean error. Failures: `1` `--values` not a JSON
object / unknown header or column, `2` table index out of range or no row matches
`--key`, `3` Word busy.

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

## `table set-heading-row`

```
wordlive table set-heading-row --table INDEX [--row R]
    [--heading | --no-heading] [--allow-break | --no-allow-break] [--doc DOC_NAME]
```

Mark a 1-based row (default `1`) as a **repeating heading row** — it repeats at
the top of every page a multi-page table spans (`Row.HeadingFormat`). `--row`
selects which row; `--no-heading` clears the flag. `--allow-break` controls
whether the row may split across a page (`Row.AllowBreakAcrossPages`), defaulting
to off for a heading row so the header stays intact. Atomic-undo.

```bash
$ wordlive table set-heading-row --table 1
{"ok": true, "table": 1, "row": 1, "heading": true}
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

## `table autofit`

```
wordlive table autofit --table INDEX [--mode content|window|fixed] [--doc DOC_NAME]
```

Resize a table's columns. `--mode content` (the default) shrinks/grows each
column to fit its cell contents, `window` stretches the table to the page width,
and `fixed` pins the current widths so Word stops auto-sizing. Atomic-undo —
a clean way to tidy a table whose columns drifted after edits.

```bash
$ wordlive table autofit --table 1 --mode content
{"ok": true, "table": 1, "mode": "content"}
```

Failures: `2` table index out of range, `3` Word busy.

## `table set-style`

```
wordlive table set-style --table INDEX --style NAME [--doc DOC_NAME]
```

Restyle an **existing** table — the post-creation counterpart of `table create
--style`. `--style` is any built-in or custom table style (`style list` filtered
to `type==table` discovers them). Applying a style reapplies its conditional
formatting and **overwrites direct cell shading**, so restyle *first*, then layer
cell-level overrides. Atomic-undo.

```bash
$ wordlive table set-style --table 1 --style "Grid Table 4 - Accent 1"
{"ok": true, "table": 1, "style": "Grid Table 4 - Accent 1"}
```

Failures: `2` table index out of range or unknown style, `3` Word busy.

## `table set-alignment`

```
wordlive table set-alignment --table INDEX --alignment left|center|right [--doc DOC_NAME]
```

Align the whole table across the page width (distinct from the text alignment
*inside* cells). Atomic-undo.

```bash
$ wordlive table set-alignment --table 1 --alignment center
{"ok": true, "table": 1, "alignment": "center"}
```

## `table set-borders`

```
wordlive table set-borders --table INDEX [--sides all] [--style single] [--weight 0.5] [--color C] [--doc DOC_NAME]
```

Draw borders across the **whole table grid** in one call — the table-wide
counterpart of the per-cell `borders` verb. `--sides` is `all`/`box`, a single
outer edge, the interior gridlines `horizontal`/`vertical`, or a comma-separated
list (`box,horizontal,vertical` rules every line). `--style` is a line style
(`single`, `double`, `dot`, … or `none`); `--weight` snaps to Word's set
(0.25/0.5/0.75/1/1.5/2.25/3 pt); `--color` is a name/hex/r,g,b. Atomic-undo.

```bash
$ wordlive table set-borders --table 1 --sides box,horizontal,vertical --style single --weight 1
{"ok": true, "table": 1, "applied": {"sides": ["box", "horizontal", "vertical"], "style": "single", "weight": 1.0, "color": null}}
```

## `table set-banding`

```
wordlive table set-banding --table INDEX [--first-row/--no-first-row] [--last-row/--no-last-row]
    [--first-column/--no-first-column] [--last-column/--no-last-column]
    [--banded-rows/--no-banded-rows] [--banded-columns/--no-banded-columns] [--doc DOC_NAME]
```

Toggle the "Table Style Options" (header row, total row, first/last column, banded
rows/columns) of the **applied table style**. Each flag is tri-state — pass it to
set, omit it to leave untouched. Only shows once a real table style is applied
(pair with `table set-style`). Atomic-undo.

```bash
$ wordlive table set-banding --table 1 --first-row --no-banded-rows
{"ok": true, "table": 1, "applied": {"first_row": true, "banded_rows": false}}
```

## `cell-valign`

```
wordlive cell-valign --anchor-id table:N:R:C --align top|center|bottom [--doc DOC_NAME]
```

Set a table cell's vertical alignment. Atomic-undo. (For whole-row / whole-column
styling, address `table:N:row:R` / `table:N:col:C` and use the `shading` /
`borders` / `apply-style` / `format-run` verbs.)

```bash
$ wordlive cell-valign --anchor-id table:1:2:2 --align bottom
{"ok": true, "anchor_id": "table:1:2:2", "align": "bottom"}
```

Failures: `1` not a cell anchor or bad `--align`, `2` table/cell out of range.

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

## `list format --anchor-id ID --levels JSON`

```
wordlive list format --anchor-id ID --levels '[{...}, …]' [--continue] [--doc DOC_NAME]
```

Author a **custom multi-level list** and apply it — the richer counterpart of
`list apply` (which only applies a gallery default). `--levels` is a JSON array
of per-level specs (1-based); each object's keys are all optional except a bullet
level's glyph:

- `kind` — `"number"` (default) or `"bullet"`.
- `format` — a number level's marker template (`"%1."`, `"%1)"`, `"%1.%2"`;
  `%N` references level N), default `"%{level}."`; for a bullet, the glyph.
- `style` — a number scheme: `arabic`, `upper-roman`, `lower-roman`,
  `upper-letter`, `lower-letter`, `ordinal`, … .
- `bullet` / `font` — a bullet's glyph and marker font (default `Symbol`).
- `start_at`, `number_position`, `text_position` (points or `"0.5in"`),
  `trailing` (`tab`/`space`/`none`), `alignment` (`left`/`center`/`right`),
  `bold`, `italic`, `color`.

More than one level mints an outline template. Atomic-undo.

```bash
$ wordlive list format --anchor-id range:0-40 \
    --levels '[{"kind":"number","format":"%1)","style":"lower-letter","trailing":"space"}]'
{"ok": true, "anchor_id": "range:0-40", "levels": 1}
```

Failures: `1` `--levels` not a JSON array / a bad spec, `2` anchor not found,
`3` Word busy.

## `list levels --anchor-id ID`

```
wordlive list levels --anchor-id ID [--doc DOC_NAME]
```

Read the per-level format of the list at an anchor (read-only) — the read mirror
of `list format`. Returns one `{level, kind, format, style, trailing,
number_position, text_position, font}` per level of the applied template, or an
empty list if the anchor isn't in a list.

```bash
$ wordlive list levels --anchor-id range:0-40
{"anchor_id": "range:0-40", "levels": [{"level": 1, "kind": "number", "format": "%1)", "style": "lower-letter", "trailing": "space", "number_position": 18.0, "text_position": 36.0, "font": ""}]}
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

## `sections`

```
wordlive sections [--doc DOC_NAME]
```

List the document's sections with each one's page setup. (Previously
`section list`, now a hidden deprecated alias.)

```bash
$ wordlive sections
[{"index": 1,
  "page_setup": {"orientation": "portrait",
                 "top_margin": 72.0, "bottom_margin": 72.0,
                 "left_margin": 72.0, "right_margin": 72.0,
                 "page_width": 612.0, "page_height": 792.0}}]
```

Margins and page dimensions are in **points**. Headers and footers live in the
`header` / `footer` commands. Failures: `3` Word busy, `4` Word not running.

## `page-setup [--section N] [...]`

```
wordlive page-setup [--section N]
    [--margins V] [--top-margin V] [--bottom-margin V] [--left-margin V] [--right-margin V]
    [--gutter V] [--orientation portrait|landscape] [--paper-size letter|legal|tabloid|a3|a4|a5]
    [--columns N] [--column-spacing V] [--doc DOC_NAME]
```

The write mirror of `section list` — set a section's page geometry. `--margins`
sets all four margins at once; the per-side `--*-margin` flags override it. Length
values (margins, gutter, column spacing) accept points or a unit string
(`1in`, `2.5cm`). `--orientation` and `--paper-size` reshape the page;
`--columns N` lays the section out in N equal, newspaper-style columns (the
section counterpart to `insert-break --kind column`). `--section` defaults to `1`,
which is the whole document for a single-section file. Atomic-undo. At least one
option is required.

```bash
$ wordlive page-setup --orientation landscape --margins 0.5in --columns 2
{"ok": true, "section": 1,
 "applied": {"margins": "0.5in", "orientation": "landscape", "columns": 2}}
```

Failures: `1` no option / bad input, `2` section out of range, `3` Word busy.

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
| `insert_paragraph`     | `anchor_id`, `text`                        | `where` (`after`/`before`) or `before: true`, `style`, `bind` |
| `insert_block`         | `anchor_id`, `items`                       | `where` or `before: true`, `bind` |
| `insert_section`       | `anchor_id`, `heading`, `body`             | `level`, `where` or `before: true`, `bind` |
| `insert_markdown`      | `anchor_id`, `markdown`                    | `where` or `before: true`, `bind` |
| `replace_section`      | `anchor_id`, one of `body` / `markdown`    | — |
| `delete_paragraph`     | `anchor_id`                                | — |
| `append`               | `text`                                     | `style`                           |
| `append_inline`        | `text`                                     | —                                 |
| `prepend`              | `text`                                     | `style`                           |
| `prepend_inline`       | `text`                                     | —                                 |
| `insert_image`         | `anchor_id`, `wrap`, and one of `path` / `base64` | `where` or `before: true`, `block`, `width`, `height`, `alt_text`, `lock_aspect` |
| `insert_equation`      | `anchor_id`, one of `unicodemath` / `latex` / `mathml` | `display`, `where` or `before: true` |
| `replace`              | `anchor_id`, `text`                        | —                                 |
| `find_replace`         | `find`, `text`                             | `in`, `all`, `occurrence`         |
| `apply_style`          | `anchor_id`, `name`                        | —                                 |
| `add_style`            | `name`                                     | `type`, `based_on`, `next_style`  |
| `set_style`            | `name`                                     | `based_on`, `next_style`, plus the `format_paragraph` / `format_run` formatting fields |
| `format_paragraph`     | `anchor_id`                                | `alignment`, `left_indent`, `right_indent`, `first_line_indent`, `space_before`, `space_after`, `line_spacing` (a multiple `1`/`1.5`/`2`, `single`/`1.5`/`double`, or an exact length like `14pt`), `page_break_before`, `keep_together`, `keep_with_next`, `widow_control` |
| `format_run`           | `anchor_id`                                | `bold`, `italic`, `underline`, `strikethrough`, `font`, `size`, `color`, `highlight`, `subscript`, `superscript`, `small_caps`, `all_caps`, `spacing` |
| `regularize`           | —                                          | `rules`, `within`, `dry_run`      |
| `set_shading`          | `anchor_id`                                | `fill`, `pattern`                 |
| `set_borders`          | `anchor_id`                                | `sides` (`all`/`box`/`top`/`bottom`/`left`/`right`/`horizontal`/`vertical`), `style` (a.k.a. `line_style`: `single`/`double`/`dot`/`dash`/`none`/…), `weight`, `color` |
| `drop_cap`             | `anchor_id`                                | `position` (`dropped`/`normal`/`margin`/`none`), `lines` (default 3), `distance`, `font` |
| `add_tab_stop`         | `anchor_id`, `position`                    | `align`, `leader`                 |
| `set_cell`             | `table`, `row`, `col`, `text`              | —                                 |
| `add_row`              | `table`                                    | `values`                          |
| `delete_row`           | `table`, `row`                             | —                                 |
| `set_heading_row`      | `table`                                    | `row` (default 1), `heading` (default true), `allow_break` |
| `autofit_table`        | `table`                                    | `mode` (`content`/`window`/`fixed`) |
| `append_record`        | `table`, `record`                          | —                                 |
| `update_row`           | `table`, `key`, `values`                   | `column`                          |
| `create_table`         | `anchor_id`, `rows`, `cols`                | `style`, `data` (row-major 2-D), `header`, `where` or `before: true`, `bind` (new cells default to the `Normal` paragraph style) |
| `delete_table`         | `table`                                    | —                                 |
| `insert_break`         | `anchor_id`                                | `kind` (`page`/`column`/`section_next`/`section_continuous`), `where` or `before: true` |
| `insert_field`         | `anchor_id`, `kind`                        | `text` (raw code for `kind: field`), `where` or `before: true` |
| `update_fields`        | —                                          | —                                 |
| `set_page_setup`       | `section`                                  | `margins`, `top_margin`, `bottom_margin`, `left_margin`, `right_margin`, `gutter`, `orientation`, `paper_size`, `columns`, `column_spacing` |
| `insert_footnote`      | `anchor_id`, `text`                        | `where` or `before: true`         |
| `insert_endnote`       | `anchor_id`, `text`                        | `where` or `before: true`         |
| `insert_toc`           | `anchor_id`                                | `levels` (`[upper, lower]`), `use_heading_styles`, `hyperlinks`, `where` or `before: true` |
| `mark_index_entry`     | `anchor_id`, `entry`                       | `cross_reference`, `bold`, `italic` |
| `insert_index`         | `anchor_id`                                | `columns`, `run_in`, `right_align_page_numbers`, `where` or `before: true` |
| `insert_table_of_figures` | `anchor_id`                             | `label`, `include_label`, `hyperlinks`, `right_align_page_numbers`, `where` or `before: true` |
| `set_bibliography_style` | `style`                                  | —                                 |
| `add_source`           | `source_type`                              | `tag`, `author`, `title`, `year`, `publisher`, `city`, `journal_name`, `volume`, `issue`, `pages`, `url`, `edition`, `doi`, or raw `xml` |
| `insert_citation`      | `anchor_id`, `tag`                         | `pages`, `prefix`, `suffix`, `volume`, `suppress_author`, `suppress_year`, `suppress_title`, `locale`, `where` or `before: true` |
| `insert_bibliography`  | `anchor_id`                                | `where` or `before: true` |
| `mark_citation`        | `anchor_id`, `long_citation`               | `short_citation`, `category`, `where` or `before: true` |
| `insert_table_of_authorities` | `anchor_id`                         | `category`, `passim`, `keep_entry_formatting`, `entry_separator`, `page_range_separator`, `where` or `before: true` |
| `add_bookmark`         | `name`, `anchor_id`                        | —                                 |
| `pin`                  | `anchor_id`                                | `name` (a readable slug)          |
| `pin_outline`          | —                                          | `levels` (an `[lo, hi]` band)     |
| `add_hyperlink`        | `anchor_id`, and one of `url` / `bookmark` | `text`, `screen_tip`              |
| `set_hyperlink`        | `index` (1-based, from `hyperlinks`)       | `address`, `sub_address`, `text`, `screen_tip` (pass ≥1; `sub_address`/`screen_tip` clear with `""`, `address`/`text` can't be emptied) |
| `insert_cross_reference` | `anchor_id`, `target`                    | `kind` (`text`/`page`/`number`/`above_below`), `hyperlink`, `where` or `before: true` |
| `insert_caption`       | `anchor_id`                                | `label`, `text`, `position` (`above`/`below`; default above for `Table`, else below) |
| `create_content_control` | `anchor_id`                              | `kind`, `title`, `tag`, `items`, `where` (`wrap`/`before`/`after`), `lock_contents`, `lock_control` |
| `set_cc_properties`    | `anchor_id` (`cc:NAME`)                    | `title`, `tag`, `lock_contents`, `lock_control` (pass ≥1; `""` clears `title`/`tag`) |
| `set_cc_items`         | `anchor_id` (`cc:NAME`), `items`           | — (replaces the combo_box/dropdown choice list) |
| `apply_theme`          | `theme`                                    | —                                 |
| `set_theme_colors`     | —                                          | `scheme`, `colors`                |
| `set_theme_fonts`      | —                                          | `scheme`, `major`, `minor`        |
| `set_property`         | `name`, `value`                            | `custom`                          |
| `delete_property`      | `name`                                     | —                                 |
| `set_variable`         | `name`, `value`                            | —                                 |
| `delete_variable`      | `name`                                     | —                                 |
| `add_comment`          | `anchor_id`, `text`                        | `author`                          |
| `resolve_comment`      | `index`                                    | —                                 |
| `delete_comment`       | `index`                                    | —                                 |
| `accept_revision`      | `index`                                    | —                                 |
| `reject_revision`      | `index`                                    | —                                 |
| `accept_all_revisions` | —                                          | `anchor_id` (scope to its range)  |
| `reject_all_revisions` | —                                          | `anchor_id`                       |
| `set_watermark`        | `text`                                     | `font`, `color`, `layout`, `semitransparent` |
| `remove_watermark`     | —                                          | —                                 |
| `insert_text_box`      | `anchor_id`, `text`                        | `width`, `height`, `wrap`, `font`, `size`, `bold`, `italic`, `alignment`, `fill`, `border`, `before`/`after` |
| `apply_list`           | `anchor_id`                                | `type` (`bulleted`/`numbered`/`outline`), `continue` |
| `apply_list_format`    | `anchor_id`, `levels`                      | `continue` (`levels` = per-level specs; see `list format`) |
| `remove_list`          | `anchor_id`                                | —                                 |
| `restart_numbering`    | `anchor_id`                                | —                                 |
| `indent_list`          | `anchor_id`                                | —                                 |
| `outdent_list`         | `anchor_id`                                | —                                 |
| `write_header`         | `section`, `text`                          | `which` (`primary`/`first`/`even`) |
| `write_footer`         | `section`, `text`                          | `which`                           |

The `find_replace` op mirrors `wordlive replace --find …` — fuzzy whitespace
+ smart-quote match, optional `in` anchor to scope it, and either `all` or
`occurrence` to handle multi-match. Ambiguous-match failures surface in the
batch response's `failure.matches` so the LLM can rewrite the op and retry. To
edit text **inside a table**, scope the replace to the cell anchor
(`"in": "table:N:R:C"`); a match resolved through a whole-document scope that
can't be verified raises a `replace_verification` failure rather than risk
overwriting the wrong cell.

`insert_paragraph` mirrors the `insert` command: a new paragraph relative to
any anchor, with placement defaulting to `after` and an optional `style` that's
validated before the batch mutates anything. Placement accepts either the
verbose `"where": "before"|"after"` or the boolean `"before": true` — the latter
mirrors the command's `--before`/`--after` flags, so the same intent encodes the
same way whether you type it or batch it. (`insert_image` accepts both forms too.)

`append` adds a new final **paragraph** at the very end of the document
(optional `style`, validated first) — no anchor to resolve, equivalent to an
`insert_paragraph` op targeting the `end` anchor. `append_inline` instead
**continues** the document's last paragraph and takes `text` only (no `style`).
`prepend` / `prepend_inline` are their start-of-document mirrors (the `start`
anchor). `append_paragraph` / `prepend_paragraph` remain as explicit synonyms
of `append` / `prepend`.

Any field an op doesn't recognise (a typo, or a `style` handed to an inline
append) is reported in a top-level `warnings` array on the batch result — the op
still runs, but the ignored field is surfaced rather than silently dropped, so a
successful-looking response can't mask a payload you got wrong.

`insert_image` mirrors `insert-image`. Supply the image with either a `path`
(read from disk) or `base64` (inline data — the natural choice in a JSON op,
with no command-line length limit). `wrap` is required; the optional fields
match the command's flags, including `block` (place the image on its own new
`Normal` line instead of in the anchor's text run). A bad image source surfaces
as the batch's `failure` with `type: "ImageSourceError"`.

`apply_style` and `format_paragraph` are the same as their dedicated CLI
verbs — the style must already exist in the document, indent and spacing
values are in points, alignment is one of `left`/`center`/`right`/`justify`.
`format_paragraph`'s `page_break_before` (a bool) forces or clears a
reflow-safe page break before the paragraph — the clean way to page-break a
style without a stray break character. The `keep_together`, `keep_with_next`,
and `widow_control` bools are the matching pagination controls (keep a
paragraph's lines together, keep it with the next paragraph, suppress widows/
orphans).

`regularize` runs the [`regularize`](#regularize) command inside the batch: it
applies the fixable `lint` findings (selected by the optional `rules` / `within`
fields) as one step of the batch's atomic undo, and returns its
`{applied, skipped, findings}` report in the batch outputs. With `dry_run: true`
it plans without writing — the natural lead-off op when you want to audit-then-fix
in a single round-trip.

`set_cell`, `add_row`, `delete_row`, and `set_heading_row` operate on tables by
1-based `table` index. `set_cell` is shorthand for a `replace` on a `table:N:R:C`
anchor; `add_row`'s optional `values` is a JSON array matched to columns;
`set_heading_row` marks a row (default 1) as a repeating header across pages.
All the table ops join the same atomic-undo scope as the rest of the batch.

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

`insert_field` inserts a self-updating field (`kind` = `page` / `numpages` /
`date` / `time` / `filename` / `author` / `title`, or `field` with a raw code in
`text`) at an `anchor_id` — put page numbers in a `footer:S:WHICH`. `update_fields`
(no args) recomputes the document's fields. `set_page_setup` mirrors `page-setup`,
taking a 1-based `section` plus any of `margins`, the per-side `*_margin` fields,
`gutter`, `orientation`, `paper_size`, `columns`, and `column_spacing`.

`insert_footnote` / `insert_endnote` attach a note to an `anchor_id` and report
the new `footnote:N` / `endnote:N` in the batch's `outputs`. `insert_toc` inserts
a table of contents (`levels` is a `[upper, lower]` pair, default `[1, 3]`);
follow it with an `update_fields` op so its page numbers populate.

`add_bookmark` names a range (the prerequisite for the rest). `pin` plants a
**durable handle** on an `anchor_id` and returns its `pin:CODE` (random, or a
readable `name` slug); `pin_outline` pins every heading at once, returning the
`{heading:N: pin:CODE}` map (idempotent; `levels` restricts the band). These are
the batch-side mirror of the `pin` / `pin-outline` verbs — reach for them when a
positional id would renumber mid-batch.

**Durable handles in a batch.** Add `bind: "slug"` (or `bind: true` for a random
code) to an `insert` / `insert_block` / `insert_section` / `insert_markdown` /
`create_table` op to mint a `pin:` on the new content — it comes back as `pin` in
that op's `outputs` entry. And any op field of the exact form `$ops[N].field` is
replaced with an earlier op's recorded output *before* the op runs, so a batch can
create then target without a round-trip — e.g. `create_table` at op 0, then
`{"op": "set_cell", "table": "$ops[0].table", …}`. A forward / unknown reference
fails that op (and the batch) cleanly.

`add_hyperlink`
links an `anchor_id` to exactly one of `url` (external) or `bookmark` (internal),
with optional `text`. `insert_cross_reference` references a `target` anchor id
(`bookmark:` / `heading:` / `footnote:` / `endnote:`) — an unresolvable target is
an `anchor_not_found` failure. `insert_caption` adds an auto-numbered caption in
its own `Caption`-styled paragraph (`label` defaults to `Figure`); `position`
overrides the default placement (above for a `Table`, below otherwise), and on a
`table:N:R:C` anchor the caption attaches to the whole table. Refresh
cross-reference page numbers with an `update_fields` op after the document
settles.

`set_hyperlink` retargets / relabels an existing link in place (addressed by its
1-based `index` from the `hyperlinks` reader) rather than delete-and-reinsert —
pass any of `address` / `sub_address` / `text` / `screen_tip`. It retargets, it
doesn't unlink: `sub_address` / `screen_tip` clear with `""`, but `address` /
`text` can't be emptied (Word keeps a link pointing somewhere).
Likewise `set_cc_properties` re-sets a content control's `title` / `tag` /
`lock_contents` / `lock_control`, and `set_cc_items` replaces a combo_box /
dropdown's choice list — both addressed by the control's `cc:NAME` anchor.

`add_comment`, `resolve_comment`, and `delete_comment` mirror the `comment`
verbs — `add_comment` attaches a side-channel annotation to an `anchor_id`
without touching the text, while `resolve_comment` / `delete_comment` take a
1-based `index`. Since deletes re-index, ordering matters within a batch.

`accept_revision` / `reject_revision` resolve the tracked change at a 1-based
`index`; `accept_all_revisions` / `reject_all_revisions` resolve them all
(optionally scoped to an `anchor_id`'s range). Like comment deletes, accepting /
rejecting re-indexes the rest, so order within a batch matters — the bulk ops are
safer when resolving several. `set_watermark` / `remove_watermark` stamp or clear
a text watermark behind every page, and `insert_text_box` drops a floating pull
quote at an `anchor_id` (these mirror the `watermark` / `insert-text-box` verbs).

`apply_list`, `apply_list_format`, `remove_list`, `restart_numbering`,
`indent_list`, and `outdent_list` mirror the `list` verbs — all take an
`anchor_id`, and `apply_list`'s optional `type` defaults to `bulleted`.
`apply_list_format` takes a `levels` array of per-level specs (the `list format`
verb) to author a custom numbered/bulleted list. `write_header` /
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

## `about`

```
wordlive --about      # or -A
wordlive --version    # or -v
```

`--about`/`-A` prints a colourful banner with the package **version**, author
(**Tom Villani, Ph.D.**), **license** (MIT), and the repository URL, then exits.
On a terminal the "word" half is blue and the "live" half a lighter cyan; piped
or redirected, the ANSI is stripped to clean ASCII. `--version`/`-v` prints just
`wordlive <version>`.
Both are eager top-level flags — no subcommand needed and Word is never touched.

```bash
$ wordlive --about

                       _ _ _
                      | | (_)
__      _____  _ __ __| | |___   _____
\ \ /\ / / _ \| '__/ _` | | \ \ / / _ \
 \ V  V / (_) | | | (_| | | |\ V /  __/
  \_/\_/ \___/|_|  \__,_|_|_| \_/ \___|

  Drive a running Microsoft Word instance with LLM agents

  version 0.16.0
  author  Tom Villani, Ph.D.
  license MIT
  repo    https://github.com/thomas-villani/wordlive
```

## `llm-help`

```
wordlive llm-help [--python]
```

Print a full **agent guide** — a bundled `SKILL.md` — to stdout: the anchor
model, every verb, image insertion, the `exec` batch format, and the exit-code
taxonomy. `wordlive --help` points an agent straight here, so a model can get
everything it needs in one call without an install step. Defaults to the **CLI**
guide; `--python` prints the **Python-API** (`import wordlive as wl`) guide
instead.

Unlike every other command, the output is raw Markdown rather than JSON (and is
unaffected by `--json/--text`) — it's documentation, exactly like `--help`
itself, meant to read cleanly into a model's context. The YAML frontmatter that
fronts the installed skill is stripped. Offline: it never touches Word.

```bash
$ wordlive llm-help
# wordlive (CLI)

`wordlive` drives a **running** Microsoft Word instance over COM (Windows only).
...
```

This is the same content [`install-skill`](#install-skill) writes to disk; reach
for `llm-help` when you just want it in context now, and `install-skill` when you
want coding tools to discover it on their own.

## `install-skill`

```
wordlive install-skill [--cli | --python | --both] [--system] [--force]
```

Install wordlive's bundled **agent skills** so LLM coding tools can pick up how
to drive it. wordlive ships two: `wordlive-cli` (the command-line workflow) and
`wordlive-python` (the `import wordlive as wl` API). By default only the **CLI**
skill is installed; pass `--python` for just the Python one, or `--both` for
both. They land under the current project at `./.agents/skills/<name>/SKILL.md`
(or `~/.agents/skills/<name>/` with `--system`). Offline — it never touches
Word, and refuses to clobber an existing file unless you pass `--force`.

```bash
$ wordlive install-skill
{"ok": true, "scope": "local", "installed": [
  {"kind": "cli", "name": "wordlive-cli", "path": ".../.agents/skills/wordlive-cli/SKILL.md", "bytes": 6172}]}

$ wordlive install-skill --both
{"ok": true, "scope": "local", "installed": [
  {"kind": "cli",    "name": "wordlive-cli",    "path": ".../.agents/skills/wordlive-cli/SKILL.md",    "bytes": 6172},
  {"kind": "python", "name": "wordlive-python", "path": ".../.agents/skills/wordlive-python/SKILL.md", "bytes": 7460}]}

$ wordlive install-skill --python --system --force
{"ok": true, "scope": "system", "installed": [
  {"kind": "python", "name": "wordlive-python", "path": "/home/you/.agents/skills/wordlive-python/SKILL.md", "bytes": 7460}]}
```

Failures: `1` if a target can't be written (every target is checked up front, so
a missing `--force` fails before anything is written).

## `install-mcp`

```
wordlive install-mcp [--client claude-desktop|claude-code] [--name NAME]
    [--directory DIR] [--config PATH] [--print] [--force]
```

Register wordlive's [MCP server](mcp.md) in an agent's config so a client can
drive your open document. It merges an `mcpServers.<name>` entry (default name
`wordlive`) that launches the stdio server with
`uvx --from "wordlive[mcp,snapshot]" wordlive-mcp` — no separate install step,
and the `snapshot` extra enables the vision tool. Offline: it only edits config,
never touches Word — restart the client to load the change.

- `--client claude-desktop` (default) writes the OS-specific
  `claude_desktop_config.json`; `--client claude-code` writes a project-local
  `./.mcp.json`.
- `--directory DIR` registers a **local checkout** via
  `uv run --directory DIR wordlive-mcp` (for development) instead of the PyPI
  `uvx` form.
- `--config PATH` targets a specific config file; `--print` just emits the JSON
  snippet (writing nothing) so you can paste it into any client.
- It refuses to overwrite an existing server entry unless you pass `--force`.

```bash
$ wordlive install-mcp --print
{
  "mcpServers": {
    "wordlive": {
      "command": "uvx",
      "args": ["--from", "wordlive[mcp,snapshot]", "wordlive-mcp"]
    }
  }
}

$ wordlive install-mcp
{"ok": true, "client": "claude-desktop", "path": ".../Claude/claude_desktop_config.json",
 "server": "wordlive", "action": "created", "entry": {"command": "uvx", "args": ["--from", "wordlive[mcp,snapshot]", "wordlive-mcp"]}}
```

Failures: `1` if the config can't be read/written, isn't a JSON object, or the
server entry already exists without `--force`. For the bundle (`.mcpb`) and a
full tool reference, see the [MCP server page](mcp.md).

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
