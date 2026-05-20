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
| `1`  | Other / unclassified   | `WordliveError` (default), `DocumentNotFoundError` |
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
wordlive outline [--doc DOC_NAME]
```

Heading outline of the target document, with addressable anchor IDs.

```bash
$ wordlive outline
[{"level": 1, "text": "Introduction", "anchor_id": "heading:1"},
 {"level": 2, "text": "Context",      "anchor_id": "heading:3"},
 {"level": 1, "text": "Risks",        "anchor_id": "heading:8"}]
```

This is the entry point for LLM workflows that need to discover what's
addressable in the document. The emitted `anchor_id` strings are exactly
what `replace`, `go-to`, and `exec` consume.

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

## `insert --after-heading "…" --text "…"`

```
wordlive insert --after-heading "Introduction" --text "..." [--style "Body Text"] [--doc DOC_NAME]
```

Insert a new paragraph immediately after the named heading.

```bash
$ wordlive insert --after-heading "Risks" --text "New risk identified."
{"ok": true, "after_heading": "Risks", "style": null}
```

`--style` is optional; if given it must be a Word style name that exists in
the document — the style is validated before the paragraph is inserted, so a
typo never partially mutates the document. Use `wordlive style list` to see
the available names. Failures: `2` heading not found or style not found, `3`
Word busy.

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
    [--alignment left|center|right|justify]
    [--left-indent POINTS] [--right-indent POINTS] [--first-line-indent POINTS]
    [--space-before POINTS] [--space-after POINTS]
    [--doc DOC_NAME]
```

Set paragraph-formatting properties on the anchor's range. At least one
formatting flag is required. Indent and spacing values are in **points** —
the unit Word's COM API uses natively for these fields. Atomic-undo.

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

## `exec --script ops.json`

```
wordlive exec --script ops.json [--doc DOC_NAME]
```

Apply a batch of operations under a single atomic-undo scope. This is the
most useful command for LLM tool-use: one round-trip per *intent*, not one
per *operation*.

Script shape:

```json
{
  "label": "Update report",
  "ops": [
    {"op": "write_bookmark",       "name": "Address",     "text": "123 Main St"},
    {"op": "write_cc",             "name": "Signatory",   "text": "Jane Doe"},
    {"op": "insert_after_heading", "heading": "Risks",    "text": "New risk paragraph.",
                                   "style":   "Body Text"},
    {"op": "replace",              "anchor_id": "heading:3", "text": "Updated section text"}
  ]
}
```

### Supported ops

| `op`                   | Required fields                            | Optional                          |
| ---------------------- | ------------------------------------------ | --------------------------------- |
| `write_bookmark`       | `name`, `text`                             | —                                 |
| `write_cc`             | `name`, `text`                             | —                                 |
| `insert_after_heading` | `heading`, `text`                          | `style`                           |
| `replace`              | `anchor_id`, `text`                        | —                                 |
| `find_replace`         | `find`, `text`                             | `in`, `all`, `occurrence`         |
| `apply_style`          | `anchor_id`, `name`                        | —                                 |
| `format_paragraph`     | `anchor_id`                                | `alignment`, `left_indent`, `right_indent`, `first_line_indent`, `space_before`, `space_after` |

The `find_replace` op mirrors `wordlive replace --find …` — fuzzy whitespace
+ smart-quote match, optional `in` anchor to scope it, and either `all` or
`occurrence` to handle multi-match. Ambiguous-match failures surface in the
batch response's `failure.matches` so the LLM can rewrite the op and retry.

`apply_style` and `format_paragraph` are the same as their dedicated CLI
verbs — the style must already exist in the document, indent and spacing
values are in points, alignment is one of `left`/`center`/`right`/`justify`.

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
    "op": {"op": "insert_after_heading", "heading": "Risks", "text": "…"},
    "error": "heading not found: 'Risks'",
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
