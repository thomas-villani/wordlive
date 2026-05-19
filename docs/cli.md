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
| `--json/--text`  | `--json`    | Output format. `--text` falls back to a human-readable repr. |
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
| `2`  | Anchor not found       | `AnchorNotFoundError`      |
| `3`  | Word busy / modal      | `WordBusyError` (retryable) |
| `4`  | Word not running       | `WordNotRunningError`      |

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
the document. Failures: `2` heading not found, `3` Word busy.

## `replace --anchor-id ID --text "…"`

```
wordlive replace --anchor-id ID --text "..." [--doc DOC_NAME]
```

Replace the text at an anchor identified by [anchor ID](concepts.md#anchor-ids).
Works across all three anchor kinds — this is the general-purpose write.

```bash
$ wordlive replace --anchor-id heading:3 --text "Updated section text"
{"ok": true,
 "anchor_id": "heading:3",
 "anchor": {"kind": "heading", "name": "Context"}}
```

The response's `anchor.name` resolves the ID back to a human-readable name
(the heading text, bookmark name, or content control title) — useful for
logging.

Failures: `2` anchor not found, `3` Word busy.

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

| `op`                   | Required fields                            | Optional       |
| ---------------------- | ------------------------------------------ | -------------- |
| `write_bookmark`       | `name`, `text`                             | —              |
| `write_cc`             | `name`, `text`                             | —              |
| `insert_after_heading` | `heading`, `text`                          | `style`        |
| `replace`              | `anchor_id`, `text`                        | —              |

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
