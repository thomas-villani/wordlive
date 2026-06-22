# Design

This page gives the condensed rationale. The canonical, longer design
document is [`spec.md`](https://github.com/thomas-villani/wordlive/blob/main/spec.md)
in the repo root; the roadmap is in
[`feature-plan.md`](https://github.com/thomas-villani/wordlive/blob/main/feature-plan.md).

## Why wordlive exists

There is no good Python library for **driving a live Microsoft Word
session**. The options today are:

| Library        | Target                              | Mechanism             |
| -------------- | ----------------------------------- | --------------------- |
| `python-docx`  | `.docx` file on disk                | OOXML I/O             |
| `docx-plus`    | `.docx` file on disk (docx extender)| OOXML I/O             |
| **`wordlive`** | **Running `winword.exe`**           | **COM (`pywin32`)**   |

File-side libraries can't help when the user has the document open — Word
holds the lock, and any change you make on disk is invisible until the user
closes and re-opens. COM is the only path. And raw `pywin32` is brutally
LLM-hostile: magic integer constants, untyped late-bound dispatch, modal
dialog footguns, STA threading rules.

`xlwings` exists for Excel. wordlive is the equivalent for Word, with one
extra goal: be *first-class* for LLM tool use, not retrofitted.

## Design principles

The four principles, in priority order:

1. **Politeness first.** Default behaviour preserves the user's `Selection`,
   view, and scroll. They keep editing alongside your script. Operations
   that *must* move the cursor say so explicitly
   ([`doc.go_to(...)`](python-api.md#wordlive.Document),
   [`scope.allow_cursor_move()`](python-api.md#wordlive.EditScope)).
2. **Semantic anchors over `Selection`.** Operations target named handles —
   bookmarks, content controls, headings — not the live cursor. Anchors
   are stable across edits and visible to an LLM as JSON strings; the
   cursor is neither.
3. **Atomic undo.** Every [`doc.edit()`](python-api.md#wordlive.Document)
   block opens a Word `UndoRecord`, so one Ctrl-Z reverts the whole
   intent. A 10-op `exec` script is one undo step, not ten.
4. **Structured I/O.** Reads return dataclasses / dicts; the CLI emits one
   JSON object per invocation; exit codes are deterministic. No string
   scraping anywhere in the pipeline. See the
   [Errors page](errors.md#cli-exit-codes) for the exit-code contract.

Underlying all four: an **escape hatch**. Every wrapper exposes `.com`. When
wordlive doesn't cover something, drop to raw COM rather than giving up.

## What's out of scope

- **Cross-platform support.** COM is Windows-only. We don't pretend
  otherwise.
- **Cloud co-authoring.** Microsoft Graph / WOPI is a different stack and a
  different problem.
- **Full Word object-model coverage.** Anything we don't cover is one
  `.com` access away.
- **Replacing `python-docx`.** Different surface, different problem.
- **Embedding the Word window as a child HWND.** Separate problem, out of
  scope.

## Architecture at a glance

```
your code / LLM
       │
       ▼
┌───────────────────────────────────────────────────┐
│  wordlive public API                              │
│    attach / connect  →  Word                      │
│                          │                        │
│                          ▼                        │
│                       Document                    │
│                          │                        │
│            ┌─────────────┼─────────────┐          │
│            ▼             ▼             ▼          │
│      bookmarks   content_controls   headings      │
│            │             │             │          │
│            ▼             ▼             ▼          │
│         Bookmark  ContentControl    Heading       │
│            └─────────────┴─────────────┘          │
│                          │                        │
│                          ▼                        │
│                  Anchor (text, set_text,          │
│                   insert_before/after, delete)    │
└───────────────────────────────────────────────────┘
                          │
                          ▼
              EditScope (UndoRecord + SelectionSnapshot)
                          │
                          ▼
        pywin32  →  Word.Application (COM, STA-threaded)
```

The library is intentionally flat: ~10 modules, no plugin system, no
hierarchy beyond Word → Document → Anchor.

## What comes next

The roadmap lives in
[`feature-plan.md`](https://github.com/thomas-villani/wordlive/blob/main/feature-plan.md).
The current release covers the politeness/anchors/EditScope core, the LLM-first
CLI, fuzzy find/replace (and fuzzy paragraph search), document-scoped styles +
paragraph formatting (with a `format_info` read mirror), tables (cells as
`table:N:R:C` anchors, plus row/column anchors `table:N:row:R` / `table:N:col:C`,
add/delete-column, merge/split cells, restyle, banding, and autofit), the
collaboration surface (review comments, scoped track-changes with
accept/reject, and arbitrary `range:START-END` anchors), document structure —
bullet/numbered lists (including custom multi-level list templates) and section
headers/footers (`header:S:WHICH` / `footer:S:WHICH` anchors), full paragraph
addressing (every paragraph is a `para:N` anchor — `doc.paragraphs`,
`outline --all`), and **durable handles** (`pin:` bookmarks that survive
renumbering). The content surface now spans **image insertion + restyle**
(inline `image:N` and the floating-shape model `shape:N` — text boxes, WordArt,
floating images, watermarks, with crop / rotate / wrap / z-order / group),
**Excel-backed charts** ([`anchor.insert_chart(...)`](python-api.md#wordlive.Anchor)
with a deep post-insert formatting surface — axes, trendlines, error bars,
series/point styling), **equations** (UnicodeMath / LaTeX / MathML),
**citations & bibliography, indexes, tables of figures/authorities**, **document
themes** (colours + fonts, theme-aware), **table creation / deletion**, **page /
column / section breaks**, and **page / section rendering to PNG** for vision
models ([`Document.snapshot`](python-api.md#snapshots), via the optional
`snapshot` extra). On the read/agent-ergonomics side it ships **Markdown / HTML
export** and a **token-budgeted whole-document read** (`doc.read(budget=…)`),
**checkpoint + diff** ("what changed this session"), and a **document linter +
regularizer** (`doc.lint` / `doc.regularize`). wordlive also ships two LLM-facing
**agent skills** — a CLI guide and an `import wordlive as wl` Python guide that
`wordlive install-skill` drops into `.agents/skills/` — and an **MCP server**
(`wordlive-mcp`, registered with `wordlive install-mcp` or the one-click `.mcpb`
bundle) that exposes the same surface as a handful of dispatch tools (see
[MCP](mcp.md)). Still ahead: a co-editing / change-watch surface built on event
sinks (`WindowSelectionChange`, `DocumentBeforeSave`) and an async wrapper
around the sync core.

## How it's tested

Because wordlive drives a real, stateful application over COM, the test suite is
layered so that most of it still runs on a Linux CI box while the parts that can
only be trusted against live Word are exercised on demand:

- **Unit tests** (the default `uv run pytest`) run against a `fake_word`
  COM fixture that quacks like `Word.Application`. They cover everything that
  doesn't need real Word — anchor resolution, fuzzy find/replace math, the
  `exec` op vocabulary, CLI argument parsing, and exit-code mapping — and run
  anywhere, including CI across Python 3.10–3.15.
- **Smoke tests** (`uv run pytest -m smoke`) attach to a running Word and assert
  *real* behaviour per feature — the gate that catches a wrong `Wd*` constant or
  a COM boundary crash a mock can't model. Windows + Word only.
- **End-to-end tests** (`uv run pytest -m e2e`) shell out to the actual CLI
  (`python -m wordlive …`) as a subprocess against live Word and walk one
  continuous document lifecycle — build via `exec` and individual verbs, read it
  back, save and export under the path gate, then close, reopen from disk, and
  verify the content survived the round-trip. This is the only layer that
  exercises the *whole* stack a user or LLM actually hits: argument parsing →
  COM → live Word → JSON on stdout → process exit code.

The `smoke` and `e2e` tiers are excluded from the default run (`addopts =
-m 'not smoke'`), so a plain `uv run pytest` stays fast and Word-free.

## Full design document

For the unabridged version — including the original motivation, the error
taxonomy in more detail, the rejected alternatives, and a list of open
questions — see
[`spec.md`](https://github.com/thomas-villani/wordlive/blob/main/spec.md) in
the repo root.
