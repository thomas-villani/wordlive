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
┌──────────────────────────────────────────────────┐
│ wordlive public API                              │
│   attach / connect  →  Word                      │
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
└──────────────────────────────────────────────────┘
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
CLI, fuzzy find/replace, and document-scoped styles + paragraph formatting.
Likely next steps: tables, comments and track-changes (collaboration features),
event sinks (`WindowSelectionChange`, `DocumentBeforeSave`), and an async wrapper
around the sync core.

## Full design document

For the unabridged version — including the original motivation, the error
taxonomy in more detail, the rejected alternatives, and a list of open
questions — see
[`spec.md`](https://github.com/thomas-villani/wordlive/blob/main/spec.md) in
the repo root.
