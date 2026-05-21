# wordlive — Spec (working draft)

> Working name: `wordlive`. Final name TBD (see Open Questions).
> Status: sketch — written to seed a separate repo, not to live here.

## Overview

A small Python library + CLI for **driving a running Microsoft Word instance**
from Python, designed for both human scripting and LLM agents. Sibling to (not
part of) `docx-plus`:

| Library      | Target                                | Mechanism                |
| ------------ | ------------------------------------- | ------------------------ |
| python-docx  | `.docx` file on disk                  | OOXML I/O                |
| docx-plus    | `.docx` file on disk (extends docx)   | OOXML I/O                |
| **wordlive** | **Running Word.exe**                  | **COM automation (pywin32)** |

Windows-only by nature. Use it when the user already has the document open and
you want to edit it *live* — or when you want an LLM agent to collaborate with
a human inside the same Word session.

## Motivation

1. **xlwings exists; nothing equivalent for Word does.** Driving Word from
   Python today means raw pywin32: no type hints, magic integer constants,
   late-bound string lookups for everything, STA threading footguns.
2. **LLM agents need a small, semantic surface.** Exposing the full Word
   object model to a model is hopeless. A focused set of structured,
   idempotent operations is exactly what tool use wants.
3. **File-level libraries can't help when the user is editing.** If Word has
   the document open, python-docx / docx-plus can't safely touch it on disk.
   COM is the only path.
4. **"Polite" editing is a real engineering problem.** Naïve scripts stomp
   the user's selection, scroll position, undo stack, and focus. A wrapper
   that handles this once is broadly valuable.

## Non-Goals

- **Cross-platform.** COM is Windows-only; we don't pretend otherwise.
- **Multi-user cloud co-authoring.** That's WOPI / Microsoft Graph — totally
  different stack.
- **Full Word object-model coverage.** We expose what's needed for common
  edits; raw COM stays accessible as an escape hatch.
- **Replacing python-docx / docx-plus.** Different surface, different problem.
- **Embedding the Word window as a child HWND.** Separate problem, out of
  scope.

## Design Principles

1. **Politeness first.** Default behaviour preserves the user's `Selection`,
   view, and scroll. Operations that *must* move the cursor say so in their
   name.
2. **Semantic anchors over `Selection`.** Operations target bookmarks,
   content controls, named ranges, or heading-anchored ranges — never the
   live `Selection` unless explicitly requested.
3. **Idempotent where possible.** `set_bookmark_text("Address", "...")` is
   repeatable; `append_paragraph(...)` is not (and is documented as such).
4. **Atomic undo.** Every public operation opens an `UndoRecord` so one
   Ctrl+Z reverts the whole thing cleanly.
5. **Structured I/O.** Reads return dataclasses/dicts; CLI emits JSON. No
   string scraping.
6. **Synchronous core, optional event hooks.** COM is STA — we don't fight
   it. Events surface via a thread-safe callback layer; user code stays sync.
7. **Escape hatch always available.** Every wrapper exposes a `.com` property
   for the raw COM object, so users are never blocked by missing coverage.

## Architecture

### Attachment model

```python
import wordlive

# attach to a running instance; raises if none
with wordlive.attach() as word:
    ...

# attach or launch
with wordlive.connect(launch_if_missing=True, visible=True) as word:
    ...
```

- Uses `GetActiveObject("Word.Application")` first, falls back to `Dispatch`
  if `launch_if_missing=True`.
- Context manager handles `pythoncom.CoInitialize` / `CoUninitialize`.
- Releases COM references on exit; **never** closes Word (it's the user's app).

### Threading

- All operations require an STA thread; the library calls `CoInitialize` on
  entry to the context.
- Event callbacks are pumped on a dedicated worker thread and dispatched to a
  callback registry, so user code never has to think about STA.

### Error taxonomy

| Exception                | Meaning                                                 |
| ------------------------ | ------------------------------------------------------- |
| `WordNotRunningError`    | no instance, `launch_if_missing=False`                  |
| `DocumentNotFoundError`  | named or active document missing                        |
| `AnchorNotFoundError`    | bookmark / content control / heading not present; also raised for zero matches in fuzzy `find_replace` |
| `AmbiguousMatchError`    | fuzzy `find_replace` matched more than one occurrence without disambiguation |
| `WordBusyError`          | Word in modal dialog or rejected the RPC (retryable)    |
| `ComError`               | generic wrap of `pywintypes.com_error` w/ decoded HRESULT |

## Python API Sketch

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active              # or word.documents["Report.docx"]

    # --- READS (structured, side-effect-free) ---
    outline   = doc.outline()                # [{level, text, anchor_id}, ...]
    bookmarks = doc.bookmarks.list()
    cc_value  = doc.content_controls["Date"].text
    sel_info  = doc.selection.info()         # range, page, line, etc.

    # --- ANCHORS (cheap, lazy) ---
    intro = doc.heading("Introduction")      # raises AnchorNotFoundError
    addr  = doc.bookmarks["Address"]
    cc    = doc.content_controls["Signatory"]

    # --- POLITE WRITES (preserves Selection) ---
    with doc.edit("Update address block"):   # opens an UndoRecord
        addr.set_text("123 Main St\nAnytown")
        cc.set_text("Jane Doe")
        intro.insert_paragraph_after("New context paragraph.", style="Body Text")

    # --- EXPLICITLY MOVE THE USER (rare, opt-in) ---
    doc.go_to(intro, scroll=True)

    # --- RAW ESCAPE HATCH ---
    doc.com.Range(0, 10).Font.Bold = True
```

### Key abstractions

- **`Word`** — application handle.
- **`Document`** — wraps a `Document` COM object; exposes `bookmarks`,
  `content_controls`, `headings`, `sections`, `selection`, `outline()`,
  `edit()`.
- **`Anchor`** — base type for `Bookmark`, `ContentControl`, `Heading`,
  `NamedRange`. All anchors support `text`, `set_text`, `insert_before`,
  `insert_after`, `delete`.
- **`EditScope`** — context manager opened by `doc.edit(label)`. Bundles an
  `UndoRecord` and a `Selection` / scroll-position snapshot+restore.

## CLI Sketch

CLI is **LLM-first**: JSON in, JSON out, deterministic exit codes, every
command idempotent or clearly marked otherwise.

```
wordlive status                            # which docs are open, which active
wordlive outline [--doc NAME]              # JSON: [{level, text, anchor_id}]
wordlive read bookmark NAME [--doc NAME]
wordlive read cc NAME [--doc NAME]
wordlive read section HEADING              # body under a heading
wordlive write bookmark NAME --text "..."
wordlive write cc NAME --text "..."
wordlive outline [--all]                   # --all lists every paragraph (para:N)
wordlive paragraphs                        # every paragraph: para:N, level, offsets, text
wordlive insert --anchor-id ID --text "..." [--before|--after] [--style "Body Text"]
wordlive cursor read                       # explicit, opt-in cursor surface
wordlive cursor write --text "..." [--no-replace]
wordlive find --text "..." [--in ANCHOR_ID]
wordlive replace --anchor-id ID --text "..."
wordlive replace --find OLD --text NEW [--in ID] [--all|--occurrence N]
wordlive exec --script ops.json            # batch a list of ops in one UndoRecord
```

Conventions:

- `--json` (default) or `--text` for the human / piping case (each command
  emits its own format: indented outline tree, bare text for reads, one-line
  acks for writes).
- Exit codes: `0` ok, `2` anchor-not-found (incl. zero `find` matches), `3`
  Word-busy, `4` Word-not-running, `5` ambiguous match, `1` other.
- One JSON object on stdout per invocation; logs go to stderr.

This makes wiring it up as an LLM tool trivial:

```json
{ "tool": "wordlive", "args": ["insert", "--anchor-id", "heading:8", "--text", "..."] }
```

## Key Technical Concerns

- **`UndoRecord`** — `Application.UndoRecord` is exactly the primitive we
  need; wrap it in `EditScope`.
- **Selection preservation** — snapshot `Selection.Range.Start`/`.End` and
  `ActiveWindow.View` scroll position on enter; restore on exit unless the
  operation explicitly moved them.
- **Event sinks** — `WithEvents(word.com, Handler)` for
  `DocumentBeforeSave`, `WindowSelectionChange`, etc. Marshal events to a
  callback registry so user code stays sync.
- **Modal dialog handling** — COM calls fail with specific HRESULTs when
  Word is in a modal dialog; surface as `WordBusyError` with optional retry.
- **`Selection` vs `Range`** — anchors always operate on `Range`, never
  `Selection`. Documented loudly.
- **Magic constants** — ship a typed enum module
  (`wl.constants.WdParagraphAlignment.CENTER`) so users never see
  `1`/`2`/`3` literals.

## Open Questions

1. **Name.** Working name `wordlive`. Alternatives: `wordwings` (xlwings
   parallel, instantly readable, but brand-association risk), `livedocx`,
   `docx-live`. Decide before first commit — affects package name, import
   name, CLI binary.
2. **Async API.** Do we offer an `asyncio` wrapper around the sync core?
   Probably yes eventually (events fit naturally), but v0 stays sync.
3. **Read-model caching.** Should `doc.outline()` be cached with
   invalidation on `WindowSelectionChange`/`DocumentChange`, or always live?
   Live is simpler; caching matters for LLM workflows that re-read often.
4. **Multi-document scope.** First-class support for multiple open
   documents, or assume single-active-doc and require explicit naming
   otherwise?
5. **Test strategy.** COM tests need Word installed. Pattern: small smoke
   suite on a Windows CI runner with Word, plus a mockable wrapper layer
   for unit tests of the politeness logic.

## Inspirations / Prior Art

- **xlwings** — API ergonomics for COM-driven Office automation.
- **pywin32 / comtypes** — the underlying COM layer; wordlive sits on top.
- **python-docx / docx-plus** — file-side counterparts; shape the
  abstraction (anchors, paragraphs, styles) similarly where possible so
  users moving between them feel at home.

## v0 Scope (suggested)

Minimum viable cut for a useful first release:

- `attach` / `connect` / context manager
- `Document.outline()`, `bookmarks`, `content_controls`, `headings`
  (read + `set_text`)
- `doc.edit()` with `UndoRecord` + Selection preservation
- CLI: `status`, `outline`, `read`, `write` (bookmark/cc), `insert
  --after-heading`
- Typed exceptions, magic-constant enums
- Smoke tests against a real Word install

Defer to later: events, async, raw range manipulation, styles/numbering
edits, full Sections/Headers/Footers API, `exec` batch script.
