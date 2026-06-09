# Errors & exit codes

wordlive translates pywin32's `pywintypes.com_error` into a small, typed
exception hierarchy. The CLI maps those exceptions to deterministic exit
codes so LLM tool-use loops can branch on the failure mode without parsing
error text.

## Exception hierarchy

```
Exception
└── WordliveError
    ├── WordNotRunningError
    ├── DocumentNotFoundError
    ├── AnchorNotFoundError
    │   └── StyleNotFoundError
    ├── AmbiguousMatchError
    ├── ReplaceVerificationError
    ├── ImageSourceError
    ├── PathNotAllowedError
    ├── SnapshotError
    ├── WordBusyError
    ├── OpError
    └── ComError
```

`WordliveError` is the catch-all base — `except wordlive.WordliveError`
catches every typed error wordlive raises. Anything that wasn't a COM error
in the first place (e.g. a `ValueError` from your own code) bubbles up
unchanged.

## Reference

### `WordliveError`
Base class. Catch this if you want one `try` for every wordlive failure.

### `WordNotRunningError`
No Word instance is running. Raised by [`attach()`](python-api.md#wordlive.attach)
and by [`connect(launch_if_missing=False)`](python-api.md#wordlive.connect).
**Not retryable** within a session — Word has to actually be running.

### `DocumentNotFoundError`
The requested document isn't open. Raised by `word.documents[name]` and by
`word.documents.active` when no document is active. The missing name is on
the exception's `.name` attribute.

### `AnchorNotFoundError`
A bookmark, content control, heading, paragraph, table cell, comment, range,
list, section, or header/footer you asked for doesn't exist — or a
`find`/`replace --find` pattern matched zero occurrences (in that case
`.kind == "find"` and `.name` is the search string). `.kind` names the thing
that was missing (`"bookmark"`, `"heading"`, `"paragraph"`, `"table cell"`,
`"comment"`, `"range"`, `"list"`, `"section"`, `"header"`, `"footer"`, …) and
`.name` is what you asked for.
**Retryable after refreshing the outline / bookmark list or reading the current
content** — the document may have changed since you last looked.

### `StyleNotFoundError`
A paragraph or character style you asked for isn't defined in the document.
Subclass of [`AnchorNotFoundError`](#anchornotfounderror) — it shares the same
exit code (2) and the same retry guidance, and `except AnchorNotFoundError`
catches it too. `.kind` is always `"style"` and `.name` is the requested style
name. Raised by `Document.styles[name]`, `Anchor.apply_style(name)`, and
`Anchor.insert_paragraph_before/after(text, style=name)`. **Retryable after
reading `doc.styles.list()`** to see what's actually defined. To MCP clients it
surfaces a distinct `code: "style_not_found"` (the CLI exit code is still `2`),
so a missing style is told apart from a missing bookmark/heading.

### `AmbiguousMatchError`
A fuzzy `find_replace` matched more than one occurrence and the caller didn't
say `all=True` or pass an `occurrence`. The exception carries `.find` (the
search string) and `.matches` (a list of `{anchor_id, start, end, text}`
dicts) so an agent can pick a specific occurrence and retry. **Retryable** by
narrowing the call with `occurrence=N` or `all=True`.

### `ReplaceVerificationError`
A fuzzy `find_replace` resolved a write target whose text didn't match what was
located, so wordlive refused to write rather than corrupt the document. This
guards the case where `Range.Text` offsets diverge from Word's document positions
across table structure — a whole-document replace could otherwise overwrite a
neighbouring cell while returning success. Carries `.find`, `.expected` (the
located text), `.resolved` (what the target actually held), and `.anchor_id`. It
maps to the generic exit code (1) and `code: "replace_verification"` for MCP
clients. **Not retryable as-is** — re-scope the replace to the cell anchor
(`scope=doc.anchor_by_id("table:N:R:C")`), which addresses one cell at a time.

### `ImageSourceError`
The image handed to [`insert_image`](python-api.md#wordlive.Anchor) couldn't be
turned into an embeddable file: a missing or unreadable path, malformed base64,
or bytes whose format isn't a recognised raster image (PNG/JPEG/GIF/BMP/TIFF).
It's a *bad-input* error — not a missing named thing — so it maps to the
generic exit code (1) rather than reusing the anchor-not-found code.
**Not retryable**: fix the input.

### `PathNotAllowedError`
A filesystem path was refused by the gated CLI / MCP surface's default-deny
policy. The Python API is trusted and ungated, but the CLI and MCP surfaces —
whose inputs can be prompt-injected — run every path through a whitelist. Raised
when a **save / save-as / export-pdf** target falls outside the configured
save-directory whitelist (or none is set, so saving is off), or when an
**image-source path** is a non-local form (UNC `\\…`, `file://`, a URL) or sits
outside the optional image-directory allowlist. It's a policy denial / bad-input
error, so it maps to the generic exit code (1) and surfaces `code:
"path_not_allowed"` to MCP clients. **Not retryable**: configure a whitelist
(`--save-dir` / `WORDLIVE_SAVE_DIRS`, `--image-dir` / `WORDLIVE_IMAGE_DIRS`) or
pass a local path inside it.

### `SnapshotError`
A page/section [`snapshot`](python-api.md#snapshots) couldn't be rendered —
almost always because the optional PDF backend (PyMuPDF) isn't installed, or
because rasterising the exported PDF failed. The PDF export itself goes through
Word's COM, so a busy/modal Word surfaces as [`WordBusyError`](#wordbusyerror),
not this. It's an environment/dependency problem, so it maps to the generic exit
code (1). Fix by installing the extra: `pip install "wordlive[snapshot]"` (or
`uv add "wordlive[snapshot]"`). **Not retryable** until the backend is present.

### `WordBusyError`
Word rejected the COM RPC. This usually means a modal dialog is open (Save
As, Find & Replace, etc.) or Word is mid-operation. **Retryable** with
exponential back-off. The HRESULT is on `.hresult`; `.retryable` is always
`True` so callers can pattern-match generically.

### `OpError`
A batch/`exec` op — or a single dispatched write — was malformed: an unknown op
kind, a missing required field, or a mutually exclusive pair given together
(e.g. both `path` and `base64` for an image, or both `text` and `runs` on an
`insert`). It's a bad-input error — fix the request — so it maps to the generic
exit code (1) and `code: "error"` for MCP clients. **Not retryable** as-is.

### `ComError`
Catch-all for any other classified COM error. Carries `.hresult` and
`.description` (when pywin32 surfaces one). Not retryable in general; treat
as a bug in your code or a Word-side problem.

## HRESULT mapping

Only one HRESULT family is special-cased: the "Word is momentarily
unavailable" codes that map to [`WordBusyError`](#wordbusyerror). Everything
else becomes a generic [`ComError`](#comerror) with the HRESULT preserved.

| HRESULT       | Mnemonic                         | wordlive exception |
| ------------- | -------------------------------- | ------------------ |
| `0x80010001`  | `RPC_E_CALL_REJECTED`            | `WordBusyError`    |
| `0x80010005`  | `RPC_E_SERVERCALL_REJECTED`      | `WordBusyError`    |
| `0x8001010A`  | `RPC_E_SERVERCALL_RETRYLATER`    | `WordBusyError`    |
| any other     | —                                | `ComError`         |

The classification logic lives in
[`src/wordlive/exceptions.py:99`](https://github.com/thomas-villani/wordlive/blob/main/src/wordlive/exceptions.py#L99).
If you find a code that should be treated as busy/retryable, it goes in the
`_BUSY_HRESULTS` set in that file.

## CLI exit codes

The CLI maps the exception hierarchy onto six exit codes, defined in
[`src/wordlive/cli/main.py`](https://github.com/thomas-villani/wordlive/blob/main/src/wordlive/cli/main.py):

| Exit | Exception(s)                                | Meaning                          | Retry?                |
| ---- | ------------------------------------------- | -------------------------------- | --------------------- |
| `0`  | —                                           | success                          | —                     |
| `1`  | `WordliveError` (default), `DocumentNotFoundError`, `ImageSourceError`, `PathNotAllowedError`, `SnapshotError`, `ReplaceVerificationError`, `OpError` | other / unclassified | depends on cause |
| `2`  | `AnchorNotFoundError`, `StyleNotFoundError`  | bookmark / cc / heading / style missing, or `find` had zero matches | yes, after re-reading content |
| `3`  | `WordBusyError`                              | modal dialog or busy RPC         | **yes**, with back-off |
| `4`  | `WordNotRunningError`                        | no Word instance                 | only if user launches Word |
| `5`  | `AmbiguousMatchError`                        | `replace --find` matched more than one occurrence | **yes**, after picking `--occurrence N` or passing `--all` |

## Retry guidance

The only exception explicitly designed to be retryable is
[`WordBusyError`](#wordbusyerror). A typical retry loop:

```python
import time
import wordlive as wl

def with_retry(fn, *, attempts=4, base=0.5):
    for i in range(attempts):
        try:
            return fn()
        except wl.WordBusyError:
            if i == attempts - 1:
                raise
            time.sleep(base * (2 ** i))   # 0.5, 1, 2, 4 seconds


def update_address():
    with wl.attach() as word:
        doc = word.documents.active
        with doc.edit("Update address"):
            doc.bookmarks["Address"].set_text("123 Main St")


with_retry(update_address)
```

For the CLI:

```bash
for i in 1 2 3 4; do
    wordlive write bookmark Address --text "123 Main St" && break
    rc=$?
    [ "$rc" = "3" ] || exit "$rc"      # only retry exit code 3
    sleep $((i * i))                   # quadratic-ish back-off
done
```

[`AnchorNotFoundError`](#anchornotfounderror) is *also* effectively retryable
— but only after you've re-read `outline()` or the bookmark list, since the
document state has demonstrably changed since your last call.
