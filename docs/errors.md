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
    ├── WordBusyError
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
A bookmark, content control, or heading you asked for doesn't exist — or a
`find`/`replace --find` pattern matched zero occurrences (in that case
`.kind == "find"` and `.name` is the search string). The exception always
carries both `.kind` and `.name`. **Retryable after refreshing the outline /
bookmark list or reading the current content** — the document may have
changed since you last looked.

### `StyleNotFoundError`
A paragraph or character style you asked for isn't defined in the document.
Subclass of [`AnchorNotFoundError`](#anchornotfounderror) — it shares the same
exit code (2) and the same retry guidance, and `except AnchorNotFoundError`
catches it too. `.kind` is always `"style"` and `.name` is the requested style
name. Raised by `Document.styles[name]`, `Anchor.apply_style(name)`, and
`Heading.insert_paragraph_after(text, style=name)`. **Retryable after reading
`doc.styles.list()`** to see what's actually defined.

### `AmbiguousMatchError`
A fuzzy `find_replace` matched more than one occurrence and the caller didn't
say `all=True` or pass an `occurrence`. The exception carries `.find` (the
search string) and `.matches` (a list of `{anchor_id, start, end, text}`
dicts) so an agent can pick a specific occurrence and retry. **Retryable** by
narrowing the call with `occurrence=N` or `all=True`.

### `WordBusyError`
Word rejected the COM RPC. This usually means a modal dialog is open (Save
As, Find & Replace, etc.) or Word is mid-operation. **Retryable** with
exponential back-off. The HRESULT is on `.hresult`; `.retryable` is always
`True` so callers can pattern-match generically.

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
| `1`  | `WordliveError` (default), `DocumentNotFoundError` | other / unclassified      | depends on cause      |
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
