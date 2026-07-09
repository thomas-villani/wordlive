"""Shared MCP helpers: error translation, param access, the instructions blob."""

from __future__ import annotations

import json
from typing import Any

from ..exceptions import OpError, WordliveError, classify

_INSTRUCTIONS = """\
wordlive drives the Microsoft Word document open right now on this Windows machine
(over COM). Edits are *polite*: the user's cursor, selection, and scroll are
preserved, and each write is a single Ctrl-Z.

Orient yourself first: `word_read(command="status")` to confirm Word is reachable,
then `word_read(command="outline")` for the heading tree (heading:N ids) or
`command="paragraphs"` for every paragraph (para:N ids + offsets), and
`command="find"` to locate text (returns range:START-END ids).

Address content by ANCHOR id, never the live cursor:
  heading:N · para:N · bookmark:NAME · cc:NAME · table:N:R:C · range:START-END ·
  header:S:WHICH · footer:S:WHICH · start · end

Make single edits with `word_write` (dispatch on `command`); batch several into one
atomic undo with `word_exec(ops=[...])`. Render a page or section to an image with
`word_snapshot`. For the full op vocabulary, anchor model, and field reference,
call `word_read(command="guide")` first (it needs neither Word nor a document) —
the same text is also the `wordlive://guide` resource where your client surfaces it.
""".strip()


def _error_payload(exc: WordliveError) -> dict[str, Any]:
    """Build the structured error body for a failed tool call."""
    code, retryable = classify(exc)
    payload: dict[str, Any] = {
        "error": str(exc),
        "code": code,
        "retryable": retryable,
        "type": type(exc).__name__,
    }
    matches = getattr(exc, "matches", None)
    if matches is not None:
        payload["matches"] = matches
    return payload


def _tool_error(exc: WordliveError, **extra: Any):  # noqa: ANN201 — needs the extra
    """Wrap a WordliveError as an MCP ToolError carrying a JSON error payload."""
    from mcp.server.fastmcp.exceptions import ToolError

    payload = _error_payload(exc)
    payload.update(extra)
    return ToolError(json.dumps(payload, ensure_ascii=False))


def _need(p: dict[str, Any], key: str, command: str) -> Any:
    """Return p[key] or raise a clean OpError naming the command and field."""
    value = p.get(key)
    if value is None:
        raise OpError(f"command {command!r} requires {key!r}")
    return value


def _image_format(mime: str | None) -> str:
    """The FastMCP `Image(format=...)` token for an `image/*` MIME type.

    `read_image` only yields raster types (PNG/JPEG/GIF/BMP/TIFF), so the subtype
    after the slash is the format token Office and FastMCP both expect — `image/
    jpeg` → ``"jpeg"``. Falls back to ``"png"`` for anything unexpected."""
    if mime and "/" in mime:
        return mime.rsplit("/", 1)[-1] or "png"
    return "png"
