"""The wordlive MCP server: four dispatch tools over a COM worker thread.

Tools (all prefixed `word_`):
  - `word_read`     — every read, dispatched on `command`.
  - `word_write`    — every single atomic-undo write, dispatched on `command`.
  - `word_exec`     — apply a batch of ops as one atomic undo (the power tool).
  - `word_snapshot` — render page(s) to PNG so the model can *see* the layout.

Plus a `wordlive://guide` resource holding the full agent guide.

A handful of dispatch tools (rather than one tool per verb) keeps the client's
tool list — and the context cost of its schemas — small. The op vocabulary for
`word_exec` and the anchor model are taught by the guide resource, not by dozens
of schemas.

All Word access funnels through a single `ComWorker` thread (see `_worker`), so
COM stays on one apartment-initialised thread and concurrent calls serialise.
`WordliveError`s are translated to MCP tool errors carrying a stable `code` and
a `retryable` hint (the same taxonomy as the CLI's exit codes).
"""

from __future__ import annotations

import base64
import json
from typing import TYPE_CHECKING, Any, Literal

from .. import attach
from .._anchors import Heading
from .._guide import skill_body
from .._ops import _PAGE_SETUP_FIELDS, _PARA_FIELDS, _STYLE_RUN_FIELDS, pick_doc, run_batch
from .._paths import PathPolicy
from ..exceptions import OpError, WordliveError, classify
from ._worker import ComWorker, Worker

if TYPE_CHECKING:
    from mcp.server.fastmcp import FastMCP


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


# ---------------------------------------------------------------------------
# Error translation
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# Implementations (worker-bound, MCP-type-free — unit-testable directly)
# ---------------------------------------------------------------------------


def _need(p: dict[str, Any], key: str, command: str) -> Any:
    """Return p[key] or raise a clean OpError naming the command and field."""
    value = p.get(key)
    if value is None:
        raise OpError(f"command {command!r} requires {key!r}")
    return value


def _read_impl(worker: Worker, command: str, p: dict[str, Any]) -> Any:
    if command == "guide":
        # The full agent guide — the same text served by the wordlive://guide
        # resource, but reachable as a tool call (resources aren't surfaced by
        # every MCP client). Needs neither Word nor the worker thread.
        return {"guide": skill_body()}

    def job() -> Any:
        with attach() as word:
            if command == "status":
                return word.documents.list()
            doc = pick_doc(word, p.get("doc"))
            if command == "outline":
                return doc.paragraphs.list() if p.get("all_paragraphs") else doc.outline()
            if command == "paragraphs":
                items = doc.paragraphs.list()
                start, count = p.get("start"), p.get("count")
                if start is not None or count is not None:
                    s = (start or 1) - 1
                    items = items[s : s + count] if count is not None else items[s:]
                return items
            if command == "find":
                text = p.get("text")
                if text is None:
                    raise OpError("read command 'find' requires 'text'")
                scope = doc.anchor_by_id(p["in_anchor"]) if p.get("in_anchor") else None
                return doc.find(text, scope=scope)
            if command == "read_bookmark":
                return {"text": doc.bookmarks[_need(p, "name", command)].text}
            if command == "read_cc":
                return {"text": doc.content_controls[_need(p, "name", command)].text}
            if command == "read_section":
                anchor_id, heading = p.get("anchor_id"), p.get("heading")
                if (anchor_id is None) == (heading is None):
                    raise OpError(
                        "read command 'read_section' requires exactly one of "
                        "'anchor_id' or 'heading'"
                    )
                if anchor_id is not None:
                    h = doc.anchor_by_id(anchor_id)
                    if not isinstance(h, Heading):
                        raise OpError("'anchor_id' must reference a heading")
                else:
                    assert heading is not None  # guaranteed by the exactly-one check above
                    h = doc.heading(heading)
                return {
                    "heading": h.text,
                    "anchor_id": h.anchor_id,
                    "level": h.level,
                    "text": h.section_text(),
                }
            if command == "table_list":
                return doc.tables.list()
            if command == "table_read":
                return doc.tables[_need(p, "table", command)].read()
            if command == "table_records":
                return doc.tables[_need(p, "table", command)].records()
            if command == "location":
                anchor_id = _need(p, "anchor_id", command)
                return {"anchor_id": anchor_id, **doc.anchor_by_id(anchor_id).location()}
            if command == "stats":
                return doc.stats()
            if command == "styles":
                return doc.styles.list()
            if command == "comments":
                return doc.comments.list()
            if command == "revisions":
                return doc.revisions.list()
            if command == "track":
                return {"track_changes": doc.track_changes}
            if command == "sections":
                return doc.sections.list()
            if command == "footnotes":
                return doc.footnotes.list()
            if command == "endnotes":
                return doc.endnotes.list()
            if command == "images":
                return doc.images.list()
            if command == "read_image":
                anchor_id = _need(p, "anchor_id", command)
                data, mime = doc.anchor_by_id(anchor_id).read_image()
                return {
                    "anchor_id": anchor_id,
                    "mime": mime,
                    "bytes": len(data),
                    "base64": base64.b64encode(data).decode("ascii"),
                }
            raise OpError(f"unknown read command: {command!r}")

    return worker.run_on_word(job)


def _build_write_op(command: str, p: dict[str, Any]) -> dict[str, Any]:
    """Translate a `word_write` command + params into a single `apply_op` op dict."""

    def need(key: str) -> Any:
        return _need(p, key, command)

    if command == "insert":
        # Either `text` (literal) or `runs` (inline-formatted spans) — exactly one.
        if (p.get("text") is None) == (p.get("runs") is None):
            raise OpError("insert requires exactly one of 'text' or 'runs'")
        op = {
            "op": "insert_paragraph",
            "anchor_id": need("anchor_id"),
            "before": bool(p.get("before", False)),
        }
        if p.get("runs") is not None:
            op["runs"] = p["runs"]
        else:
            op["text"] = p["text"]
        if p.get("style") is not None:
            op["style"] = p["style"]
        return op
    if command == "insert_block":
        return {
            "op": "insert_block",
            "anchor_id": need("anchor_id"),
            "items": need("items"),
            "before": bool(p.get("before", False)),
        }
    if command == "insert_section":
        op = {
            "op": "insert_section",
            "anchor_id": need("anchor_id"),
            "heading": need("heading"),
            "body": need("body"),
            "before": bool(p.get("before", False)),
        }
        if p.get("level") is not None:
            op["level"] = p["level"]
        return op
    if command == "insert_markdown":
        return {
            "op": "insert_markdown",
            "anchor_id": need("anchor_id"),
            "markdown": need("markdown"),
            "before": bool(p.get("before", False)),
        }
    if command == "replace_section":
        if (p.get("body") is None) == (p.get("markdown") is None):
            raise OpError("replace_section requires exactly one of 'body' or 'markdown'")
        op = {"op": "replace_section", "anchor_id": need("anchor_id")}
        if p.get("markdown") is not None:
            op["markdown"] = p["markdown"]
        else:
            op["body"] = p["body"]
        return op
    if command == "delete_paragraph":
        return {"op": "delete_paragraph", "anchor_id": need("anchor_id")}
    if command in ("append", "prepend"):
        text = need("text")
        if p.get("paragraph", True):
            op = {"op": f"{command}_paragraph", "text": text}
            if p.get("style") is not None:
                op["style"] = p["style"]
            return op
        # paragraph=False → the inline "continue the adjacent paragraph" variant.
        return {"op": f"{command}_inline", "text": text}
    if command == "replace":
        text = need("text")
        if p.get("find") is not None:
            op = {"op": "find_replace", "find": p["find"], "text": text, "all": bool(p.get("all"))}
            if p.get("occurrence") is not None:
                op["occurrence"] = p["occurrence"]
            if p.get("in_anchor") is not None:
                op["in"] = p["in_anchor"]
            return op
        return {"op": "replace", "anchor_id": need("anchor_id"), "text": text}
    if command == "write_bookmark":
        return {"op": "write_bookmark", "name": need("name"), "text": need("text")}
    if command == "write_cc":
        return {"op": "write_cc", "name": need("name"), "text": need("text")}
    if command == "apply_style":
        return {"op": "apply_style", "anchor_id": need("anchor_id"), "name": need("name")}
    if command == "format_paragraph":
        op = {"op": "format_paragraph", "anchor_id": need("anchor_id")}
        for k in (
            "alignment",
            "left_indent",
            "right_indent",
            "first_line_indent",
            "space_before",
            "space_after",
            "page_break_before",
            "keep_together",
            "keep_with_next",
            "widow_control",
        ):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "format_run":
        op = {"op": "format_run", "anchor_id": need("anchor_id")}
        for k in (
            "bold",
            "italic",
            "underline",
            "strikethrough",
            "font",
            "size",
            "color",
            "highlight",
            "subscript",
            "superscript",
            "small_caps",
            "all_caps",
            "spacing",
        ):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "set_shading":
        op = {"op": "set_shading", "anchor_id": need("anchor_id")}
        for k in ("fill", "pattern"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "set_borders":
        op = {"op": "set_borders", "anchor_id": need("anchor_id")}
        # `line_style` is the MCP param name (avoids colliding with the `style`
        # param used by apply_style/create_table); the op field is `style`.
        if p.get("line_style") is not None:
            op["style"] = p["line_style"]
        for k in ("sides", "weight", "color"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "add_tab_stop":
        op = {"op": "add_tab_stop", "anchor_id": need("anchor_id"), "position": need("position")}
        for k in ("align", "leader"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "add_style":
        op = {"op": "add_style", "name": need("name")}
        for k in ("type", "based_on", "next_style"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "set_style":
        op = {"op": "set_style", "name": need("name")}
        for k in (*_STYLE_RUN_FIELDS, *_PARA_FIELDS, "based_on", "next_style"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "insert_break":
        return {
            "op": "insert_break",
            "anchor_id": need("anchor_id"),
            "kind": p.get("kind") or "page",
            "before": bool(p.get("before", False)),
        }
    if command == "insert_field":
        op = {
            "op": "insert_field",
            "anchor_id": need("anchor_id"),
            "kind": need("kind"),
            "before": bool(p.get("before", False)),
        }
        if p.get("text") is not None:
            op["text"] = p["text"]
        return op
    if command == "update_fields":
        return {"op": "update_fields"}
    if command in ("insert_footnote", "insert_endnote"):
        return {
            "op": command,
            "anchor_id": need("anchor_id"),
            "text": need("text"),
            "before": bool(p.get("before", False)),
        }
    if command == "insert_toc":
        op = {
            "op": "insert_toc",
            "anchor_id": p.get("anchor_id") or "start",
            "before": bool(p.get("before", False)),
        }
        if p.get("levels") is not None:
            op["levels"] = p["levels"]
        for k in ("use_heading_styles", "hyperlinks"):
            if p.get(k) is not None:
                op[k] = bool(p[k])
        return op
    if command == "add_bookmark":
        return {"op": "add_bookmark", "name": need("name"), "anchor_id": need("anchor_id")}
    if command == "add_hyperlink":
        if (p.get("url") is None) == (p.get("bookmark") is None):
            raise OpError("add_hyperlink requires exactly one of 'url' or 'bookmark'")
        op = {"op": "add_hyperlink", "anchor_id": need("anchor_id")}
        for k in ("url", "bookmark", "text", "screen_tip"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "insert_cross_reference":
        op = {
            "op": "insert_cross_reference",
            "anchor_id": need("anchor_id"),
            "target": need("target"),
            "before": bool(p.get("before", False)),
        }
        if p.get("kind") is not None:
            op["kind"] = p["kind"]
        if p.get("hyperlink") is not None:
            op["hyperlink"] = bool(p["hyperlink"])
        return op
    if command == "insert_caption":
        op = {"op": "insert_caption", "anchor_id": need("anchor_id")}
        for k in ("label", "text", "position"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "page_setup":
        op = {"op": "set_page_setup", "section": need("section")}
        for k in _PAGE_SETUP_FIELDS:
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "list":
        mapping = {
            "apply": "apply_list",
            "remove": "remove_list",
            "restart": "restart_numbering",
            "indent": "indent_list",
            "outdent": "outdent_list",
        }
        action = need("action")
        if action not in mapping:
            raise OpError(f"unknown list action: {action!r}")
        op = {"op": mapping[action], "anchor_id": need("anchor_id")}
        if action == "apply" and p.get("type") is not None:
            op["type"] = p["type"]
        return op
    if command == "comment":
        action = need("action")
        if action == "add":
            op = {"op": "add_comment", "anchor_id": need("anchor_id"), "text": need("text")}
            if p.get("author") is not None:
                op["author"] = p["author"]
            return op
        if action == "resolve":
            return {"op": "resolve_comment", "index": need("index")}
        if action == "delete":
            return {"op": "delete_comment", "index": need("index")}
        raise OpError(f"unknown comment action: {action!r}")
    if command == "table":
        action = need("action")
        if action == "set_cell":
            return {
                "op": "set_cell",
                "table": need("table"),
                "row": need("row"),
                "col": need("col"),
                "text": need("text"),
            }
        if action == "add_row":
            op = {"op": "add_row", "table": need("table")}
            if p.get("values") is not None:
                op["values"] = p["values"]
            return op
        if action == "delete_row":
            return {"op": "delete_row", "table": need("table"), "row": need("row")}
        if action == "append_record":
            return {"op": "append_record", "table": need("table"), "record": need("record")}
        if action == "update_row":
            op = {
                "op": "update_row",
                "table": need("table"),
                "key": need("key"),
                "values": need("values"),
            }
            if p.get("column") is not None:
                op["column"] = p["column"]
            return op
        if action == "set_heading_row":
            op = {"op": "set_heading_row", "table": need("table")}
            if p.get("row") is not None:
                op["row"] = p["row"]
            if p.get("heading") is not None:
                op["heading"] = bool(p["heading"])
            if p.get("allow_break") is not None:
                op["allow_break"] = bool(p["allow_break"])
            return op
        if action == "create":
            # rows/cols are optional when `data` is given (inferred from it);
            # required otherwise.
            if p.get("data") is None and (p.get("rows") is None or p.get("cols") is None):
                raise OpError("table create requires 'rows' and 'cols' unless 'data' is given")
            op = {
                "op": "create_table",
                "anchor_id": need("anchor_id"),
                "before": bool(p.get("before", False)),
            }
            if p.get("rows") is not None:
                op["rows"] = p["rows"]
            if p.get("cols") is not None:
                op["cols"] = p["cols"]
            if p.get("style") is not None:
                op["style"] = p["style"]
            if p.get("header") is not None:
                op["header"] = bool(p["header"])
            if p.get("data") is not None:
                op["data"] = p["data"]
            return op
        if action == "delete":
            return {"op": "delete_table", "table": need("table")}
        raise OpError(f"unknown table action: {action!r}")
    if command == "header":
        return {
            "op": "write_header",
            "section": need("section"),
            "text": need("text"),
            "which": p.get("which", "primary"),
        }
    if command == "footer":
        return {
            "op": "write_footer",
            "section": need("section"),
            "text": need("text"),
            "which": p.get("which", "primary"),
        }
    if command == "insert_image":
        if (p.get("image_base64") is None) == (p.get("path") is None):
            raise OpError("insert_image requires exactly one of 'image_base64' or 'path'")
        op = {
            "op": "insert_image",
            "anchor_id": need("anchor_id"),
            "wrap": need("wrap"),
            "before": bool(p.get("before", False)),
        }
        if p.get("image_base64") is not None:
            op["base64"] = p["image_base64"]
        else:
            op["path"] = p["path"]
        if p.get("block"):
            op["block"] = True
        for k in ("width", "height", "alt_text", "lock_aspect"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    raise OpError(f"unknown write command: {command!r}")


def _write_impl(
    worker: Worker, command: str, p: dict[str, Any], *, policy: PathPolicy | None = None
) -> dict[str, Any]:
    # Default to a deny-all policy: saving is off and non-local image paths are
    # rejected unless the caller (build_server) supplies a configured policy.
    policy = policy if policy is not None else PathPolicy()

    def job() -> dict[str, Any]:
        with attach() as word:
            doc = pick_doc(word, p.get("doc"))
            if command == "track":
                on = p.get("on")
                if on is None:
                    raise OpError("write command 'track' requires 'on' (bool)")
                doc.track_changes = bool(on)
                return {"ok": True, "command": "track", "track_changes": bool(on)}
            # Persistence — terminal side-effects, not undoable ops, so they bypass
            # run_batch (like `track`). Gated: the target must sit inside the
            # configured save-directory whitelist (WORDLIVE_SAVE_DIRS) or saving is off.
            if command == "save":
                policy.resolve_save_target(doc.path)
                return {"ok": True, "command": "save", "path": doc.save(), "saved": True}
            if command == "save_as":
                target = policy.resolve_save_target(_need(p, "path", command))
                written = doc.save_as(target, overwrite=bool(p.get("overwrite", False)))
                return {"ok": True, "command": "save_as", "path": written}
            if command == "export_pdf":
                target = policy.resolve_save_target(_need(p, "path", command))
                written = doc.export_pdf(
                    target, from_page=p.get("from_page"), to_page=p.get("to_page")
                )
                return {"ok": True, "command": "export_pdf", "path": written}
            op = _build_write_op(command, p)
            # Screen an insert_image op's path before it reaches COM/filesystem
            # (a UNC path's existence probe would authenticate to a remote SMB host).
            policy.screen_op_image_paths([op])
            result, exc = run_batch(doc, [op], label=f"MCP: {command}")
            if exc is not None:
                raise exc
            out = {"ok": True, "command": command, "ops_run": result["ops_run"]}
            if result.get("outputs"):
                # e.g. table create reports the new table's 1-based index.
                out["result"] = result["outputs"][0]
            if result.get("warnings"):
                # Surface ignored/unknown fields (e.g. a style passed to an
                # op that doesn't take one) instead of swallowing them.
                out["warnings"] = result["warnings"]
            return out

    return worker.run_on_word(job)


def _exec_impl(
    worker: Worker,
    ops: list[dict[str, Any]],
    *,
    doc: str | None,
    label: str | None,
    tracked: bool,
    policy: PathPolicy | None = None,
) -> tuple[dict[str, Any], WordliveError | None]:
    pol = policy if policy is not None else PathPolicy()

    def job() -> tuple[dict[str, Any], WordliveError | None]:
        with attach() as word:
            d = pick_doc(word, doc)
            # Vet image-source paths before any COM/filesystem access.
            pol.screen_op_image_paths(ops)
            return run_batch(d, ops, label=label or "MCP: exec", tracked=tracked)

    return worker.run_on_word(job)


def _parse_pages(pages: str | None) -> int | tuple[int, int] | None:
    """`"4"` -> 4, `"2-5"` -> (2, 5), None/"" -> None (whole document)."""
    if not pages:
        return None
    s = str(pages).strip()
    if "-" in s:
        a, _, b = s.partition("-")
        try:
            return int(a), int(b)
        except ValueError as e:
            raise OpError(f"invalid pages range: {pages!r}") from e
    try:
        return int(s)
    except ValueError as e:
        raise OpError(f"invalid pages value: {pages!r}") from e


def _snapshot_impl(
    worker: Worker,
    *,
    doc: str | None,
    pages: str | None,
    anchor: str | None,
    dpi: int,
    markup: str,
    max_dim: int | None = None,
) -> list[tuple[int, bytes]]:
    dpi = max(72, min(300, int(dpi)))
    md = max(1, int(max_dim)) if max_dim is not None else None
    pages_arg = _parse_pages(pages)

    def job() -> list[tuple[int, bytes]]:
        with attach() as word:
            d = pick_doc(word, doc)
            if anchor:
                snaps = d.snapshot_anchor(
                    d.anchor_by_id(anchor), dpi=dpi, max_dim=md, markup=markup
                )
            else:
                snaps = d.snapshot(pages=pages_arg, dpi=dpi, max_dim=md, markup=markup)
            return [(s.page, s.png) for s in snaps]

    return worker.run_on_word(job)


# ---------------------------------------------------------------------------
# Server assembly
# ---------------------------------------------------------------------------


def build_server(worker: Worker | None = None) -> FastMCP:
    """Build the FastMCP server. `worker` is injectable for tests (defaults to a
    real COM worker thread). Importing the `mcp` extra is deferred to here."""
    try:
        from mcp.server.fastmcp import FastMCP, Image
        from mcp.types import TextContent
    except ImportError as e:  # pragma: no cover - exercised only without the extra
        raise RuntimeError(
            "the wordlive MCP server requires the 'mcp' extra: "
            'pip install "wordlive[mcp]" (or "wordlive[mcp,snapshot]" for snapshots)'
        ) from e

    w: Worker = worker if worker is not None else ComWorker()
    # Saving is default-deny: the operator opts in by configuring WORDLIVE_SAVE_DIRS
    # at launch (image-source paths optionally restricted via WORDLIVE_IMAGE_DIRS).
    policy = PathPolicy.from_env()
    mcp = FastMCP("wordlive", instructions=_INSTRUCTIONS)

    @mcp.tool()
    def word_read(
        command: Literal[
            "status",
            "guide",
            "outline",
            "paragraphs",
            "find",
            "read_bookmark",
            "read_cc",
            "read_section",
            "table_list",
            "table_read",
            "table_records",
            "styles",
            "comments",
            "revisions",
            "track",
            "sections",
            "footnotes",
            "endnotes",
            "images",
            "read_image",
            "location",
            "stats",
        ],
        doc: str | None = None,
        name: str | None = None,
        text: str | None = None,
        in_anchor: str | None = None,
        anchor_id: str | None = None,
        heading: str | None = None,
        table: int | None = None,
        all_paragraphs: bool = False,
        start: int | None = None,
        count: int | None = None,
    ) -> Any:
        """Read from the open Word document. Dispatch on `command`:

        guide (no Word needed — returns the full agent guide: anchor model, the
        word_exec op vocabulary, and every field; read this first) ·
        status (no doc needed; reports name/path/saved/is_active per open doc) ·
        outline [all_paragraphs] · paragraphs [start,count] ·
        find {text,[in_anchor]} · read_bookmark {name} · read_cc {name} ·
        read_section {heading | anchor_id} · table_list · table_read {table} ·
        table_records {table} (body rows as dicts keyed by the header row — the
        read mirror of building a table from records) ·
        styles · comments · revisions (tracked changes: type/author/text/range per
        change) · track (is Track Changes on?) · sections · footnotes · endnotes ·
        images (embedded pictures: image:N id, mime, size, alt, para) ·
        read_image {anchor_id} (an embedded picture's bytes as base64 + mime — pass
        an image:N id or any single-image anchor) ·
        location {anchor_id} (where an anchor sits in the laid-out document:
        page/end_page span, line, column, in_table — "what page is this on"
        without a snapshot) ·
        stats (one-shot document summary: page/word/character/paragraph/line
        counts plus section/heading/table/image/comment/revision counts and saved).
        `doc` targets a document by name (default: active).
        """
        params = {
            "doc": doc,
            "name": name,
            "text": text,
            "in_anchor": in_anchor,
            "anchor_id": anchor_id,
            "heading": heading,
            "table": table,
            "all_paragraphs": all_paragraphs,
            "start": start,
            "count": count,
        }
        try:
            return _read_impl(w, command, params)
        except WordliveError as exc:
            raise _tool_error(exc) from exc

    @mcp.tool()
    def word_write(
        command: Literal[
            "insert",
            "insert_block",
            "insert_section",
            "insert_markdown",
            "replace_section",
            "delete_paragraph",
            "append",
            "prepend",
            "replace",
            "write_bookmark",
            "write_cc",
            "apply_style",
            "format_paragraph",
            "format_run",
            "set_shading",
            "set_borders",
            "add_tab_stop",
            "add_style",
            "set_style",
            "list",
            "comment",
            "table",
            "header",
            "footer",
            "track",
            "insert_image",
            "insert_break",
            "insert_field",
            "update_fields",
            "insert_footnote",
            "insert_endnote",
            "insert_toc",
            "add_bookmark",
            "add_hyperlink",
            "insert_cross_reference",
            "insert_caption",
            "page_setup",
            "save",
            "save_as",
            "export_pdf",
        ],
        doc: str | None = None,
        anchor_id: str | None = None,
        text: str | None = None,
        runs: list[Any] | None = None,
        items: list[Any] | None = None,
        name: str | None = None,
        style: str | None = None,
        before: bool = False,
        paragraph: bool = True,
        find: str | None = None,
        all: bool = False,
        occurrence: int | None = None,
        in_anchor: str | None = None,
        action: str | None = None,
        type: str | None = None,
        author: str | None = None,
        index: int | None = None,
        table: int | None = None,
        row: int | None = None,
        col: int | None = None,
        rows: int | None = None,
        cols: int | None = None,
        data: list[Any] | None = None,
        header: bool | None = None,
        heading: str | bool | None = None,
        body: list[Any] | str | None = None,
        markdown: str | None = None,
        level: int | None = None,
        allow_break: bool | None = None,
        values: list[Any] | dict[str, Any] | None = None,
        record: dict[str, Any] | None = None,
        key: str | None = None,
        column: str | None = None,
        section: int | None = None,
        which: str = "primary",
        on: bool | None = None,
        alignment: str | None = None,
        left_indent: float | None = None,
        right_indent: float | None = None,
        first_line_indent: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        page_break_before: bool | None = None,
        keep_together: bool | None = None,
        keep_with_next: bool | None = None,
        widow_control: bool | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikethrough: bool | None = None,
        font: str | None = None,
        size: str | float | None = None,
        color: str | None = None,
        highlight: str | None = None,
        subscript: bool | None = None,
        superscript: bool | None = None,
        small_caps: bool | None = None,
        all_caps: bool | None = None,
        spacing: str | float | None = None,
        fill: str | None = None,
        pattern: str | None = None,
        sides: str | None = None,
        line_style: str | None = None,
        weight: float | None = None,
        position: str | float | None = None,
        align: str | None = None,
        leader: str | None = None,
        based_on: str | None = None,
        next_style: str | None = None,
        kind: str | None = None,
        wrap: str | None = None,
        image_base64: str | None = None,
        path: str | None = None,
        block: bool | None = None,
        width: float | None = None,
        height: float | None = None,
        alt_text: str | None = None,
        lock_aspect: bool | None = None,
        margins: str | float | None = None,
        top_margin: str | float | None = None,
        bottom_margin: str | float | None = None,
        left_margin: str | float | None = None,
        right_margin: str | float | None = None,
        gutter: str | float | None = None,
        orientation: str | None = None,
        paper_size: str | None = None,
        columns: int | None = None,
        column_spacing: str | float | None = None,
        levels: list[int] | None = None,
        use_heading_styles: bool | None = None,
        hyperlinks: bool | None = None,
        url: str | None = None,
        bookmark: str | None = None,
        screen_tip: str | None = None,
        target: str | None = None,
        hyperlink: bool | None = None,
        label: str | None = None,
        overwrite: bool = False,
        from_page: int | None = None,
        to_page: int | None = None,
    ) -> dict[str, Any]:
        """Make one atomic-undo edit to the open Word document. Dispatch on `command`:

        insert {anchor_id, text|runs, [before,style]} — text is literal; runs is
            [{text,bold?,italic?,underline?,style?}] for inline-formatted spans ·
        insert_block {anchor_id, items, [before]} — a contiguous run of styled
            paragraphs in one op; each item is "plain text" or {text|runs, style?}
            (text carries **bold**/*italic* markdown); returns the block's
            range:START-END so you can apply_list/comment over the whole run ·
        insert_section {anchor_id, heading, body, [level=1, before]} — a Heading {level}
            paragraph plus its body (body = insert_block items shape) in one op ·
        insert_markdown {anchor_id, markdown, [before]} — a constrained-Markdown block
            as real Word structure: #/##/### headings, -/* bullets, 1. numbers,
            blank-line paragraphs, inline **bold**/*italic* (a subset, not CommonMark) ·
        replace_section {anchor_id=heading:N, body | markdown} — rewrite a heading's body
            (up to the next same-or-higher heading), keeping the heading; one of body/markdown ·
        delete_paragraph {anchor_id} — remove the paragraph(s) at an anchor, mark included ·
        append/prepend {text,[style]} — new final/first paragraph; pass paragraph=false
            to continue the adjacent paragraph inline (an inline append takes no style) ·
        replace {text, find|anchor_id, [all,occurrence,in_anchor]} ·
        write_bookmark/write_cc {name,text} · apply_style {anchor_id,name} ·
        format_paragraph {anchor_id,[alignment,*_indent,space_*,page_break_before,keep_together,keep_with_next,widow_control]} ·
        format_run {anchor_id,[bold,italic,underline,strikethrough,font,size,color,
            highlight,subscript,superscript,small_caps,all_caps,spacing]} — colour is a
            name/hex; highlight is a named palette colour; size/spacing accept unit strings ·
        set_shading {anchor_id,fill} — fill colour of a range/cell ·
        set_borders {anchor_id,[sides=all|box|top|bottom|left|right|horizontal|vertical,
            line_style=single|double|dot|dash|none,weight,color]} ·
        add_tab_stop {anchor_id,position,[align=left|center|right|decimal|bar,
            leader=dots|dashes|lines|none]} ·
        add_style {name,[type=paragraph|character|table|list,based_on,next_style]} —
            define a new style ·
        set_style {name,[bold,italic,underline,font,size,color,alignment,space_*,
            based_on,next_style]} — set an existing style's font/paragraph defaults ·
        list {anchor_id,action=apply|remove|restart|indent|outdent,[type]} ·
        comment {action=add|resolve|delete,...} ·
        table {action=set_cell|add_row|append_record|update_row|delete_row|set_heading_row|create|delete,
               create needs anchor_id and [rows,cols] (optional when data is given —
               inferred from it),[style,header,data,before]; data is a 2-D array OR
               records (a list of objects whose keys become a header row);
               append_record {table,record} — append a row from a {header: value} object;
               update_row {table,key,values,[column]} — set cells (values={header: value})
               on the first row whose key-column (column=, default first) equals key;
               set_heading_row {table,[row=1,heading=true,allow_break]} — repeating header row} ·
        insert_break {anchor_id,[kind=page|column|section_next|section_continuous,before]} ·
        insert_field {anchor_id,kind=page|numpages|date|time|filename|author|title|field,[text,before]} —
            a self-updating field; put page numbers in a footer; kind=field takes a raw code in text ·
        update_fields {} — recompute the document's fields (page numbers, refs, dates) ·
        insert_footnote {anchor_id,text,[before]} / insert_endnote {anchor_id,text,[before]} —
            a note anchored to a range; the new footnote:N/endnote:N is in result ·
        insert_toc {[anchor_id=start],levels=[upper,lower],use_heading_styles,hyperlinks,[before]} —
            a table of contents; run update_fields after to populate page numbers ·
        add_bookmark {name,anchor_id} — create a named bookmark over an anchor's range ·
        add_hyperlink {anchor_id, url | bookmark, [text,screen_tip]} — external URL or internal
            bookmark jump; text sets the visible link text ·
        insert_cross_reference {anchor_id,target,[kind=text|page|number|above_below,hyperlink,before]} —
            target is a bookmark:/heading:/footnote:/endnote: id ·
        insert_caption {anchor_id,[label=Figure,text,position=above|below]} — a numbered
            caption in its own paragraph (Table defaults above, else below) ·
        page_setup {[section=1],margins|top_margin|bottom_margin|left_margin|right_margin|gutter,
            orientation=portrait|landscape,paper_size=letter|legal|tabloid|a3|a4|a5,columns,column_spacing} —
            section page geometry; lengths accept unit strings ·
        header/footer {section,text,[which]} · track {on} ·
        insert_image {anchor_id,wrap, image_base64|path,
            [before,block,width,height,alt_text,lock_aspect]} — block puts the image on its
            own new line instead of in the anchor's text run ·
        save {} — save to the document's existing file (must already be saved) ·
        save_as {path,[overwrite]} — save a .docx to path ·
        export_pdf {path,[from_page,to_page]} — export a PDF (the deliverable path).
            save/save_as/export_pdf are GATED: they only write inside the server's
            configured save directories (WORDLIVE_SAVE_DIRS); with none set, saving is off.

        For several edits in one undo step, use word_exec instead. Call
        word_read(command="guide") for the full anchor model and field reference.
        """
        params = {
            "doc": doc,
            "anchor_id": anchor_id,
            "text": text,
            "runs": runs,
            "items": items,
            "name": name,
            "style": style,
            "before": before,
            "paragraph": paragraph,
            "find": find,
            "all": all,
            "occurrence": occurrence,
            "in_anchor": in_anchor,
            "action": action,
            "type": type,
            "author": author,
            "index": index,
            "table": table,
            "row": row,
            "col": col,
            "rows": rows,
            "cols": cols,
            "data": data,
            "header": header,
            "heading": heading,
            "body": body,
            "markdown": markdown,
            "level": level,
            "allow_break": allow_break,
            "values": values,
            "record": record,
            "key": key,
            "column": column,
            "section": section,
            "which": which,
            "on": on,
            "alignment": alignment,
            "left_indent": left_indent,
            "right_indent": right_indent,
            "first_line_indent": first_line_indent,
            "space_before": space_before,
            "space_after": space_after,
            "page_break_before": page_break_before,
            "keep_together": keep_together,
            "keep_with_next": keep_with_next,
            "widow_control": widow_control,
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "strikethrough": strikethrough,
            "font": font,
            "size": size,
            "color": color,
            "highlight": highlight,
            "subscript": subscript,
            "superscript": superscript,
            "small_caps": small_caps,
            "all_caps": all_caps,
            "spacing": spacing,
            "fill": fill,
            "pattern": pattern,
            "sides": sides,
            "line_style": line_style,
            "weight": weight,
            "position": position,
            "align": align,
            "leader": leader,
            "based_on": based_on,
            "next_style": next_style,
            "kind": kind,
            "wrap": wrap,
            "image_base64": image_base64,
            "path": path,
            "block": block,
            "width": width,
            "height": height,
            "alt_text": alt_text,
            "lock_aspect": lock_aspect,
            "margins": margins,
            "top_margin": top_margin,
            "bottom_margin": bottom_margin,
            "left_margin": left_margin,
            "right_margin": right_margin,
            "gutter": gutter,
            "orientation": orientation,
            "paper_size": paper_size,
            "columns": columns,
            "column_spacing": column_spacing,
            "levels": levels,
            "use_heading_styles": use_heading_styles,
            "hyperlinks": hyperlinks,
            "url": url,
            "bookmark": bookmark,
            "screen_tip": screen_tip,
            "target": target,
            "hyperlink": hyperlink,
            "label": label,
            "overwrite": overwrite,
            "from_page": from_page,
            "to_page": to_page,
        }
        try:
            return _write_impl(w, command, params, policy=policy)
        except WordliveError as exc:
            raise _tool_error(exc) from exc

    @mcp.tool()
    def word_exec(
        ops: list[dict[str, Any]],
        doc: str | None = None,
        label: str | None = None,
        tracked: bool = False,
    ) -> dict[str, Any]:
        """Apply a batch of ops to the open document as a SINGLE atomic undo.

        Each op is `{"op": "<kind>", ...}` — e.g.
        {"op":"write_bookmark","name":"Addr","text":"…"},
        {"op":"insert_paragraph","anchor_id":"heading:2","text":"…","style":"Body Text"},
        {"op":"find_replace","find":"Q3","text":"Q4","all":true}. Set `tracked`
        true to record the batch as tracked changes. Stops at the first failing
        op and reports `failure` (its `index`, `error`, `type`). Fields an op
        doesn't use are reported in `warnings`, not silently dropped.

        Anchor ids (the `anchor_id` of placement ops):
          heading:N · para:N (any paragraph) · bookmark:NAME · cc:NAME ·
          table:N:R:C (a cell) · range:START-END (what `find` emits — for
          replace/comments, NOT a placement target) · header:S:WHICH ·
          footer:S:WHICH · start · end. heading:N / para:N are positional and
          renumber on structural inserts — re-read outline/paragraphs after one.
          bookmark:/cc: are name-based and survive edits.

        Ops (required fields → behaviour):
          write_bookmark {name,text} · write_cc {name,text} ·
          insert_paragraph {anchor_id, text|runs, [style,before]} — new paragraph by an anchor;
            text is literal, runs is [{text,bold?,italic?,underline?,style?}] for inline spans ·
          insert_block {anchor_id, items, [before]} — a contiguous run of styled paragraphs in one
            op; items are "plain text" or {text|runs, style?} (text takes **bold**/*italic*);
            returns the block's range:START-END in outputs ·
          insert_section {anchor_id, heading, body, [level=1,before]} — a Heading {level} + its body
            (body = insert_block items) in one op; returns the section's range:START-END ·
          insert_markdown {anchor_id, markdown, [before]} — a constrained-Markdown block as Word
            structure (#/##/### headings, -/* bullets, 1. numbers, paragraphs, inline **bold**/*italic*) ·
          replace_section {anchor_id=heading:N, body|markdown} — rewrite a heading's body, keep the heading ·
          delete_paragraph {anchor_id} — remove the paragraph(s) at an anchor, mark included ·
          append {text,[style]} / prepend {text,[style]} — new final/first paragraph ·
          append_inline {text} / prepend_inline {text} — continue the last/first paragraph (NO style) ·
          append_paragraph / prepend_paragraph — explicit synonyms of append/prepend ·
          replace {anchor_id,text} · find_replace {find,text,[all,occurrence,in]} ·
          apply_style {anchor_id,name} · format_paragraph {anchor_id,[alignment,*_indent,space_*,page_break_before,keep_together,keep_with_next,widow_control]} ·
          insert_image {anchor_id,wrap, path|base64, [before,block,width,height,alt_text,lock_aspect]} ·
          insert_break {anchor_id,[kind=page|column|section_next|section_continuous,before]} ·
          insert_field {anchor_id,kind,[text,before]} · update_fields {} · set_page_setup {section,[margins,*_margin,gutter,orientation,paper_size,columns,column_spacing]} ·
          insert_footnote/insert_endnote {anchor_id,text,[before]} — returns the new footnote:N/endnote:N in outputs ·
          insert_toc {anchor_id,[levels=[upper,lower],use_heading_styles,hyperlinks,before]} — update_fields after to fill page numbers ·
          add_bookmark {name,anchor_id} · add_hyperlink {anchor_id, url|bookmark, [text,screen_tip]} ·
          insert_cross_reference {anchor_id,target,[kind,hyperlink,before]} — target is a bookmark:/heading:/footnote:/endnote: id ·
          insert_caption {anchor_id,[label,text,position=above|below]} — own-paragraph caption ·
          create_table {anchor_id, [rows,cols] (optional when data given — inferred),[style,data,header,before]} —
            data is a 2-D array OR records (objects whose keys become a header row); cells default to Normal; returns the new index in outputs ·
          set_cell {table,row,col,text} · add_row {table,[values]} · delete_row {table,row} ·
          append_record {table,record} — append a row from a {header: value} object ·
          update_row {table,key,values,[column]} — set cells (values={header: value}) on the first row whose key-column equals key ·
          set_heading_row {table,[row=1,heading,allow_break]} — repeating header row on a multi-page table · delete_table {table} ·
          add_comment {anchor_id,text,[author]} · resolve_comment {index} · delete_comment {index} ·
          apply_list {anchor_id,[type=bulleted|numbered|outline,continue_previous]} · remove_list/restart_numbering/indent_list/outdent_list {anchor_id} ·
          write_header/write_footer {section,text,[which=primary|first|even]}.

        Call word_read(command="guide") for the full field reference.
        """
        try:
            result, exc = _exec_impl(w, ops, doc=doc, label=label, tracked=tracked, policy=policy)
        except WordliveError as setup_exc:
            raise _tool_error(setup_exc) from setup_exc
        if exc is not None:
            raise _tool_error(exc, result=result)
        return result

    # structured_output=False: this tool returns MCP content blocks (image + text
    # labels), not structured data. Without it FastMCP infers a wrapped output
    # schema from the `-> list[Any]` annotation and *re-serialises every block —
    # including the base64 PNG bytes — into structuredContent*, sending each image
    # on the wire twice (a large, silent token cost on hosts that forward
    # structuredContent). Suppressing the schema sends the image exactly once.
    @mcp.tool(structured_output=False)
    def word_snapshot(
        doc: str | None = None,
        pages: str | None = None,
        anchor: str | None = None,
        dpi: int = 150,
        max_dim: int | None = None,
        markup: str = "none",
    ) -> list[Any]:
        """Render page(s) of the open document to PNG so you can SEE the layout.

        Pick at most one target: `anchor` (the page(s) an anchor occupies — a
        heading expands to its whole section), or `pages` ("4" or "2-5"). With
        neither, the whole document renders. `max_dim` caps each page's long edge
        to that many pixels (only ever lowering resolution) — pair it with no page
        target to check the WHOLE document's layout cheaply: a vision model is
        billed on pixel area, so the cap is a predictable per-page token budget
        regardless of paper size (~1000 stays legible for "did my styling land").
        `markup` is "none" (the final document) or "all" (show tracked changes and
        comments as visible revision marks / balloons — pair with
        word_read(command="revisions") for the structured list). Returns image
        content (and a "page N" label per page) inline, so a vision model sees the
        render directly — no filesystem path that a remote/sandboxed host couldn't
        open. Needs the snapshot extra (PyMuPDF).
        """
        try:
            rendered = _snapshot_impl(
                w, doc=doc, pages=pages, anchor=anchor, dpi=dpi, max_dim=max_dim, markup=markup
            )
        except WordliveError as exc:
            raise _tool_error(exc) from exc
        content: list[Any] = []
        for page, png in rendered:
            content.append(TextContent(type="text", text=f"page {page}"))
            # Convert to an ImageContent block explicitly: FastMCP won't serialise
            # a bare Image when it's one element of a mixed content list.
            content.append(Image(data=png, format="png").to_image_content())
        return content

    @mcp.resource("wordlive://guide", mime_type="text/markdown")
    def guide() -> str:
        """The full wordlive agent guide: anchor model, every verb, the op vocabulary."""
        return skill_body()

    return mcp


def main() -> None:
    """Launch the server over stdio (the transport Claude Desktop spawns)."""
    build_server().run()
