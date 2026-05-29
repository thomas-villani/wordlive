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

import json
from typing import TYPE_CHECKING, Any, Literal

from .. import attach
from .._anchors import Heading
from .._guide import skill_body
from .._ops import pick_doc, run_batch
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
`word_snapshot`. Read the `wordlive://guide` resource for the full op vocabulary
and field reference.
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
            if command == "styles":
                return doc.styles.list()
            if command == "comments":
                return doc.comments.list()
            if command == "sections":
                return doc.sections.list()
            raise OpError(f"unknown read command: {command!r}")

    return worker.run_on_word(job)


def _build_write_op(command: str, p: dict[str, Any]) -> dict[str, Any]:
    """Translate a `word_write` command + params into a single `apply_op` op dict."""

    def need(key: str) -> Any:
        return _need(p, key, command)

    if command == "insert":
        op = {
            "op": "insert_paragraph",
            "anchor_id": need("anchor_id"),
            "text": need("text"),
            "before": bool(p.get("before", False)),
        }
        if p.get("style") is not None:
            op["style"] = p["style"]
        return op
    if command in ("append", "prepend"):
        text = need("text")
        if p.get("paragraph", True):
            op = {"op": f"{command}_paragraph", "text": text}
            if p.get("style") is not None:
                op["style"] = p["style"]
            return op
        return {"op": command, "text": text}
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
        ):
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
        if action == "create":
            op = {
                "op": "create_table",
                "anchor_id": need("anchor_id"),
                "rows": need("rows"),
                "cols": need("cols"),
                "before": bool(p.get("before", False)),
            }
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
        for k in ("width", "height", "alt_text", "lock_aspect"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    raise OpError(f"unknown write command: {command!r}")


def _write_impl(worker: Worker, command: str, p: dict[str, Any]) -> dict[str, Any]:
    def job() -> dict[str, Any]:
        with attach() as word:
            doc = pick_doc(word, p.get("doc"))
            if command == "track":
                on = p.get("on")
                if on is None:
                    raise OpError("write command 'track' requires 'on' (bool)")
                doc.track_changes = bool(on)
                return {"ok": True, "command": "track", "track_changes": bool(on)}
            op = _build_write_op(command, p)
            result, exc = run_batch(doc, [op], label=f"MCP: {command}")
            if exc is not None:
                raise exc
            out = {"ok": True, "command": command, "ops_run": result["ops_run"]}
            if result.get("outputs"):
                # e.g. table create reports the new table's 1-based index.
                out["result"] = result["outputs"][0]
            return out

    return worker.run_on_word(job)


def _exec_impl(
    worker: Worker,
    ops: list[dict[str, Any]],
    *,
    doc: str | None,
    label: str | None,
    tracked: bool,
) -> tuple[dict[str, Any], WordliveError | None]:
    def job() -> tuple[dict[str, Any], WordliveError | None]:
        with attach() as word:
            d = pick_doc(word, doc)
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
) -> list[tuple[int, bytes]]:
    dpi = max(72, min(300, int(dpi)))
    pages_arg = _parse_pages(pages)

    def job() -> list[tuple[int, bytes]]:
        with attach() as word:
            d = pick_doc(word, doc)
            if anchor:
                snaps = d.snapshot_anchor(d.anchor_by_id(anchor), dpi=dpi)
            else:
                snaps = d.snapshot(pages=pages_arg, dpi=dpi)
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
    mcp = FastMCP("wordlive", instructions=_INSTRUCTIONS)

    @mcp.tool()
    def word_read(
        command: Literal[
            "status",
            "outline",
            "paragraphs",
            "find",
            "read_bookmark",
            "read_cc",
            "read_section",
            "table_list",
            "table_read",
            "styles",
            "comments",
            "sections",
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

        status (no doc needed) · outline [all_paragraphs] · paragraphs [start,count] ·
        find {text,[in_anchor]} · read_bookmark {name} · read_cc {name} ·
        read_section {heading | anchor_id} · table_list · table_read {table} ·
        styles · comments · sections. `doc` targets a document by name (default:
        active).
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
            "append",
            "prepend",
            "replace",
            "write_bookmark",
            "write_cc",
            "apply_style",
            "format_paragraph",
            "list",
            "comment",
            "table",
            "header",
            "footer",
            "track",
            "insert_image",
            "insert_break",
        ],
        doc: str | None = None,
        anchor_id: str | None = None,
        text: str | None = None,
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
        values: list[Any] | None = None,
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
        kind: str | None = None,
        wrap: str | None = None,
        image_base64: str | None = None,
        path: str | None = None,
        width: float | None = None,
        height: float | None = None,
        alt_text: str | None = None,
        lock_aspect: bool | None = None,
    ) -> dict[str, Any]:
        """Make one atomic-undo edit to the open Word document. Dispatch on `command`:

        insert {anchor_id,text,[before,style]} · append/prepend {text,[paragraph,style]} ·
        replace {text, find|anchor_id, [all,occurrence,in_anchor]} ·
        write_bookmark/write_cc {name,text} · apply_style {anchor_id,name} ·
        format_paragraph {anchor_id,[alignment,*_indent,space_*,page_break_before]} ·
        list {anchor_id,action=apply|remove|restart|indent|outdent,[type]} ·
        comment {action=add|resolve|delete,...} ·
        table {action=set_cell|add_row|delete_row|create|delete,
               create needs anchor_id,rows,cols,[style,header,data,before]} ·
        insert_break {anchor_id,[kind=page|column|section_next|section_continuous,before]} ·
        header/footer {section,text,[which]} · track {on} ·
        insert_image {anchor_id,wrap, image_base64|path, [before,width,height,alt_text,lock_aspect]}.

        For several edits in one undo step, use word_exec instead. See the
        wordlive://guide resource for every field.
        """
        params = {
            "doc": doc,
            "anchor_id": anchor_id,
            "text": text,
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
            "values": values,
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
            "kind": kind,
            "wrap": wrap,
            "image_base64": image_base64,
            "path": path,
            "width": width,
            "height": height,
            "alt_text": alt_text,
            "lock_aspect": lock_aspect,
        }
        try:
            return _write_impl(w, command, params)
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
        op and reports it. The full op vocabulary is in the wordlive://guide
        resource.
        """
        try:
            result, exc = _exec_impl(w, ops, doc=doc, label=label, tracked=tracked)
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
    ) -> list[Any]:
        """Render page(s) of the open document to PNG so you can SEE the layout.

        Pick at most one target: `anchor` (the page(s) an anchor occupies — a
        heading expands to its whole section), or `pages` ("4" or "2-5"). With
        neither, the whole document renders. Returns image content (and a "page N"
        label per page) inline, so a vision model sees the render directly — no
        filesystem path that a remote/sandboxed host couldn't open. Needs the
        snapshot extra (PyMuPDF).
        """
        try:
            rendered = _snapshot_impl(w, doc=doc, pages=pages, anchor=anchor, dpi=dpi)
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
