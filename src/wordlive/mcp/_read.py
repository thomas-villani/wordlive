"""`word_read` dispatch implementation."""

from __future__ import annotations

import base64
import json
from typing import Any

from .. import attach
from .._anchors import Heading
from .._guide import skill_body
from .._ops import (
    pick_doc,
)
from .._suggest import unknown_value_message
from ..exceptions import OpError
from ._commands import READ_COMMANDS
from ._common import _need
from ._worker import Worker

_GUIDE_POINTER = 'call word_read(command="guide")'


def _read_impl(worker: Worker, command: str, p: dict[str, Any]) -> Any:
    # Validate before the worker attaches to Word: a mistyped command should cost
    # a suggestion, not a COM round-trip. (`command` is a plain `str` on the tool
    # signature precisely so a near-miss lands here rather than in pydantic.)
    if command not in READ_COMMANDS:
        raise OpError(
            unknown_value_message("read command", command, READ_COMMANDS, fallback=_GUIDE_POINTER)
        )
    if command == "guide":
        # The MCP-native agent guide (not the CLI one) — the same text served by
        # the wordlive://guide resource, but reachable as a tool call (resources
        # aren't surfaced by every MCP client). Needs neither Word nor the worker
        # thread.
        return {"guide": skill_body("mcp")}

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
                # `params` always carries every key, so an omitted `mode` arrives
                # as None and `.get`'s default never fires — coalesce instead.
                return doc.find(text, scope=scope, mode=p.get("mode") or "fuzzy")
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
            if command == "to_markdown":
                return {"markdown": doc.to_markdown(within=p.get("within"))}
            if command == "to_html":
                return {"html": doc.to_html(within=p.get("within"))}
            if command == "digest":
                digest_kwargs: dict[str, Any] = {}
                if p.get("budget") is not None:
                    digest_kwargs["budget"] = p["budget"]
                if p.get("depth") is not None:
                    digest_kwargs["depth"] = p["depth"]
                return {"digest": doc.read(**digest_kwargs)}
            if command == "between":
                start_id = _need(p, "start_anchor", command)
                end_id = _need(p, "end_anchor", command)
                inclusive = bool(p.get("inclusive", False))
                span = doc.between(start_id, end_id, inclusive=inclusive)
                return {
                    "start": start_id,
                    "end": end_id,
                    "inclusive": inclusive,
                    "anchor_id": span.anchor_id,
                    "text": span.text,
                }
            if command == "nearest_heading":
                anchor_id = _need(p, "anchor_id", command)
                direction = p.get("direction") or "before"
                row = doc.nearest_heading(anchor_id, direction=direction)
                return {"anchor_id": anchor_id, "direction": direction, "heading": row}
            if command == "find_paragraphs":
                kwargs: dict[str, Any] = {}
                if p.get("limit") is not None:
                    kwargs["limit"] = p["limit"]
                if p.get("min_score") is not None:
                    kwargs["min_score"] = p["min_score"]
                return doc.find_paragraphs(_need(p, "text", command), **kwargs)
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
            if command == "read_text":
                anchor_id = _need(p, "anchor_id", command)
                view = p.get("view") or "raw"
                anchor = doc.anchor_by_id(anchor_id)
                if view == "segments":
                    return {"anchor_id": anchor_id, "segments": anchor.revision_segments()}
                if view not in ("raw", "final", "original"):
                    raise OpError("read 'read_text' view must be raw / final / original / segments")
                text = {
                    "raw": lambda: anchor.text,
                    "final": lambda: anchor.text_final,
                    "original": lambda: anchor.text_original,
                }[view]()
                return {"anchor_id": anchor_id, "view": view, "text": text}
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
            if command == "equations":
                return doc.equations.list()
            if command == "charts":
                return doc.charts.list()
            if command == "shapes":
                return doc.shapes.list()
            if command == "hyperlinks":
                return doc.hyperlinks.list()
            if command == "fields":
                return doc.fields.list()
            if command == "properties":
                return doc.properties.read()
            if command == "watermark":
                info = doc.watermark()
                return info.to_dict() if info is not None else None
            if command == "variables":
                return doc.variables.list()
            if command == "theme":
                return doc.theme.to_dict()
            if command == "themes":
                return doc.theme.list_available()
            if command == "proofing":
                return doc.proofing()
            if command == "lint":
                return doc.lint(
                    rules=p.get("rules"), within=p.get("within"), profile=p.get("profile")
                )
            if command == "checkpoint":
                cp = doc.checkpoint(
                    include=p.get("include") or "text+style",
                    within=p.get("within"),
                )
                return json.loads(cp.to_json())
            if command == "diff":
                if p.get("checkpoint") is not None:
                    return doc.changes_since(p["checkpoint"])
                a, b = p.get("cp_a"), p.get("cp_b")
                if a is None or b is None:
                    raise OpError(
                        "read 'diff' needs `checkpoint` (vs the document now), "
                        "or both `cp_a` and `cp_b`"
                    )
                return doc.diff(a, b)
            if command == "format_info":
                anchor_id = _need(p, "anchor_id", command)
                return doc.anchor_by_id(anchor_id).format_info()
            if command == "list_levels":
                anchor_id = _need(p, "anchor_id", command)
                return {"levels": doc.anchor_by_id(anchor_id).read_list_levels()}
            if command == "read_image":
                anchor_id = _need(p, "anchor_id", command)
                data, mime = doc.anchor_by_id(anchor_id).read_image()
                return {
                    "anchor_id": anchor_id,
                    "mime": mime,
                    "bytes": len(data),
                    "base64": base64.b64encode(data).decode("ascii"),
                }
            # Unreachable via the guard above: this fires only if a name is added
            # to READ_COMMANDS without a branch here. `test_dispatch_errors` pins it.
            raise OpError(f"read command {command!r} is declared but not implemented")

    return worker.run_on_word(job)
