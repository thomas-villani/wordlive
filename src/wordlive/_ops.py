"""Click-free batch-op core, shared by the CLI (`exec`) and the MCP server.

`apply_op` applies one op dict to a document; `run_batch` wraps a list of them
in a single atomic-undo (and optional Track Changes) scope. Both the CLI and the
MCP `word_exec` tool funnel through here, so the op vocabulary, validation, and
failure reporting stay identical across surfaces.

This module deliberately imports no Click — malformed input raises `OpError`
(generic exit code 1 / `code="error"`), which the CLI's `_run` and the MCP error
wrapper both already handle.
"""

from __future__ import annotations

from contextlib import nullcontext
from pathlib import Path
from typing import TYPE_CHECKING, Any

from .exceptions import AmbiguousMatchError, OpError, WordliveError

if TYPE_CHECKING:
    from ._app import Word
    from ._document import Document


def pick_doc(word: Word, doc_name: str | None) -> Document:
    """Resolve the target document: the named one, or the active document."""
    if doc_name is None:
        return word.documents.active
    return word.documents[doc_name]


# Required fields per op kind. Validated up-front so a malformed payload raises a
# clean OpError ("op 'write_bookmark' is missing required field(s): 'name'")
# instead of a Python KeyError traceback that would land a tool-use loop on a
# generic error with no actionable signal.
OP_REQUIRED_FIELDS: dict[str, tuple[str, ...]] = {
    "write_bookmark": ("name", "text"),
    "write_cc": ("name", "text"),
    "insert_paragraph": ("anchor_id", "text"),
    "append_paragraph": ("text",),
    "append": ("text",),
    "append_inline": ("text",),
    "prepend_paragraph": ("text",),
    "prepend": ("text",),
    "prepend_inline": ("text",),
    "insert_image": ("anchor_id", "wrap"),
    "replace": ("anchor_id", "text"),
    "find_replace": ("find", "text"),
    "apply_style": ("anchor_id", "name"),
    "format_paragraph": ("anchor_id",),
    "format_run": ("anchor_id",),
    "set_shading": ("anchor_id",),
    "set_borders": ("anchor_id",),
    "add_tab_stop": ("anchor_id", "position"),
    "add_style": ("name",),
    "set_style": ("name",),
    "insert_field": ("anchor_id", "kind"),
    "set_page_setup": ("section",),
    "update_fields": (),
    "insert_footnote": ("anchor_id", "text"),
    "insert_endnote": ("anchor_id", "text"),
    "insert_toc": ("anchor_id",),
    "add_bookmark": ("name", "anchor_id"),
    "add_hyperlink": ("anchor_id",),
    "insert_cross_reference": ("anchor_id", "target"),
    "insert_caption": ("anchor_id",),
    "set_cell": ("table", "row", "col", "text"),
    "add_row": ("table",),
    "delete_row": ("table", "row"),
    "create_table": ("anchor_id", "rows", "cols"),
    "delete_table": ("table",),
    "insert_break": ("anchor_id",),
    "add_comment": ("anchor_id", "text"),
    "resolve_comment": ("index",),
    "delete_comment": ("index",),
    "apply_list": ("anchor_id",),
    "remove_list": ("anchor_id",),
    "restart_numbering": ("anchor_id",),
    "indent_list": ("anchor_id",),
    "outdent_list": ("anchor_id",),
    "write_header": ("section", "text"),
    "write_footer": ("section", "text"),
}

# `before` / `after` / `where` all steer an insert op's side (see `op_before`),
# so any op that calls it accepts the trio.
_WHERE_FIELDS = ("before", "after", "where")

# Character-formatting kwargs shared by `format_run` (and `set_style`).
_RUN_FIELDS = (
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
)

# Paragraph-formatting kwargs shared by `format_paragraph` (and `set_style`).
_PARA_FIELDS = (
    "alignment",
    "left_indent",
    "right_indent",
    "first_line_indent",
    "space_before",
    "space_after",
    "page_break_before",
)

# `set_style` accepts the run + paragraph vocab (minus highlight, which a style's
# font can't carry) plus the style-chaining fields.
_STYLE_RUN_FIELDS = tuple(f for f in _RUN_FIELDS if f != "highlight")
_SET_STYLE_FIELDS = (*_STYLE_RUN_FIELDS, *_PARA_FIELDS, "based_on", "next_style")

# Page-geometry kwargs shared by `set_page_setup` (the `Section.set_page_setup`
# vocab). Reused by apply_op, the optional-field map, and the MCP param list.
_PAGE_SETUP_FIELDS = (
    "margins",
    "top_margin",
    "bottom_margin",
    "left_margin",
    "right_margin",
    "gutter",
    "orientation",
    "paper_size",
    "columns",
    "column_spacing",
)

# Optional fields each op *reads*. Combined with the required set (and the
# implicit `op` key), this is the full vocabulary an op understands — anything
# else in the payload is silently ignored by `apply_op`, which is exactly the
# silent-success footgun the warnings below surface. Keep an entry for every op
# in OP_REQUIRED_FIELDS so `unexpected_fields` can flag stray keys on all of them.
OP_OPTIONAL_FIELDS: dict[str, tuple[str, ...]] = {
    "write_bookmark": (),
    "write_cc": (),
    "insert_paragraph": ("style", *_WHERE_FIELDS),
    "append_paragraph": ("style",),
    "append": ("style",),
    "append_inline": (),
    "prepend_paragraph": ("style",),
    "prepend": ("style",),
    "prepend_inline": (),
    "insert_image": (
        "path",
        "base64",
        "block",
        "width",
        "height",
        "alt_text",
        "lock_aspect",
        *_WHERE_FIELDS,
    ),
    "replace": (),
    "find_replace": ("in", "all", "occurrence"),
    "apply_style": (),
    "format_paragraph": (
        "alignment",
        "left_indent",
        "right_indent",
        "first_line_indent",
        "space_before",
        "space_after",
        "page_break_before",
    ),
    "format_run": _RUN_FIELDS,
    "set_shading": ("fill", "pattern"),
    "set_borders": ("sides", "style", "weight", "color"),
    "add_tab_stop": ("align", "leader"),
    "add_style": ("type", "based_on", "next_style"),
    "set_style": _SET_STYLE_FIELDS,
    "insert_field": ("text", *_WHERE_FIELDS),
    "set_page_setup": _PAGE_SETUP_FIELDS,
    "update_fields": (),
    "insert_footnote": _WHERE_FIELDS,
    "insert_endnote": _WHERE_FIELDS,
    "insert_toc": ("levels", "use_heading_styles", "hyperlinks", *_WHERE_FIELDS),
    "add_bookmark": (),
    "add_hyperlink": ("url", "bookmark", "text", "screen_tip"),
    "insert_cross_reference": ("kind", "hyperlink", *_WHERE_FIELDS),
    "insert_caption": ("label", "text", *_WHERE_FIELDS),
    "set_cell": (),
    "add_row": ("values",),
    "delete_row": (),
    "create_table": ("style", "data", "header", *_WHERE_FIELDS),
    "delete_table": (),
    "insert_break": ("kind", *_WHERE_FIELDS),
    "add_comment": ("author",),
    "resolve_comment": (),
    "delete_comment": (),
    "apply_list": ("type", "continue_previous", "continue"),
    "remove_list": (),
    "restart_numbering": (),
    "indent_list": (),
    "outdent_list": (),
    "write_header": ("which",),
    "write_footer": ("which",),
}


def unexpected_fields(op: dict[str, Any], kind: str) -> list[str]:
    """Fields present on `op` that its `kind` neither requires nor reads.

    `apply_op` ignores unknown keys, so a typo (``anchorid``) or a field that
    doesn't apply to this op (``style`` on a bare ``append_inline``) would
    otherwise vanish with no signal — silent success with wrong output. The
    caller turns each returned name into a soft warning in the batch result.
    """
    known = {"op", *OP_REQUIRED_FIELDS.get(kind, ()), *OP_OPTIONAL_FIELDS.get(kind, ())}
    return [k for k in op if k not in known]


def op_before(op: dict[str, Any]) -> bool:
    """Whether an insert op targets *before* its anchor (default: after).

    Accepts either the verbose `"where": "before"|"after"` or the boolean
    `"before": true` / `"after": true` — the latter mirrors the CLI's
    `--before/--after` flags, so the natural JSON encoding works regardless of
    which form an LLM reaches for. An explicit `"before"` wins if both appear.
    """
    if "before" in op:
        return bool(op["before"])
    if "after" in op:
        return not bool(op["after"])
    return op.get("where") == "before"


def validate_op(op: dict[str, Any]) -> str:
    """Return the op kind after asserting it's known and required keys exist."""
    if not isinstance(op, dict):
        raise OpError(f"each op must be an object; got {type(op).__name__}")
    kind = op.get("op")
    if kind is None:
        raise OpError("op is missing the 'op' field")
    if kind not in OP_REQUIRED_FIELDS:
        raise OpError(f"unknown op: {kind!r}")
    missing = [k for k in OP_REQUIRED_FIELDS[kind] if k not in op]
    if missing:
        raise OpError(
            f"op {kind!r} is missing required field(s): {', '.join(repr(m) for m in missing)}"
        )
    return kind


def apply_op(doc: Document, op: dict[str, Any]) -> dict[str, Any] | None:
    """Apply a single op from an exec batch. Raises WordliveError on bad input.

    Most ops return `None`. Ops that *create* addressable structure return a
    small result dict so the batch can report it — `create_table` returns
    `{"table": N, "rows": R, "columns": C}` for the new table's 1-based index.
    """
    kind = validate_op(op)
    if kind == "write_bookmark":
        doc.bookmarks[op["name"]].set_text(op["text"])
    elif kind == "write_cc":
        doc.content_controls[op["name"]].set_text(op["text"])
    elif kind == "insert_paragraph":
        anchor = doc.anchor_by_id(op["anchor_id"])
        if op_before(op):
            anchor.insert_paragraph_before(op["text"], style=op.get("style"))
        else:
            anchor.insert_paragraph_after(op["text"], style=op.get("style"))
    elif kind in ("append", "append_paragraph"):
        # `append` is the natural name for the common case — a new final
        # paragraph (matching its description and `append_paragraph`). The
        # inline "continue the last paragraph" variant is `append_inline`.
        doc.append_paragraph(op["text"], style=op.get("style"))
    elif kind == "append_inline":
        doc.append(op["text"])
    elif kind in ("prepend", "prepend_paragraph"):
        doc.prepend_paragraph(op["text"], style=op.get("style"))
    elif kind == "prepend_inline":
        doc.prepend(op["text"])
    elif kind == "insert_image":
        if ("path" in op) == ("base64" in op):
            raise OpError("op 'insert_image' requires exactly one of 'path' or 'base64'")
        image: str | Path = Path(op["path"]) if "path" in op else op["base64"]
        kwargs = {
            k: op[k] for k in ("block", "width", "height", "alt_text", "lock_aspect") if k in op
        }
        doc.anchor_by_id(op["anchor_id"]).insert_image(
            image, wrap=op["wrap"], where=("before" if op_before(op) else "after"), **kwargs
        )
    elif kind == "replace":
        doc.anchor_by_id(op["anchor_id"]).set_text(op["text"])
    elif kind == "find_replace":
        scope = doc.anchor_by_id(op["in"]) if op.get("in") else None
        doc.find_replace(
            op["find"],
            op["text"],
            scope=scope,
            all=bool(op.get("all", False)),
            occurrence=op.get("occurrence"),
        )
    elif kind == "apply_style":
        doc.anchor_by_id(op["anchor_id"]).apply_style(op["name"])
    elif kind == "format_paragraph":
        kwargs = {
            k: op[k]
            for k in (
                "alignment",
                "left_indent",
                "right_indent",
                "first_line_indent",
                "space_before",
                "space_after",
                "page_break_before",
            )
            if k in op
        }
        doc.anchor_by_id(op["anchor_id"]).format_paragraph(**kwargs)
    elif kind == "format_run":
        kwargs = {k: op[k] for k in _RUN_FIELDS if k in op}
        doc.anchor_by_id(op["anchor_id"]).format_run(**kwargs)
    elif kind == "set_shading":
        kwargs = {k: op[k] for k in ("fill", "pattern") if k in op}
        doc.anchor_by_id(op["anchor_id"]).set_shading(**kwargs)
    elif kind == "set_borders":
        kwargs = {k: op[k] for k in ("sides", "style", "weight", "color") if k in op}
        doc.anchor_by_id(op["anchor_id"]).set_borders(**kwargs)
    elif kind == "add_tab_stop":
        kwargs = {k: op[k] for k in ("align", "leader") if k in op}
        doc.anchor_by_id(op["anchor_id"]).add_tab_stop(op["position"], **kwargs)
    elif kind == "add_style":
        kwargs = {k: op[k] for k in ("type", "based_on", "next_style") if k in op}
        style = doc.styles.add(op["name"], **kwargs)
        return {"style": style.name}
    elif kind == "set_style":
        style = doc.styles[op["name"]]
        run_kwargs = {k: op[k] for k in _STYLE_RUN_FIELDS if k in op}
        para_kwargs = {k: op[k] for k in _PARA_FIELDS if k in op}
        if run_kwargs:
            style.format_run(**run_kwargs)
        if para_kwargs:
            style.format_paragraph(**para_kwargs)
        if "based_on" in op:
            style.base_style = op["based_on"]
        if "next_style" in op:
            style.next_paragraph_style = op["next_style"]
    elif kind == "insert_field":
        kwargs = {"text": op["text"]} if "text" in op else {}
        doc.anchor_by_id(op["anchor_id"]).insert_field(
            op["kind"], where=("before" if op_before(op) else "after"), **kwargs
        )
    elif kind == "set_page_setup":
        kwargs = {k: op[k] for k in _PAGE_SETUP_FIELDS if k in op}
        doc.sections[op["section"]].set_page_setup(**kwargs)
    elif kind == "update_fields":
        doc.update_fields()
    elif kind == "insert_footnote":
        note = doc.anchor_by_id(op["anchor_id"]).insert_footnote(
            op["text"], where=("before" if op_before(op) else "after")
        )
        return {"footnote": note.index, "anchor_id": note.anchor_id}
    elif kind == "insert_endnote":
        note = doc.anchor_by_id(op["anchor_id"]).insert_endnote(
            op["text"], where=("before" if op_before(op) else "after")
        )
        return {"endnote": note.index, "anchor_id": note.anchor_id}
    elif kind == "insert_toc":
        kwargs = {k: op[k] for k in ("levels", "use_heading_styles", "hyperlinks") if k in op}
        doc.anchor_by_id(op["anchor_id"]).insert_toc(
            where=("before" if op_before(op) else "after"), **kwargs
        )
        return {"toc": True}
    elif kind == "add_bookmark":
        doc.bookmarks.add(op["name"], op["anchor_id"])
        return {"bookmark": op["name"]}
    elif kind == "add_hyperlink":
        if ("url" in op) == ("bookmark" in op):
            raise OpError("op 'add_hyperlink' requires exactly one of 'url' or 'bookmark'")
        doc.anchor_by_id(op["anchor_id"]).link_to(
            address=op.get("url"),
            bookmark=op.get("bookmark"),
            text=op.get("text"),
            screen_tip=op.get("screen_tip"),
        )
    elif kind == "insert_cross_reference":
        kwargs = {k: op[k] for k in ("kind", "hyperlink") if k in op}
        doc.anchor_by_id(op["anchor_id"]).insert_cross_reference(
            op["target"], where=("before" if op_before(op) else "after"), **kwargs
        )
    elif kind == "insert_caption":
        kwargs = {k: op[k] for k in ("label", "text") if k in op}
        doc.anchor_by_id(op["anchor_id"]).insert_caption(
            where=("before" if op_before(op) else "after"), **kwargs
        )
    elif kind == "set_cell":
        doc.tables[op["table"]].cell(op["row"], op["col"]).set_text(op["text"])
    elif kind == "add_row":
        doc.tables[op["table"]].add_row(op.get("values"))
    elif kind == "delete_row":
        doc.tables[op["table"]].delete_row(op["row"])
    elif kind == "create_table":
        anchor = doc.anchor_by_id(op["anchor_id"])
        kwargs = {k: op[k] for k in ("style", "data", "header") if k in op}
        table = anchor.insert_table(
            int(op["rows"]),
            int(op["cols"]),
            where=("before" if op_before(op) else "after"),
            **kwargs,
        )
        return {"table": table.index, "rows": table.row_count, "columns": table.column_count}
    elif kind == "delete_table":
        doc.tables[op["table"]].delete()
    elif kind == "insert_break":
        doc.anchor_by_id(op["anchor_id"]).insert_break(
            op.get("kind", "page"),
            where=("before" if op_before(op) else "after"),
        )
    elif kind == "add_comment":
        anchor = doc.anchor_by_id(op["anchor_id"])
        doc.comments.add(anchor, op["text"], author=op.get("author"))
    elif kind == "resolve_comment":
        doc.comments[op["index"]].resolve()
    elif kind == "delete_comment":
        doc.comments[op["index"]].delete()
    elif kind == "apply_list":
        continue_previous = bool(op.get("continue_previous", op.get("continue", False)))
        doc.anchor_by_id(op["anchor_id"]).apply_list(
            op.get("type", "bulleted"), continue_previous=continue_previous
        )
    elif kind == "remove_list":
        doc.anchor_by_id(op["anchor_id"]).remove_list()
    elif kind == "restart_numbering":
        doc.anchor_by_id(op["anchor_id"]).restart_numbering()
    elif kind == "indent_list":
        doc.anchor_by_id(op["anchor_id"]).indent_list()
    elif kind == "outdent_list":
        doc.anchor_by_id(op["anchor_id"]).outdent_list()
    elif kind == "write_header":
        doc.sections[op["section"]].header(op.get("which", "primary")).set_text(op["text"])
    elif kind == "write_footer":
        doc.sections[op["section"]].footer(op.get("which", "primary")).set_text(op["text"])
    # Only the structure-creating ops (create_table) return early with a result
    # dict; every other op falls through here and reports nothing.
    return None


def run_batch(
    doc: Document,
    ops: list[dict[str, Any]],
    *,
    label: str,
    tracked: bool = False,
) -> tuple[dict[str, Any], WordliveError | None]:
    """Apply `ops` to `doc` in one atomic-undo (and optional tracked) scope.

    Stops at the first failing op, recording a `failure` dict (its index, the op,
    the error message and type, plus `matches` for an ambiguous find).

    Returns `(result, failure_exc)`:
      - `result` is `{"ok", "ops_run", "label", "failure"?}`, always present and
        JSON-serialisable — the payload to emit/return verbatim.
      - `failure_exc` is the original `WordliveError` on failure, else `None`, so
        the caller can map it to the right CLI exit code or MCP error code
        (re-raise it after emitting `result`).

    The successful prefix of ops still rolls back as one undo step.
    """
    ops_run = 0
    outputs: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    failure_exc: WordliveError | None = None
    failure_meta: dict[str, Any] | None = None
    tracking = doc.tracked_changes() if tracked else nullcontext()
    with tracking, doc.edit(label):
        for i, op in enumerate(ops):
            try:
                out = apply_op(doc, op)
            except WordliveError as exc:
                failure_exc = exc
                failure_meta = {
                    "index": i,
                    "op": op,
                    "error": str(exc),
                    "type": type(exc).__name__,
                }
                if isinstance(exc, AmbiguousMatchError):
                    failure_meta["matches"] = exc.matches
                break
            # The op applied, so its kind is valid; flag any stray fields it
            # ignored so a silent-but-wrong payload (e.g. `style` on an inline
            # append, or a typo'd key) doesn't pass as a clean success.
            kind = op.get("op")
            for field in unexpected_fields(op, str(kind)):
                warnings.append(
                    {
                        "index": i,
                        "op": kind,
                        "field": field,
                        "message": f"op {kind!r} does not use field {field!r}; it was ignored",
                    }
                )
            if out is not None:
                outputs.append({"index": i, "op": op.get("op"), **out})
            ops_run += 1

    if failure_exc is None:
        result: dict[str, Any] = {"ok": True, "ops_run": ops_run, "label": label}
        if outputs:
            # Only on success: a failed batch rolls back as one undo step, so
            # any structure a prior op created no longer exists to report.
            result["outputs"] = outputs
        if warnings:
            result["warnings"] = warnings
        return result, None

    assert failure_meta is not None  # set together with failure_exc
    result = {"ok": False, "ops_run": ops_run, "label": label, "failure": failure_meta}
    if warnings:
        result["warnings"] = warnings
    return result, failure_exc
