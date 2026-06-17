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

import re
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
    "insert_paragraph": ("anchor_id",),  # exactly one of text/runs (checked in apply_op)
    "insert_block": ("anchor_id", "items"),
    "insert_section": ("anchor_id", "heading", "body"),
    "insert_markdown": ("anchor_id", "markdown"),
    "replace_section": ("anchor_id",),  # exactly one of body/markdown (checked in apply_op)
    "delete_paragraph": ("anchor_id",),
    "append_paragraph": ("text",),
    "append": ("text",),
    "append_inline": ("text",),
    "prepend_paragraph": ("text",),
    "prepend": ("text",),
    "prepend_inline": ("text",),
    "insert_image": ("anchor_id", "wrap"),
    "insert_equation": ("anchor_id",),  # exactly one of unicodemath/latex/mathml (apply_op)
    "insert_chart": ("anchor_id", "kind", "data"),
    "format_chart": ("anchor_id",),
    "format_axis": ("anchor_id", "which"),
    "add_trendline": ("anchor_id",),
    "set_series_color": ("anchor_id", "color"),
    "replace": ("anchor_id", "text"),
    "find_replace": ("find", "text"),
    "apply_style": ("anchor_id", "name"),
    "format_paragraph": ("anchor_id",),
    "format_run": ("anchor_id",),
    "set_shading": ("anchor_id",),
    "set_borders": ("anchor_id",),
    "drop_cap": ("anchor_id",),
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
    "pin": ("anchor_id",),
    "pin_outline": (),
    "add_hyperlink": ("anchor_id",),
    "insert_cross_reference": ("anchor_id", "target"),
    "insert_caption": ("anchor_id",),
    "create_content_control": ("anchor_id",),
    "mark_index_entry": ("anchor_id", "entry"),
    "insert_index": ("anchor_id",),
    "insert_table_of_figures": ("anchor_id",),
    "set_bibliography_style": ("style",),
    "add_source": ("source_type",),
    "insert_citation": ("anchor_id", "tag"),
    "insert_bibliography": ("anchor_id",),
    "mark_citation": ("anchor_id", "long_citation"),
    "insert_table_of_authorities": ("anchor_id",),
    "apply_theme": ("theme",),
    "set_theme_colors": (),
    "set_theme_fonts": (),
    "set_property": ("name", "value"),
    "delete_property": ("name",),
    "set_variable": ("name", "value"),
    "delete_variable": ("name",),
    "set_cell": ("table", "row", "col", "text"),
    "autofit_table": ("table",),
    "add_row": ("table",),
    "append_record": ("table", "record"),
    "update_row": ("table", "key", "values"),
    "delete_row": ("table", "row"),
    "set_heading_row": ("table",),
    "create_table": ("anchor_id",),  # rows/cols required only without data (apply_op)
    "delete_table": ("table",),
    "insert_break": ("anchor_id",),
    "add_comment": ("anchor_id", "text"),
    "resolve_comment": ("index",),
    "delete_comment": ("index",),
    "accept_revision": ("index",),
    "reject_revision": ("index",),
    "accept_all_revisions": (),
    "reject_all_revisions": (),
    "set_watermark": ("text",),
    "remove_watermark": (),
    "insert_text_box": ("anchor_id", "text"),
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
    "line_spacing",
    "page_break_before",
    "keep_together",
    "keep_with_next",
    "widow_control",
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
    "insert_paragraph": ("text", "runs", "style", "bind", *_WHERE_FIELDS),
    "insert_block": ("items", "bind", *_WHERE_FIELDS),
    "insert_section": ("level", "bind", *_WHERE_FIELDS),
    "insert_markdown": ("bind", *_WHERE_FIELDS),
    "replace_section": ("body", "markdown"),
    "delete_paragraph": (),
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
    "insert_equation": ("unicodemath", "latex", "mathml", "display", *_WHERE_FIELDS),
    "insert_chart": ("title", *_WHERE_FIELDS),
    "format_chart": (
        "title",
        "legend",
        "legend_position",
        "chart_style",
        "background",
        "plot_background",
        "font",
        "font_size",
        "font_color",
        "data_labels",
        "data_label_format",
        "chart_type",
    ),
    "format_axis": ("title", "minimum", "maximum", "scale", "number_format", "gridlines"),
    "add_trendline": (
        "series",
        "kind",
        "display_equation",
        "display_r_squared",
        "forward",
        "backward",
    ),
    "set_series_color": ("series", "point"),
    "replace": (),
    "find_replace": ("in", "all", "occurrence"),
    "apply_style": (),
    "format_paragraph": _PARA_FIELDS,
    "format_run": _RUN_FIELDS,
    "set_shading": ("fill", "pattern"),
    # `line_style` is an accepted alias for `style` (the line style) — it is the
    # name the MCP `set_borders` command and its `word_write` schema use, so a
    # batch author who learned it there gets it honoured here instead of warned-
    # and-ignored. See the alias handling in `apply_op`.
    "set_borders": ("sides", "style", "line_style", "weight", "color"),
    "drop_cap": ("lines", "position", "distance", "font"),
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
    "pin": ("name",),
    "pin_outline": ("levels",),
    "add_hyperlink": ("url", "bookmark", "text", "screen_tip"),
    "insert_cross_reference": ("kind", "hyperlink", *_WHERE_FIELDS),
    "insert_caption": ("label", "text", "position", *_WHERE_FIELDS),
    "create_content_control": (
        "kind",
        "title",
        "tag",
        "items",
        "where",
        "lock_contents",
        "lock_control",
    ),
    "mark_index_entry": ("cross_reference", "bold", "italic"),
    "insert_index": ("columns", "run_in", "right_align_page_numbers", *_WHERE_FIELDS),
    "insert_table_of_figures": (
        "label",
        "include_label",
        "hyperlinks",
        "right_align_page_numbers",
        *_WHERE_FIELDS,
    ),
    "set_bibliography_style": (),
    "add_source": (
        "tag",
        "author",
        "title",
        "year",
        "publisher",
        "city",
        "journal_name",
        "volume",
        "issue",
        "pages",
        "url",
        "edition",
        "doi",
        "xml",
    ),
    "insert_citation": (
        "pages",
        "prefix",
        "suffix",
        "volume",
        "suppress_author",
        "suppress_year",
        "suppress_title",
        "locale",
        *_WHERE_FIELDS,
    ),
    "insert_bibliography": _WHERE_FIELDS,
    "mark_citation": ("short_citation", "category", *_WHERE_FIELDS),
    "insert_table_of_authorities": (
        "category",
        "passim",
        "keep_entry_formatting",
        "entry_separator",
        "page_range_separator",
        *_WHERE_FIELDS,
    ),
    "apply_theme": (),
    "set_theme_colors": ("scheme", "colors"),
    "set_theme_fonts": ("scheme", "major", "minor"),
    "set_property": ("custom",),
    "delete_property": (),
    "set_variable": (),
    "delete_variable": (),
    "set_cell": (),
    "autofit_table": ("mode",),
    "add_row": ("values",),
    "append_record": (),
    "update_row": ("column",),
    "delete_row": (),
    "set_heading_row": ("row", "heading", "allow_break"),
    "create_table": ("rows", "cols", "style", "data", "header", "bind", *_WHERE_FIELDS),
    "delete_table": (),
    "insert_break": ("kind", *_WHERE_FIELDS),
    "add_comment": ("author",),
    "resolve_comment": (),
    "delete_comment": (),
    "accept_revision": (),
    "reject_revision": (),
    "accept_all_revisions": ("anchor_id",),
    "reject_all_revisions": ("anchor_id",),
    "set_watermark": ("font", "color", "layout", "semitransparent"),
    "remove_watermark": (),
    "insert_text_box": (
        "width",
        "height",
        "wrap",
        "font",
        "size",
        "bold",
        "italic",
        "alignment",
        "fill",
        "border",
        *_WHERE_FIELDS,
    ),
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


def _apply_bind(doc: Document, op: dict[str, Any], rng: Any) -> dict[str, Any]:
    """Mint a durable handle over a just-inserted range when the op carries `bind`.

    `bind` is `true` (auto random code) or a readable slug string. Returns
    `{"pin": "pin:<code>"}` to merge into the op's output dict, or `{}` when
    there's no usable `bind`. `rng` is the `Anchor` the insert returned.
    """
    spec = op.get("bind")
    if spec in (None, False, ""):
        return {}
    name = spec if isinstance(spec, str) else None
    return {"pin": doc.pin(rng, name=name)["pin"]}


# A whole-string reference to an earlier op's output, e.g. `$ops[0].table`.
_OP_REF_RE = re.compile(r"^\$ops\[(\d+)\]\.([A-Za-z_]\w*)$")


def _resolve_op_refs(op: dict[str, Any], results: dict[int, dict[str, Any]]) -> dict[str, Any]:
    """Substitute `$ops[N].field` references in `op` with prior ops' outputs.

    A batch op can target what an earlier op produced — e.g. after a
    `create_table` at index 0, ``{"op": "set_cell", "table": "$ops[0].table", …}``
    reuses the new table's index without a second round-trip. Only a *whole*
    string value of the exact form ``$ops[N].field`` is substituted (no
    interpolation inside a larger string); the walk recurses into nested list /
    dict values. Returns a new dict — the caller's op is left untouched.

    Raises `OpError` (a `WordliveError`, so the batch records it as this op's
    failure) for a forward / self / failed-op reference or an unknown field.
    """

    def resolve(value: Any) -> Any:
        if isinstance(value, str):
            m = _OP_REF_RE.match(value)
            if not m:
                return value
            idx, field = int(m.group(1)), m.group(2)
            if idx not in results:
                raise OpError(
                    f"reference {value!r} points at op {idx}, which has not produced output "
                    "yet (only earlier, successful ops can be referenced)"
                )
            out = results[idx]
            if field not in out:
                available = ", ".join(sorted(out)) or "none"
                raise OpError(
                    f"reference {value!r}: op {idx} has no output field {field!r} "
                    f"(available: {available})"
                )
            return out[field]
        if isinstance(value, list):
            return [resolve(v) for v in value]
        if isinstance(value, dict):
            return {k: resolve(v) for k, v in value.items()}
        return value

    return resolve(op)


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
        if ("text" in op) == ("runs" in op):
            raise OpError("op 'insert_paragraph' requires exactly one of 'text' or 'runs'")
        anchor = doc.anchor_by_id(op["anchor_id"])
        where = "before" if op_before(op) else "after"
        if "runs" in op:
            # Structured inline formatting — route through the block primitive
            # (one item). `text` stays a literal plain insert (no markdown sugar);
            # markdown lives in insert_block's item text.
            rng = anchor.insert_block([{"runs": op["runs"], "style": op.get("style")}], where=where)
        elif op.get("bind"):
            # `bind` needs a handle on the new paragraph; route the literal text
            # through insert_block as a single run (no markdown sugar) to get a
            # RangeAnchor to pin, preserving the plain-text-is-literal contract.
            rng = anchor.insert_block(
                [{"runs": [{"text": op["text"]}], "style": op.get("style")}], where=where
            )
        else:
            if op_before(op):
                anchor.insert_paragraph_before(op["text"], style=op.get("style"))
            else:
                anchor.insert_paragraph_after(op["text"], style=op.get("style"))
            return None
        return _apply_bind(doc, op, rng) or None
    elif kind == "insert_block":
        rng = doc.anchor_by_id(op["anchor_id"]).insert_block(
            op["items"], where=("before" if op_before(op) else "after")
        )
        return {
            "anchor_id": rng.anchor_id,
            "paragraphs": len(op["items"]),
            **_apply_bind(doc, op, rng),
        }
    elif kind == "insert_section":
        rng = doc.anchor_by_id(op["anchor_id"]).insert_section(
            op["heading"],
            op["body"],
            level=int(op.get("level", 1)),
            where=("before" if op_before(op) else "after"),
        )
        return {"anchor_id": rng.anchor_id, **_apply_bind(doc, op, rng)}
    elif kind == "insert_markdown":
        rng = doc.anchor_by_id(op["anchor_id"]).insert_markdown(
            op["markdown"], where=("before" if op_before(op) else "after")
        )
        return {"anchor_id": rng.anchor_id, **_apply_bind(doc, op, rng)}
    elif kind == "replace_section":
        if ("body" in op) == ("markdown" in op):
            raise OpError("op 'replace_section' requires exactly one of 'body' or 'markdown'")
        anchor = doc.anchor_by_id(op["anchor_id"])
        if not hasattr(anchor, "replace_section_body"):
            raise OpError(
                f"replace_section needs a heading anchor; {op['anchor_id']!r} is a {anchor.kind}"
            )
        if "markdown" in op:
            rng = anchor.replace_section_body(op["markdown"], markdown=True)
        else:
            rng = anchor.replace_section_body(op["body"])
        return {"anchor_id": rng.anchor_id}
    elif kind == "delete_paragraph":
        doc.delete_paragraph(op["anchor_id"])
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
    elif kind == "insert_equation":
        eq_kwargs = {k: op[k] for k in ("unicodemath", "latex", "mathml", "display") if k in op}
        equation = doc.anchor_by_id(op["anchor_id"]).insert_equation(
            where=("before" if op_before(op) else "after"), **eq_kwargs
        )
        return {"equation": equation.index, "anchor_id": equation.anchor_id}
    elif kind == "insert_chart":
        chart = doc.anchor_by_id(op["anchor_id"]).insert_chart(
            op["kind"],
            op["data"],
            title=op.get("title"),
            where=("before" if op_before(op) else "after"),
        )
        return {"chart": chart.index, "anchor_id": chart.anchor_id}
    elif kind in ("format_chart", "format_axis", "add_trendline", "set_series_color"):
        from ._anchors import ChartAnchor  # lazy: avoid an _ops → _anchors import cycle

        anchor = doc.anchor_by_id(op["anchor_id"])
        if not isinstance(anchor, ChartAnchor):
            raise OpError(f"{op['anchor_id']!r} is not a chart; {kind} needs a chart:N anchor")
        fields = OP_OPTIONAL_FIELDS[kind]
        kwargs = {k: op[k] for k in fields if k in op}
        if kind == "format_chart":
            anchor.format(**kwargs)
        elif kind == "format_axis":
            anchor.set_axis(op["which"], **kwargs)
        elif kind == "add_trendline":
            anchor.add_trendline(**kwargs)
        else:  # set_series_color
            anchor.set_series_color(op["color"], **kwargs)
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
        kwargs = {k: op[k] for k in _PARA_FIELDS if k in op}
        doc.anchor_by_id(op["anchor_id"]).format_paragraph(**kwargs)
    elif kind == "format_run":
        kwargs = {k: op[k] for k in _RUN_FIELDS if k in op}
        doc.anchor_by_id(op["anchor_id"]).format_run(**kwargs)
    elif kind == "set_shading":
        kwargs = {k: op[k] for k in ("fill", "pattern") if k in op}
        doc.anchor_by_id(op["anchor_id"]).set_shading(**kwargs)
    elif kind == "set_borders":
        kwargs = {k: op[k] for k in ("sides", "weight", "color") if k in op}
        # Accept `line_style` as an alias for `style` (the MCP / word_write name);
        # an explicit `style` wins if somehow both are present.
        if "style" in op:
            kwargs["style"] = op["style"]
        elif "line_style" in op:
            kwargs["style"] = op["line_style"]
        doc.anchor_by_id(op["anchor_id"]).set_borders(**kwargs)
    elif kind == "drop_cap":
        kwargs = {k: op[k] for k in ("position", "distance", "font") if k in op}
        if "lines" in op:
            kwargs["lines"] = int(op["lines"])
        doc.anchor_by_id(op["anchor_id"]).drop_cap(**kwargs)
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
    elif kind == "pin":
        return doc.pin(op["anchor_id"], name=op.get("name"))
    elif kind == "pin_outline":
        return {"pins": doc.pin_outline(levels=op.get("levels"))}
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
        if "position" in op:
            position = op["position"]
        elif "before" in op:  # back-compat: before=True meant "above"
            position = "above" if op["before"] else "below"
        else:
            position = None
        doc.anchor_by_id(op["anchor_id"]).insert_caption(position=position, **kwargs)
    elif kind == "create_content_control":
        cc_kwargs = {
            k: op[k] for k in ("title", "tag", "items", "lock_contents", "lock_control") if k in op
        }
        cc = doc.anchor_by_id(op["anchor_id"]).insert_content_control(
            op.get("kind", "rich_text"), where=op.get("where", "wrap"), **cc_kwargs
        )
        return {
            "content_control": cc.name or None,
            "anchor_id": cc.anchor_id if cc.name else None,
        }
    elif kind == "mark_index_entry":
        mk = {k: op[k] for k in ("cross_reference", "bold", "italic") if k in op}
        doc.anchor_by_id(op["anchor_id"]).mark_index_entry(op["entry"], **mk)
    elif kind == "insert_index":
        ik = {k: op[k] for k in ("columns", "run_in", "right_align_page_numbers") if k in op}
        doc.anchor_by_id(op["anchor_id"]).insert_index(
            where=("before" if op_before(op) else "after"), **ik
        )
        return {"index": True}
    elif kind == "insert_table_of_figures":
        tk = {
            k: op[k]
            for k in ("label", "include_label", "hyperlinks", "right_align_page_numbers")
            if k in op
        }
        doc.anchor_by_id(op["anchor_id"]).insert_table_of_figures(
            where=("before" if op_before(op) else "after"), **tk
        )
        return {"table_of_figures": True}
    elif kind == "set_bibliography_style":
        doc.bibliography_style = op["style"]
    elif kind == "add_source":
        if "xml" in op:
            src = doc.sources.add_xml(op["xml"])
        else:
            sk = {
                k: op[k]
                for k in (
                    "tag",
                    "author",
                    "title",
                    "year",
                    "publisher",
                    "city",
                    "journal_name",
                    "volume",
                    "issue",
                    "pages",
                    "url",
                    "edition",
                    "doi",
                )
                if k in op
            }
            src = doc.sources.add(op["source_type"], **sk)
        return {"source": src.tag}
    elif kind == "insert_citation":
        ck = {
            k: op[k]
            for k in (
                "pages",
                "prefix",
                "suffix",
                "volume",
                "suppress_author",
                "suppress_year",
                "suppress_title",
                "locale",
            )
            if k in op
        }
        doc.anchor_by_id(op["anchor_id"]).insert_citation(
            op["tag"], where=("before" if op_before(op) else "after"), **ck
        )
        return {"citation": op["tag"]}
    elif kind == "insert_bibliography":
        doc.anchor_by_id(op["anchor_id"]).insert_bibliography(
            where=("before" if op_before(op) else "after")
        )
        return {"bibliography": True}
    elif kind == "mark_citation":
        mc = {k: op[k] for k in ("short_citation", "category") if k in op}
        doc.anchor_by_id(op["anchor_id"]).mark_citation(
            op["long_citation"], where=("before" if op_before(op) else "after"), **mc
        )
    elif kind == "insert_table_of_authorities":
        ak = {
            k: op[k]
            for k in (
                "category",
                "passim",
                "keep_entry_formatting",
                "entry_separator",
                "page_range_separator",
            )
            if k in op
        }
        doc.anchor_by_id(op["anchor_id"]).insert_table_of_authorities(
            where=("before" if op_before(op) else "after"), **ak
        )
        return {"table_of_authorities": True}
    elif kind == "apply_theme":
        return {"theme": doc.theme.apply(op["theme"])}
    elif kind == "set_theme_colors":
        colors = op.get("colors") or {}
        return {"colors": doc.theme.set_colors(scheme=op.get("scheme"), **colors)}
    elif kind == "set_theme_fonts":
        return doc.theme.set_fonts(
            scheme=op.get("scheme"), major=op.get("major"), minor=op.get("minor")
        )
    elif kind == "set_property":
        doc.properties.set(op["name"], op["value"], custom=bool(op.get("custom", False)))
    elif kind == "delete_property":
        doc.properties.delete(op["name"])
    elif kind == "set_variable":
        doc.variables.set(op["name"], op["value"])
    elif kind == "delete_variable":
        doc.variables.delete(op["name"])
    elif kind == "set_cell":
        doc.tables[op["table"]].cell(op["row"], op["col"]).set_text(op["text"])
    elif kind == "autofit_table":
        doc.tables[op["table"]].autofit(op.get("mode", "content"))
    elif kind == "add_row":
        doc.tables[op["table"]].add_row(op.get("values"))
    elif kind == "append_record":
        doc.tables[op["table"]].append_record(op["record"])
    elif kind == "update_row":
        kwargs = {"column": op["column"]} if "column" in op else {}
        doc.tables[op["table"]].update_row(op["key"], op["values"], **kwargs)
    elif kind == "delete_row":
        doc.tables[op["table"]].delete_row(op["row"])
    elif kind == "set_heading_row":
        kwargs = {k: op[k] for k in ("heading", "allow_break") if k in op}
        doc.tables[op["table"]].set_heading_row(int(op.get("row", 1)), **kwargs)
    elif kind == "create_table":
        anchor = doc.anchor_by_id(op["anchor_id"])
        kwargs = {k: op[k] for k in ("style", "data", "header") if k in op}
        # rows/cols are optional when `data` is present — insert_table infers
        # them (and raises OpError if they're missing with nothing to infer from).
        table = anchor.insert_table(
            int(op["rows"]) if "rows" in op else None,
            int(op["cols"]) if "cols" in op else None,
            where=("before" if op_before(op) else "after"),
            **kwargs,
        )
        result = {"table": table.index, "rows": table.row_count, "columns": table.column_count}
        if op.get("bind"):
            trng = table.com.Range
            result.update(_apply_bind(doc, op, doc.range(int(trng.Start), int(trng.End))))
        return result
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
    elif kind == "accept_revision":
        doc.revisions[op["index"]].accept()
    elif kind == "reject_revision":
        doc.revisions[op["index"]].reject()
    elif kind == "accept_all_revisions":
        within = doc.anchor_by_id(op["anchor_id"]) if op.get("anchor_id") else None
        doc.revisions.accept_all(within=within)
    elif kind == "reject_all_revisions":
        within = doc.anchor_by_id(op["anchor_id"]) if op.get("anchor_id") else None
        doc.revisions.reject_all(within=within)
    elif kind == "set_watermark":
        doc.set_watermark(
            op["text"],
            font=op.get("font", "Calibri"),
            color=op.get("color", "#C0C0C0"),
            layout=op.get("layout", "diagonal"),
            semitransparent=bool(op.get("semitransparent", True)),
        )
    elif kind == "remove_watermark":
        doc.remove_watermark()
    elif kind == "insert_text_box":
        doc.anchor_by_id(op["anchor_id"]).insert_text_box(
            op["text"],
            width=op.get("width", 200),
            height=op.get("height", 100),
            wrap=op.get("wrap", "square"),
            where=("before" if op_before(op) else "after"),
            font=op.get("font"),
            size=op.get("size"),
            bold=op.get("bold"),
            italic=op.get("italic"),
            alignment=op.get("alignment"),
            fill=op.get("fill"),
            border=op.get("border"),
        )
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
    # Each successful op's raw output dict, keyed by index — the source for any
    # `$ops[N].field` references a later op makes (see `_resolve_op_refs`).
    results_by_index: dict[int, dict[str, Any]] = {}
    failure_exc: WordliveError | None = None
    failure_meta: dict[str, Any] | None = None
    tracking = doc.tracked_changes() if tracked else nullcontext()
    with tracking, doc.edit(label):
        for i, op in enumerate(ops):
            try:
                out = apply_op(doc, _resolve_op_refs(op, results_by_index))
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
            results_by_index[i] = out or {}
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
