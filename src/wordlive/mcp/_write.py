"""`word_write` dispatch: command→op translation and application."""

from __future__ import annotations

from typing import Any

from .. import attach
from .._ops import (
    _PAGE_SETUP_FIELDS,
    _PARA_FIELDS,
    _STYLE_RUN_FIELDS,
    OP_OPTIONAL_FIELDS,
    pick_doc,
    run_batch,
)
from .._paths import PathPolicy
from ..exceptions import OpError
from ._common import _need
from ._worker import Worker


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
        if p.get("bind") is not None:
            op["bind"] = p["bind"]
        return op
    if command == "insert_block":
        op = {
            "op": "insert_block",
            "anchor_id": need("anchor_id"),
            "items": need("items"),
            "before": bool(p.get("before", False)),
        }
        if p.get("bind") is not None:
            op["bind"] = p["bind"]
        return op
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
        if p.get("bind") is not None:
            op["bind"] = p["bind"]
        return op
    if command == "insert_markdown":
        op = {
            "op": "insert_markdown",
            "anchor_id": need("anchor_id"),
            "markdown": need("markdown"),
            "before": bool(p.get("before", False)),
        }
        if p.get("bind") is not None:
            op["bind"] = p["bind"]
        return op
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
            if p.get("mode") is not None:
                op["mode"] = p["mode"]
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
            "line_spacing",
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
        # param used by apply_style/create_table); the canonical op field is
        # `style`, but the exec op also accepts `line_style` as an alias, so a
        # hand-built word_exec batch reusing this name is honoured too.
        if p.get("line_style") is not None:
            op["style"] = p["line_style"]
        for k in ("sides", "weight", "color"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "cell_valign":
        return {
            "op": "set_cell_vertical_alignment",
            "anchor_id": need("anchor_id"),
            "align": need("align"),
        }
    if command == "drop_cap":
        op = {"op": "drop_cap", "anchor_id": need("anchor_id")}
        for k in ("lines", "position", "distance", "font"):
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
    if command == "regularize":
        op = {"op": "regularize"}
        for key in ("rules", "within", "profile", "dry_run", "allow_content"):
            if p.get(key) is not None:
                op[key] = p[key]
        return op
    if command == "set_property":
        op = {"op": "set_property", "name": need("name"), "value": need("value")}
        if p.get("custom") is not None:
            op["custom"] = bool(p["custom"])
        return op
    if command == "delete_property":
        return {"op": "delete_property", "name": need("name")}
    if command == "set_variable":
        return {"op": "set_variable", "name": need("name"), "value": need("value")}
    if command == "delete_variable":
        return {"op": "delete_variable", "name": need("name")}
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
    if command == "pin":
        op = {"op": "pin", "anchor_id": need("anchor_id")}
        if p.get("name") is not None:
            op["name"] = p["name"]
        return op
    if command == "pin_outline":
        op = {"op": "pin_outline"}
        if p.get("levels") is not None:
            op["levels"] = p["levels"]
        return op
    if command == "add_hyperlink":
        if (p.get("url") is None) == (p.get("bookmark") is None):
            raise OpError("add_hyperlink requires exactly one of 'url' or 'bookmark'")
        op = {"op": "add_hyperlink", "anchor_id": need("anchor_id")}
        for k in ("url", "bookmark", "text", "screen_tip"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "set_hyperlink":
        # Vocabulary parity with add_hyperlink: url -> address, bookmark ->
        # sub_address. The op/CLI layer uses Word's own terms.
        op = {"op": "set_hyperlink", "index": need("index")}
        for mcp_key, op_key in (
            ("url", "address"),
            ("bookmark", "sub_address"),
            ("text", "text"),
            ("screen_tip", "screen_tip"),
        ):
            if p.get(mcp_key) is not None:
                op[op_key] = p[mcp_key]
        if len(op) == 2:  # only "op" and "index" set
            raise OpError(
                "set_hyperlink needs at least one of 'url', 'bookmark', 'text', 'screen_tip'"
            )
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
    if command == "create_content_control":
        op = {"op": "create_content_control", "anchor_id": need("anchor_id")}
        if p.get("kind") is not None:
            op["kind"] = p["kind"]
        if p.get("where") is not None:
            op["where"] = p["where"]
        for k in ("title", "tag", "items", "lock_contents", "lock_control"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "set_cc_properties":
        op = {"op": "set_cc_properties", "anchor_id": need("anchor_id")}
        for k in OP_OPTIONAL_FIELDS["set_cc_properties"]:
            if p.get(k) is not None:
                op[k] = p[k]
        if len(op) == 2:  # only "op" and "anchor_id" set
            raise OpError(
                "set_cc_properties needs at least one of 'title', 'tag', "
                "'lock_contents', 'lock_control'"
            )
        return op
    if command == "set_cc_items":
        return {"op": "set_cc_items", "anchor_id": need("anchor_id"), "items": need("items")}
    if command == "mark_index_entry":
        op = {"op": "mark_index_entry", "anchor_id": need("anchor_id"), "entry": need("entry")}
        for k in ("cross_reference", "bold", "italic"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "insert_index":
        op = {
            "op": "insert_index",
            "anchor_id": p.get("anchor_id") or "end",
            "before": bool(p.get("before", False)),
        }
        for k in ("columns", "run_in", "right_align_page_numbers"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "insert_table_of_figures":
        op = {
            "op": "insert_table_of_figures",
            "anchor_id": p.get("anchor_id") or "start",
            "before": bool(p.get("before", False)),
        }
        for k in ("label", "include_label", "hyperlinks", "right_align_page_numbers"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "set_bibliography_style":
        return {"op": "set_bibliography_style", "style": need("style")}
    if command == "add_source":
        op = {"op": "add_source", "source_type": p.get("source_type") or "book"}
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
            "xml",
        ):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "insert_citation":
        op = {"op": "insert_citation", "anchor_id": need("anchor_id"), "tag": need("tag")}
        op["before"] = bool(p.get("before", False))
        for k in (
            "pages",
            "prefix",
            "suffix",
            "volume",
            "suppress_author",
            "suppress_year",
            "suppress_title",
            "locale",
        ):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "insert_bibliography":
        return {
            "op": "insert_bibliography",
            "anchor_id": p.get("anchor_id") or "end",
            "before": bool(p.get("before", False)),
        }
    if command == "mark_citation":
        op = {
            "op": "mark_citation",
            "anchor_id": need("anchor_id"),
            "long_citation": need("long_citation"),
            "before": bool(p.get("before", False)),
        }
        for k in ("short_citation", "category"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "insert_table_of_authorities":
        op = {
            "op": "insert_table_of_authorities",
            "anchor_id": p.get("anchor_id") or "end",
            "before": bool(p.get("before", False)),
        }
        for k in (
            "category",
            "passim",
            "keep_entry_formatting",
            "entry_separator",
            "page_range_separator",
        ):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "apply_theme":
        return {"op": "apply_theme", "theme": need("theme")}
    if command == "set_theme_colors":
        op = {"op": "set_theme_colors"}
        if p.get("scheme") is not None:
            op["scheme"] = p["scheme"]
        if p.get("colors") is not None:
            op["colors"] = p["colors"]
        return op
    if command == "set_theme_fonts":
        op = {"op": "set_theme_fonts"}
        for k in ("scheme", "major", "minor"):
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
        action = need("action")
        if action == "format":
            op = {
                "op": "apply_list_format",
                "anchor_id": need("anchor_id"),
                "levels": need("levels"),
            }
            if p.get("continue_previous") is not None:
                op["continue_previous"] = bool(p["continue_previous"])
            return op
        mapping = {
            "apply": "apply_list",
            "remove": "remove_list",
            "restart": "restart_numbering",
            "indent": "indent_list",
            "outdent": "outdent_list",
        }
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
    if command == "revision":
        action = need("action")
        if action == "accept":
            return {"op": "accept_revision", "index": need("index")}
        if action == "reject":
            return {"op": "reject_revision", "index": need("index")}
        if action in ("accept_all", "reject_all"):
            op = {"op": f"{action}_revisions"}
            if p.get("anchor_id") is not None:
                op["anchor_id"] = p["anchor_id"]
            return op
        raise OpError(f"unknown revision action: {action!r}")
    if command == "watermark":
        if p.get("remove"):
            return {"op": "remove_watermark"}
        op = {"op": "set_watermark", "text": need("text")}
        for k in ("font", "color", "layout", "semitransparent"):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "text_box":
        op = {"op": "insert_text_box", "anchor_id": need("anchor_id"), "text": need("text")}
        if p.get("before") is not None:
            op["before"] = bool(p["before"])
        for k in (
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
        ):
            if p.get(k) is not None:
                op[k] = p[k]
        return op
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
        if action == "add_column":
            op = {"op": "add_column", "table": need("table")}
            if p.get("values") is not None:
                op["values"] = p["values"]
            return op
        if action == "delete_column":
            return {"op": "delete_column", "table": need("table"), "column": need("column")}
        if action == "merge_cells":
            return {
                "op": "merge_cells",
                "table": need("table"),
                "from": need("from"),
                "to": need("to"),
            }
        if action == "split_cell":
            op = {"op": "split_cell", "table": need("table"), "cell": need("cell")}
            if p.get("rows") is not None:
                op["rows"] = p["rows"]
            if p.get("cols") is not None:
                op["cols"] = p["cols"]
            return op
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
            if p.get("bind") is not None:
                op["bind"] = p["bind"]
            return op
        if action == "autofit":
            op = {"op": "autofit_table", "table": need("table")}
            if p.get("mode") is not None:
                op["mode"] = p["mode"]
            return op
        if action == "delete":
            return {"op": "delete_table", "table": need("table")}
        if action == "set_style":
            return {"op": "set_table_style", "table": need("table"), "style": need("style")}
        if action == "set_alignment":
            return {
                "op": "set_table_alignment",
                "table": need("table"),
                "alignment": need("alignment"),
            }
        if action == "set_borders":
            op = {"op": "set_table_borders", "table": need("table")}
            for key in ("sides", "style", "line_style", "weight", "color"):
                if p.get(key) is not None:
                    op[key] = p[key]
            return op
        if action == "set_banding":
            op = {"op": "set_table_banding", "table": need("table")}
            for key in (
                "first_row",
                "last_row",
                "first_column",
                "last_column",
                "banded_rows",
                "banded_columns",
            ):
                if p.get(key) is not None:
                    op[key] = bool(p[key])
            return op
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
    if command == "insert_equation":
        given = [k for k in ("unicodemath", "latex", "mathml") if p.get(k) is not None]
        if len(given) != 1:
            raise OpError(
                "insert_equation requires exactly one of 'unicodemath', 'latex', or 'mathml'"
            )
        op = {
            "op": "insert_equation",
            "anchor_id": need("anchor_id"),
            "before": bool(p.get("before", False)),
        }
        op[given[0]] = p[given[0]]
        if p.get("display") is not None:
            op["display"] = bool(p["display"])
        return op
    if command == "insert_chart":
        op = {
            "op": "insert_chart",
            "anchor_id": need("anchor_id"),
            "kind": need("kind"),
            "data": need("data"),
            "before": bool(p.get("before", False)),
        }
        if p.get("title") is not None:
            op["title"] = p["title"]
        return op
    if command in (
        "format_chart",
        "format_axis",
        "add_trendline",
        "set_series_color",
        "format_series",
        "add_error_bars",
    ):
        op = {"op": command, "anchor_id": need("anchor_id")}
        if command == "format_axis":
            op["which"] = need("which")
        if command == "set_series_color":
            op["color"] = need("color")
        for k in OP_OPTIONAL_FIELDS[command]:
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command in (
        "set_shape_wrap",
        "set_shape_position",
        "set_shape_size",
        "format_shape",
        "set_shape_alt_text",
        "set_shape_text",
        "set_shape_rotation",
        "set_shape_z_order",
        "set_shape_text_frame",
        "delete_shape",
        "ungroup_shape",
    ):
        op = {"op": command, "anchor_id": need("anchor_id")}
        if command in ("set_shape_alt_text", "set_shape_text"):
            op["text"] = need("text")
        if command == "set_shape_rotation":
            op["degrees"] = need("degrees")
        if command == "set_shape_z_order":
            op["order"] = need("order")
        for k in OP_OPTIONAL_FIELDS[command]:
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command in ("set_shape_crop", "set_image_crop"):
        # crop_* params keep the model from confusing crop edges with a shape's
        # position left/top; map them onto the op's left/top/right/bottom fields.
        op = {"op": command, "anchor_id": need("anchor_id")}
        for src, dst in (
            ("crop_left", "left"),
            ("crop_top", "top"),
            ("crop_right", "right"),
            ("crop_bottom", "bottom"),
        ):
            if p.get(src) is not None:
                op[dst] = p[src]
        return op
    if command in ("set_image_alt_text", "set_image_size"):
        op = {"op": command, "anchor_id": need("anchor_id")}
        if command == "set_image_alt_text":
            op["text"] = need("text")
        for k in OP_OPTIONAL_FIELDS[command]:
            if p.get(k) is not None:
                op[k] = p[k]
        return op
    if command == "group_shapes":
        shapes = p.get("shapes")
        if not isinstance(shapes, list) or len(shapes) < 2:
            raise OpError("group_shapes requires 'shapes': a list of two or more shape:N ids")
        return {"op": "group_shapes", "shapes": shapes}
    if command == "replace_shape_image":
        if (p.get("image_base64") is None) == (p.get("path") is None):
            raise OpError("replace_shape_image requires exactly one of 'image_base64' or 'path'")
        op = {"op": "replace_shape_image", "anchor_id": need("anchor_id")}
        if p.get("image_base64") is not None:
            op["base64"] = p["image_base64"]
        else:
            op["path"] = p["path"]
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
