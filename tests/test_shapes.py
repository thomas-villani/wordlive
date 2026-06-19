"""Floating-shape anchor model — `shape:N`, the restyle handle for text boxes,
floating images, and WordArt.

Floating shapes are COM-heavy; these pin the wiring against the `fake_word`
MagicMock — that the right anchor comes back from `insert_text_box` / a floating
`insert_image`, resolves via `doc.shapes` / `anchor_by_id`, and that each mutator
op reaches the right COM property across Python / CLI / exec / MCP. Visual
correctness (and the delete+reinsert `replace_image` swap) is a live-Word concern,
covered by the smoke pass.
"""

from __future__ import annotations

import base64
import json

import pytest

import wordlive
from wordlive.cli.main import EXIT_ANCHOR_NOT_FOUND, EXIT_OK, main
from wordlive.exceptions import AnchorNotFoundError, OpError
from wordlive.mcp._worker import InlineWorker
from wordlive.mcp.server import _build_write_op, _read_impl, _write_impl

# A 1x1 transparent PNG (same as test_images).
_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYPhfDwAChwGA"
    "60e6kgAAAABJRU5ErkJggg=="
)
_B64 = base64.b64encode(_PNG).decode("ascii")

W = InlineWorker()


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


@pytest.fixture
def png_file(tmp_path):
    p = tmp_path / "pic.png"
    p.write_bytes(_PNG)
    return p


# --- insert returns a resolvable shape anchor -----------------------------------


def test_insert_text_box_returns_shape_anchor(fake_word):
    from wordlive._anchors import ShapeAnchor

    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = doc.bookmarks["Address"].insert_text_box("Pull quote")
    assert isinstance(shape, ShapeAnchor)
    assert shape.anchor_id == "shape:1"
    assert shape.shape_type == "text_box"
    # Resolves both ways.
    assert doc.shapes[1].anchor_id == "shape:1"
    assert doc.anchor_by_id("shape:1").shape_type == "text_box"


def test_insert_image_floating_returns_shape_anchor(fake_word, png_file):
    from wordlive._anchors import ShapeAnchor

    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
    assert isinstance(shape, ShapeAnchor)
    assert shape.shape_type == "picture"


def test_insert_image_inline_returns_none(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            result = doc.bookmarks["Address"].insert_image(str(png_file), wrap="inline")
    assert result is None


# --- mutators -------------------------------------------------------------------


def _text_box(doc):
    return doc.bookmarks["Address"].insert_text_box("Body")


def test_set_wrap(fake_word):
    from wordlive.constants import WdWrapType

    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_wrap("tight")
    assert fake_word.ActiveDocument.Shapes.Item(1).WrapFormat.Type == int(WdWrapType.TIGHT)


def test_set_position_sets_left_top_and_frame(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_position(left="1in", top="2in", relative_to="page")
    com = fake_word.ActiveDocument.Shapes.Item(1)
    assert com.Left == 72.0 and com.Top == 144.0
    assert com.RelativeHorizontalPosition == 1 and com.RelativeVerticalPosition == 1


def test_set_position_center(fake_word):
    from wordlive.constants import WdShapePosition

    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_position(left="center")
    assert fake_word.ActiveDocument.Shapes.Item(1).Left == float(WdShapePosition.CENTER)


def test_set_size_honours_both_dims(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_size(width="3in", height="1in", lock_aspect=False)
    com = fake_word.ActiveDocument.Shapes.Item(1)
    assert com.Width == 216.0 and com.Height == 72.0


def test_format_fill_and_border(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.format(fill="#eeeeff", border=False)
    com = fake_word.ActiveDocument.Shapes.Item(1)
    com.Fill.Solid.assert_called()
    assert com.Line.Visible == 0  # border=False hides the outline (MsoTriState.FALSE)


def test_set_alt_text(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_alt_text("A pull quote")
    assert fake_word.ActiveDocument.Shapes.Item(1).AlternativeText == "A pull quote"
    assert doc.shapes[1].alt_text == "A pull quote"


def test_set_text_replaces_contents(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_text("New body")
    assert fake_word.ActiveDocument.Shapes.Item(1).TextFrame.TextRange.Text == "New body"
    assert doc.shapes[1].text == "New body"


def test_set_text_on_picture_raises(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
        with pytest.raises(OpError):
            shape.set_text("nope")


def test_replace_image_needs_a_picture(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        with pytest.raises(OpError):
            shape.replace_image(_PNG)


def test_replace_image_swaps_in_place(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
        before = fake_word.ActiveDocument.Shapes.Count
        shape.replace_image(_PNG)
    # Delete + reinsert keeps a single shape (no orphan).
    assert fake_word.ActiveDocument.Shapes.Count == before == 1
    assert doc.shapes[1].shape_type == "picture"


def test_delete_removes_the_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        assert fake_word.ActiveDocument.Shapes.Count == 1
        shape.delete()
    assert fake_word.ActiveDocument.Shapes.Count == 0


# --- collections + resolution ---------------------------------------------------


def test_shapes_list_fields(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            doc.bookmarks["Address"].insert_text_box("Quote", width="2in", height="1in")
        rows = doc.shapes.list()
    assert len(rows) == 1
    row = rows[0]
    assert row["anchor_id"] == "shape:1"
    assert row["shape_type"] == "text_box"
    assert row["width"] == 144.0 and row["height"] == 72.0


def test_text_boxes_filter_keeps_unfiltered_id(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("mix"):
            # shape:1 = floating image, shape:2 = text box
            doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
            doc.bookmarks["Address"].insert_text_box("Quote")
        boxes = list(doc.text_boxes)
    assert len(boxes) == 1
    # The lone text box keeps its canonical shape:N id (its position among ALL shapes).
    assert boxes[0].anchor_id == "shape:2"
    assert doc.text_boxes[1].anchor_id == "shape:2"


def test_unknown_shape_index_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("shape:99")


def test_watermark_excluded_from_body_shapes(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("mix"):
            doc.set_watermark("DRAFT")  # header story
            doc.bookmarks["Address"].insert_text_box("Quote")  # body
        rows = doc.shapes.list()
    # Only the body text box; the header-story watermark is excluded.
    assert [r["shape_type"] for r in rows] == ["text_box"]


def test_body_name_prefixed_watermark_excluded(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            doc.bookmarks["Address"].insert_text_box("Quote")
        # A stray body shape named like Word's watermark is filtered by the guard.
        com = fake_word.ActiveDocument.Shapes
        com.AddTextbox(Anchor=fake_word.ActiveDocument.Range(0, 0))
        com.Item(2).Name = "PowerPlusWaterMarkObject9"
        rows = doc.shapes.list()
    assert len(rows) == 1


# --- exec ops -------------------------------------------------------------------


def test_exec_set_shape_wrap(fake_word):
    from wordlive.constants import WdWrapType

    script = json.dumps(
        [
            {"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "Q"},
            {"op": "set_shape_wrap", "anchor_id": "shape:1", "wrap": "behind"},
        ]
    )
    code, out = _invoke(["--json", "exec", "--ops", script])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 2
    assert fake_word.ActiveDocument.Shapes.Item(1).WrapFormat.Type == int(WdWrapType.BEHIND)


def test_exec_insert_text_box_returns_shape_output(fake_word):
    code, out = _invoke(
        [
            "--json",
            "exec",
            "--ops",
            '[{"op": "insert_text_box", "anchor_id": "start", "text": "Q"}]',
        ]
    )
    assert code == EXIT_OK
    outputs = json.loads(out)["outputs"]
    assert outputs[0]["anchor_id"] == "shape:1"


def test_exec_shape_op_on_non_shape_raises(fake_word):
    code, out = _invoke(
        [
            "--json",
            "exec",
            "--ops",
            '[{"op": "set_shape_wrap", "anchor_id": "para:1", "wrap": "tight"}]',
        ]
    )
    assert code != EXIT_OK
    assert "not a shape" in out


# --- CLI ------------------------------------------------------------------------


def test_cli_shapes_list(fake_word):
    _invoke(
        [
            "exec",
            "--ops",
            '[{"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "Q"}]',
        ]
    )
    code, out = _invoke(["--json", "shapes"])
    assert code == EXIT_OK
    rows = json.loads(out)
    assert rows[0]["anchor_id"] == "shape:1"


def test_cli_set_shape_size(fake_word):
    _invoke(
        [
            "exec",
            "--ops",
            '[{"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "Q"}]',
        ]
    )
    code, out = _invoke(["--json", "set-shape-size", "--anchor-id", "shape:1", "--width", "3in"])
    assert code == EXIT_OK
    assert fake_word.ActiveDocument.Shapes.Item(1).Width == 216.0


def test_cli_format_shape_requires_an_option(fake_word):
    _invoke(
        [
            "exec",
            "--ops",
            '[{"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "Q"}]',
        ]
    )
    code, _ = _invoke(["format-shape", "--anchor-id", "shape:1"])
    assert code != EXIT_OK  # no formatting option passed


def test_cli_set_shape_wrap_bad_anchor_exits_2(fake_word):
    code, _ = _invoke(["set-shape-wrap", "--anchor-id", "shape:99", "--wrap", "tight"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_cli_set_shape_wrap_on_non_shape_is_usage_error(fake_word):
    code, _ = _invoke(["set-shape-wrap", "--anchor-id", "para:1", "--wrap", "tight"])
    assert code != EXIT_OK  # not a shape


# --- MCP ------------------------------------------------------------------------


def test_mcp_read_shapes(fake_word):
    _write_impl(W, "text_box", {"anchor_id": "bookmark:Address", "text": "Q"})
    rows = _read_impl(W, "shapes", {})
    assert rows[0]["anchor_id"] == "shape:1"


def test_mcp_build_set_shape_position_op():
    op = _build_write_op(
        "set_shape_position", {"anchor_id": "shape:1", "left": "1in", "relative_to": "page"}
    )
    assert op == {
        "op": "set_shape_position",
        "anchor_id": "shape:1",
        "left": "1in",
        "relative_to": "page",
    }


def test_mcp_build_replace_shape_image_requires_one_source():
    with pytest.raises(OpError):
        _build_write_op("replace_shape_image", {"anchor_id": "shape:1"})
    op = _build_write_op("replace_shape_image", {"anchor_id": "shape:1", "image_base64": _B64})
    assert op == {"op": "replace_shape_image", "anchor_id": "shape:1", "base64": _B64}


def test_mcp_write_set_shape_wrap(fake_word):
    _write_impl(W, "text_box", {"anchor_id": "bookmark:Address", "text": "Q"})
    from wordlive.constants import WdWrapType

    _write_impl(W, "set_shape_wrap", {"anchor_id": "shape:1", "wrap": "behind"})
    assert fake_word.ActiveDocument.Shapes.Item(1).WrapFormat.Type == int(WdWrapType.BEHIND)


# --- shape depth: rotation / z-order / text-frame -------------------------------


def test_set_rotation(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_rotation("30")
    assert fake_word.ActiveDocument.Shapes.Item(1).Rotation == 30.0
    assert doc.shapes[1].rotation == 30.0


def test_set_rotation_bad_value_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        with pytest.raises(OpError):
            shape.set_rotation("sideways")


def test_set_z_order_brings_to_front(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("two"):
            doc.bookmarks["Address"].insert_text_box("A")  # shape:1 (backmost)
            doc.bookmarks["Address"].insert_text_box("B")  # shape:2 (frontmost)
        # shape:1 starts at the back; bring it to the front.
        assert doc.shapes[1].z_order == 1
        doc.shapes[1].set_z_order("front")
    # The shape formerly at shape:1 is now frontmost (highest ZOrderPosition).
    assert fake_word.ActiveDocument.Shapes.Item(2).ZOrderPosition == 2


def test_set_text_frame_margins_and_wrap(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        shape.set_text_frame(margin_left="0.1in", margin_top=5, word_wrap=False)
    frame = fake_word.ActiveDocument.Shapes.Item(1).TextFrame
    assert frame.MarginLeft == 7.2 and frame.MarginTop == 5.0
    assert frame.WordWrap == 0  # word_wrap=False -> MsoTriState.FALSE


def test_set_text_frame_on_picture_raises(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            shape = doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
        with pytest.raises(OpError):
            shape.set_text_frame(margin_left=5)


def test_shapes_list_includes_rotation_and_z_order(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            _text_box(doc)
        rows = doc.shapes.list()
    assert rows[0]["rotation"] == 0.0
    assert rows[0]["z_order"] == 1


# --- group / ungroup ------------------------------------------------------------


def _two_boxes(doc):
    doc.bookmarks["Address"].insert_text_box("A")
    doc.bookmarks["Address"].insert_text_box("B")


def test_group_collapses_to_one_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("grp"):
            _two_boxes(doc)
            group = doc.group_shapes("shape:1", "shape:2")
    assert group.shape_type == "group"
    # Two boxes became one group shape.
    assert len(doc.shapes) == 1
    assert doc.shapes[1].shape_type == "group"


def test_group_needs_two_shapes(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            _text_box(doc)
        with pytest.raises(OpError):
            doc.group_shapes("shape:1")


def test_group_rejects_non_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            _text_box(doc)
        with pytest.raises(OpError):
            doc.group_shapes("shape:1", "para:1")


def test_ungroup_restores_members(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("grp"):
            _two_boxes(doc)
            group = doc.group_shapes("shape:1", "shape:2")
        with doc.edit("ungrp"):
            members = group.ungroup()
    assert len(members) == 2
    assert {m.anchor_id for m in members} == {"shape:1", "shape:2"}
    assert len(doc.shapes) == 2


def test_ungroup_on_non_group_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tb"):
            shape = _text_box(doc)
        with pytest.raises(OpError):
            shape.ungroup()


# --- inline image:N restyle -----------------------------------------------------


def test_image_set_alt_text(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("alt"):
            doc.images[1].set_alt_text("New caption")
    assert fake_word.ActiveDocument.InlineShapes.Item(1).AlternativeText == "New caption"
    assert doc.anchor_by_id("image:1").alt_text == "New caption"


def test_image_set_size(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("size"):
            doc.images[1].set_size(width="3in")
    assert fake_word.ActiveDocument.InlineShapes.Item(1).Width == 216.0


def test_image_op_on_non_image_raises(fake_word):
    code, out = _invoke(
        [
            "--json",
            "exec",
            "--ops",
            '[{"op": "set_image_size", "anchor_id": "para:1", "width": "1in"}]',
        ]
    )
    assert code != EXIT_OK
    assert "not an image" in out


# --- textbox:N alias ------------------------------------------------------------


def test_textbox_alias_resolves_to_canonical_shape_id(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("mix"):
            doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")  # shape:1
            doc.bookmarks["Address"].insert_text_box("Q")  # shape:2
        anchor = doc.anchor_by_id("textbox:1")
    # The alias resolves to the text box but reports its canonical shape:N id.
    assert anchor.anchor_id == "shape:2"
    assert anchor.shape_type == "text_box"


def test_textbox_alias_unknown_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("textbox:9")


# --- new exec / CLI / MCP wiring ------------------------------------------------


def test_exec_set_shape_rotation(fake_word):
    script = json.dumps(
        [
            {"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "Q"},
            {"op": "set_shape_rotation", "anchor_id": "shape:1", "degrees": 45},
        ]
    )
    code, out = _invoke(["--json", "exec", "--ops", script])
    assert code == EXIT_OK
    assert fake_word.ActiveDocument.Shapes.Item(1).Rotation == 45.0


def test_exec_group_and_ungroup(fake_word):
    group_script = json.dumps(
        [
            {"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "A"},
            {"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "B"},
            {"op": "group_shapes", "shapes": ["shape:1", "shape:2"]},
        ]
    )
    code, out = _invoke(["--json", "exec", "--ops", group_script])
    assert code == EXIT_OK
    outputs = json.loads(out)["outputs"]
    assert outputs[-1]["anchor_id"] == "shape:1"  # the new group
    assert fake_word.ActiveDocument.Shapes.Count == 1

    code, out = _invoke(
        ["--json", "exec", "--ops", '[{"op": "ungroup_shape", "anchor_id": "shape:1"}]']
    )
    assert code == EXIT_OK
    assert json.loads(out)["outputs"][0]["count"] == 2


def test_cli_set_shape_z_order(fake_word):
    _invoke(
        [
            "exec",
            "--ops",
            '[{"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "Q"}]',
        ]
    )
    code, _ = _invoke(["--json", "set-shape-z-order", "--anchor-id", "shape:1", "--order", "back"])
    assert code == EXIT_OK


def test_cli_group_shapes_needs_two(fake_word):
    _invoke(
        [
            "exec",
            "--ops",
            '[{"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "Q"}]',
        ]
    )
    code, _ = _invoke(["group-shapes", "--anchor-id", "shape:1"])
    assert code != EXIT_OK  # one shape isn't a group


def test_mcp_build_set_shape_rotation_needs_degrees():
    with pytest.raises(OpError):
        _build_write_op("set_shape_rotation", {"anchor_id": "shape:1"})
    op = _build_write_op("set_shape_rotation", {"anchor_id": "shape:1", "degrees": 30})
    assert op == {"op": "set_shape_rotation", "anchor_id": "shape:1", "degrees": 30}


def test_mcp_build_group_shapes_requires_two():
    with pytest.raises(OpError):
        _build_write_op("group_shapes", {"shapes": ["shape:1"]})
    op = _build_write_op("group_shapes", {"shapes": ["shape:1", "shape:2"]})
    assert op == {"op": "group_shapes", "shapes": ["shape:1", "shape:2"]}


def test_mcp_build_set_image_size_op():
    op = _build_write_op("set_image_size", {"anchor_id": "image:1", "width": "2in"})
    assert op == {"op": "set_image_size", "anchor_id": "image:1", "width": "2in"}
