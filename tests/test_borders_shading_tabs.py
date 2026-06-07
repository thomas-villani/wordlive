"""Borders, shading, and tab stops — the `set_shading` / `set_borders` /
`add_tab_stop` anchor methods and their exec ops.

Round-trips against the `fake_word` MagicMock. Shading and TabStops auto-create;
`Range.Borders` uses the memoising `_FakeBorders` fixture so per-edge asserts
stay distinct.
"""

from __future__ import annotations

from unittest.mock import call

import pytest

import wordlive
from wordlive._format import to_bgr
from wordlive._ops import run_batch
from wordlive.constants import WdLineStyle, WdTabAlignment, WdTabLeader
from wordlive.exceptions import OpError


def _addr_range(fake_word):
    return fake_word.ActiveDocument.Bookmarks("Address").Range


# --- shading -------------------------------------------------------------------


def test_set_shading_fill_hex(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("shade"):
            doc.bookmarks["Address"].set_shading(fill="FFFF00")
    # Yellow #FFFF00 -> BGR long 65535.
    assert _addr_range(fake_word).Shading.BackgroundPatternColor == 65535


def test_set_shading_fill_named(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].set_shading(fill="red")
    assert _addr_range(fake_word).Shading.BackgroundPatternColor == 255


def test_set_shading_bad_color_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].set_shading(fill="ultraviolet")


def test_set_shading_no_fill_is_noop(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].set_shading()


# --- borders -------------------------------------------------------------------


def test_set_borders_all_sides(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("border"):
            doc.bookmarks["Address"].set_borders(style="double", weight=1.0, color="000000")
    borders = _addr_range(fake_word).Borders
    # all -> the four outer edges: top(-1) bottom(-3) left(-2) right(-4).
    for edge in (-1, -2, -3, -4):
        b = borders(edge)
        assert b.LineStyle == int(WdLineStyle.DOUBLE)
        assert b.Color == 0
    # 1.0pt snaps to WdLineWidth 8 (points x 8).
    assert borders(-1).LineWidth == 8


def test_set_borders_single_side(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].set_borders(sides="bottom", style="single")
    borders = _addr_range(fake_word).Borders
    assert borders(-3).LineStyle == int(WdLineStyle.SINGLE)
    # The top edge was never touched — its LineStyle stays a bare mock, not int.
    assert not isinstance(borders(-1).LineStyle, int)


def test_set_borders_list_of_sides(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].set_borders(sides=["top", "bottom"], style="dot")
    borders = _addr_range(fake_word).Borders
    assert borders(-1).LineStyle == int(WdLineStyle.DOT)
    assert borders(-3).LineStyle == int(WdLineStyle.DOT)


def test_set_borders_weight_snaps(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].set_borders(sides="top", weight=0.25)
    # 0.25pt -> WdLineWidth 2.
    assert _addr_range(fake_word).Borders(-1).LineWidth == 2


def test_set_borders_unknown_side_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].set_borders(sides="diagonal")


# --- tab stops -----------------------------------------------------------------


def test_add_tab_stop_positional_args(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tab"):
            doc.bookmarks["Address"].add_tab_stop("1in", align="right", leader="dots")
    tabstops = _addr_range(fake_word).ParagraphFormat.TabStops
    # 1in -> 72.0pt; positional Position, Alignment, Leader.
    assert tabstops.Add.call_args == call(72.0, int(WdTabAlignment.RIGHT), int(WdTabLeader.DOTS))


def test_add_tab_stop_default_leader(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].add_tab_stop(36)
    tabstops = _addr_range(fake_word).ParagraphFormat.TabStops
    assert tabstops.Add.call_args == call(36.0, int(WdTabAlignment.LEFT), int(WdTabLeader.SPACES))


def test_add_tab_stop_bad_align_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].add_tab_stop(36, align="sideways")


# --- cell shading (a Cell is an Anchor) ----------------------------------------


def test_set_shading_on_table_cell(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("cell shade"):
            doc.tables[1].cell(1, 1).set_shading(fill="green")
    # The cell's backing range is persistent in the fake, so the shading lands.
    rng = fake_word.ActiveDocument.Tables(1).Cell(1, 1).Range
    assert rng.Shading.BackgroundPatternColor == to_bgr("green")


# --- exec ops ------------------------------------------------------------------


def test_exec_set_shading(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_shading", "anchor_id": "bookmark:Address", "fill": "FFFF00"}],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert _addr_range(fake_word).Shading.BackgroundPatternColor == 65535


def test_exec_set_borders(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "set_borders",
                    "anchor_id": "bookmark:Address",
                    "sides": "top",
                    "style": "dash",
                }
            ],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert _addr_range(fake_word).Borders(-1).LineStyle == int(WdLineStyle.DASH_LARGE_GAP)


def test_exec_add_tab_stop(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "add_tab_stop", "anchor_id": "bookmark:Address", "position": "2in"}],
            label="test",
        )
    assert exc is None and result["ok"] is True
    tabstops = _addr_range(fake_word).ParagraphFormat.TabStops
    assert tabstops.Add.call_args == call(144.0, int(WdTabAlignment.LEFT), int(WdTabLeader.SPACES))


def test_exec_add_tab_stop_missing_position_fails_cleanly(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "add_tab_stop", "anchor_id": "bookmark:Address"}], label="bad"
        )
    assert exc is not None and result["ok"] is False
    assert "position" in result["failure"]["error"]
