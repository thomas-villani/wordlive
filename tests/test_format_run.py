"""Character formatting — `Anchor.format_run` and the `format_run` exec op.

Round-trips against the `fake_word` MagicMock: call the method, then read the
fake COM `Range.Font.*` / `Range.HighlightColorIndex` back. MagicMock
auto-creates `.Font`, so no fixture change is needed.
"""

from __future__ import annotations

import pytest

import wordlive
from wordlive._format import to_bgr
from wordlive._ops import run_batch
from wordlive.constants import WdColorIndex, WdUnderline
from wordlive.exceptions import OpError


def _addr_font(fake_word):
    return fake_word.ActiveDocument.Bookmarks("Address").Range.Font


def _addr_range(fake_word):
    return fake_word.ActiveDocument.Bookmarks("Address").Range


def test_format_run_bold_italic(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("bold"):
            doc.bookmarks["Address"].format_run(bold=True, italic=True)
    font = _addr_font(fake_word)
    assert font.Bold is True
    assert font.Italic is True


def test_format_run_size_unit_string(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(size="12pt")
    assert _addr_font(fake_word).Size == 12.0


def test_format_run_size_bare_number_is_points(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(size=14)
    assert _addr_font(fake_word).Size == 14.0


def test_format_run_color_hex_is_bgr(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(color="FF0000")
    # Red #FF0000 -> BGR long 255.
    assert _addr_font(fake_word).Color == 255


def test_format_run_color_rgb_tuple(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(color=(0, 0, 255))
    # Blue -> BGR long 16711680.
    assert _addr_font(fake_word).Color == 16711680


def test_format_run_underline_true_is_single(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(underline=True)
    assert _addr_font(fake_word).Underline == int(WdUnderline.SINGLE)


def test_format_run_underline_false_is_none(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(underline=False)
    assert _addr_font(fake_word).Underline == int(WdUnderline.NONE)


def test_format_run_highlight_named(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(highlight="yellow")
    assert _addr_range(fake_word).HighlightColorIndex == int(WdColorIndex.YELLOW)


def test_format_run_spacing_unit_string(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(spacing="1pt")
    assert _addr_font(fake_word).Spacing == 1.0


def test_format_run_caps_and_scripts(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(
            strikethrough=True, small_caps=True, all_caps=False, superscript=True
        )
    font = _addr_font(fake_word)
    assert font.StrikeThrough is True
    assert font.SmallCaps is True
    assert font.AllCaps is False
    assert font.Superscript is True


def test_format_run_font_name(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run(font="Calibri")
    assert _addr_font(fake_word).Name == "Calibri"


def test_format_run_no_kwargs_is_noop(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_run()


def test_format_run_bad_color_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].format_run(color="not-a-colour")


# --- cross-kind (lives on the base Anchor) -------------------------------------


def test_format_run_on_content_control(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.content_controls["Signatory"].format_run(bold=True, color="navy")
    cc = next(c for c in fake_word.ActiveDocument.ContentControls if c.Title == "Signatory")
    assert cc.Range.Font.Bold is True
    assert cc.Range.Font.Color == to_bgr("navy")


def test_format_run_on_heading(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.heading("Risks").format_run(italic=True)
    para = next(
        p for p in fake_word.ActiveDocument.Paragraphs if p.Range.Text.rstrip("\r\n\x07") == "Risks"
    )
    assert para.Range.Font.Italic is True


# --- exec op -------------------------------------------------------------------


def test_exec_format_run(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "format_run", "anchor_id": "bookmark:Address", "bold": True, "size": "10pt"}],
            label="test",
        )
    assert exc is None
    assert result["ok"] is True
    font = _addr_font(fake_word)
    assert font.Bold is True
    assert font.Size == 10.0


def test_exec_format_run_bad_color_fails_cleanly(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "format_run", "anchor_id": "bookmark:Address", "color": "chartreusey"}],
            label="bad",
        )
    assert exc is not None
    assert result["ok"] is False
    assert result["failure"]["type"] == "OpError"


def test_exec_format_run_missing_anchor_fails_cleanly(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "format_run", "bold": True}], label="bad")
    assert exc is not None
    assert result["ok"] is False
    assert "anchor_id" in result["failure"]["error"]
