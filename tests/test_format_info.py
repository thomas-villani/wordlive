"""The format read mirror — `Anchor.format_info` and its reverse helpers.

Round-trips against the `fake_word` MagicMock, whose ranges now carry real
`Font` / `ParagraphFormat` values and an applied `ParagraphStyle` baseline
(see `conftest._fake_font` / `_fake_paragraph_format`). The effective values and
the style baseline share the same defaults, so a clean range reads no overrides;
a test pokes the effective `Font` / `ParagraphFormat` to create a direct override
and asserts `format_info` flags it.
"""

from __future__ import annotations

import wordlive
from wordlive._anchors import _line_spacing_repr, _read_font, _read_paragraph_format
from wordlive._format import to_bgr
from wordlive.constants import WD_UNDEFINED, WdLineSpacing


def _addr_range(fake_word):
    return fake_word.ActiveDocument.Bookmarks("Address").Range


# --- pure reverse helpers ---------------------------------------------------


def test_line_spacing_repr_named_rules():
    assert _line_spacing_repr(int(WdLineSpacing.SINGLE), 12.0) == "single"
    assert _line_spacing_repr(int(WdLineSpacing.ONE_POINT_FIVE), 18.0) == "1.5"
    assert _line_spacing_repr(int(WdLineSpacing.DOUBLE), 24.0) == "double"


def test_line_spacing_repr_multiple_and_exact():
    # MULTIPLE stores points-per-line; 13.8 / 12 -> 1.15.
    assert _line_spacing_repr(int(WdLineSpacing.MULTIPLE), 13.8) == "1.15"
    assert _line_spacing_repr(int(WdLineSpacing.EXACTLY), 14.0) == "14pt"
    assert _line_spacing_repr(int(WdLineSpacing.AT_LEAST), 14.0) == "at_least:14pt"


def test_read_font_mixed_size_goes_to_mixed():
    from types import SimpleNamespace

    font = SimpleNamespace(
        Name="Aptos",
        Size=WD_UNDEFINED,
        Bold=0,
        Italic=0,
        Underline=0,
        StrikeThrough=0,
        Color=0,
        Subscript=0,
        Superscript=0,
        SmallCaps=0,
        AllCaps=0,
        Spacing=0.0,
    )
    values, mixed = _read_font(font)
    assert values["size"] is None
    assert "size" in mixed
    assert values["name"] == "Aptos"


def test_read_paragraph_format_defaults():
    from types import SimpleNamespace

    pf = SimpleNamespace(
        Alignment=3,
        LeftIndent=0.0,
        RightIndent=0.0,
        FirstLineIndent=0.0,
        SpaceBefore=0.0,
        SpaceAfter=8.0,
        LineSpacingRule=int(WdLineSpacing.SINGLE),
        LineSpacing=12.0,
        PageBreakBefore=0,
        KeepTogether=0,
        KeepWithNext=-1,
        WidowControl=-1,
    )
    out = _read_paragraph_format(pf)
    assert out["alignment"] == "justify"
    assert out["space_after"] == 8.0
    assert out["keep_with_next"] is True
    assert out["line_spacing"] == "single"


# --- the public method ------------------------------------------------------


def test_format_info_clean_reports_no_overrides(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        info = doc.bookmarks["Address"].format_info()
    assert info["anchor_id"] == "bookmark:Address"
    assert info["style"] == "Normal"
    assert info["font"]["size"]["override"] is False
    assert info["font"]["size"]["value"] == 12.0
    assert info["font"]["mixed"] == []
    assert info["paragraph"]["space_after"]["override"] is False


def test_format_info_size_override_detected(fake_word):
    _addr_range(fake_word).Font.Size = 15.0
    with wordlive.attach() as word:
        doc = word.documents.active
        info = doc.bookmarks["Address"].format_info()
    size = info["font"]["size"]
    assert size["value"] == 15.0
    assert size["style"] == 12.0
    assert size["override"] is True


def test_format_info_color_override_as_hex(fake_word):
    _addr_range(fake_word).Font.Color = to_bgr("red")  # 255
    with wordlive.attach() as word:
        doc = word.documents.active
        info = doc.bookmarks["Address"].format_info()
    color = info["font"]["color"]
    assert color["value"] == "#FF0000"
    assert color["override"] is True


def test_format_info_mixed_runs_listed(fake_word):
    _addr_range(fake_word).Font.Size = WD_UNDEFINED
    with wordlive.attach() as word:
        doc = word.documents.active
        info = doc.bookmarks["Address"].format_info()
    assert info["font"]["size"]["value"] is None
    # A mixed field is never reported as an override.
    assert info["font"]["size"]["override"] is False
    assert "size" in info["font"]["mixed"]


def test_format_info_paragraph_override_detected(fake_word):
    _addr_range(fake_word).ParagraphFormat.SpaceAfter = 20.0
    with wordlive.attach() as word:
        doc = word.documents.active
        info = doc.bookmarks["Address"].format_info()
    sp = info["paragraph"]["space_after"]
    assert sp["value"] == 20.0
    assert sp["style"] == 8.0
    assert sp["override"] is True
