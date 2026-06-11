"""Page / column / section breaks — insert_break + format_paragraph(page_break_before)."""

from __future__ import annotations

import pytest

import wordlive
from wordlive._ops import run_batch
from wordlive.constants import WdBreakType, WdLineSpacing

# ---------------------------------------------------------------------------
# insert_break — the anchor verb
# ---------------------------------------------------------------------------


def test_insert_break_page_after_calls_insertbreak(fake_word):
    # Bookmark "Address" spans (13, 24); where="after" inserts at its End.
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("break"):
            doc.bookmarks["Address"].insert_break("page")
    rng = fake_word.ActiveDocument.Range(24, 24)
    rng.InsertBreak.assert_called_once_with(Type=int(WdBreakType.PAGE))


def test_insert_break_before_inserts_at_start(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_break("column", where="before")
    rng = fake_word.ActiveDocument.Range(13, 13)
    rng.InsertBreak.assert_called_once_with(Type=int(WdBreakType.COLUMN))


def test_insert_break_default_kind_is_page(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_break()
    rng = fake_word.ActiveDocument.Range(24, 24)
    rng.InsertBreak.assert_called_once_with(Type=int(WdBreakType.PAGE))


@pytest.mark.parametrize(
    ("kind", "expected"),
    [
        ("page", WdBreakType.PAGE),
        ("column", WdBreakType.COLUMN),
        ("section_next", WdBreakType.SECTION_NEXT_PAGE),
        ("section_continuous", WdBreakType.SECTION_CONTINUOUS),
    ],
)
def test_insert_break_maps_each_kind(fake_word, kind, expected):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_break(kind)
    rng = fake_word.ActiveDocument.Range(24, 24)
    rng.InsertBreak.assert_called_once_with(Type=int(expected))


def test_insert_break_section_resets_break_paragraph_to_normal(fake_word):
    # A section break creates a new paragraph that inherits the anchor's style;
    # it must be reset to Normal so it doesn't pollute the heading outline.
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("break"):
            doc.bookmarks["Address"].insert_break("section_continuous")
    rng = fake_word.ActiveDocument.Range(24, 24)
    style = rng.Paragraphs(1).Range.Style
    assert style is not None and style.NameLocal == "Normal"


def test_insert_break_page_does_not_touch_paragraph_style(fake_word):
    # Page breaks are an in-paragraph character — no new paragraph, no reset.
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("break"):
            doc.bookmarks["Address"].insert_break("page")
    rng = fake_word.ActiveDocument.Range(24, 24)
    # Style was never assigned a real style object (still a bare auto-mock).
    assert not isinstance(getattr(rng.Paragraphs(1).Range.Style, "NameLocal", None), str)


def test_insert_break_unknown_kind_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError, match="unknown break kind"):
            doc.bookmarks["Address"].insert_break("paragraph")


def test_insert_break_bad_where_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError, match="where must be"):
            doc.bookmarks["Address"].insert_break("page", where="above")


# ---------------------------------------------------------------------------
# format_paragraph(page_break_before=...) — the clean, reflow-safe primitive
# ---------------------------------------------------------------------------


def test_format_paragraph_page_break_before_true(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_paragraph(page_break_before=True)
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert pf.PageBreakBefore is True


def test_format_paragraph_page_break_before_false_clears(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_paragraph(page_break_before=False)
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert pf.PageBreakBefore is False


# ---------------------------------------------------------------------------
# format_paragraph pagination controls — keep_together / keep_with_next /
# widow_control (tri-state: None leaves the property untouched)
# ---------------------------------------------------------------------------


def test_format_paragraph_pagination_flags(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_paragraph(
            keep_together=True, keep_with_next=True, widow_control=False
        )
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert pf.KeepTogether is True
    assert pf.KeepWithNext is True
    assert pf.WidowControl is False


def test_format_paragraph_pagination_flags_untouched_when_none(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_paragraph(keep_with_next=True)
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    # Only the explicitly-passed flag is written; the others keep their auto-mock
    # default (never assigned True/False).
    assert pf.KeepWithNext is True
    assert pf.KeepTogether is not True and pf.KeepTogether is not False
    assert pf.WidowControl is not True and pf.WidowControl is not False


def test_exec_format_paragraph_pagination_flags(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "format_paragraph",
                    "anchor_id": "bookmark:Address",
                    "keep_together": True,
                    "keep_with_next": True,
                    "widow_control": False,
                }
            ],
            label="test",
        )
    assert exc is None
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert pf.KeepTogether is True
    assert pf.KeepWithNext is True
    assert pf.WidowControl is False


# ---------------------------------------------------------------------------
# format_paragraph(line_spacing=...) — leading within the paragraph
# ---------------------------------------------------------------------------


def test_coerce_line_spacing_variants():
    from wordlive._anchors import _coerce_line_spacing

    # A number is a multiple of single spacing (rule MULTIPLE, value × 12pt).
    assert _coerce_line_spacing(1.5) == (int(WdLineSpacing.MULTIPLE), 18.0)
    assert _coerce_line_spacing(2) == (int(WdLineSpacing.MULTIPLE), 24.0)
    # Named multiples carry no companion value.
    assert _coerce_line_spacing("single") == (int(WdLineSpacing.SINGLE), None)
    assert _coerce_line_spacing("1.5") == (int(WdLineSpacing.ONE_POINT_FIVE), None)
    assert _coerce_line_spacing("double") == (int(WdLineSpacing.DOUBLE), None)
    # A length string with a unit is an exact line height.
    assert _coerce_line_spacing("14pt") == (int(WdLineSpacing.EXACTLY), 14.0)
    # A unitless numeric string is a multiple, like a bare number.
    assert _coerce_line_spacing("1.15") == (int(WdLineSpacing.MULTIPLE), pytest.approx(13.8))
    with pytest.raises(ValueError, match="line_spacing"):
        _coerce_line_spacing("loose")
    with pytest.raises(TypeError):
        _coerce_line_spacing(True)


def test_format_paragraph_line_spacing_multiple(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_paragraph(line_spacing=2)
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert int(pf.LineSpacingRule) == int(WdLineSpacing.MULTIPLE)
    assert pf.LineSpacing == 24.0


def test_format_paragraph_line_spacing_exact_points(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_paragraph(line_spacing="14pt")
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert int(pf.LineSpacingRule) == int(WdLineSpacing.EXACTLY)
    assert pf.LineSpacing == 14.0


def test_exec_format_paragraph_line_spacing(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "format_paragraph", "anchor_id": "bookmark:Address", "line_spacing": "double"}],
            label="test",
        )
    assert exc is None and "warnings" not in result
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert int(pf.LineSpacingRule) == int(WdLineSpacing.DOUBLE)


# ---------------------------------------------------------------------------
# exec ops
# ---------------------------------------------------------------------------


def test_exec_insert_break(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_break", "anchor_id": "bookmark:Address", "kind": "section_next"}],
            label="test",
        )
    assert exc is None
    assert result["ok"] is True
    rng = fake_word.ActiveDocument.Range(24, 24)
    rng.InsertBreak.assert_called_once_with(Type=int(WdBreakType.SECTION_NEXT_PAGE))


def test_exec_insert_break_defaults_to_page(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "insert_break", "anchor_id": "bookmark:Address"}], label="test"
        )
    assert exc is None
    rng = fake_word.ActiveDocument.Range(24, 24)
    rng.InsertBreak.assert_called_once_with(Type=int(WdBreakType.PAGE))


def test_exec_insert_break_before(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        run_batch(
            doc,
            [{"op": "insert_break", "anchor_id": "bookmark:Address", "before": True}],
            label="test",
        )
    rng = fake_word.ActiveDocument.Range(13, 13)
    rng.InsertBreak.assert_called_once_with(Type=int(WdBreakType.PAGE))


def test_exec_insert_break_missing_anchor_fails_cleanly(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "insert_break", "kind": "page"}], label="bad")
    assert exc is not None
    assert result["ok"] is False
    assert "anchor_id" in result["failure"]["error"]


def test_exec_format_paragraph_page_break_before(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "format_paragraph",
                    "anchor_id": "bookmark:Address",
                    "page_break_before": True,
                }
            ],
            label="test",
        )
    assert exc is None
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert pf.PageBreakBefore is True
