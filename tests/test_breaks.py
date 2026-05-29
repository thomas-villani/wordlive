"""Page / column / section breaks — insert_break + format_paragraph(page_break_before)."""

from __future__ import annotations

import pytest

import wordlive
from wordlive._ops import run_batch
from wordlive.constants import WdBreakType

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
