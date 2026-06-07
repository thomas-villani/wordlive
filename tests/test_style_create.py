"""Style creation & modification — `doc.styles.add` and the writable `Style`
setters, plus the `add_style` / `set_style` exec ops.

Round-trips against the `fake_word` MagicMock; `_FakeStyles.Add` appends a style
the direct-then-iterate `Style.com` lookup can then resolve by name.
"""

from __future__ import annotations

import pytest

import wordlive
from wordlive._format import to_bgr
from wordlive._ops import run_batch
from wordlive.constants import WdParagraphAlignment, WdStyleType
from wordlive.exceptions import OpError, StyleNotFoundError


def _style_mock(fake_word, name):
    return fake_word.ActiveDocument.Styles(name)


def test_styles_add_records_add_call(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("add"):
            new = doc.styles.add("Brand")
    assert new.name == "Brand"
    # Default type is paragraph (1).
    assert _style_mock(fake_word, "Brand").Type == int(WdStyleType.PARAGRAPH)


def test_styles_add_character_type(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.styles.add("Inline Code", type="character")
    assert _style_mock(fake_word, "Inline Code").Type == int(WdStyleType.CHARACTER)


def test_styles_add_unknown_type_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.styles.add("Whatever", type="bogus")


def test_styles_add_unknown_based_on_raises_style_not_found(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(StyleNotFoundError):
            doc.styles.add("Brand", based_on="NoSuchStyle")


def test_styles_add_then_format_run(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("add+style"):
            new = doc.styles.add("Brand")
            new.format_run(bold=True, color="navy")
    sm = _style_mock(fake_word, "Brand")
    assert sm.Font.Bold is True
    assert sm.Font.Color == to_bgr("navy")


def test_styles_add_then_format_paragraph(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        new = doc.styles.add("Brand")
        new.format_paragraph(space_before=12, alignment="center")
    sm = _style_mock(fake_word, "Brand")
    assert sm.ParagraphFormat.SpaceBefore == 12.0
    assert sm.ParagraphFormat.Alignment == int(WdParagraphAlignment.CENTER)


def test_style_format_run_rejects_highlight(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        new = doc.styles.add("Brand")
        with pytest.raises(OpError):
            new.format_run(highlight="yellow")


def test_style_base_style_setter(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        new = doc.styles.add("Brand")
        new.base_style = "Normal"
    # BaseStyle was assigned the Normal style mock.
    assert _style_mock(fake_word, "Brand").BaseStyle.NameLocal == "Normal"


def test_style_base_style_setter_missing_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        new = doc.styles.add("Brand")
        with pytest.raises(StyleNotFoundError):
            new.base_style = "NoSuchStyle"


def test_style_next_paragraph_style_setter(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        new = doc.styles.add("Brand")
        new.next_paragraph_style = "Body Text"
    assert _style_mock(fake_word, "Brand").NextParagraphStyle.NameLocal == "Body Text"


# --- exec ops ------------------------------------------------------------------


def test_exec_add_style_returns_output(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "add_style", "name": "Brand", "based_on": "Normal"}],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"] == [{"index": 0, "op": "add_style", "style": "Brand"}]


def test_exec_set_style(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        # add then set in one batch (atomic undo).
        result, exc = run_batch(
            doc,
            [
                {"op": "add_style", "name": "Brand"},
                {
                    "op": "set_style",
                    "name": "Brand",
                    "bold": True,
                    "size": "14pt",
                    "alignment": "center",
                    "next_style": "Body Text",
                },
            ],
            label="test",
        )
    assert exc is None and result["ok"] is True
    sm = _style_mock(fake_word, "Brand")
    assert sm.Font.Bold is True
    assert sm.Font.Size == 14.0
    assert sm.ParagraphFormat.Alignment == int(WdParagraphAlignment.CENTER)
    assert sm.NextParagraphStyle.NameLocal == "Body Text"


def test_exec_set_style_unknown_style_fails_cleanly(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "set_style", "name": "NoSuchStyle", "bold": True}], label="bad"
        )
    assert exc is not None and result["ok"] is False
    assert result["failure"]["type"] == "StyleNotFoundError"
