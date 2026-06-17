"""Publishing flourishes — `doc.set_watermark` / `remove_watermark` and
`anchor.insert_text_box`.

Floating shapes are COM-heavy and version-sensitive; these tests pin the wiring
against the `fake_word` MagicMock (the right shapes get added to the right story,
named the watermark convention, and removed again). Visual correctness is a
smoke / live-Word concern.
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive.cli.main import EXIT_OK, main
from wordlive.exceptions import OpError


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


def _primary_header_shapes(fake_word):
    return fake_word.ActiveDocument.Sections(1).Headers(1).Shapes


# --- watermark ------------------------------------------------------------------


def test_set_watermark_adds_wordart_to_header(fake_word):
    with wordlive.attach() as word:
        n = word.documents.active.set_watermark("DRAFT")
    assert n == 1
    shapes = _primary_header_shapes(fake_word)
    assert shapes.Count == 1
    shape = shapes(1)
    assert shape.Text == "DRAFT"
    assert shape.Name.startswith("PowerPlusWaterMarkObject")


def test_set_watermark_replaces_prior(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.set_watermark("DRAFT")
        doc.set_watermark("FINAL")
    shapes = _primary_header_shapes(fake_word)
    # Calling twice doesn't stack — the first watermark was cleared.
    assert shapes.Count == 1
    assert shapes(1).Text == "FINAL"


def test_remove_watermark_clears_and_counts(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.set_watermark("CONFIDENTIAL")
        removed = doc.remove_watermark()
    assert removed == 1
    assert _primary_header_shapes(fake_word).Count == 0


def test_remove_watermark_none_present_is_zero(fake_word):
    with wordlive.attach() as word:
        assert word.documents.active.remove_watermark() == 0


def test_set_watermark_bad_layout_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(OpError):
            word.documents.active.set_watermark("X", layout="sideways")


def test_cli_watermark_set_and_remove(fake_word):
    code, out = _invoke(["--json", "watermark", "--text", "DRAFT"])
    assert code == EXIT_OK
    assert json.loads(out)["sections"] == 1

    code, out = _invoke(["--json", "watermark", "--remove"])
    assert code == EXIT_OK
    assert json.loads(out)["removed"] == 1


def test_cli_watermark_requires_text_or_remove(fake_word):
    code, out = _invoke(["watermark"])
    assert code != EXIT_OK  # neither --text nor --remove


def test_exec_set_watermark(fake_word):
    code, out = _invoke(["--json", "exec", "--ops", '[{"op": "set_watermark", "text": "DRAFT"}]'])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    assert _primary_header_shapes(fake_word).Count == 1


# --- text box -------------------------------------------------------------------


def test_insert_text_box_adds_floating_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_text_box("Pull quote", width="3in")
    shapes = fake_word.ActiveDocument.Shapes
    assert shapes.Count == 1
    # The text landed in the box's text frame.
    assert shapes(1).TextFrame.TextRange.Text == "Pull quote"


def test_insert_text_box_unknown_wrap_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(ValueError):
            word.documents.active.bookmarks["Address"].insert_text_box("x", wrap="inline")


def test_cli_insert_text_box(fake_word):
    code, out = _invoke(
        [
            "--json",
            "insert-text-box",
            "--anchor-id",
            "bookmark:Address",
            "--text",
            "Sidebar",
            "--fill",
            "#eeeeff",
            "--no-border",
        ]
    )
    assert code == EXIT_OK
    payload = json.loads(out)
    assert payload["ok"] is True
    assert fake_word.ActiveDocument.Shapes.Count == 1


def test_cli_insert_text_box_border_conflict(fake_word):
    code, _ = _invoke(
        [
            "insert-text-box",
            "--anchor-id",
            "bookmark:Address",
            "--text",
            "x",
            "--no-border",
            "--border-color",
            "red",
        ]
    )
    assert code != EXIT_OK


def test_exec_insert_text_box(fake_word):
    code, out = _invoke(
        [
            "--json",
            "exec",
            "--ops",
            '[{"op": "insert_text_box", "anchor_id": "bookmark:Address", "text": "Quote"}]',
        ]
    )
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    assert fake_word.ActiveDocument.Shapes.Count == 1
