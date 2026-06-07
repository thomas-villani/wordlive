"""PageSetup writes + multi-column — `Section.set_page_setup`, its exec op, and CLI.

Round-trips against the `fake_word` MagicMock. `_FakePageSetup` carries
`PaperSize` / `Gutter` / `TextColumns` so the write surface has something to set;
`TextColumns.SetCount` is a MagicMock so the column count is assertable.
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive._ops import run_batch
from wordlive.cli.main import EXIT_OK, main
from wordlive.constants import WdOrientation, WdPaperSize
from wordlive.exceptions import OpError


def _ps(fake_word, section: int = 1):
    return fake_word.ActiveDocument.Sections(section).PageSetup


# --- the Section method --------------------------------------------------------


def test_set_orientation_landscape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("page setup"):
            doc.sections[1].set_page_setup(orientation="landscape")
    assert _ps(fake_word).Orientation == int(WdOrientation.LANDSCAPE)


def test_set_margins_aggregate_unit_string(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.sections[1].set_page_setup(margins="1in")
    ps = _ps(fake_word)
    assert ps.TopMargin == 72.0
    assert ps.BottomMargin == 72.0
    assert ps.LeftMargin == 72.0
    assert ps.RightMargin == 72.0


def test_per_side_margin_overrides_aggregate(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.sections[1].set_page_setup(margins="1in", top_margin="2in")
    ps = _ps(fake_word)
    assert ps.TopMargin == 144.0
    assert ps.LeftMargin == 72.0  # still the aggregate


def test_set_paper_size_a4(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.sections[1].set_page_setup(paper_size="a4")
    assert _ps(fake_word).PaperSize == int(WdPaperSize.A4)


def test_set_gutter(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.sections[1].set_page_setup(gutter="0.5in")
    assert _ps(fake_word).Gutter == 36.0


def test_set_columns(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.sections[1].set_page_setup(columns=2, column_spacing="0.25in")
    ps = _ps(fake_word)
    ps.TextColumns.SetCount.assert_called_once_with(2)
    assert ps.TextColumns.Spacing == 18.0


def test_bad_orientation_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.sections[1].set_page_setup(orientation="sideways")


def test_bad_paper_size_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.sections[1].set_page_setup(paper_size="a7")


def test_bad_columns_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.sections[1].set_page_setup(columns=0)


def test_bad_margin_unit_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.sections[1].set_page_setup(margins="1league")


def test_no_kwargs_is_noop(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.sections[1].set_page_setup()
    # Untouched: still the default 72pt margins from the fixture.
    assert _ps(fake_word).TopMargin == 72.0


# --- exec op -------------------------------------------------------------------


def test_exec_set_page_setup(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_page_setup", "section": 1, "orientation": "landscape", "columns": 3}],
            label="test",
        )
    assert exc is None and result["ok"] is True
    ps = _ps(fake_word)
    assert ps.Orientation == int(WdOrientation.LANDSCAPE)
    ps.TextColumns.SetCount.assert_called_once_with(3)


def test_exec_set_page_setup_missing_section_fails(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "set_page_setup"}], label="bad")
    assert exc is not None and result["ok"] is False
    assert "section" in result["failure"]["error"]


def test_exec_set_page_setup_bad_paper_size_fails_cleanly(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_page_setup", "section": 1, "paper_size": "a7"}],
            label="bad",
        )
    assert exc is not None and result["ok"] is False
    assert result["failure"]["type"] == "OpError"


# --- CLI -----------------------------------------------------------------------


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


def test_cli_page_setup(fake_word):
    code, out = _invoke(["--json", "page-setup", "--orientation", "landscape", "--columns", "2"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["applied"] == {"orientation": "landscape", "columns": 2}
    ps = _ps(fake_word)
    assert ps.Orientation == int(WdOrientation.LANDSCAPE)
    ps.TextColumns.SetCount.assert_called_once_with(2)


def test_cli_page_setup_requires_an_option(fake_word):
    code, _ = _invoke(["--json", "page-setup"])
    assert code != EXIT_OK
