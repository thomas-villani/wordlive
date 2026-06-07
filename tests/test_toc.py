"""Table of contents — `Anchor.insert_toc`, `Document.add_toc`, `Toc.update`,
and the `insert_toc` op.

`doc.TablesOfContents.Add` is a recording MagicMock (see `_FakeTOCs`), so the
positional arguments wordlive passes — `UpperHeadingLevel`, `LowerHeadingLevel`,
and `UseHyperlinks` in particular — are assertable without a real Word.
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive._ops import run_batch
from wordlive.cli.main import EXIT_OK, main
from wordlive.exceptions import OpError


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


def _toc_add(fake_word):
    return fake_word.ActiveDocument.TablesOfContents.Add


# Positional layout of TablesOfContents.Add (see Anchor.insert_toc):
# 0 Range · 1 UseHeadingStyles · 2 Upper · 3 Lower · 4 UseFields · 5 TableID ·
# 6 RightAlign · 7 IncludePageNumbers · 8 AddedStyles · 9 UseHyperlinks ·
# 10 HidePageNumbersInWeb · 11 UseOutlineLevels


# --- the Anchor method ---------------------------------------------------------


def test_insert_toc_passes_levels_and_hyperlinks(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("toc"):
            toc = doc.start.insert_toc(levels=(1, 2), hyperlinks=False)
        assert isinstance(toc, wordlive.Toc)
    args = _toc_add(fake_word).call_args.args
    assert args[1] is True  # use_heading_styles default
    assert args[2] == 1  # upper
    assert args[3] == 2  # lower
    assert args[9] is False  # hyperlinks


def test_add_toc_defaults(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("toc"):
            doc.add_toc()
    args = _toc_add(fake_word).call_args.args
    assert args[2] == 1 and args[3] == 3  # default levels (1, 3)
    assert args[9] is True  # hyperlinks default


def test_insert_toc_bad_levels_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.start.insert_toc(levels=(3, 1))  # upper > lower


def test_insert_toc_bad_where_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.start.insert_toc(where="sideways")


def test_toc_update_calls_com(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        toc = doc.add_toc()
        toc.update()
        toc.update_page_numbers()
    assert toc.com.Update.call_count == 1
    assert toc.com.UpdatePageNumbers.call_count == 1


# --- exec op -------------------------------------------------------------------


def test_exec_insert_toc(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_toc", "anchor_id": "start", "levels": [1, 4]}],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["toc"] is True
    args = _toc_add(fake_word).call_args.args
    assert args[2] == 1 and args[3] == 4


def test_exec_insert_toc_bad_levels_fails(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_toc", "anchor_id": "start", "levels": [5, 2]}],
            label="bad",
        )
    assert exc is not None and result["ok"] is False


# --- CLI -----------------------------------------------------------------------


def test_cli_insert_toc(fake_word):
    code, out = _invoke(["--json", "insert-toc", "--levels", "1-2", "--no-hyperlinks"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["applied"]["levels"] == [1, 2]
    assert data["applied"]["hyperlinks"] is False
    args = _toc_add(fake_word).call_args.args
    assert args[2] == 1 and args[3] == 2 and args[9] is False


def test_cli_insert_toc_bad_levels(fake_word):
    code, _ = _invoke(["--json", "insert-toc", "--levels", "abc"])
    assert code != EXIT_OK
