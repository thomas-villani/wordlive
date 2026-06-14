"""Table of figures — `Anchor.insert_table_of_figures`, the
`insert_table_of_figures` op, and the CLI.

`doc.TablesOfFigures.Add` is a recording MagicMock (see `_FakeTablesOfFigures`).
wordlive calls it with `Range` + `Caption` positional and the flags as keyword
args (the optional string Variants reject positional ""), so the assertions read
the keyword args.
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


def _tof_add(fake_word):
    return fake_word.ActiveDocument.TablesOfFigures.Add


# --- the Anchor method ---------------------------------------------------------


def test_insert_table_of_figures_passes_label_and_flags(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tof"):
            tof = doc.start.insert_table_of_figures(label="Table", hyperlinks=False)
        assert isinstance(tof, wordlive.TableOfFigures)
    call = _tof_add(fake_word).call_args
    assert call.args[1] == "Table"  # Caption is positional
    assert call.kwargs["UseHyperlinks"] is False
    assert call.kwargs["IncludeLabel"] is True


def test_insert_table_of_figures_no_label(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tof"):
            doc.start.insert_table_of_figures(include_label=False)
    assert _tof_add(fake_word).call_args.kwargs["IncludeLabel"] is False


def test_insert_table_of_figures_bad_where_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.start.insert_table_of_figures(where="sideways")


def test_table_of_figures_update_calls_com(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("tof"):
            tof = doc.start.insert_table_of_figures()
        tof.update()
        tof.update_page_numbers()
    assert tof.com.Update.call_count == 1
    assert tof.com.UpdatePageNumbers.call_count == 1


# --- exec op -------------------------------------------------------------------


def test_exec_insert_table_of_figures(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_table_of_figures", "anchor_id": "start", "label": "Equation"}],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["table_of_figures"] is True
    assert _tof_add(fake_word).call_args.args[1] == "Equation"


# --- CLI -----------------------------------------------------------------------


def test_cli_table_of_figures(fake_word):
    code, out = _invoke(["--json", "table-of-figures", "--label", "Table", "--no-hyperlinks"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["applied"]["label"] == "Table"
    assert data["applied"]["hyperlinks"] is False
    assert _tof_add(fake_word).call_args.kwargs["UseHyperlinks"] is False
