"""Back-of-book index — `Anchor.mark_index_entry` / `insert_index`,
`Document.add_index`, the `mark_index_entry` / `insert_index` ops, and the CLI.

`doc.Indexes.MarkEntry` and `doc.Indexes.Add` are recording MagicMocks (see
`_FakeIndexes`), so the positional arguments wordlive passes are assertable
without real Word.
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive._ops import run_batch
from wordlive.cli.main import EXIT_OK, main
from wordlive.constants import WdIndexType
from wordlive.exceptions import OpError


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


def _indexes(fake_word):
    return fake_word.ActiveDocument.Indexes


# Positional layout of Indexes.Add (see Anchor.insert_index):
# 0 Range · 1 HeadingSeparator · 2 RightAlignPageNumbers · 3 Type · 4 NumberOfColumns
# Positional layout of Indexes.MarkEntry:
# 0 Range · 1 Entry · 2 EntryAutoText · 3 CrossReference · ... · 6 Bold · 7 Italic


# --- mark_index_entry ----------------------------------------------------------


def test_mark_index_entry_passes_entry_and_flags(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("mark"):
            doc.range(0, 5).mark_index_entry("widgets:gadgets", bold=True)
    args = _indexes(fake_word).MarkEntry.call_args.args
    assert args[1] == "widgets:gadgets"
    assert args[6] is True  # bold
    assert args[7] is False  # italic


def test_mark_index_entry_cross_reference(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("mark"):
            doc.range(0, 5).mark_index_entry("foo", cross_reference="bar")
    assert _indexes(fake_word).MarkEntry.call_args.args[3] == "bar"


def test_mark_index_entry_empty_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.range(0, 5).mark_index_entry("   ")


# --- insert_index --------------------------------------------------------------


def test_insert_index_passes_columns_and_type(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("index"):
            idx = doc.end.insert_index(columns=3, run_in=True, right_align_page_numbers=True)
        assert isinstance(idx, wordlive.Index)
    args = _indexes(fake_word).Add.call_args.args
    assert args[1] == 0  # HeadingSeparator (none) — must not be ""
    assert args[2] is True  # right_align
    assert args[3] == int(WdIndexType.RUNIN)
    assert args[4] == 3  # columns


def test_add_index_defaults_indent(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("index"):
            doc.add_index()
    args = _indexes(fake_word).Add.call_args.args
    assert args[3] == int(WdIndexType.INDENT)
    assert args[4] == 2  # default columns


def test_insert_index_bad_columns_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.end.insert_index(columns=0)


def test_index_update_calls_com(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        idx = doc.add_index()
        idx.update()
    assert idx.com.Update.call_count == 1


# --- exec ops ------------------------------------------------------------------


def test_exec_mark_then_insert_index(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {"op": "mark_index_entry", "anchor_id": "range:0-5", "entry": "foo"},
                {"op": "insert_index", "anchor_id": "end", "columns": 1},
            ],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][-1]["index"] is True
    assert _indexes(fake_word).Add.call_args.args[4] == 1


# --- CLI -----------------------------------------------------------------------


def test_cli_mark_index_entry(fake_word):
    code, out = _invoke(
        ["--json", "mark-index-entry", "--anchor-id", "range:0-5", "--entry", "topic", "--italic"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["applied"]["entry"] == "topic"
    assert _indexes(fake_word).MarkEntry.call_args.args[7] is True  # italic


def test_cli_insert_index(fake_word):
    code, out = _invoke(["--json", "insert-index", "--columns", "3", "--run-in"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["applied"]["columns"] == 3
    assert data["applied"]["run_in"] is True
    assert _indexes(fake_word).Add.call_args.args[3] == int(WdIndexType.RUNIN)
