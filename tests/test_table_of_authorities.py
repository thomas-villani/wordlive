"""Table of authorities — `Anchor.mark_citation` / `insert_table_of_authorities`,
`Document.add_table_of_authorities`, the matching ops, and the CLI.

`mark_citation` inserts a `TA` field (assert via `<anchor>.Range.Duplicate.Fields.Add`,
as in `test_fields.py`); `insert_table_of_authorities` calls
`doc.TablesOfAuthorities.Add` (a recording `_FakeTablesOfAuthorities`).
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


def _field_add(fake_word):
    return fake_word.ActiveDocument.Bookmarks("Address").Range.Duplicate.Fields.Add


def _toa_add(fake_word):
    return fake_word.ActiveDocument.TablesOfAuthorities.Add


# --- mark_citation -------------------------------------------------------------


def test_mark_citation_default_short_and_category(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("mark"):
            doc.bookmarks["Address"].mark_citation("Smith v. Jones, 1 U.S. 1")
    code = _field_add(fake_word).call_args.args[2]
    # short defaults to long; category "cases" -> 1.
    assert code == 'TA \\l "Smith v. Jones, 1 U.S. 1" \\s "Smith v. Jones, 1 U.S. 1" \\c 1'


def test_mark_citation_short_and_named_category(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("mark"):
            doc.bookmarks["Address"].mark_citation(
                "Securities Act of 1933", short_citation="1933 Act", category="statutes"
            )
    code = _field_add(fake_word).call_args.args[2]
    assert code == 'TA \\l "Securities Act of 1933" \\s "1933 Act" \\c 2'


def test_mark_citation_numeric_category(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("mark"):
            doc.bookmarks["Address"].mark_citation("X", category=5)
    assert _field_add(fake_word).call_args.args[2].endswith("\\c 5")


def test_mark_citation_empty_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].mark_citation("   ")


def test_mark_citation_bad_category_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].mark_citation("X", category="bogus")


# --- insert_table_of_authorities -----------------------------------------------


def test_insert_table_of_authorities_passes_category_and_flags(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("toa"):
            toa = doc.end.insert_table_of_authorities(category="cases", passim=False)
        assert isinstance(toa, wordlive.TableOfAuthorities)
    call = _toa_add(fake_word).call_args
    assert call.args[1] == 1  # category "cases" -> 1 (positional)
    assert call.kwargs["Passim"] is False
    assert call.kwargs["KeepEntryFormatting"] is True


def test_add_table_of_authorities_defaults_all(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("toa"):
            doc.add_table_of_authorities()
    assert _toa_add(fake_word).call_args.args[1] == 0  # "all" -> 0


def test_insert_table_of_authorities_separators_keyword(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("toa"):
            doc.end.insert_table_of_authorities(entry_separator="\t", page_range_separator="-")
    kw = _toa_add(fake_word).call_args.kwargs
    assert kw["EntrySeparator"] == "\t"
    assert kw["PageRangeSeparator"] == "-"


def test_table_of_authorities_update_calls_com(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("toa"):
            toa = doc.end.insert_table_of_authorities()
        toa.update()
    assert toa.com.Update.call_count == 1


# --- exec ops ------------------------------------------------------------------


def test_exec_mark_then_insert_toa(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {"op": "mark_citation", "anchor_id": "bookmark:Address", "long_citation": "Foo"},
                {"op": "insert_table_of_authorities", "anchor_id": "end", "category": "cases"},
            ],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][-1]["table_of_authorities"] is True
    assert _toa_add(fake_word).call_args.args[1] == 1


# --- CLI -----------------------------------------------------------------------


def test_cli_mark_citation(fake_word):
    code, out = _invoke(
        [
            "--json",
            "mark-citation",
            "--anchor-id",
            "bookmark:Address",
            "--long",
            "Smith v. Jones",
            "--category",
            "cases",
        ]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["applied"]["long_citation"] == "Smith v. Jones"
    assert _field_add(fake_word).call_args.args[2].endswith("\\c 1")


def test_cli_table_of_authorities(fake_word):
    code, out = _invoke(["--json", "table-of-authorities", "--category", "statutes", "--no-passim"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["applied"]["category"] == "statutes"
    assert data["applied"]["passim"] is False
    assert _toa_add(fake_word).call_args.args[1] == 2
