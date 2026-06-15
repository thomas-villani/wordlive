"""Bibliography sources — `doc.sources` (`add`/`add_xml`/lookup), `Source`,
`doc.bibliography_style`, the `add_source` / `set_bibliography_style` ops, and CLI.

`doc.Bibliography` is a `_FakeBibliography`: `Sources.Add(xml)` records the XML and
parses its `<b:Tag>`, and `BibliographyStyle` is a settable attribute. The XML
builder is asserted by substring on the recorded `Sources.Add` argument.
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


def _bib(fake_word):
    return fake_word.ActiveDocument.Bibliography


def _sources_add(fake_word):
    return fake_word.ActiveDocument.Bibliography.Sources.Add


# --- sources.add ---------------------------------------------------------------


def test_add_builds_source_xml(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("src"):
            src = doc.sources.add("book", author="Smith, John", title="The Work", year=2020)
        assert isinstance(src, wordlive.Source)
        assert src.tag == "Smith2020"  # auto-derived: surname + year
    xml = _sources_add(fake_word).call_args.args[0]
    assert "<b:SourceType>Book</b:SourceType>" in xml
    assert "<b:Tag>Smith2020</b:Tag>" in xml
    assert "<b:Last>Smith</b:Last><b:First>John</b:First>" in xml
    assert "<b:Title>The Work</b:Title>" in xml


def test_add_explicit_tag_and_journal_type(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("src"):
            doc.sources.add(
                "journal_article",
                tag="Doe19",
                author="Doe, Jane",
                journal_name="J. Things",
                volume="5",
            )
    xml = _sources_add(fake_word).call_args.args[0]
    assert "<b:SourceType>JournalArticle</b:SourceType>" in xml
    assert "<b:Tag>Doe19</b:Tag>" in xml
    assert "<b:JournalName>J. Things</b:JournalName>" in xml
    assert "<b:Volume>5</b:Volume>" in xml


def test_add_extra_field_passthrough(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("src"):
            doc.sources.add("misc", tag="X1", Medium="Web")
    assert "<b:Medium>Web</b:Medium>" in _sources_add(fake_word).call_args.args[0]


def test_add_bad_source_type_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.sources.add("banana", tag="X")


def test_add_no_tag_no_author_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.sources.add("book", title="Untitled")


# --- sources.add_xml -----------------------------------------------------------

_RAW = (
    '<b:Source xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography">'
    "<b:SourceType>Book</b:SourceType><b:Tag>Raw1</b:Tag></b:Source>"
)


def test_add_xml_passthrough(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("src"):
            src = doc.sources.add_xml(_RAW)
        assert src.tag == "Raw1"
    assert _sources_add(fake_word).call_args.args[0] == _RAW


def test_add_xml_without_tag_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.sources.add_xml(
                '<b:Source xmlns:b="x"><b:SourceType>Book</b:SourceType></b:Source>'
            )


def test_add_xml_malformed_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.sources.add_xml("<b:Source>oops")


# --- collection ----------------------------------------------------------------


def test_collection_list_contains_getitem_len(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("src"):
            doc.sources.add("book", tag="A1", author="A, A", year=2001)
            doc.sources.add("book", tag="B2", author="B, B", year=2002)
        assert doc.sources.list() == ["A1", "B2"]
        assert "A1" in doc.sources
        assert "ZZ" not in doc.sources
        assert len(doc.sources) == 2
        assert doc.sources["B2"].tag == "B2"
        assert [s.tag for s in doc.sources] == ["A1", "B2"]


def test_getitem_missing_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(wordlive.AnchorNotFoundError):
            doc.sources["nope"]


def test_source_delete_calls_com(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("src"):
            src = doc.sources.add("book", tag="Del1", author="D, D", year=2020)
        src.delete()
        assert src.cited is True
    assert _bib(fake_word).Sources(1).Delete.call_count == 1


# --- bibliography_style --------------------------------------------------------


def test_bibliography_style_get_set(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.bibliography_style == "APA"
        doc.bibliography_style = "MLA"
        assert doc.bibliography_style == "MLA"
    assert _bib(fake_word).BibliographyStyle == "MLA"


def test_bibliography_style_empty_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bibliography_style = "  "


# --- exec ops ------------------------------------------------------------------


def test_exec_add_source_typed(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "add_source",
                    "source_type": "book",
                    "tag": "Ex1",
                    "author": "Ex, Ample",
                    "year": 2020,
                }
            ],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["source"] == "Ex1"
    assert "<b:Tag>Ex1</b:Tag>" in _sources_add(fake_word).call_args.args[0]


def test_exec_add_source_xml(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "add_source", "source_type": "book", "xml": _RAW}], label="test"
        )
    assert exc is None
    assert result["outputs"][0]["source"] == "Raw1"


def test_exec_set_bibliography_style(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "set_bibliography_style", "style": "IEEE"}], label="test"
        )
    assert exc is None and result["ok"] is True
    assert _bib(fake_word).BibliographyStyle == "IEEE"


# --- CLI -----------------------------------------------------------------------


def test_cli_add_source(fake_word):
    code, out = _invoke(
        [
            "--json",
            "add-source",
            "--type",
            "book",
            "--tag",
            "Cli1",
            "--author",
            "Cli, Author",
            "--year",
            "2021",
        ]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["source"] == "Cli1"
    assert "<b:Tag>Cli1</b:Tag>" in _sources_add(fake_word).call_args.args[0]


def test_cli_bibliography_style(fake_word):
    code, out = _invoke(["--json", "bibliography-style", "--style", "Chicago"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["applied"]["style"] == "Chicago"
    assert _bib(fake_word).BibliographyStyle == "Chicago"
