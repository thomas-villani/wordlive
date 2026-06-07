"""Footnotes & endnotes — `Anchor.insert_footnote`/`insert_endnote`, the
`footnote:N`/`endnote:N` anchors, `doc.footnotes`/`doc.endnotes`, and the ops.

Round-trips against the `fake_word` MagicMock: `doc.Footnotes`/`doc.Endnotes`
are `_FakeNotes` whose `Add` is a recording MagicMock that appends a note (the
`fake_word` doc is pre-seeded with one footnote whose mark sits at offset 20,
inside the body paragraph).
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive._ops import run_batch
from wordlive.cli.main import EXIT_OK, main
from wordlive.exceptions import AnchorNotFoundError, OpError


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


# --- the Anchor methods --------------------------------------------------------


def test_insert_footnote_adds_note_and_returns_anchor(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        before = len(doc.footnotes)
        with doc.edit("footnote"):
            note = doc.bookmarks["Address"].insert_footnote("See appendix.")
        assert isinstance(note, wordlive.Footnote)
        assert note.anchor_id == f"footnote:{note.index}"
        assert len(doc.footnotes) == before + 1
    add = fake_word.ActiveDocument.Footnotes.Add
    # Positional: empty Reference auto-numbers; the body text is the 3rd arg.
    assert add.call_args.args[1] == ""
    assert add.call_args.args[2] == "See appendix."


def test_insert_footnote_after_targets_end(fake_word):
    # bookmark:Address is (13, 24); "after" puts the reference mark at End=24.
    with wordlive.attach() as word:
        doc = word.documents.active
        note = doc.bookmarks["Address"].insert_footnote("x")
        assert int(note._note().Reference.Start) == 24


def test_insert_footnote_before_targets_start(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        note = doc.bookmarks["Address"].insert_footnote("x", where="before")
        assert int(note._note().Reference.Start) == 13


def test_insert_endnote_uses_endnotes_collection(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("endnote"):
            note = doc.bookmarks["Address"].insert_endnote("Endnote body.")
        assert isinstance(note, wordlive.Endnote)
        assert len(doc.endnotes) == 1
    assert fake_word.ActiveDocument.Endnotes.Add.call_args.args[2] == "Endnote body."


def test_insert_footnote_bad_where_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].insert_footnote("x", where="sideways")


# --- footnote:N addressing + set_text + delete ---------------------------------


def test_footnote_anchor_resolves(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        anchor = doc.anchor_by_id("footnote:1")
        assert isinstance(anchor, wordlive.Footnote)
        assert anchor.text == "A seeded footnote."


def test_footnote_anchor_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("footnote:9")


def test_footnote_set_text_writes_body(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("edit note"):
            doc.footnotes[1].set_text("Rewritten.")
        assert doc.footnotes[1].text == "Rewritten."


def test_footnote_delete_removes_reference(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        note = doc.footnotes[1]
        ref_delete = note._note().Reference.Delete
        with doc.edit("delete note"):
            note.delete()
        ref_delete.assert_called_once()


# --- discovery list ------------------------------------------------------------


def test_footnotes_list_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        rows = doc.footnotes.list()
    assert rows == [
        {
            "index": 1,
            "anchor_id": "footnote:1",
            "marker": "1",
            "text": "A seeded footnote.",
            "para": "para:2",  # mark at offset 20 sits in the body paragraph
        }
    ]


# --- exec ops ------------------------------------------------------------------


def test_exec_insert_footnote_reports_output(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_footnote", "anchor_id": "bookmark:Address", "text": "note"}],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["footnote"] == result["outputs"][0]["footnote"]
    assert result["outputs"][0]["anchor_id"].startswith("footnote:")


def test_exec_insert_endnote(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_endnote", "anchor_id": "bookmark:Address", "text": "end"}],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["endnote"] >= 1


def test_exec_insert_footnote_missing_text_fails(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "insert_footnote", "anchor_id": "bookmark:Address"}], label="bad"
        )
    assert exc is not None and result["ok"] is False
    assert "text" in result["failure"]["error"]


# --- CLI -----------------------------------------------------------------------


def test_cli_insert_footnote(fake_word):
    code, out = _invoke(
        ["--json", "insert-footnote", "--anchor-id", "bookmark:Address", "--text", "cli note"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["note_id"].startswith("footnote:")
    assert fake_word.ActiveDocument.Footnotes.Add.call_args.args[2] == "cli note"


def test_cli_footnotes_list(fake_word):
    code, out = _invoke(["--json", "footnotes"])
    assert code == EXIT_OK
    rows = json.loads(out)
    assert rows[0]["anchor_id"] == "footnote:1"
    assert rows[0]["text"] == "A seeded footnote."
