"""Revisions reader — `doc.revisions`, the `revisions` CLI command.

Round-trips against the `fake_word` MagicMock, whose `doc.Revisions` is a
`_FakeRevisions` seeded with one tracked insertion (`type=1`, author "Reviewer",
text "Body " over range 13–18).
"""

from __future__ import annotations

import json
from unittest.mock import MagicMock

import pytest

import wordlive
from wordlive.cli.main import EXIT_OK, main
from wordlive.exceptions import AnchorNotFoundError


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


# --- the reader ----------------------------------------------------------------


def test_revisions_seeded(fake_word):
    with wordlive.attach() as word:
        revs = word.documents.active.revisions
        assert len(revs) == 1
        r = revs[1]
        assert r.index == 1
        assert r.type == "insert"
        assert r.author == "Reviewer"
        assert r.text == "Body"


def test_revisions_list_shape(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.revisions.list()
    assert len(rows) == 1
    row = rows[0]
    assert row["index"] == 1
    assert row["type"] == "insert"
    assert row["author"] == "Reviewer"
    assert row["anchor_id"] == "range:13-18"
    assert row["start"] == 13 and row["end"] == 18
    assert row["date"]  # the seeded ISO string round-trips


def test_revision_type_mapping(fake_word):
    # type=2 -> delete, an unknown int -> "other".
    fake_word.ActiveDocument.Revisions = _revisions_with(
        {"type": 2, "author": "X", "text": "old", "start": 5, "end": 8},
        {"type": 99, "author": "Y", "text": "?", "start": 8, "end": 9},
    )
    with wordlive.attach() as word:
        rows = word.documents.active.revisions.list()
    assert [r["type"] for r in rows] == ["delete", "other"]


def test_revision_index_out_of_range(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError):
            word.documents.active.revisions[5]


def test_revisions_empty(fake_word):
    fake_word.ActiveDocument.Revisions = _revisions_with()
    with wordlive.attach() as word:
        assert word.documents.active.revisions.list() == []


# --- CLI -----------------------------------------------------------------------


def test_cli_revisions_json(fake_word):
    code, out = _invoke(["--json", "revisions"])
    assert code == EXIT_OK
    rows = json.loads(out)
    assert rows[0]["type"] == "insert"
    assert rows[0]["author"] == "Reviewer"


def test_cli_revisions_text(fake_word):
    code, out = _invoke(["revisions"])
    assert code == EXIT_OK
    assert "insert" in out and "Reviewer" in out


def test_cli_track_status(fake_word):
    code, out = _invoke(["--json", "track", "status"])
    assert code == EXIT_OK
    assert json.loads(out) == {"tracked": False}


# --- helper --------------------------------------------------------------------


def _revisions_with(*specs):
    """A minimal stand-in for doc.Revisions: Count, 1-based call, iteration."""
    items = []
    for s in specs:
        rev = MagicMock(name="Revision")
        rev.Type = s["type"]
        rev.Author = s["author"]
        rev.Date = "2026-06-08T12:00:00"
        rng = MagicMock(name="RevRange")
        rng.Start, rng.End, rng.Text = s["start"], s["end"], s["text"]
        rev.Range = rng
        items.append(rev)
    coll = MagicMock(name="Revisions")
    coll.Count = len(items)
    coll.side_effect = lambda i: items[i - 1]
    coll.__iter__ = MagicMock(side_effect=lambda: iter(items))
    return coll
