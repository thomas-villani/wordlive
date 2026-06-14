"""Proofing — `doc.proofing()`, the `proofing` CLI command, MCP read.

Round-trips against `fake_word`, seeded with one spelling error ("teh"), one
grammar error ("is are"), and a handful of readability statistics.
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive.cli.main import EXIT_OK, main


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


def _make_doc(**kwargs):
    # Build a bespoke document with whatever proofing seeds a test needs.
    from tests.conftest import _make_application, _make_document

    return _make_application([_make_document(**kwargs)])


# --- the reader ----------------------------------------------------------------


def test_proofing_shape(fake_word):
    with wordlive.attach() as word:
        data = word.documents.active.proofing()
    assert data["spelling"]["count"] == 1
    assert data["spelling"]["errors"][0]["text"] == "teh"
    assert data["spelling"]["errors"][0]["anchor_id"] == "range:14-17"
    assert data["grammar"]["count"] == 1
    assert data["grammar"]["errors"][0]["text"] == "is are"


def test_readability_slugged(fake_word):
    with wordlive.attach() as word:
        read = word.documents.active.proofing()["readability"]
    assert read["flesch_reading_ease"] == 65.5
    assert read["flesch_kincaid_grade_level"] == 7.2
    assert read["passive_sentences"] == 12.0
    assert read["words_per_sentence"] == 15.3


def test_proofing_unavailable_is_graceful(monkeypatch):
    # A checker that raises (proofing off / protected doc) -> count None, [].
    from tests.conftest import _FakeProofErrors, _FakeReadability

    app = _make_doc()
    doc = app.ActiveDocument
    doc.SpellingErrors = _FakeProofErrors([], raises=True)
    doc.GrammaticalErrors = _FakeProofErrors([], raises=True)
    doc.ReadabilityStatistics = _FakeReadability({}, raises=True)
    from wordlive import _com

    monkeypatch.setattr(_com, "get_active_word", lambda: app)
    monkeypatch.setattr(_com, "launch_word", lambda visible=True: app)
    with wordlive.attach() as word:
        data = word.documents.active.proofing()
    assert data["spelling"]["count"] is None
    assert data["spelling"]["errors"] == []
    assert data["readability"] == {}


# --- CLI -----------------------------------------------------------------------


def test_cli_proofing(fake_word):
    code, out = _invoke(["proofing"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["spelling"]["count"] == 1
    assert data["readability"]["flesch_reading_ease"] == 65.5


def test_cli_proofing_text(fake_word):
    code, out = _invoke(["--text", "proofing"])
    assert code == EXIT_OK
    assert "spelling errors: 1" in out
    assert "flesch_reading_ease" in out


# --- MCP -----------------------------------------------------------------------


def test_mcp_proofing(fake_word):
    pytest.importorskip("mcp")
    from wordlive.mcp._worker import InlineWorker
    from wordlive.mcp.server import _read_impl

    data = _read_impl(InlineWorker(), "proofing", {})
    assert data["spelling"]["count"] == 1
    assert data["readability"]["flesch_kincaid_grade_level"] == 7.2
