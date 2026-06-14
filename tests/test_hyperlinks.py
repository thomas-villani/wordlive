"""Hyperlinks reader — `doc.hyperlinks`, the `hyperlinks` CLI command, MCP.

Round-trips against `fake_word`, seeded with one external link ("Acme" ->
https://acme.example) whose range sits in the body paragraph (13–29).
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive.cli.main import EXIT_OK, main
from wordlive.exceptions import AnchorNotFoundError


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


def _make_app(**kwargs):
    from tests.conftest import _make_application, _make_document

    return _make_application([_make_document(**kwargs)])


# --- the reader ----------------------------------------------------------------


def test_seeded(fake_word):
    with wordlive.attach() as word:
        links = word.documents.active.hyperlinks
        assert len(links) == 1
        h = links[1]
        assert h.index == 1
        assert h.text == "Acme"
        assert h.address == "https://acme.example"
        assert h.sub_address == ""


def test_list_shape(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.hyperlinks.list()
    assert len(rows) == 1
    row = rows[0]
    assert row["index"] == 1
    assert row["text"] == "Acme"
    assert row["address"] == "https://acme.example"
    assert row["anchor_id"] == "range:15-19"
    assert row["para"] == "para:2"  # the body paragraph (13–29)


def test_internal_link_sub_address(monkeypatch):
    app = _make_app(
        hyperlinks=[{"text": "See Risks", "sub_address": "Risks", "start": 5, "end": 9}],
    )
    from wordlive import _com

    monkeypatch.setattr(_com, "get_active_word", lambda: app)
    monkeypatch.setattr(_com, "launch_word", lambda visible=True: app)
    with wordlive.attach() as word:
        row = word.documents.active.hyperlinks.list()[0]
    assert row["address"] == ""
    assert row["sub_address"] == "Risks"


def test_index_out_of_range(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError):
            _ = word.documents.active.hyperlinks[2]


# --- CLI -----------------------------------------------------------------------


def test_cli_hyperlinks(fake_word):
    code, out = _invoke(["hyperlinks"])
    assert code == EXIT_OK
    rows = json.loads(out)
    assert rows[0]["address"] == "https://acme.example"


def test_cli_hyperlinks_text(fake_word):
    code, out = _invoke(["--text", "hyperlinks"])
    assert code == EXIT_OK
    assert "https://acme.example" in out


# --- MCP -----------------------------------------------------------------------


def test_mcp_hyperlinks(fake_word):
    pytest.importorskip("mcp")
    from wordlive.mcp._worker import InlineWorker
    from wordlive.mcp.server import _read_impl

    rows = _read_impl(InlineWorker(), "hyperlinks", {})
    assert rows[0]["text"] == "Acme"
