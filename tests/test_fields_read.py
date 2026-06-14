"""Fields reader — `doc.fields`, the `fields` CLI command, MCP read.

The read mirror of `insert_field` (whose write side lives in test_fields.py).
Round-trips against `fake_word`, seeded with one PAGE field whose code range
sits in the body paragraph (13–29).
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive._fields import field_kind
from wordlive.cli.main import EXIT_OK, main
from wordlive.exceptions import AnchorNotFoundError


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


def _make_app(**kwargs):
    from tests.conftest import _make_application, _make_document

    return _make_application([_make_document(**kwargs)])


# --- the kind parser -----------------------------------------------------------


def test_field_kind_parsing():
    assert field_kind("PAGE", 33) == "PAGE"
    assert field_kind(" REF myBookmark \\h ", 3) == "REF"
    assert field_kind("docvariable Client", 64) == "DOCVARIABLE"
    assert field_kind("", 99) == "field:99"


# --- the reader ----------------------------------------------------------------


def test_seeded(fake_word):
    with wordlive.attach() as word:
        fields = word.documents.active.fields
        assert len(fields) == 1
        f = fields[1]
        assert f.index == 1
        assert f.kind == "PAGE"
        assert f.code == "PAGE"
        assert f.result == "1"
        assert f.type == 33


def test_list_shape(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.fields.list()
    row = rows[0]
    assert row["index"] == 1
    assert row["kind"] == "PAGE"
    assert row["code"] == "PAGE"
    assert row["result"] == "1"
    assert row["type"] == 33
    assert row["locked"] is False
    assert row["anchor_id"] == "range:16-17"
    assert row["para"] == "para:2"


def test_ref_field_kind(monkeypatch):
    app = _make_app(
        fields=[{"code": "REF Risks \\h", "result": "Risks", "type": 3, "start": 5, "end": 6}],
    )
    from wordlive import _com

    monkeypatch.setattr(_com, "get_active_word", lambda: app)
    monkeypatch.setattr(_com, "launch_word", lambda visible=True: app)
    with wordlive.attach() as word:
        row = word.documents.active.fields.list()[0]
    assert row["kind"] == "REF"
    assert row["code"] == "REF Risks \\h"


def test_index_out_of_range(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError):
            _ = word.documents.active.fields[5]


# --- CLI -----------------------------------------------------------------------


def test_cli_fields(fake_word):
    code, out = _invoke(["fields"])
    assert code == EXIT_OK
    rows = json.loads(out)
    assert rows[0]["kind"] == "PAGE"


def test_cli_fields_text(fake_word):
    code, out = _invoke(["--text", "fields"])
    assert code == EXIT_OK
    assert "PAGE" in out


# --- MCP -----------------------------------------------------------------------


def test_mcp_fields(fake_word):
    pytest.importorskip("mcp")
    from wordlive.mcp._worker import InlineWorker
    from wordlive.mcp.server import _read_impl

    rows = _read_impl(InlineWorker(), "fields", {})
    assert rows[0]["kind"] == "PAGE"
