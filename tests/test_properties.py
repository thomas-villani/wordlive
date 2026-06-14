"""Document properties — `doc.properties`, the `properties` CLI group, MCP.

Round-trips against `fake_word`, whose document is seeded with built-in
properties (Title, Author, a read-only "Word count", an unreadable "Last print
date") and one custom property (Project=Apollo).
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


# --- the reader ----------------------------------------------------------------


def test_read_builtin_and_custom(fake_word):
    with wordlive.attach() as word:
        data = word.documents.active.properties.read()
    # Title/Author/Word count are readable; the unset "Last print date" is skipped.
    assert data["builtin"]["Title"] == "Quarterly Report"
    assert data["builtin"]["Author"] == "Jane Doe"
    assert data["builtin"]["Word count"] == 1234
    assert "Last print date" not in data["builtin"]
    assert data["custom"] == {"Project": "Apollo"}


def test_get_builtin_then_custom(fake_word):
    with wordlive.attach() as word:
        props = word.documents.active.properties
        assert props.get("Title") == "Quarterly Report"
        assert props.get("Project") == "Apollo"
        with pytest.raises(AnchorNotFoundError):
            props.get("Nonexistent")


# --- writes --------------------------------------------------------------------


def test_set_builtin(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.properties.set("Title", "New Title")
        assert doc.properties.builtin()["Title"] == "New Title"


def test_set_custom_new_and_update(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.properties.set("Client", "Acme", custom=True)  # new
            doc.properties.set("Project", "Gemini", custom=True)  # update
        custom = doc.properties.custom()
        assert custom["Client"] == "Acme"
        assert custom["Project"] == "Gemini"


def test_set_unknown_builtin_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError, match="not a built-in"):
            doc.properties.set("Bogus", "x")


def test_set_readonly_builtin_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError, match="read-only"):
            doc.properties.set("Word count", 5)


def test_delete_custom_and_missing(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("t"):
            doc.properties.delete("Project")
        assert "Project" not in doc.properties.custom()
        with pytest.raises(AnchorNotFoundError):
            doc.properties.delete("Project")


# --- exec ops ------------------------------------------------------------------


def test_op_set_and_delete_property(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {"op": "set_property", "name": "Subject", "value": "Q3"},
                {"op": "set_property", "name": "Client", "value": "Acme", "custom": True},
                {"op": "delete_property", "name": "Project"},
            ],
            label="t",
        )
        assert exc is None and result["ok"]
        assert doc.properties.builtin()["Subject"] == "Q3"
        assert doc.properties.custom() == {"Client": "Acme"}


# --- CLI -----------------------------------------------------------------------


def test_cli_properties_list(fake_word):
    code, out = _invoke(["properties", "list"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["builtin"]["Title"] == "Quarterly Report"
    assert data["custom"]["Project"] == "Apollo"


def test_cli_properties_set_and_delete(fake_word):
    code, out = _invoke(["properties", "set", "--name", "Author", "--value", "X"])
    assert code == EXIT_OK and json.loads(out)["ok"]
    code, out = _invoke(["properties", "set", "--name", "C", "--value", "v", "--custom"])
    assert code == EXIT_OK and json.loads(out)["custom"] is True
    code, out = _invoke(["properties", "delete", "--name", "Project"])
    assert code == EXIT_OK and json.loads(out)["ok"]


# --- MCP -----------------------------------------------------------------------


def test_mcp_properties(fake_word):
    pytest.importorskip("mcp")
    from wordlive.mcp._worker import InlineWorker
    from wordlive.mcp.server import _read_impl, _write_impl

    w = InlineWorker()
    data = _read_impl(w, "properties", {})
    assert data["builtin"]["Title"] == "Quarterly Report"
    r = _write_impl(w, "set_property", {"name": "Title", "value": "Z"})
    assert r["ok"]
    r = _write_impl(w, "set_property", {"name": "K", "value": "v", "custom": True})
    assert r["ok"]
    r = _write_impl(w, "delete_property", {"name": "Project"})
    assert r["ok"]
