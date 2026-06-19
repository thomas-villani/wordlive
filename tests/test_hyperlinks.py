"""Hyperlinks reader — `doc.hyperlinks`, the `hyperlinks` CLI command, MCP.

Round-trips against `fake_word`, seeded with one external link ("Acme" ->
https://acme.example) whose range sits in the body paragraph (13–29).
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


# --- the setters ---------------------------------------------------------------


def test_set_address_and_text(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("link"):
            doc.hyperlinks[1].set_address("https://new.example").set_text("New")
        h = doc.hyperlinks[1]
        assert h.address == "https://new.example"
        assert h.text == "New"


def test_update_multi_field_and_chainable(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("link"):
            ret = doc.hyperlinks[1].update(address="https://x.example", text="X", screen_tip="go")
        assert isinstance(ret, wordlive.Hyperlink)
        h = doc.hyperlinks[1]
        assert h.address == "https://x.example"
        assert h.text == "X"


def test_set_none_leaves_field_untouched(fake_word):
    # None = leave (matches the op/CLI/MCP `is not None` filter); no-op here.
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("link"):
            doc.hyperlinks[1].update(address=None, screen_tip="tip")
        h = doc.hyperlinks[1]
        assert h.address == "https://acme.example"  # untouched
        assert h.to_dict()["screen_tip"] == "tip"


def test_sub_address_clears_with_empty_string(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("link"):
            doc.hyperlinks[1].set_sub_address("Risks")
        assert doc.hyperlinks[1].sub_address == "Risks"
        with doc.edit("link"):
            doc.hyperlinks[1].set_sub_address("")  # clearable (Word allows)
        assert doc.hyperlinks[1].sub_address == ""


def test_address_and_text_cannot_be_cleared(fake_word):
    # Word keeps a link pointing somewhere with visible text; "" raises.
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.hyperlinks[1].set_address("")
        with pytest.raises(OpError):
            doc.hyperlinks[1].set_text("")


def test_set_bad_index_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError):
            word.documents.active.hyperlinks[2].set_text("x")


# --- exec op -------------------------------------------------------------------


def test_op_set_hyperlink(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_hyperlink", "index": 1, "address": "https://op.example"}],
            label="test",
        )
        assert exc is None and result["ok"] is True
        assert doc.hyperlinks[1].address == "https://op.example"


def test_op_set_hyperlink_no_fields_fails(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "set_hyperlink", "index": 1}], label="test")
        assert isinstance(exc, OpError) and result["ok"] is False


# --- CLI -----------------------------------------------------------------------


def test_cli_set_hyperlink(fake_word):
    code, out = _invoke(["set-hyperlink", "--index", "1", "--address", "https://cli.example"])
    assert code == EXIT_OK
    payload = json.loads(out)
    assert payload["ok"] is True
    assert payload["applied"]["address"] == "https://cli.example"


def test_cli_set_hyperlink_no_fields_errors(fake_word):
    code, _ = _invoke(["set-hyperlink", "--index", "1"])
    assert code != EXIT_OK


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


def test_mcp_set_hyperlink_aliases_url(fake_word):
    pytest.importorskip("mcp")
    from wordlive.mcp.server import _build_write_op

    # MCP vocabulary parity: url -> address, bookmark -> sub_address.
    op = _build_write_op("set_hyperlink", {"index": 1, "url": "https://m.example"})
    assert op == {"op": "set_hyperlink", "index": 1, "address": "https://m.example"}
    op2 = _build_write_op("set_hyperlink", {"index": 2, "bookmark": "Risks"})
    assert op2 == {"op": "set_hyperlink", "index": 2, "sub_address": "Risks"}


def test_mcp_set_hyperlink_no_fields_raises(fake_word):
    pytest.importorskip("mcp")
    from wordlive.mcp.server import _build_write_op

    with pytest.raises(OpError):
        _build_write_op("set_hyperlink", {"index": 1})
