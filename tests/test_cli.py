"""CLI shape: JSON output and exit codes."""

from __future__ import annotations

import json

from click.testing import CliRunner

from wordlive.cli.main import (
    EXIT_ANCHOR_NOT_FOUND,
    EXIT_OK,
    EXIT_WORD_NOT_RUNNING,
    main,
)


def _invoke(args: list[str]) -> tuple[int, str, str]:
    runner = CliRunner()
    result = runner.invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout, result.stderr


def test_status_lists_active_doc(fake_word):
    code, out, _ = _invoke(["status"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert isinstance(data, list)
    assert data[0]["name"] == "Test.docx"
    assert data[0]["is_active"] is True


def test_outline_returns_headings(fake_word):
    code, out, _ = _invoke(["outline"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert [item["text"] for item in data] == ["Introduction", "Risks"]


def test_read_bookmark_success(fake_word):
    code, out, _ = _invoke(["read", "bookmark", "Address"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert "text" in data


def test_read_bookmark_missing(fake_word):
    code, _, err = _invoke(["read", "bookmark", "DoesNotExist"])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "bookmark" in err.lower()


def test_read_cc_success(fake_word):
    code, out, _ = _invoke(["read", "cc", "Signatory"])
    assert code == EXIT_OK
    assert json.loads(out) == {"text": "Jane Doe"}


def test_write_bookmark(fake_word):
    code, out, _ = _invoke(["write", "bookmark", "Address", "--text", "123 Main"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor"]["kind"] == "bookmark"


def test_insert_after_heading(fake_word):
    code, out, _ = _invoke(["insert", "--after-heading", "Introduction", "--text", "new para"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["after_heading"] == "Introduction"


def test_insert_after_missing_heading(fake_word):
    code, _, err = _invoke(["insert", "--after-heading", "Nope", "--text", "x"])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "heading" in err.lower()


def test_word_not_running_returns_exit_4(no_word):
    code, _, err = _invoke(["status"])
    assert code == EXIT_WORD_NOT_RUNNING
    assert "not running" in err.lower() or "word" in err.lower()
