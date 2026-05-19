"""CLI shape: JSON output and exit codes."""

from __future__ import annotations

import json
from pathlib import Path

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


# ---------------------------------------------------------------------------
# v0.1: replace / go-to / exec
# ---------------------------------------------------------------------------


def test_replace_by_heading_id(fake_word):
    code, out, _ = _invoke(["replace", "--anchor-id", "heading:1", "--text", "New intro"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor"]["kind"] == "heading"


def test_replace_by_bookmark_id(fake_word):
    code, out, _ = _invoke(["replace", "--anchor-id", "bookmark:Address", "--text", "123 Main"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor"]["kind"] == "bookmark"
    assert data["anchor"]["name"] == "Address"


def test_replace_missing_anchor(fake_word):
    code, _, err = _invoke(["replace", "--anchor-id", "bookmark:Nope", "--text", "..."])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "bookmark" in err.lower()


def test_replace_bad_scheme(fake_word):
    code, _, _ = _invoke(["replace", "--anchor-id", "table:1", "--text", "..."])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_go_to_heading(fake_word):
    code, out, _ = _invoke(["go-to", "--anchor-id", "heading:1"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor_id"] == "heading:1"


def test_go_to_missing_anchor(fake_word):
    code, _, _ = _invoke(["go-to", "--anchor-id", "bookmark:Nope"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_exec_all_ops_succeed(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "label": "Batch update",
                "ops": [
                    {"op": "write_bookmark", "name": "Address", "text": "123 Main"},
                    {"op": "write_cc", "name": "Signatory", "text": "Jane Doe"},
                    {"op": "insert_after_heading", "heading": "Introduction", "text": "New para"},
                    {"op": "replace", "anchor_id": "heading:3", "text": "Updated Risks"},
                ],
            }
        ),
        encoding="utf-8",
    )

    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["ops_run"] == 4
    assert data["label"] == "Batch update"


def test_exec_wraps_ops_in_single_undo_record(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "write_bookmark", "name": "Address", "text": "A"},
                    {"op": "write_bookmark", "name": "Address", "text": "B"},
                ]
            }
        ),
        encoding="utf-8",
    )

    code, _, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    # All ops should ride a single UndoRecord — Start/End called exactly once each.
    fake_word.UndoRecord.StartCustomRecord.assert_called_once()
    fake_word.UndoRecord.EndCustomRecord.assert_called_once()


def test_exec_stops_at_first_failure_and_reports_partial(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "write_bookmark", "name": "Address", "text": "ok-1"},
                    {"op": "write_bookmark", "name": "Nope", "text": "boom"},
                    {"op": "write_bookmark", "name": "Address", "text": "never-runs"},
                ]
            }
        ),
        encoding="utf-8",
    )

    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_ANCHOR_NOT_FOUND
    data = json.loads(out)
    assert data["ok"] is False
    assert data["ops_run"] == 1
    assert data["failure"]["index"] == 1
    assert data["failure"]["type"] == "AnchorNotFoundError"


def test_exec_unknown_op_is_click_error(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps({"ops": [{"op": "drop_table", "name": "anything"}]}),
        encoding="utf-8",
    )

    # Unknown op is a ClickException — non-zero exit, but not our taxonomy.
    code, _, err = _invoke(["exec", "--script", str(script)])
    assert code != EXIT_OK
    assert "drop_table" in err
