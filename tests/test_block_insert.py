"""Block insert + inline runs — Python API, exec op, CLI, and MCP surfaces.

The markdown parser itself is covered in `test_runs.py`; here we check the four
surfaces wire `insert_block` / `insert_paragraph runs` through correctly and that
the spanning range comes back. Visible-formatting correctness against real Word
lives in the smoke tests.
"""

from __future__ import annotations

import json
from typing import Any

import pytest
from click.testing import CliRunner

import wordlive
from wordlive._anchors import RangeAnchor
from wordlive._ops import run_batch
from wordlive.cli.main import EXIT_OK, main
from wordlive.exceptions import OpError, StyleNotFoundError
from wordlive.mcp._worker import InlineWorker
from wordlive.mcp.server import _write_impl

W = InlineWorker()


def _invoke(args: list[str], *, input: str | None = None) -> tuple[int, str, str]:
    result = CliRunner().invoke(main, args, input=input, catch_exceptions=False)
    return result.exit_code, result.stdout, result.stderr


# ---------------------------------------------------------------------------
# Python API — Anchor.insert_block
# ---------------------------------------------------------------------------


class TestInsertBlockApi:
    def test_returns_range_anchor_spanning_block(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("block"):
                rng = doc.headings["Introduction"].insert_block(
                    [{"text": "one"}, {"text": "two"}, {"text": "three"}]
                )
            assert isinstance(rng, RangeAnchor)
            # 3 paragraphs of 3/3/5 chars joined by 2 CRs -> span = 3+1+3+1+5 = 13.
            assert rng.end - rng.start == len("one") + 1 + len("two") + 1 + len("three")

    def test_unknown_paragraph_style_raises_before_mutation(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(StyleNotFoundError):
                doc.headings["Introduction"].insert_block([{"text": "x", "style": "No Such"}])

    def test_unknown_run_style_raises_before_mutation(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(StyleNotFoundError):
                doc.headings["Introduction"].insert_block(
                    [{"runs": [{"text": "x", "style": "No Such"}]}]
                )

    def test_bad_where_raises(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(ValueError, match="before.*after"):
                doc.headings["Introduction"].insert_block([{"text": "x"}], where="sideways")

    def test_malformed_items_raise_operror(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(OpError):
                doc.headings["Introduction"].insert_block([])


# ---------------------------------------------------------------------------
# exec op — insert_block / insert_paragraph runs
# ---------------------------------------------------------------------------


class TestInsertBlockOp:
    def test_reports_range_and_paragraph_count(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(
                doc,
                [
                    {
                        "op": "insert_block",
                        "anchor_id": "heading:1",
                        "items": ["a", {"text": "**b**"}, {"runs": [{"text": "c"}]}],
                    }
                ],
                label="b",
            )
        assert exc is None
        out = result["outputs"][0]
        assert out["op"] == "insert_block"
        assert out["paragraphs"] == 3
        assert out["anchor_id"].startswith("range:")

    def test_insert_paragraph_with_runs(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(
                doc,
                [
                    {
                        "op": "insert_paragraph",
                        "anchor_id": "heading:1",
                        "runs": [{"text": "L", "bold": True}, {"text": " rest"}],
                    }
                ],
                label="r",
            )
        assert exc is None and result["ok"] is True

    def test_insert_paragraph_text_xor_runs(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(
                doc,
                [
                    {
                        "op": "insert_paragraph",
                        "anchor_id": "heading:1",
                        "text": "a",
                        "runs": [{"text": "b"}],
                    }
                ],
                label="bad",
            )
        assert exc is not None
        assert "exactly one of 'text' or 'runs'" in result["failure"]["error"]

    def test_insert_block_missing_items_fails_cleanly(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(doc, [{"op": "insert_block", "anchor_id": "end"}], label="bad")
        assert exc is not None
        assert result["ok"] is False
        assert "items" in result["failure"]["error"]

    def test_plain_insert_paragraph_still_works(self, fake_word: Any) -> None:
        # Back-compat: text-only insert_paragraph keeps its literal behaviour.
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(
                doc,
                [{"op": "insert_paragraph", "anchor_id": "heading:1", "text": "plain"}],
                label="p",
            )
        assert exc is None and result["ok"] is True


# ---------------------------------------------------------------------------
# CLI — insert-block / insert --runs
# ---------------------------------------------------------------------------


class TestInsertBlockCli:
    def test_insert_block_reports_range(self, fake_word: Any) -> None:
        items = json.dumps([{"text": "**A** b", "style": "Heading 1"}, "plain"])
        code, out, _ = _invoke(["insert-block", "--anchor-id", "heading:1", "--items", items])
        assert code == EXIT_OK
        data = json.loads(out)
        assert data["ok"] is True
        assert data["paragraphs"] == 2
        assert data["anchor_id"].startswith("range:")

    def test_insert_block_items_from_stdin(self, fake_word: Any) -> None:
        items = json.dumps(["one", "two"])
        code, out, _ = _invoke(["insert-block", "--anchor-id", "end", "--items", "-"], input=items)
        assert code == EXIT_OK
        assert json.loads(out)["paragraphs"] == 2

    def test_insert_block_bad_json_is_usage_error(self, fake_word: Any) -> None:
        code, _, err = _invoke(["insert-block", "--anchor-id", "end", "--items", "{not json"])
        assert code != EXIT_OK
        assert "JSON" in err

    def test_insert_with_runs(self, fake_word: Any) -> None:
        runs = json.dumps([{"text": "L", "bold": True}, {"text": " rest"}])
        code, out, _ = _invoke(["insert", "--anchor-id", "heading:1", "--runs", runs])
        assert code == EXIT_OK
        assert json.loads(out)["ok"] is True

    def test_insert_text_and_runs_is_usage_error(self, fake_word: Any) -> None:
        code, _, err = _invoke(["insert", "--anchor-id", "end", "--text", "a", "--runs", "[]"])
        assert code != EXIT_OK
        assert "exactly one" in err

    def test_insert_neither_text_nor_runs_is_usage_error(self, fake_word: Any) -> None:
        code, _, err = _invoke(["insert", "--anchor-id", "end"])
        assert code != EXIT_OK
        assert "exactly one" in err


# ---------------------------------------------------------------------------
# MCP — word_write insert_block / insert runs
# ---------------------------------------------------------------------------


class TestInsertBlockMcp:
    def test_insert_block(self, fake_word: Any) -> None:
        r = _write_impl(
            W,
            "insert_block",
            {"anchor_id": "heading:1", "items": ["a", {"text": "**b**"}]},
        )
        assert r["ok"] is True and r["command"] == "insert_block"

    def test_insert_with_runs(self, fake_word: Any) -> None:
        r = _write_impl(W, "insert", {"anchor_id": "heading:1", "runs": [{"text": "x"}]})
        assert r["ok"] is True

    def test_insert_text_and_runs_errors(self, fake_word: Any) -> None:
        with pytest.raises(Exception, match="exactly one of 'text' or 'runs'"):
            _write_impl(W, "insert", {"anchor_id": "end", "text": "a", "runs": [{"text": "b"}]})
