"""Compose helpers — insert_section / insert_markdown / replace_section_body.

Checks the four surfaces (Python API, exec op, CLI, MCP) wire the helpers
through and return the spanning range. Markdown classification is covered in
test_markdown.py; visible Word formatting (real list markers, applied styles) in
the smoke tests.
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
from wordlive.exceptions import OpError
from wordlive.mcp._worker import InlineWorker
from wordlive.mcp.server import _write_impl

W = InlineWorker()


def _invoke(args: list[str], *, input: str | None = None) -> tuple[int, str, str]:
    result = CliRunner().invoke(main, args, input=input, catch_exceptions=False)
    return result.exit_code, result.stdout, result.stderr


# ---------------------------------------------------------------------------
# Python API
# ---------------------------------------------------------------------------


class TestInsertSectionApi:
    def test_returns_range_spanning_heading_and_body(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("sec"):
                rng = doc.headings["Introduction"].insert_section(
                    "New Section", ["body one", "body two"], level=2
                )
            assert isinstance(rng, RangeAnchor)
            # heading + 2 body paras, joined by CRs: 11 + 1 + 8 + 1 + 8.
            assert rng.end - rng.start == len("New Section") + 1 + len("body one") + 1 + len(
                "body two"
            )

    def test_bare_string_body_is_one_paragraph(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("sec"):
                rng = doc.headings["Introduction"].insert_section("H", "just one")
            assert rng.end - rng.start == len("H") + 1 + len("just one")

    def test_bad_level_raises_value_error(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(ValueError, match="level"):
                doc.headings["Introduction"].insert_section("H", ["b"], level=0)

    def test_missing_heading_style_raises_before_mutation(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            # Heading 9 isn't seeded in the fake's styles → fails up front.
            from wordlive.exceptions import StyleNotFoundError

            with pytest.raises(StyleNotFoundError):
                doc.headings["Introduction"].insert_section("H", ["b"], level=9)


class TestInsertMarkdownApi:
    def test_returns_range_anchor(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("md"):
                rng = doc.headings["Introduction"].insert_markdown(
                    "# Title\n\nA para.\n\n- one\n- two"
                )
            assert isinstance(rng, RangeAnchor)
            assert rng.end > rng.start

    def test_bulleted_segment_gets_list_formatting(self, fake_word: Any) -> None:
        # A pure-bullet doc is one segment, so the returned range == that list's
        # span; the fake caches ranges by offset, so list_info round-trips.
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("md"):
                rng = doc.headings["Introduction"].insert_markdown("- a\n- b\n- c")
            assert rng.list_info()["type"] == "bulleted"

    def test_numbered_segment_gets_list_formatting(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("md"):
                rng = doc.headings["Introduction"].insert_markdown("1. a\n2. b")
            assert rng.list_info()["type"] == "numbered"

    def test_empty_markdown_raises(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(OpError):
                doc.headings["Introduction"].insert_markdown("   \n\n")


class TestReplaceSectionBodyApi:
    def test_items_body_returns_range(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("rewrite"):
                rng = doc.headings["Introduction"].replace_section_body(["fresh body"])
            assert isinstance(rng, RangeAnchor)

    def test_markdown_body_returns_range(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("rewrite"):
                rng = doc.headings["Introduction"].replace_section_body(
                    "## Sub\n\n- a\n- b", markdown=True
                )
            assert isinstance(rng, RangeAnchor)

    def test_markdown_true_requires_string(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(OpError):
                doc.headings["Introduction"].replace_section_body(["x"], markdown=True)


# ---------------------------------------------------------------------------
# exec ops
# ---------------------------------------------------------------------------


class TestComposeOps:
    def test_insert_section_op(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(
                doc,
                [
                    {
                        "op": "insert_section",
                        "anchor_id": "heading:1",
                        "heading": "H",
                        "body": ["b1", "b2"],
                        "level": 2,
                    }
                ],
                label="s",
            )
        assert exc is None
        out = result["outputs"][0]
        assert out["op"] == "insert_section"
        assert out["anchor_id"].startswith("range:")

    def test_insert_markdown_op(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(
                doc,
                [{"op": "insert_markdown", "anchor_id": "heading:1", "markdown": "# H\n\n- a"}],
                label="m",
            )
        assert exc is None
        assert result["outputs"][0]["anchor_id"].startswith("range:")

    def test_replace_section_op_with_markdown(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(
                doc,
                [{"op": "replace_section", "anchor_id": "heading:1", "markdown": "new body"}],
                label="r",
            )
        assert exc is None
        assert result["outputs"][0]["anchor_id"].startswith("range:")

    def test_replace_section_needs_body_xor_markdown(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(
                doc,
                [
                    {
                        "op": "replace_section",
                        "anchor_id": "heading:1",
                        "body": ["a"],
                        "markdown": "b",
                    }
                ],
                label="bad",
            )
        assert exc is not None
        assert "exactly one of 'body' or 'markdown'" in result["failure"]["error"]

    def test_replace_section_rejects_non_heading_anchor(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            result, exc = run_batch(
                doc,
                [{"op": "replace_section", "anchor_id": "para:2", "body": ["x"]}],
                label="bad",
            )
        assert exc is not None
        assert "heading" in result["failure"]["error"]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


class TestComposeCli:
    def test_insert_section(self, fake_word: Any) -> None:
        body = json.dumps(["b1", "b2"])
        code, out, _ = _invoke(
            ["insert-section", "--anchor-id", "heading:1", "--heading", "H", "--body", body]
        )
        assert code == EXIT_OK
        assert json.loads(out)["anchor_id"].startswith("range:")

    def test_insert_markdown_from_stdin(self, fake_word: Any) -> None:
        code, out, _ = _invoke(
            ["insert-markdown", "--anchor-id", "end", "--markdown", "-"],
            input="# H\n\n- a\n- b",
        )
        assert code == EXIT_OK
        assert json.loads(out)["ok"] is True

    def test_replace_section_with_markdown(self, fake_word: Any) -> None:
        code, out, _ = _invoke(
            ["replace-section", "--anchor-id", "heading:1", "--markdown", "new body"]
        )
        assert code == EXIT_OK
        assert json.loads(out)["anchor_id"].startswith("range:")

    def test_replace_section_both_inputs_is_usage_error(self, fake_word: Any) -> None:
        code, _, err = _invoke(
            ["replace-section", "--anchor-id", "heading:1", "--body", "[]", "--markdown", "x"]
        )
        assert code != EXIT_OK
        assert "exactly one" in err

    def test_replace_section_non_heading_is_error(self, fake_word: Any) -> None:
        code, _, err = _invoke(["replace-section", "--anchor-id", "para:2", "--body", '["x"]'])
        assert code != EXIT_OK
        assert "heading" in err


# ---------------------------------------------------------------------------
# MCP
# ---------------------------------------------------------------------------


class TestComposeMcp:
    def test_insert_section(self, fake_word: Any) -> None:
        r = _write_impl(
            W, "insert_section", {"anchor_id": "heading:1", "heading": "H", "body": ["b1"]}
        )
        assert r["ok"] is True and r["command"] == "insert_section"

    def test_insert_markdown(self, fake_word: Any) -> None:
        r = _write_impl(W, "insert_markdown", {"anchor_id": "heading:1", "markdown": "# H\n- a"})
        assert r["ok"] is True and r["command"] == "insert_markdown"

    def test_replace_section(self, fake_word: Any) -> None:
        r = _write_impl(W, "replace_section", {"anchor_id": "heading:1", "markdown": "body"})
        assert r["ok"] is True and r["command"] == "replace_section"

    def test_replace_section_both_inputs_errors(self, fake_word: Any) -> None:
        with pytest.raises(Exception, match="exactly one of 'body' or 'markdown'"):
            _write_impl(
                W, "replace_section", {"anchor_id": "heading:1", "body": [], "markdown": "x"}
            )
