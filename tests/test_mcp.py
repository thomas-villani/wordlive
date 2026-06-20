"""Tests for the MCP server (`wordlive.mcp`).

Skipped cleanly when the optional `mcp` extra isn't installed. The Word-touching
tools are driven through the existing `fake_word` fixture (which patches
`_com.get_active_word`), so no real Word is needed; an `InlineWorker` runs tool
bodies on the test thread, sidestepping the COM worker thread.
"""

from __future__ import annotations

import asyncio
import json
import threading
from typing import Any

import pytest

pytest.importorskip("mcp")

from mcp.shared.memory import (  # noqa: E402
    create_connected_server_and_client_session,
)

from wordlive.exceptions import (  # noqa: E402
    AnchorNotFoundError,
    OpError,
    WordBusyError,
    WordNotRunningError,
)
from wordlive.mcp._worker import ComWorker, InlineWorker  # noqa: E402
from wordlive.mcp.server import (  # noqa: E402
    _error_payload,
    _exec_impl,
    _image_format,
    _read_impl,
    _snapshot_impl,
    _write_impl,
    build_server,
)

W = InlineWorker()


def _texts(result: Any) -> str:
    # Match on the text attribute rather than isinstance: depending on import
    # paths the SDK's TextContent class object may not be identical to the one
    # imported here, which would make isinstance silently miss every block.
    return "\n".join(c.text for c in result.content if getattr(c, "text", None) is not None)


def _call(coro_factory: Any) -> Any:
    return asyncio.run(coro_factory())


# ---------------------------------------------------------------------------
# Worker
# ---------------------------------------------------------------------------


class TestWorker:
    def test_inline_runs_on_caller_thread(self) -> None:
        assert InlineWorker().run_on_word(lambda: 7) == 7

    def test_com_worker_serialises_on_one_thread(self) -> None:
        worker = ComWorker()
        try:
            t1 = worker.run_on_word(threading.get_ident)
            t2 = worker.run_on_word(threading.get_ident)
            assert t1 == t2  # same dedicated thread every call
            assert t1 != threading.get_ident()  # not the caller's thread
        finally:
            worker.shutdown()

    def test_com_worker_marshals_exceptions(self) -> None:
        worker = ComWorker()

        def boom() -> None:
            raise ValueError("kaboom")

        try:
            with pytest.raises(ValueError, match="kaboom"):
                worker.run_on_word(boom)
        finally:
            worker.shutdown()


# ---------------------------------------------------------------------------
# word_read
# ---------------------------------------------------------------------------


class TestReadImpl:
    def test_status(self, fake_word: Any) -> None:
        rows = _read_impl(W, "status", {})
        assert rows[0]["name"] == "Test.docx"
        assert rows[0]["is_active"] is True
        # Always a usable identifier + a saved flag so the agent can confirm
        # its target before writing.
        assert rows[0]["name"]
        assert "saved" in rows[0]

    def test_guide_needs_no_word(self, no_word: Any) -> None:
        # The guide is fetchable as a tool call even when Word isn't running —
        # it never touches COM.
        out = _read_impl(W, "guide", {})
        assert "guide" in out and "anchor" in out["guide"].lower()

    def test_outline(self, fake_word: Any) -> None:
        items = _read_impl(W, "outline", {})
        assert [i["text"] for i in items] == ["Introduction", "Risks"]

    def test_outline_all_paragraphs(self, fake_word: Any) -> None:
        items = _read_impl(W, "outline", {"all_paragraphs": True})
        assert len(items) == 3  # every paragraph, not just headings

    def test_paragraphs_window(self, fake_word: Any) -> None:
        items = _read_impl(W, "paragraphs", {"start": 2, "count": 1})
        assert len(items) == 1
        assert items[0]["text"] == "Body text here."

    def test_find(self, fake_word: Any) -> None:
        matches = _read_impl(W, "find", {"text": "Body"})
        assert matches and matches[0]["anchor_id"].startswith("range:")

    def test_read_bookmark(self, fake_word: Any) -> None:
        assert "text" in _read_impl(W, "read_bookmark", {"name": "Address"})

    def test_read_cc(self, fake_word: Any) -> None:
        assert _read_impl(W, "read_cc", {"name": "Signatory"}) == {"text": "Jane Doe"}

    def test_read_section_by_heading(self, fake_word: Any) -> None:
        out = _read_impl(W, "read_section", {"heading": "Introduction"})
        assert out["anchor_id"] == "heading:1"

    def test_collections(self, fake_word: Any) -> None:
        assert isinstance(_read_impl(W, "styles", {}), list)
        assert isinstance(_read_impl(W, "comments", {}), list)
        assert isinstance(_read_impl(W, "sections", {}), list)
        assert isinstance(_read_impl(W, "table_list", {}), list)

    def test_revisions(self, fake_word: Any) -> None:
        rows = _read_impl(W, "revisions", {})
        assert isinstance(rows, list)
        assert rows[0]["type"] == "insert" and rows[0]["author"] == "Reviewer"

    def test_track_status(self, fake_word: Any) -> None:
        assert _read_impl(W, "track", {}) == {"track_changes": False}

    def test_images_list(self, fake_word: Any) -> None:
        rows = _read_impl(W, "images", {})
        assert rows[0]["anchor_id"] == "image:1"
        assert rows[0]["mime"] == "image/png"

    def test_read_image_base64(self, fake_word: Any) -> None:
        import base64

        out = _read_impl(W, "read_image", {"anchor_id": "image:1"})
        assert out["mime"] == "image/png"
        assert base64.b64decode(out["base64"]) == b"\x89PNG\r\n\x1a\nSEEDED"

    def test_read_image_requires_anchor_id(self, fake_word: Any) -> None:
        with pytest.raises(OpError):
            _read_impl(W, "read_image", {})

    def test_image_format_from_mime(self) -> None:
        assert _image_format("image/png") == "png"
        assert _image_format("image/jpeg") == "jpeg"
        assert _image_format(None) == "png"  # fallback
        assert _image_format("image/") == "png"  # empty subtype → fallback

    def test_missing_bookmark_raises(self, fake_word: Any) -> None:
        with pytest.raises(AnchorNotFoundError):
            _read_impl(W, "read_bookmark", {"name": "DoesNotExist"})

    def test_unknown_command_raises_op_error(self, fake_word: Any) -> None:
        with pytest.raises(OpError):
            _read_impl(W, "nonsense", {})

    def test_find_without_text_raises_op_error(self, fake_word: Any) -> None:
        with pytest.raises(OpError):
            _read_impl(W, "find", {})


# ---------------------------------------------------------------------------
# word_write
# ---------------------------------------------------------------------------


class TestWriteImpl:
    def test_write_bookmark(self, fake_word: Any) -> None:
        r = _write_impl(W, "write_bookmark", {"name": "Address", "text": "123 Main"})
        assert r["ok"] is True and r["command"] == "write_bookmark"

    def test_insert(self, fake_word: Any) -> None:
        r = _write_impl(W, "insert", {"anchor_id": "heading:1", "text": "new para"})
        assert r["ok"] is True

    def test_delete_paragraph(self, fake_word: Any) -> None:
        r = _write_impl(W, "delete_paragraph", {"anchor_id": "para:1"})
        assert r["ok"] is True and r["command"] == "delete_paragraph"
        fake_word.ActiveDocument.Range(0, 13).Delete.assert_called_once()

    def test_replace_find(self, fake_word: Any) -> None:
        r = _write_impl(W, "replace", {"find": "Body", "text": "Corpus"})
        assert r["ok"] is True

    def test_append_defaults_to_new_paragraph(self, fake_word: Any) -> None:
        r = _write_impl(W, "append", {"text": "Tail."})
        assert r["ok"] is True
        assert fake_word.ActiveDocument.Range(34, 34).Text == "\rTail."

    def test_append_paragraph_false_is_inline(self, fake_word: Any) -> None:
        r = _write_impl(W, "append", {"text": " more", "paragraph": False})
        assert r["ok"] is True
        fake_word.ActiveDocument.Content.InsertAfter.assert_called_once_with(" more")

    def test_track(self, fake_word: Any) -> None:
        r = _write_impl(W, "track", {"on": True})
        assert r["track_changes"] is True

    def test_insert_break(self, fake_word: Any) -> None:
        r = _write_impl(
            W, "insert_break", {"anchor_id": "bookmark:Address", "kind": "section_next"}
        )
        assert r["ok"] is True and r["command"] == "insert_break"
        # Address ends at 24; wdSectionBreakNextPage = 2.
        fake_word.ActiveDocument.Range(24, 24).InsertBreak.assert_called_once_with(Type=2)

    def test_insert_break_defaults_to_page(self, fake_word: Any) -> None:
        r = _write_impl(W, "insert_break", {"anchor_id": "bookmark:Address"})
        assert r["ok"] is True
        fake_word.ActiveDocument.Range(24, 24).InsertBreak.assert_called_once_with(Type=7)

    def test_format_paragraph_page_break_before(self, fake_word: Any) -> None:
        r = _write_impl(
            W, "format_paragraph", {"anchor_id": "bookmark:Address", "page_break_before": True}
        )
        assert r["ok"] is True
        pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
        assert pf.PageBreakBefore is True

    def test_missing_required_field_raises(self, fake_word: Any) -> None:
        with pytest.raises(OpError):
            _write_impl(W, "write_bookmark", {"name": "Address"})  # no text

    def test_unknown_command_raises(self, fake_word: Any) -> None:
        with pytest.raises(OpError):
            _write_impl(W, "nope", {})

    def test_missing_anchor_raises_anchor_error(self, fake_word: Any) -> None:
        with pytest.raises(AnchorNotFoundError):
            _write_impl(W, "insert", {"anchor_id": "heading:99", "text": "x"})


# ---------------------------------------------------------------------------
# word_exec
# ---------------------------------------------------------------------------


class TestExecImpl:
    def test_batch_success(self, fake_word: Any) -> None:
        ops = [
            {"op": "write_bookmark", "name": "Address", "text": "A"},
            {"op": "insert_paragraph", "anchor_id": "heading:1", "text": "B"},
        ]
        result, exc = _exec_impl(W, ops, doc=None, label="t", tracked=False)
        assert exc is None
        assert result["ok"] is True and result["ops_run"] == 2

    def test_batch_warns_on_ignored_field(self, fake_word: Any) -> None:
        ops = [{"op": "append_inline", "text": "x", "style": "Heading 1"}]
        result, exc = _exec_impl(W, ops, doc=None, label="t", tracked=False)
        assert exc is None and result["ok"] is True
        assert any(w["field"] == "style" for w in result.get("warnings", []))

    def test_batch_failure_reports_first_bad_op(self, fake_word: Any) -> None:
        ops = [
            {"op": "write_bookmark", "name": "Address", "text": "A"},
            {"op": "write_bookmark", "name": "Nope", "text": "B"},
        ]
        result, exc = _exec_impl(W, ops, doc=None, label="t", tracked=False)
        assert isinstance(exc, AnchorNotFoundError)
        assert result["ok"] is False
        assert result["ops_run"] == 1
        assert result["failure"]["index"] == 1
        assert result["failure"]["type"] == "AnchorNotFoundError"


# ---------------------------------------------------------------------------
# Error mapping
# ---------------------------------------------------------------------------


class TestErrorPayload:
    def test_anchor_not_found(self) -> None:
        p = _error_payload(AnchorNotFoundError("bookmark", "X"))
        assert p["code"] == "anchor_not_found" and p["retryable"] is False

    def test_word_busy_is_retryable(self) -> None:
        p = _error_payload(WordBusyError())
        assert p["code"] == "word_busy" and p["retryable"] is True

    def test_word_not_running(self) -> None:
        p = _error_payload(WordNotRunningError("no Word"))
        assert p["code"] == "word_not_running" and p["retryable"] is False


# ---------------------------------------------------------------------------
# word_snapshot (Document.snapshot stubbed — render internals not exercised here)
# ---------------------------------------------------------------------------


class _FakeSnap:
    def __init__(self, page: int, png: bytes) -> None:
        self.page = page
        self.png = png
        self.path = None


class TestSnapshotImpl:
    def test_returns_page_png_pairs(self, fake_word: Any, monkeypatch: pytest.MonkeyPatch) -> None:
        from wordlive._document import Document

        monkeypatch.setattr(
            Document,
            "snapshot",
            lambda self, pages=None, dpi=150, max_dim=None, markup="none": [
                _FakeSnap(3, b"\x89PNGdata")
            ],
        )
        pairs = _snapshot_impl(W, doc=None, pages="3", anchor=None, dpi=150, markup="none")
        assert pairs == [(3, b"\x89PNGdata")]

    def test_markup_threads_through(self, fake_word: Any, monkeypatch: pytest.MonkeyPatch) -> None:
        from wordlive._document import Document

        seen: dict[str, Any] = {}

        def fake_snapshot(self, pages=None, dpi=150, max_dim=None, markup="none"):
            seen["markup"] = markup
            return [_FakeSnap(1, b"\x89PNGx")]

        monkeypatch.setattr(Document, "snapshot", fake_snapshot)
        _snapshot_impl(W, doc=None, pages="1", anchor=None, dpi=150, markup="all")
        assert seen["markup"] == "all"

    def test_max_dim_threads_through(self, fake_word: Any, monkeypatch: pytest.MonkeyPatch) -> None:
        from wordlive._document import Document

        seen: dict[str, Any] = {}

        def fake_snapshot(self, pages=None, dpi=150, max_dim=None, markup="none"):
            seen["max_dim"] = max_dim
            return [_FakeSnap(1, b"\x89PNGx")]

        monkeypatch.setattr(Document, "snapshot", fake_snapshot)
        _snapshot_impl(W, doc=None, pages="1", anchor=None, dpi=150, markup="none", max_dim=1000)
        assert seen["max_dim"] == 1000


# ---------------------------------------------------------------------------
# In-memory client/server session (end-to-end through FastMCP)
# ---------------------------------------------------------------------------


class TestSession:
    def test_lists_exactly_the_four_tools(self, fake_word: Any) -> None:
        async def go() -> set[str]:
            server = build_server(InlineWorker())._mcp_server
            async with create_connected_server_and_client_session(server) as client:
                tools = await client.list_tools()
                return {t.name for t in tools.tools}

        assert _call(go) == {"word_read", "word_write", "word_exec", "word_snapshot"}

    def test_read_outline_roundtrip(self, fake_word: Any) -> None:
        async def go() -> Any:
            server = build_server(InlineWorker())._mcp_server
            async with create_connected_server_and_client_session(server) as client:
                return await client.call_tool("word_read", {"command": "outline"})

        result = _call(go)
        assert result.isError is False
        assert "Introduction" in _texts(result)

    def test_error_call_is_flagged_and_coded(self, fake_word: Any) -> None:
        async def go() -> Any:
            server = build_server(InlineWorker())._mcp_server
            async with create_connected_server_and_client_session(server) as client:
                return await client.call_tool(
                    "word_read", {"command": "read_bookmark", "name": "DoesNotExist"}
                )

        result = _call(go)
        assert result.isError is True
        assert "anchor_not_found" in _texts(result)

    def test_guide_resource(self, fake_word: Any) -> None:
        async def go() -> str:
            server = build_server(InlineWorker())._mcp_server
            async with create_connected_server_and_client_session(server) as client:
                res = await client.read_resource("wordlive://guide")
                return "\n".join(getattr(c, "text", "") for c in res.contents)

        body = _call(go)
        assert "anchor" in body.lower()

    def test_snapshot_returns_image_content(
        self, fake_word: Any, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        from wordlive._document import Document

        monkeypatch.setattr(
            Document,
            "snapshot",
            lambda self, pages=None, dpi=150, max_dim=None, markup="none": [
                _FakeSnap(1, b"\x89PNGdata")
            ],
        )

        async def go() -> Any:
            server = build_server(InlineWorker())._mcp_server
            async with create_connected_server_and_client_session(server) as client:
                return await client.call_tool("word_snapshot", {})

        result = _call(go)
        assert result.isError is False
        assert any(getattr(c, "type", None) == "image" for c in result.content)
        # The image rides in `content` ONLY. structured_output=False keeps FastMCP
        # from inferring an output schema off `-> list[Any]` and re-serialising the
        # base64 PNG into structuredContent — which would send every image twice.
        assert result.structuredContent is None

    def test_read_image_returns_image_content(self, fake_word: Any) -> None:
        async def go() -> Any:
            server = build_server(InlineWorker())._mcp_server
            async with create_connected_server_and_client_session(server) as client:
                return await client.call_tool(
                    "word_read", {"command": "read_image", "anchor_id": "image:1"}
                )

        result = _call(go)
        assert result.isError is False
        # The picture rides back as a real image block (so a vision model SEES it),
        # alongside a compact {anchor_id,mime,bytes} text label — not base64 text.
        assert any(getattr(c, "type", None) == "image" for c in result.content)
        assert "image:1" in _texts(result)
        assert "base64" not in _texts(result)

    def test_snapshot_tool_has_no_output_schema(self, fake_word: Any) -> None:
        # Guards the structured_output=False on word_snapshot: an inferred schema
        # is what double-encodes the image into structuredContent.
        async def go() -> Any:
            server = build_server(InlineWorker())._mcp_server
            async with create_connected_server_and_client_session(server) as client:
                tools = {t.name: t for t in (await client.list_tools()).tools}
                return tools["word_snapshot"].outputSchema

        assert _call(go) is None


def test_json_error_payload_is_parseable(fake_word: Any) -> None:
    """The ToolError message is a JSON object the model can parse."""

    async def go() -> Any:
        server = build_server(InlineWorker())._mcp_server
        async with create_connected_server_and_client_session(server) as client:
            return await client.call_tool(
                "word_write", {"command": "insert", "anchor_id": "heading:99", "text": "x"}
            )

    result = _call(go)
    assert result.isError is True
    text = _texts(result)
    # The payload is embedded as JSON somewhere in the error text.
    start = text.find("{")
    payload = json.loads(text[start:])
    assert payload["code"] == "anchor_not_found"
    assert payload["retryable"] is False
