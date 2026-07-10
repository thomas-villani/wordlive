"""Error copy for unknown commands, ops, actions, and anchor ids.

Three findings from an agent's trial run drive this file: a wrong command dumped
the whole ~100-name enum with no "did you mean"; `list` demanded an `action`
without naming one; and `table:2` reported `table cell not found` while the
guidance sat unread in a docstring.

The `TestVocabularyDoesNotDrift` class is the load-bearing part. The suggestion
copy is only as good as the name lists behind it, so those are pinned against the
dispatchers and the published JSON schema rather than trusted.
"""

from __future__ import annotations

import asyncio
from typing import Any

import pytest

import wordlive
from wordlive._ops import OP_REQUIRED_FIELDS, validate_op
from wordlive._suggest import did_you_mean, or_list, unknown_value_message
from wordlive.exceptions import AnchorNotFoundError, OpError
from wordlive.mcp._commands import (
    READ_COMMANDS,
    WRITE_ACTION_ALIASES,
    WRITE_ACTIONS,
    WRITE_COMMANDS,
)
from wordlive.mcp._read import _read_impl
from wordlive.mcp._worker import InlineWorker
from wordlive.mcp._write import _build_write_op, _write_impl

W = InlineWorker()

# Raised only when a name is declared in a vocabulary but has no dispatch branch.
_DRIFT = "declared but not implemented"


class _ExplodingWorker:
    """A worker that fails the test if anything tries to reach Word."""

    def run_on_word(self, fn: Any) -> Any:
        raise AssertionError("dispatch touched Word before validating the command")


# ---------------------------------------------------------------------------
# The matcher itself
# ---------------------------------------------------------------------------


class TestDidYouMean:
    def test_ranks_a_rare_shared_token_over_a_common_prefix(self) -> None:
        # `read` appears in five read commands and says little; `format` appears
        # in one. Plain difflib picks `read_image` here — the whole reason this
        # module exists rather than a bare `difflib.get_close_matches`.
        assert did_you_mean("read_format", READ_COMMANDS)[0] == "format_info"

    def test_containment_prefers_the_closest_length(self) -> None:
        # Both contain "paragraph"; the caller plainly meant the shorter one.
        assert did_you_mean("paragraph", READ_COMMANDS)[0] == "paragraphs"

    def test_ignores_case_and_separators(self) -> None:
        assert did_you_mean("FIND", READ_COMMANDS) == ["find"]
        assert did_you_mean("Read-Text", READ_COMMANDS) == ["read_text"]
        assert did_you_mean("list levels", READ_COMMANDS)[0] == "list_levels"

    @pytest.mark.parametrize("junk", ["xyzzy", "qqqq", "", "!!!", "12345"])
    def test_suggests_nothing_for_a_name_with_no_near_neighbour(self, junk: str) -> None:
        # Three confident-looking wrong answers are worse than none.
        assert did_you_mean(junk, READ_COMMANDS) == []

    def test_caps_the_shortlist(self) -> None:
        assert len(did_you_mean("insert_paragraph", WRITE_COMMANDS)) <= 3

    def test_or_list_reads_as_english(self) -> None:
        assert or_list(["a"]) == "'a'"
        assert or_list(["a", "b"]) == "'a' or 'b'"
        assert or_list(["a", "b", "c"]) == "'a', 'b', or 'c'"


class TestUnknownValueMessage:
    def test_small_vocabulary_is_listed_in_full(self) -> None:
        # With three options a "did you mean" is pointless — just name them.
        msg = unknown_value_message("comment action", "resolv", ("add", "resolve", "delete"))
        assert "'add', 'resolve', or 'delete'" in msg
        assert "did you mean" not in msg

    def test_a_lone_confident_suggestion_needs_no_escape_hatch(self) -> None:
        msg = unknown_value_message("read command", "markdown", READ_COMMANDS, fallback="see guide")
        assert msg == "unknown read command 'markdown'; did you mean 'to_markdown'?"

    def test_a_shortlist_keeps_the_pointer_to_the_full_vocabulary(self) -> None:
        msg = unknown_value_message("read command", "read_format", READ_COMMANDS, fallback="F")
        assert "did you mean 'format_info'" in msg
        assert "45 valid — F" in msg

    def test_no_suggestion_still_points_somewhere(self) -> None:
        msg = unknown_value_message("read command", "xyzzy", READ_COMMANDS, fallback="F")
        assert "45 valid read commands exist — F" in msg
        # Never dump the whole enum: that was the original complaint.
        assert "to_markdown" not in msg


# ---------------------------------------------------------------------------
# #13 — unknown command / op
# ---------------------------------------------------------------------------


class TestUnknownReadCommand:
    def test_suggests_and_never_attaches_to_word(self) -> None:
        with pytest.raises(OpError) as excinfo:
            _read_impl(_ExplodingWorker(), "read_format", {})
        assert "did you mean 'format_info'" in str(excinfo.value)

    def test_pydantic_no_longer_swallows_the_miss(self, fake_word: Any) -> None:
        # `command` is a plain `str` so the near-miss reaches our dispatcher
        # instead of dying in a pydantic `literal_error` that names all 45.
        pytest.importorskip("mcp")
        from mcp.shared.memory import create_connected_server_and_client_session

        from wordlive.mcp.server import build_server

        async def go() -> Any:
            server = build_server(InlineWorker())._mcp_server
            async with create_connected_server_and_client_session(server) as client:
                return await client.call_tool("word_read", {"command": "read_format"})

        result = asyncio.run(go())
        text = "\n".join(c.text for c in result.content if getattr(c, "text", None))
        assert result.isError is True
        assert "did you mean 'format_info'" in text
        assert "literal_error" not in text


class TestUnknownWriteCommand:
    def test_suggests_and_never_attaches_to_word(self) -> None:
        with pytest.raises(OpError) as excinfo:
            _write_impl(_ExplodingWorker(), "insert_paragraph", {})
        assert "did you mean" in str(excinfo.value)

    @pytest.mark.parametrize(
        ("attempted", "command", "action"),
        [
            ("add_row", "table", "add_row"),
            ("apply_list", "list", "apply"),
            ("add_comment", "comment", "add"),
            ("accept_all_revisions", "revision", "accept_all"),
        ],
    )
    def test_an_exec_op_name_is_answered_exactly_not_guessed(
        self, attempted: str, command: str, action: str
    ) -> None:
        # These are `exec` ops reached through a sub-dispatcher. No string metric
        # recovers `table` from `delete_row`, so the alias table answers directly.
        with pytest.raises(OpError) as excinfo:
            _write_impl(_ExplodingWorker(), attempted, {})
        message = str(excinfo.value)
        assert f"action of command {command!r}" in message
        assert f"word_write(command={command!r}, action={action!r})" in message

    @pytest.mark.parametrize(
        ("op_name", "command"),
        [
            ("insert_paragraph", "insert"),
            ("append_paragraph", "append"),
            ("append_inline", "append"),
            ("prepend_inline", "prepend"),
            ("find_replace", "replace"),
            ("insert_text_box", "text_box"),
            ("set_watermark", "watermark"),
            ("remove_watermark", "watermark"),
            ("set_cell_vertical_alignment", "cell_valign"),
            ("set_page_setup", "page_setup"),
            ("write_header", "header"),
            ("write_footer", "footer"),
        ],
    )
    def test_a_renamed_op_is_found_by_the_matcher_alone(self, op_name: str, command: str) -> None:
        # These 13 exec ops reach `word_write` under a different name. Unlike the
        # sub-dispatcher ops they need no alias entry: containment and token
        # overlap already rank the right command first. This test is why
        # WRITE_ACTION_ALIASES stays small — if the matcher regresses, add them.
        assert did_you_mean(op_name, WRITE_COMMANDS)[0] == command


class TestUnknownExecOp:
    def test_suggests_a_near_op(self) -> None:
        with pytest.raises(OpError) as excinfo:
            validate_op({"op": "insert_paragrph", "anchor_id": "end", "text": "x"})
        assert "did you mean 'insert_paragraph'" in str(excinfo.value)

    def test_nonsense_op_points_at_the_guide(self) -> None:
        with pytest.raises(OpError) as excinfo:
            validate_op({"op": "xyzzy"})
        message = str(excinfo.value)
        assert "llm-help" in message
        assert "did you mean" not in message


# ---------------------------------------------------------------------------
# #14 — a sub-dispatcher's `action`
# ---------------------------------------------------------------------------


class TestActionErrors:
    @pytest.mark.parametrize("command", sorted(WRITE_ACTIONS))
    def test_missing_action_enumerates_every_valid_action(self, command: str) -> None:
        with pytest.raises(OpError) as excinfo:
            _build_write_op(command, {})
        message = str(excinfo.value)
        assert f"command {command!r} requires 'action'" in message
        for action in WRITE_ACTIONS[command]:
            assert repr(action) in message

    def test_mistyped_action_in_a_large_set_suggests(self) -> None:
        with pytest.raises(OpError) as excinfo:
            _build_write_op("table", {"action": "add_rows"})
        assert "did you mean 'add_row'" in str(excinfo.value)

    def test_mistyped_action_in_a_small_set_lists_them_all(self) -> None:
        with pytest.raises(OpError) as excinfo:
            _build_write_op("list", {"action": "aply"})
        assert "did you mean 'apply'" in str(excinfo.value)


# ---------------------------------------------------------------------------
# #15 — `table:2` and other anchor-id misses
# ---------------------------------------------------------------------------


class TestAnchorIdHints:
    def test_bare_table_id_explains_the_three_forms_that_resolve(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(AnchorNotFoundError) as excinfo:
                doc.anchor_by_id("table:2")
        message = str(excinfo.value)
        assert "table:2:R:C" in message
        assert "table:2:row:R" in message
        assert "table:2:col:C" in message
        assert "collection, not a single anchor" in message

    def test_malformed_table_id_gets_the_generic_shape(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(AnchorNotFoundError) as excinfo:
                doc.anchor_by_id("table:a:b:c:d")
        assert "table:N:R:C" in str(excinfo.value)

    def test_misspelled_anchor_kind_is_suggested(self, fake_word: Any) -> None:
        with wordlive.attach() as word:
            doc = word.documents.active
            with pytest.raises(AnchorNotFoundError) as excinfo:
                doc.anchor_by_id("tabel:1:1:1")
        assert "did you mean 'table'" in str(excinfo.value)


# ---------------------------------------------------------------------------
# The vocabularies must match the code they describe
# ---------------------------------------------------------------------------


class TestVocabularyDoesNotDrift:
    def test_read_commands_match_the_published_schema(self, fake_word: Any) -> None:
        pytest.importorskip("mcp")
        assert _schema_enum("word_read") == list(READ_COMMANDS)

    def test_write_commands_match_the_published_schema(self, fake_word: Any) -> None:
        pytest.importorskip("mcp")
        assert _schema_enum("word_write") == list(WRITE_COMMANDS)

    @pytest.mark.parametrize("command", READ_COMMANDS)
    def test_every_read_command_has_a_dispatch_branch(self, command: str, fake_word: Any) -> None:
        # Called with no params, so most raise "requires 'x'". We only care that
        # none falls through to the declared-but-unimplemented backstop.
        try:
            _read_impl(W, command, {})
        except Exception as exc:  # noqa: BLE001 — any error but drift is fine here
            assert _DRIFT not in str(exc), f"{command} is in READ_COMMANDS with no branch"

    @pytest.mark.parametrize("command", WRITE_COMMANDS)
    def test_every_write_command_has_a_dispatch_branch(self, command: str) -> None:
        # `save` / `save_as` / `export_pdf` / `track` are terminal side-effects
        # handled by `_write_impl` itself, so they never reach `_build_write_op`.
        if command in ("save", "save_as", "export_pdf", "track"):
            return
        try:
            _build_write_op(command, {})
        except OpError as exc:
            assert _DRIFT not in str(exc), f"{command} is in WRITE_COMMANDS with no branch"

    @pytest.mark.parametrize("command", sorted(WRITE_ACTIONS))
    def test_every_declared_action_has_a_dispatch_branch(self, command: str) -> None:
        for action in WRITE_ACTIONS[command]:
            try:
                _build_write_op(command, {"action": action})
            except OpError as exc:
                assert _DRIFT not in str(exc), f"{command}/{action} is declared with no branch"

    def test_alias_targets_are_real_commands_and_actions(self) -> None:
        for op_name, (command, action) in WRITE_ACTION_ALIASES.items():
            assert command in WRITE_ACTIONS, f"{op_name} -> unknown command {command!r}"
            assert action in WRITE_ACTIONS[command], f"{op_name} -> unknown action {action!r}"

    def test_alias_keys_are_real_exec_ops(self) -> None:
        # The point of the alias table: these names are valid in `word_exec`, so
        # an agent reasonably tries them as `word_write` commands.
        for op_name in WRITE_ACTION_ALIASES:
            assert op_name in OP_REQUIRED_FIELDS, f"{op_name!r} is not an exec op"

    def test_alias_keys_never_shadow_a_real_write_command(self) -> None:
        # `set_style` and `set_borders` exist as both a top-level command and a
        # table action; the alias keys are the *op* names, so they must not collide.
        assert not (set(WRITE_ACTION_ALIASES) & set(WRITE_COMMANDS))

    def test_every_sub_dispatcher_action_round_trips_to_its_alias(self) -> None:
        # For each declared action, the op that `_build_write_op` emits must be
        # the alias key that maps back to it — pinning the table to the dispatcher.
        params: dict[str, Any] = {
            "anchor_id": "end", "levels": [], "text": "t", "index": 1, "table": 1,
            "row": 1, "col": 1, "column": 1, "cell": "1:1", "from": "1:1", "to": "2:2",
            "record": {}, "key": "k", "values": [], "style": "Normal", "alignment": "left",
            "rows": 1, "cols": 1,
        }  # fmt: skip
        for command, actions in WRITE_ACTIONS.items():
            for action in actions:
                op = _build_write_op(command, {"action": action, **params})
                assert WRITE_ACTION_ALIASES[op["op"]] == (command, action)


def _schema_enum(tool_name: str) -> list[str]:
    from mcp.shared.memory import create_connected_server_and_client_session

    from wordlive.mcp.server import build_server

    async def go() -> list[str]:
        server = build_server(InlineWorker())._mcp_server
        async with create_connected_server_and_client_session(server) as client:
            tools = {t.name: t for t in (await client.list_tools()).tools}
            enum = tools[tool_name].inputSchema["properties"]["command"]["enum"]
            return list(enum)

    return asyncio.run(go())
