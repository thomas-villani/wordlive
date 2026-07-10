"""Keep the hand-maintained `exec` op enumerations honest.

The `exec` op vocabulary lives in `_ops.OP_REQUIRED_FIELDS`, but it is *also*
spelled out by hand in four LLM-facing places: the CLI skill's "Ops:" list, the
MCP skill's `word_exec` op list, the `exec` command's `--help` docstring, and the
MCP `word_exec` tool docstring. Those lists drift silently as ops are added (they
did — 16 ops went undocumented in the skill). These tests fail the moment an op
exists in code but is missing from a doc surface, so the surfaces stay in
agreement (CLAUDE.md: "surfaces must agree").
"""

from __future__ import annotations

import re

from wordlive._guide import bundled_skill, skill_body
from wordlive._ops import OP_REQUIRED_FIELDS
from wordlive.mcp._commands import READ_COMMANDS, WRITE_ACTIONS, WRITE_COMMANDS


def _mentions(text: str, op: str) -> bool:
    """Whether `op` appears as a whole token (so `append` ≠ `append_inline`)."""
    return re.search(rf"\b{re.escape(op)}\b", text) is not None


def test_cli_skill_lists_every_exec_op():
    skill = bundled_skill("cli")
    missing = sorted(op for op in OP_REQUIRED_FIELDS if not _mentions(skill, op))
    assert not missing, f"exec ops absent from the CLI SKILL.md: {missing}"


def test_mcp_skill_lists_every_exec_op():
    skill = bundled_skill("mcp")
    missing = sorted(op for op in OP_REQUIRED_FIELDS if not _mentions(skill, op))
    assert not missing, f"exec ops absent from the MCP SKILL.md: {missing}"


def test_mcp_skill_names_every_write_dispatcher_and_action():
    """The MCP guide teaches word_write via `command` + `action`; every
    sub-dispatcher command and each of its actions must appear, or an agent can't
    learn the shape the tool actually dispatches on."""
    skill = bundled_skill("mcp")
    missing_cmds = sorted(c for c in WRITE_ACTIONS if not _mentions(skill, c))
    assert not missing_cmds, f"write sub-dispatchers absent from the MCP SKILL.md: {missing_cmds}"
    missing_actions = sorted(
        f"{cmd}.{action}"
        for cmd, actions in WRITE_ACTIONS.items()
        for action in actions
        if not _mentions(skill, action)
    )
    assert not missing_actions, f"write actions absent from the MCP SKILL.md: {missing_actions}"


def test_mcp_skill_crosswalk_names_are_real_commands():
    """Every backticked MCP command the crosswalk names (`word_read(command="X")`
    / `word_write(command="X")`) must be a real command — a stale crosswalk that
    points at a renamed command is worse than none."""
    skill = bundled_skill("mcp")
    read_refs = set(re.findall(r'word_read\(command="([a-z_]+)"', skill))
    write_refs = set(re.findall(r'word_write\(command="([a-z_]+)"', skill))
    bad_reads = sorted(c for c in read_refs if c not in READ_COMMANDS)
    bad_writes = sorted(c for c in write_refs if c not in WRITE_COMMANDS)
    assert not bad_reads, f"MCP guide names unknown read commands: {bad_reads}"
    assert not bad_writes, f"MCP guide names unknown write commands: {bad_writes}"


def test_mcp_guide_is_served_not_the_cli_one():
    """The server's guide (skill_body default for MCP surfaces) is the MCP skill —
    it must teach the word_* tools, not the CLI verbs."""
    body = skill_body("mcp")
    assert "word_read" in body and "word_write" in body and "word_exec" in body
    assert "wordlive outline" not in body


def test_exec_help_docstring_lists_every_op():
    from wordlive.cli.main import main

    exec_cmd = main.commands["exec"]
    doc = exec_cmd.help or exec_cmd.callback.__doc__ or ""
    missing = sorted(op for op in OP_REQUIRED_FIELDS if not _mentions(doc, op))
    assert not missing, f"exec ops absent from the `exec` --help docstring: {missing}"
