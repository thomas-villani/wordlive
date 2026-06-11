"""Keep the hand-maintained `exec` op enumerations honest.

The `exec` op vocabulary lives in `_ops.OP_REQUIRED_FIELDS`, but it is *also*
spelled out by hand in three LLM-facing places: the CLI skill's "Ops:" list, the
`exec` command's `--help` docstring, and the MCP `word_exec` docstring. Those
lists drift silently as ops are added (they did — 16 ops went undocumented in the
skill). These tests fail the moment an op exists in code but is missing from a
doc surface, so the surfaces stay in agreement (CLAUDE.md: "surfaces must agree").
"""

from __future__ import annotations

import re

from wordlive._guide import bundled_skill
from wordlive._ops import OP_REQUIRED_FIELDS


def _mentions(text: str, op: str) -> bool:
    """Whether `op` appears as a whole token (so `append` ≠ `append_inline`)."""
    return re.search(rf"\b{re.escape(op)}\b", text) is not None


def test_cli_skill_lists_every_exec_op():
    skill = bundled_skill("cli")
    missing = sorted(op for op in OP_REQUIRED_FIELDS if not _mentions(skill, op))
    assert not missing, f"exec ops absent from the CLI SKILL.md: {missing}"


def test_exec_help_docstring_lists_every_op():
    from wordlive.cli.main import main

    exec_cmd = main.commands["exec"]
    doc = exec_cmd.help or exec_cmd.callback.__doc__ or ""
    missing = sorted(op for op in OP_REQUIRED_FIELDS if not _mentions(doc, op))
    assert not missing, f"exec ops absent from the `exec` --help docstring: {missing}"
