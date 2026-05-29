"""Access to the bundled agent guide (SKILL.md).

Shared by the CLI (`llm-help`, `install-skill`) and the MCP server (the
`wordlive://guide` resource and server instructions). Named `_guide` rather than
`_skill` because the package already ships a `_skill/` data directory holding the
Markdown file — a module of the same name can't coexist with it.
"""

from __future__ import annotations

from importlib.resources import files


def bundled_skill() -> str:
    """The packaged agent skill (SKILL.md) text, frontmatter and all."""
    return (files("wordlive") / "_skill" / "SKILL.md").read_text(encoding="utf-8")


def strip_frontmatter(md: str) -> str:
    """Drop a leading YAML frontmatter block (--- … ---), if present.

    The bundled SKILL.md opens with `name:` / `description:` frontmatter for the
    agent-skill loader. That metadata is noise when the doc is read straight off
    stdout or served as a resource, so callers emit just the Markdown body.
    """
    lines = md.splitlines()
    if lines and lines[0].strip() == "---":
        for i in range(1, len(lines)):
            if lines[i].strip() == "---":
                return "\n".join(lines[i + 1 :]).lstrip("\n")
    return md


def skill_body() -> str:
    """The bundled guide with its YAML frontmatter stripped."""
    return strip_frontmatter(bundled_skill())
