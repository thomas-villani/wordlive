"""Access to the bundled agent guides (the `SKILL.md` files).

Shared by the CLI (`llm-help`, `install-skill`) and the MCP server (the
`wordlive://guide` resource and server instructions). Named `_guide` rather than
`_skill` because the package already ships a `_skill/` data directory holding the
Markdown files — a module of the same name can't coexist with it.

wordlive ships **three** skills: a CLI-facing guide, a Python-API guide, and an
MCP-facing guide (the one the server surfaces). Each lives in its own
`.agents/skills/<name>/` directory when installed, so the bundled layout mirrors
that:

    _skill/wordlive-cli/SKILL.md
    _skill/wordlive-python/SKILL.md
    _skill/wordlive-mcp/SKILL.md

The MCP server serves the `mcp` guide (not the CLI one): `word_read(command=
"guide")` and the `wordlive://guide` resource both return `skill_body("mcp")`, so
an agent driving the four `word_*` dispatch tools is taught in their own terms —
never CLI verbs it can't call.
"""

from __future__ import annotations

from importlib.resources import files

# kind -> installed skill directory name (also the `name:` in each frontmatter).
SKILLS: dict[str, str] = {
    "cli": "wordlive-cli",
    "python": "wordlive-python",
    "mcp": "wordlive-mcp",
}


def skill_name(kind: str = "cli") -> str:
    """The skill's canonical name / install directory (e.g. ``wordlive-cli``)."""
    try:
        return SKILLS[kind]
    except KeyError as e:
        raise ValueError(f"unknown skill kind {kind!r}; expected one of {sorted(SKILLS)}") from e


def bundled_skill(kind: str = "cli") -> str:
    """The packaged agent skill (SKILL.md) text, frontmatter and all."""
    name = skill_name(kind)
    return (files("wordlive") / "_skill" / name / "SKILL.md").read_text(encoding="utf-8")


def strip_frontmatter(md: str) -> str:
    """Drop a leading YAML frontmatter block (--- … ---), if present.

    Each bundled SKILL.md opens with `name:` / `description:` frontmatter for the
    agent-skill loader. That metadata is noise when the doc is read straight off
    stdout or served as a resource, so callers emit just the Markdown body.
    """
    lines = md.splitlines()
    if lines and lines[0].strip() == "---":
        for i in range(1, len(lines)):
            if lines[i].strip() == "---":
                return "\n".join(lines[i + 1 :]).lstrip("\n")
    return md


def skill_body(kind: str = "cli") -> str:
    """The bundled guide with its YAML frontmatter stripped."""
    return strip_frontmatter(bundled_skill(kind))
