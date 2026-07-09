"""exec, guide, and install commands."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import click

from ... import attach
from ..._guide import bundled_skill as _bundled_skill
from ..._guide import skill_body as _skill_body
from ..._guide import skill_name as _skill_name
from ..._ops import pick_doc as _pick_doc
from ..._ops import run_batch as _run_batch
from ..main import _run, emit
from ._common import (
    _claude_desktop_config_path,
    _mcp_server_entry,
)


@click.command(name="exec")
@click.option(
    "--script",
    "script",
    default=None,
    type=click.Path(exists=True, dir_okay=False, path_type=Path),
    help="Path to an ops JSON file.",
)
@click.option(
    "--ops",
    "ops_inline",
    default=None,
    help="Inline JSON batch — the same content as a --script file, passed "
    'directly. Accepts the full {"label", "ops", …} object or a bare '
    "[…] ops array, or '-' to read JSON from stdin. Alternative to --script.",
)
@click.pass_context
def exec_(ctx: click.Context, script: Path | None, ops_inline: str | None) -> None:
    """Apply a batch of ops in a single atomic-undo scope.

    Provide the batch either as a file (`--script ops.json`) or inline
    (`--ops '{"ops": [...]}'`, or `--ops -` to read JSON from stdin — best for
    large payloads such as base64 images, which can exceed the command-line
    length limit). The script shape is
    `{"label": "…", "ops": [{"op": "...", ...}, ...]}`; a bare `[...]` array is
    accepted as shorthand for `{"ops": [...]}`. Set `"tracked": true` at the top
    level to record the whole batch as Word revisions (Track Changes is restored
    to its prior state afterwards).
    Supported ops: write_bookmark, write_cc, insert_paragraph, insert_block,
    insert_section, insert_markdown, replace_section,
    delete_paragraph, append, append_inline, prepend, prepend_inline,
    insert_image, insert_equation, insert_chart, format_chart, format_axis, add_trendline,
    set_series_color, format_series, add_error_bars,
    set_shape_wrap, set_shape_crop, set_shape_position, set_shape_size,
    format_shape, set_shape_alt_text, set_shape_text, set_shape_rotation, set_shape_z_order,
    set_shape_text_frame, replace_shape_image, delete_shape, group_shapes, ungroup_shape,
    set_image_alt_text, set_image_size, set_image_crop,
    replace, find_replace, apply_style, format_paragraph,
    format_run, set_shading, set_borders, drop_cap, add_tab_stop, add_style, set_style,
    insert_field, set_page_setup, update_fields, regularize, insert_footnote, insert_endnote,
    insert_toc, add_bookmark, pin, pin_outline, add_hyperlink, set_hyperlink,
    insert_cross_reference,
    insert_caption, create_content_control, set_cc_properties, set_cc_items,
    mark_index_entry, insert_index,
    insert_table_of_figures, set_bibliography_style, add_source, insert_citation,
    insert_bibliography, mark_citation, insert_table_of_authorities,
    apply_theme, set_theme_colors, set_theme_fonts,
    set_cell, add_row, append_record, update_row, delete_row,
    add_column, delete_column, merge_cells, split_cell,
    set_heading_row, autofit_table, create_table, delete_table,
    set_table_style, set_table_alignment, set_table_borders, set_table_banding,
    set_cell_vertical_alignment,
    set_property, delete_property, set_variable, delete_variable,
    insert_break, add_comment,
    resolve_comment, delete_comment,
    accept_revision, reject_revision, accept_all_revisions, reject_all_revisions,
    set_watermark, remove_watermark, insert_text_box,
    apply_list, apply_list_format, remove_list, restart_numbering, indent_list, outdent_list,
    write_header, write_footer. (append/prepend add a new paragraph + optional
    style; append_inline/prepend_inline continue the adjacent paragraph, text
    only. append_paragraph/prepend_paragraph remain as synonyms.) A field an op
    doesn't use is reported in the result's `warnings`, not silently dropped.

    Durable handles: `bind: "slug"` (or `true`) on insert/insert_block/
    insert_section/insert_markdown/create_table mints a `pin:` handle on the new
    content and returns it in that op's `outputs` entry. An op field of the exact
    form `$ops[N].field` is replaced with an earlier op's output before the op
    runs — e.g. create a table at op 0, then `{"op": "set_cell", "table":
    "$ops[0].table", ...}`.
    See docs/cli.md for each op's required and optional fields.
    """
    if (script is None) == (ops_inline is None):
        raise click.UsageError("provide exactly one of --script or --ops")

    def go() -> None:
        if ops_inline is not None:
            raw = click.get_text_stream("stdin").read() if ops_inline == "-" else ops_inline
            source = "inline"
        else:
            assert script is not None  # guaranteed by the validation above
            raw = script.read_text(encoding="utf-8")
            source = script.name
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.ClickException(f"ops JSON is malformed: {e}") from e
        if isinstance(payload, list):
            payload = {"ops": payload}
        if not isinstance(payload, dict):
            raise click.ClickException(
                'ops JSON must be an object {"ops": [...]} or an array of ops'
            )
        label = str(payload.get("label") or f"CLI: exec {source}")
        tracked = bool(payload.get("tracked", False))
        ops = payload.get("ops") or []
        if not isinstance(ops, list):
            raise click.ClickException("'ops' must be a list")
        # Vet image-source paths before any COM/filesystem access (a UNC path's
        # own existence probe would authenticate to a remote SMB server).
        ctx.obj["policy"].screen_op_image_paths(ops)

        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            result, failure_exc = _run_batch(doc, ops, label=label, tracked=tracked)
            if failure_exc is None:
                emit(
                    result,
                    as_text=not ctx.obj["as_json"],
                    text=f"applied {result['ops_run']} op(s): {label!r}",
                )
            else:
                failure = result["failure"]
                emit(
                    result,
                    as_text=not ctx.obj["as_json"],
                    text=f"failed at op {failure['index']}: {failure['error']}",
                )
                # Re-raise the original so _run() maps it to the right exit code
                # (e.g. anchor-not-found → 2, busy → 3, ambiguous → 5).
                raise failure_exc

    _run(ctx, go)


@click.command(name="llm-help")
@click.option(
    "--python",
    "python",
    is_flag=True,
    default=False,
    help="Print the Python-API guide instead of the CLI guide.",
)
def llm_help_cmd(python: bool) -> None:
    """Print the full wordlive agent guide (the bundled skill) to stdout.

    One-shot orientation for an LLM: the anchor model, every verb, image
    insertion, the `exec` batch format, and the exit-code taxonomy. `wordlive
    --help` points here. Defaults to the CLI guide; `--python` prints the
    Python-API (`import wordlive as wl`) guide instead. Output is raw Markdown —
    not JSON, and unaffected by `--json/--text` — so it reads cleanly straight
    into a model's context, exactly like `--help`. Offline: never touches Word.
    """
    kind = "python" if python else "cli"
    try:
        click.echo(_skill_body(kind))
    except (FileNotFoundError, ModuleNotFoundError, OSError, ValueError) as e:
        raise click.ClickException(f"could not read the bundled skill: {e}") from e


@click.command(name="install-skill")
@click.option(
    "--cli", "cli", is_flag=True, default=False, help="Install the CLI skill (the default)."
)
@click.option(
    "--python", "python", is_flag=True, default=False, help="Install only the Python-API skill."
)
@click.option(
    "--both",
    "both",
    is_flag=True,
    default=False,
    help="Install both the CLI and Python-API skills.",
)
@click.option(
    "--system",
    "system",
    is_flag=True,
    default=False,
    help="Install to ~/.agents/skills/ instead of the current project's ./.agents/skills/.",
)
@click.option(
    "--force", "force", is_flag=True, default=False, help="Overwrite an existing SKILL.md."
)
@click.pass_context
def install_skill_cmd(
    ctx: click.Context, cli: bool, python: bool, both: bool, system: bool, force: bool
) -> None:
    """Install wordlive's agent skill(s) (SKILL.md) for LLM coding tools.

    wordlive ships two skills — `wordlive-cli` (the command-line workflow) and
    `wordlive-python` (the `import wordlive as wl` API). By default only the
    **CLI** skill is installed; pass `--python` for just the Python one, or
    `--both` for both. They land under `.agents/skills/<name>/SKILL.md` in the
    current directory (default) or your home directory (`--system`). Offline —
    this doesn't touch Word.
    """
    if both or (cli and python):
        kinds = ["cli", "python"]
    elif python:
        kinds = ["python"]
    else:
        kinds = ["cli"]

    base = Path.home() if system else Path.cwd()
    scope = "system" if system else "local"
    dests = [(kind, base / ".agents" / "skills" / _skill_name(kind) / "SKILL.md") for kind in kinds]

    # Check every target up front so we never half-write when --force is absent.
    clashes = [str(dest) for _, dest in dests if dest.exists()]
    if clashes and not force:
        raise click.ClickException(
            "already exists (pass --force to overwrite): " + ", ".join(clashes)
        )

    installed = []
    try:
        for kind, dest in dests:
            content = _bundled_skill(kind)
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_text(content, encoding="utf-8")
            installed.append(
                {
                    "kind": kind,
                    "name": _skill_name(kind),
                    "path": str(dest),
                    "bytes": len(content.encode("utf-8")),
                }
            )
    except (FileNotFoundError, ModuleNotFoundError, OSError, ValueError) as e:
        raise click.ClickException(f"could not install the skill: {e}") from e

    emit(
        {"ok": True, "scope": scope, "installed": installed},
        as_text=not ctx.obj["as_json"],
        text="installed:\n" + "\n".join(f"  {r['name']} → {r['path']}" for r in installed),
    )


@click.command(name="install-mcp")
@click.option(
    "--client",
    type=click.Choice(["claude-desktop", "claude-code"]),
    default="claude-desktop",
    help="Which MCP client's config to write (default: claude-desktop).",
)
@click.option(
    "--name", "server_name", default="wordlive", help="Server key to register (default: wordlive)."
)
@click.option(
    "--directory",
    "directory",
    default=None,
    help="Register a local checkout via `uv run --directory DIR` (dev), instead of the default `uvx --from wordlive[mcp,snapshot]`.",
)
@click.option(
    "--config",
    "config_path",
    default=None,
    type=click.Path(),
    help="Write to this config file instead of the client's default location.",
)
@click.option(
    "--print",
    "print_only",
    is_flag=True,
    default=False,
    help="Print the JSON server snippet to stdout instead of writing any file.",
)
@click.option(
    "--force", "force", is_flag=True, default=False, help="Overwrite an existing server entry."
)
@click.pass_context
def install_mcp_cmd(
    ctx: click.Context,
    client: str,
    server_name: str,
    directory: str | None,
    config_path: str | None,
    print_only: bool,
    force: bool,
) -> None:
    """Register the wordlive MCP server in an agent's config.

    Merges an `mcpServers.<name>` entry into Claude Desktop's
    `claude_desktop_config.json` (default) or a Claude Code `.mcp.json`
    (`--client claude-code`, project-local). The entry launches the stdio server
    with `uvx --from "wordlive[mcp,snapshot]" wordlive-mcp` (no separate install
    needed), or `uv run --directory DIR wordlive-mcp` for a local checkout. Use
    `--print` to just emit the snippet for any client. Offline — never touches
    Word; restart the client to pick up the change.
    """
    entry = _mcp_server_entry(directory)

    if print_only:
        emit(
            {"ok": True, "server": server_name, "entry": entry, "mcpServers": {server_name: entry}},
            as_text=not ctx.obj["as_json"],
            text=json.dumps({"mcpServers": {server_name: entry}}, indent=2),
        )
        return

    if config_path is not None:
        target = Path(config_path)
    elif client == "claude-desktop":
        target = _claude_desktop_config_path()
    else:  # claude-code: portable, project-local server file
        target = Path.cwd() / ".mcp.json"

    cfg: dict[str, Any] = {}
    if target.exists():
        try:
            raw = target.read_text(encoding="utf-8").strip()
            cfg = json.loads(raw) if raw else {}
        except (OSError, json.JSONDecodeError) as e:
            raise click.ClickException(f"could not read existing config {target}: {e}") from e
        if not isinstance(cfg, dict):
            raise click.ClickException(f"existing config {target} is not a JSON object")

    servers = cfg.setdefault("mcpServers", {})
    if not isinstance(servers, dict):
        raise click.ClickException(f"'mcpServers' in {target} is not a JSON object")
    action = "updated" if server_name in servers else "created"
    if server_name in servers and not force:
        raise click.ClickException(
            f"server '{server_name}' is already in {target}; pass --force to overwrite"
        )
    servers[server_name] = entry

    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(cfg, indent=2) + "\n", encoding="utf-8")
    except OSError as e:
        raise click.ClickException(f"could not write {target}: {e}") from e

    emit(
        {
            "ok": True,
            "client": client,
            "path": str(target),
            "server": server_name,
            "action": action,
            "entry": entry,
        },
        as_text=not ctx.obj["as_json"],
        text=f"{action} server '{server_name}' → {target}\n(restart {client} to load it)",
    )
