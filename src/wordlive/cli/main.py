"""Click entry point for `wordlive`. JSON in, JSON out, deterministic exit codes."""

from __future__ import annotations

import json
import sys
from typing import Any

import click

from .._paths import PathPolicy
from ..exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    DocumentNotFoundError,
    WordBusyError,
    WordliveError,
    WordNotRunningError,
)

# Exit codes per spec.md §"CLI Sketch":
EXIT_OK = 0
EXIT_OTHER = 1
EXIT_ANCHOR_NOT_FOUND = 2
EXIT_WORD_BUSY = 3
EXIT_WORD_NOT_RUNNING = 4
EXIT_AMBIGUOUS_MATCH = 5


def emit(payload: Any, *, as_text: bool = False, text: str | None = None) -> None:
    """One JSON object on stdout per invocation.

    With `--json` (default), `payload` is dumped as JSON. With `--text`, `text`
    (if given) is echoed verbatim; otherwise we fall back to pretty-printed
    JSON of `payload` so machines and humans see the same data.
    """
    if as_text:
        if text is not None:
            click.echo(text)
        else:
            click.echo(payload if isinstance(payload, str) else json.dumps(payload, indent=2))
    else:
        click.echo(json.dumps(payload, ensure_ascii=False))


def _exit_for(exc: WordliveError) -> int:
    if isinstance(exc, AnchorNotFoundError):
        return EXIT_ANCHOR_NOT_FOUND
    if isinstance(exc, AmbiguousMatchError):
        return EXIT_AMBIGUOUS_MATCH
    if isinstance(exc, WordBusyError):
        return EXIT_WORD_BUSY
    if isinstance(exc, WordNotRunningError):
        return EXIT_WORD_NOT_RUNNING
    if isinstance(exc, DocumentNotFoundError):
        return EXIT_OTHER
    # PathNotAllowedError (policy denial) also lands on EXIT_OTHER (1), like
    # ImageSourceError / SnapshotError — the bad-input / policy bucket.
    return EXIT_OTHER


@click.group(context_settings={"help_option_names": ["-h", "--help"]})
@click.option("--json/--text", "as_json", default=True, help="Output format (default JSON).")
@click.option("--doc", "doc_name", default=None, help="Target document by name (default: active).")
@click.option(
    "--save-dir",
    "save_dirs",
    multiple=True,
    metavar="DIR",
    help="Allow save/save-as/export-pdf to write under DIR (repeatable). "
    "Default-deny: with no --save-dir (and no WORDLIVE_SAVE_DIRS), saving is off.",
)
@click.option(
    "--image-dir",
    "image_dirs",
    multiple=True,
    metavar="DIR",
    help="Restrict insert-image --path to files under DIR (repeatable; adds to "
    "WORDLIVE_IMAGE_DIRS). Non-local paths (UNC, URLs) are always rejected.",
)
@click.pass_context
def main(
    ctx: click.Context,
    as_json: bool,
    doc_name: str | None,
    save_dirs: tuple[str, ...],
    image_dirs: tuple[str, ...],
) -> None:
    """wordlive — drive a running Microsoft Word instance.

    LLM agent? Run `wordlive llm-help` for the full agent guide in one shot: the
    anchor model, every command, the exec batch format, and exit codes (add
    `--python` for the Python-API guide). `wordlive install-skill` drops those
    guides into `.agents/skills/`, and `wordlive install-mcp` registers the MCP
    server with Claude Desktop or Claude Code.

    Saving is gated: `save`/`save-as`/`export-pdf` only write under a directory
    you whitelist with `--save-dir` (or `WORDLIVE_SAVE_DIRS`); the Python API is
    ungated.
    """
    ctx.ensure_object(dict)
    ctx.obj["as_json"] = as_json
    ctx.obj["doc_name"] = doc_name
    ctx.obj["policy"] = PathPolicy.from_env(extra_save=save_dirs, extra_image=image_dirs)


def _run(ctx: click.Context, fn: Any) -> None:
    """Top-level error boundary: classify WordliveError into exit codes."""
    try:
        fn()
    except WordliveError as exc:
        click.echo(f"error: {exc}", err=True)
        sys.exit(_exit_for(exc))


# Register subcommands. Import here to avoid a circular dependency at module load.
from . import commands  # noqa: E402

commands.register(main)
