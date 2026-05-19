"""Click entry point for `wordlive`. JSON in, JSON out, deterministic exit codes."""

from __future__ import annotations

import json
import sys
from typing import Any

import click

from ..exceptions import (
    AnchorNotFoundError,
    DocumentNotFoundError,
    WordBusyError,
    WordNotRunningError,
    WordliveError,
)


# Exit codes per spec.md §"CLI Sketch":
EXIT_OK = 0
EXIT_OTHER = 1
EXIT_ANCHOR_NOT_FOUND = 2
EXIT_WORD_BUSY = 3
EXIT_WORD_NOT_RUNNING = 4


def emit(payload: Any, *, as_text: bool = False) -> None:
    """One JSON object on stdout, period. (`--text` falls back to repr for humans.)"""
    if as_text:
        click.echo(payload if isinstance(payload, str) else json.dumps(payload, indent=2))
    else:
        click.echo(json.dumps(payload, ensure_ascii=False))


def _exit_for(exc: WordliveError) -> int:
    if isinstance(exc, AnchorNotFoundError):
        return EXIT_ANCHOR_NOT_FOUND
    if isinstance(exc, WordBusyError):
        return EXIT_WORD_BUSY
    if isinstance(exc, WordNotRunningError):
        return EXIT_WORD_NOT_RUNNING
    if isinstance(exc, DocumentNotFoundError):
        return EXIT_OTHER
    return EXIT_OTHER


@click.group(context_settings={"help_option_names": ["-h", "--help"]})
@click.option("--json/--text", "as_json", default=True, help="Output format (default JSON).")
@click.option("--doc", "doc_name", default=None, help="Target document by name (default: active).")
@click.pass_context
def main(ctx: click.Context, as_json: bool, doc_name: str | None) -> None:
    """wordlive — drive a running Microsoft Word instance."""
    ctx.ensure_object(dict)
    ctx.obj["as_json"] = as_json
    ctx.obj["doc_name"] = doc_name


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
