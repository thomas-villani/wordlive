"""Document properties and variables."""

from __future__ import annotations

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _fmt_properties,
    _fmt_variables,
)


@click.group(name="properties")
def properties() -> None:
    """Read and edit document properties (metadata): built-in + custom."""


@properties.command(name="list")
@click.pass_context
def properties_list(ctx: click.Context) -> None:
    """List the document's built-in and custom properties."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data = doc.properties.read()
            emit(data, as_text=not ctx.obj["as_json"], text=_fmt_properties(data))

    _run(ctx, go)


@properties.command(name="set")
@click.option("--name", "name", required=True, help="Property name (e.g. 'Title', 'Author').")
@click.option("--value", "value", required=True, help="New value.")
@click.option(
    "--custom/--builtin",
    "custom",
    default=False,
    show_default=True,
    help="Set a custom property (created if absent) instead of a built-in one.",
)
@click.pass_context
def properties_set(ctx: click.Context, name: str, value: str, custom: bool) -> None:
    """Set a built-in (default) or custom document property (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: set property {name}"):
                doc.properties.set(name, value, custom=custom)
            label = "custom property" if custom else "property"
            emit(
                {"ok": True, "name": name, "value": value, "custom": custom},
                as_text=not ctx.obj["as_json"],
                text=f"set {label} {name!r} = {value!r}",
            )

    _run(ctx, go)


@properties.command(name="delete")
@click.option("--name", "name", required=True, help="Custom property name to delete.")
@click.pass_context
def properties_delete(ctx: click.Context, name: str) -> None:
    """Delete a custom document property (atomic-undo). Built-ins can't be removed."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: delete property {name}"):
                doc.properties.delete(name)
            emit(
                {"ok": True, "name": name},
                as_text=not ctx.obj["as_json"],
                text=f"deleted custom property {name!r}",
            )

    _run(ctx, go)


@click.group(name="variables")
def variables() -> None:
    """Read and edit document variables (invisible named string storage)."""


@variables.command(name="list")
@click.pass_context
def variables_list(ctx: click.Context) -> None:
    """List the document's variables as name: value pairs."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data = doc.variables.list()
            emit(data, as_text=not ctx.obj["as_json"], text=_fmt_variables(data))

    _run(ctx, go)


@variables.command(name="set")
@click.option("--name", "name", required=True, help="Variable name.")
@click.option("--value", "value", required=True, help="Value (stored as a string).")
@click.pass_context
def variables_set(ctx: click.Context, name: str, value: str) -> None:
    """Create or update a document variable (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: set variable {name}"):
                doc.variables.set(name, value)
            emit(
                {"ok": True, "name": name, "value": value},
                as_text=not ctx.obj["as_json"],
                text=f"set variable {name!r} = {value!r}",
            )

    _run(ctx, go)


@variables.command(name="delete")
@click.option("--name", "name", required=True, help="Variable name to delete.")
@click.pass_context
def variables_delete(ctx: click.Context, name: str) -> None:
    """Delete a document variable (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: delete variable {name}"):
                doc.variables.delete(name)
            emit(
                {"ok": True, "name": name},
                as_text=not ctx.obj["as_json"],
                text=f"deleted variable {name!r}",
            )

    _run(ctx, go)
