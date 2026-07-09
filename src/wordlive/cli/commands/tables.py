"""Table commands."""

from __future__ import annotations

import json
from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ...exceptions import OpError
from ..main import _run, emit
from ._common import (
    _fmt_table_list,
    _fmt_table_read,
    _parse_color,
    _parse_rc,
)


@click.group(name="table")
def table() -> None:
    """Read and edit tables (cells are anchors: table:N:R:C)."""


@table.command(name="list")
@click.pass_context
def table_list(ctx: click.Context) -> None:
    """List every table with its position, size, and title."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.tables.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_table_list(rows))

    _run(ctx, go)


@table.command(name="read")
@click.argument("index", type=int)
@click.pass_context
def table_read(ctx: click.Context, index: int) -> None:
    """Read table INDEX (1-based) as a grid of cells with anchor IDs."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            grid = doc.tables[index].read()
            emit(grid, as_text=not ctx.obj["as_json"], text=_fmt_table_read(grid))

    _run(ctx, go)


@table.command(name="add-row")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--values", "values", default=None, help="Optional JSON array of cell values for the new row."
)
@click.pass_context
def table_add_row(ctx: click.Context, table_index: int, values: str | None) -> None:
    """Append a row to the table (atomic-undo)."""
    parsed: list[Any] | None = None
    if values is not None:
        try:
            parsed = json.loads(values)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--values must be a JSON array: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--values must be a JSON array")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: add row to table {table_index}"):
                t.add_row(parsed)
            emit(
                {"ok": True, "table": table_index, "rows": t.row_count},
                as_text=not ctx.obj["as_json"],
                text=f"added row to table:{table_index} (now {t.row_count} rows)",
            )

    _run(ctx, go)


@table.command(name="delete-row")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--row", "row", type=int, required=True, help="1-based row to delete.")
@click.pass_context
def table_delete_row(ctx: click.Context, table_index: int, row: int) -> None:
    """Delete a row from the table (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: delete row {row} from table {table_index}"):
                t.delete_row(row)
            emit(
                {"ok": True, "table": table_index, "rows": t.row_count},
                as_text=not ctx.obj["as_json"],
                text=f"deleted row {row} from table:{table_index} (now {t.row_count} rows)",
            )

    _run(ctx, go)


@table.command(name="add-column")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--values",
    "values",
    default=None,
    help="Optional JSON array of cell values for the new column (top-to-bottom).",
)
@click.pass_context
def table_add_column(ctx: click.Context, table_index: int, values: str | None) -> None:
    """Append a column to the table (atomic-undo)."""
    parsed: list[Any] | None = None
    if values is not None:
        try:
            parsed = json.loads(values)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--values must be a JSON array: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--values must be a JSON array")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: add column to table {table_index}"):
                t.add_column(parsed)
            emit(
                {"ok": True, "table": table_index, "columns": t.column_count},
                as_text=not ctx.obj["as_json"],
                text=f"added column to table:{table_index} (now {t.column_count} columns)",
            )

    _run(ctx, go)


@table.command(name="delete-column")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--column", "column", type=int, required=True, help="1-based column to delete.")
@click.pass_context
def table_delete_column(ctx: click.Context, table_index: int, column: int) -> None:
    """Delete a column from the table (atomic-undo).

    Fails with an OpError on a table with merged / mixed-width cells — Word can't
    address an individual column there; delete its cells via table:N:R:C instead.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: delete column {column} from table {table_index}"):
                t.delete_column(column)
            emit(
                {"ok": True, "table": table_index, "columns": t.column_count},
                as_text=not ctx.obj["as_json"],
                text=f"deleted column {column} from table:{table_index} (now {t.column_count} columns)",
            )

    _run(ctx, go)


@table.command(name="merge-cells")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--from", "from_cell", required=True, help='Anchor cell "R:C" (1-based).')
@click.option("--to", "to_cell", required=True, help='Opposite cell "R:C" (1-based).')
@click.pass_context
def table_merge_cells(ctx: click.Context, table_index: int, from_cell: str, to_cell: str) -> None:
    """Merge two cells (and the rectangle they span) into one (atomic-undo)."""
    fr, fc = _parse_rc(from_cell)
    tr, tc = _parse_rc(to_cell)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: merge cells in table {table_index}"):
                t.cell(fr, fc).merge(t.cell(tr, tc))
            emit(
                {"ok": True, "table": table_index, "anchor_id": f"table:{table_index}:{fr}:{fc}"},
                as_text=not ctx.obj["as_json"],
                text=f"merged into table:{table_index}:{fr}:{fc} (table is now non-uniform)",
            )

    _run(ctx, go)


@table.command(name="split-cell")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--cell", "cell", required=True, help='Cell to split, "R:C" (1-based).')
@click.option("--rows", "rows", type=int, default=1, show_default=True, help="Rows to split into.")
@click.option(
    "--columns", "columns", type=int, default=2, show_default=True, help="Columns to split into."
)
@click.pass_context
def table_split_cell(
    ctx: click.Context, table_index: int, cell: str, rows: int, columns: int
) -> None:
    """Split one cell into a rows x columns grid (atomic-undo)."""
    cr, cc = _parse_rc(cell)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: split cell in table {table_index}"):
                t.cell(cr, cc).split(rows, columns)
            emit(
                {"ok": True, "table": table_index, "anchor_id": f"table:{table_index}:{cr}:{cc}"},
                as_text=not ctx.obj["as_json"],
                text=f"split table:{table_index}:{cr}:{cc} into {rows}x{columns} (table is now non-uniform)",
            )

    _run(ctx, go)


@table.command(name="set-heading-row")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--row", "row", type=int, default=1, show_default=True, help="1-based row.")
@click.option(
    "--heading/--no-heading",
    "heading",
    default=True,
    show_default=True,
    help="Make the row a repeating table heading (repeats on every page).",
)
@click.option(
    "--allow-break/--no-allow-break",
    "allow_break",
    default=None,
    help="Allow the row to split across a page (default: off for a heading row).",
)
@click.pass_context
def table_set_heading_row(
    ctx: click.Context, table_index: int, row: int, heading: bool, allow_break: bool | None
) -> None:
    """Mark a row as a repeating table heading row (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set heading row {row} on table {table_index}"):
                t.set_heading_row(row, heading=heading, allow_break=allow_break)
            emit(
                {"ok": True, "table": table_index, "row": row, "heading": heading},
                as_text=not ctx.obj["as_json"],
                text=(
                    f"{'set' if heading else 'cleared'} heading row {row} on table:{table_index}"
                ),
            )

    _run(ctx, go)


@table.command(name="records")
@click.argument("index", type=int)
@click.pass_context
def table_records(ctx: click.Context, index: int) -> None:
    """Read table INDEX as records — body rows as dicts keyed by the header row.

    The read mirror of `table create --data` from records: row 1 is the header,
    each row below becomes a `{header: value}` object. Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            recs = doc.tables[index].records()
            emit(recs, as_text=not ctx.obj["as_json"])

    _run(ctx, go)


@table.command(name="append-record")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--record", "record", required=True, help="JSON object mapping header names to cell values."
)
@click.pass_context
def table_append_record(ctx: click.Context, table_index: int, record: str) -> None:
    """Append a row from a JSON record, mapping keys to header columns (atomic-undo)."""
    try:
        parsed = json.loads(record)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--record must be a JSON object: {e}") from e
    if not isinstance(parsed, dict):
        raise click.UsageError("--record must be a JSON object")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: append record to table {table_index}"):
                t.append_record(parsed)
            emit(
                {"ok": True, "table": table_index, "rows": t.row_count},
                as_text=not ctx.obj["as_json"],
                text=f"appended record to table:{table_index} (now {t.row_count} rows)",
            )

    _run(ctx, go)


@table.command(name="update-row")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--key", "key", required=True, help="Value to match in the key column.")
@click.option(
    "--values", "values", required=True, help="JSON object of {header: new_value} cells to set."
)
@click.option(
    "--column",
    "column",
    default=None,
    help="Header name of the key column to match (default: first column).",
)
@click.pass_context
def table_update_row(
    ctx: click.Context, table_index: int, key: str, values: str, column: str | None
) -> None:
    """Update the first row whose key-column cell equals --key, by header (atomic-undo)."""
    try:
        parsed = json.loads(values)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--values must be a JSON object: {e}") from e
    if not isinstance(parsed, dict):
        raise click.UsageError("--values must be a JSON object")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: update row in table {table_index}"):
                t.update_row(key, parsed, column=column)
            emit(
                {"ok": True, "table": table_index, "key": key},
                as_text=not ctx.obj["as_json"],
                text=f"updated row {key!r} in table:{table_index}",
            )

    _run(ctx, go)


@table.command(name="create")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Position anchor for the new table (heading:/para:/start/end/range:…).",
)
@click.option(
    "--rows",
    "rows",
    type=int,
    default=None,
    help="Number of rows (>= 1). Optional when --data is given (inferred from it).",
)
@click.option(
    "--cols",
    "cols",
    type=int,
    default=None,
    help="Number of columns (>= 1). Optional when --data is given (inferred from it).",
)
@click.option(
    "--style",
    "style",
    default=None,
    help="Table style name (default: the built-in 'Table Grid', so borders show).",
)
@click.option(
    "--header/--no-header",
    "header",
    default=False,
    show_default=True,
    help="Bold the first row as a header.",
)
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.option(
    "--data",
    "data",
    default=None,
    help="JSON to populate cells, or '-' to read it from stdin: a row-major 2-D "
    'array (\'[["Name","Qty"],["Widget","3"]]\') OR records — a list of objects '
    '(\'[{"Name":"Widget","Qty":"3"}]\'), whose keys become a header row. '
    "Reading from stdin avoids quoting/backslash fights on Windows.",
)
@click.pass_context
def table_create(
    ctx: click.Context,
    anchor_id: str,
    rows: int | None,
    cols: int | None,
    style: str | None,
    header: bool,
    before: bool,
    data: str | None,
) -> None:
    """Create a table at an anchor (atomic-undo).

    Builds new table structure where wordlive's other verbs only edit existing
    structure. Fill cells at creation with --data — a row-major JSON array, or
    records (a list of objects whose keys become a header row), or '--data -' to
    read it from stdin; a short array leaves trailing cells empty. --rows/--cols
    are optional when --data is given (inferred from its shape), required
    otherwise. --style defaults to 'Table Grid' (visible borders); a style name
    not defined in the document raises (exit 2). Reports the new table's 1-based
    index for a follow-up `table set-cell` / `add-row`.
    """
    parsed: list[Any] | None = None
    if data is not None:
        raw = click.get_text_stream("stdin").read() if data == "-" else data
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--data must be JSON (a 2-D array or records): {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--data must be a JSON array of rows or records")
    if data is None and (rows is None or cols is None):
        raise click.UsageError("--rows and --cols are required when --data is not given")

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: create table at {anchor_id}"):
                t = anchor.insert_table(
                    rows,
                    cols,
                    where=("before" if before else "after"),
                    style=style,
                    data=parsed,
                    header=header,
                )
            emit(
                {"ok": True, "table": t.index, "rows": t.row_count, "columns": t.column_count},
                as_text=not ctx.obj["as_json"],
                text=f"created table:{t.index} ({t.row_count}x{t.column_count}) at {anchor_id}",
            )

    _run(ctx, go)


@table.command(name="delete")
@click.argument("index", type=int)
@click.pass_context
def table_delete(ctx: click.Context, index: int) -> None:
    """Delete table INDEX (1-based) and all its cells (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[index]  # AnchorNotFoundError (exit 2) if missing
            with doc.edit(f"CLI: delete table {index}"):
                t.delete()
            emit(
                {"ok": True, "deleted": index},
                as_text=not ctx.obj["as_json"],
                text=f"deleted table:{index}",
            )

    _run(ctx, go)


@table.command(name="autofit")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--mode",
    type=click.Choice(["content", "window", "fixed"]),
    default="content",
    show_default=True,
    help="content: fit columns to cells · window: stretch to page · fixed: pin widths.",
)
@click.pass_context
def table_autofit(ctx: click.Context, table_index: int, mode: str) -> None:
    """Resize a table's columns — fit to content/window, or pin them (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: autofit table {table_index}"):
                t.autofit(mode)
            emit(
                {"ok": True, "table": table_index, "mode": mode},
                as_text=not ctx.obj["as_json"],
                text=f"autofit table:{table_index} ({mode})",
            )

    _run(ctx, go)


@table.command(name="set-style")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--style", "style", required=True, help="Table style name (e.g. 'Grid Table 4 - Accent 1')."
)
@click.pass_context
def table_set_style(ctx: click.Context, table_index: int, style: str) -> None:
    """Restyle an existing table (restyle first, then layer cell overrides; atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set-style table {table_index}"):
                t.set_style(style)
            emit(
                {"ok": True, "table": table_index, "style": style},
                as_text=not ctx.obj["as_json"],
                text=f"styled table:{table_index} ({style})",
            )

    _run(ctx, go)


@table.command(name="set-alignment")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--alignment",
    "alignment",
    type=click.Choice(["left", "center", "right"]),
    required=True,
    help="Align the whole table across the page width.",
)
@click.pass_context
def table_set_alignment(ctx: click.Context, table_index: int, alignment: str) -> None:
    """Align a whole table left/center/right across the page (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set-alignment table {table_index}"):
                t.set_alignment(alignment)
            emit(
                {"ok": True, "table": table_index, "alignment": alignment},
                as_text=not ctx.obj["as_json"],
                text=f"aligned table:{table_index} ({alignment})",
            )

    _run(ctx, go)


@table.command(name="set-borders")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option(
    "--sides",
    "sides",
    default="all",
    help="Edges: all/box, top, bottom, left, right, horizontal, vertical "
    "(comma-separated for several; interior gridlines need horizontal/vertical).",
)
@click.option(
    "--style",
    "style",
    default="single",
    help="Line style: single, double, dot, dash, … or none. "
    "(In exec/MCP this field is named `line_style` to avoid colliding with a "
    "paragraph/table `style` name.)",
)
@click.option("--weight", "weight", type=float, default=0.5, help="Line width in points (snapped).")
@click.option("--color", "color", default=None, help="Border colour: name, hex, or r,g,b.")
@click.pass_context
def table_set_borders(
    ctx: click.Context,
    table_index: int,
    sides: str,
    style: str,
    weight: float,
    color: str | None,
) -> None:
    """Draw borders across a whole table grid in one call (atomic-undo)."""
    side_list = [s.strip() for s in sides.split(",") if s.strip()]
    color_value = _parse_color(color)

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set-borders table {table_index}"):
                t.set_borders(sides=side_list, style=style, weight=weight, color=color_value)
            emit(
                {
                    "ok": True,
                    "table": table_index,
                    "applied": {
                        "sides": side_list,
                        "style": style,
                        "weight": weight,
                        "color": color_value,
                    },
                },
                as_text=not ctx.obj["as_json"],
                text=f"bordered table:{table_index}: {side_list} {style}",
            )

    _run(ctx, go)


@table.command(name="set-banding")
@click.option("--table", "table_index", type=int, required=True, help="1-based table index.")
@click.option("--first-row/--no-first-row", "first_row", default=None, help="Header-row banding.")
@click.option("--last-row/--no-last-row", "last_row", default=None, help="Total-row banding.")
@click.option(
    "--first-column/--no-first-column", "first_column", default=None, help="First-column banding."
)
@click.option(
    "--last-column/--no-last-column", "last_column", default=None, help="Last-column banding."
)
@click.option(
    "--banded-rows/--no-banded-rows", "banded_rows", default=None, help="Alternating row stripes."
)
@click.option(
    "--banded-columns/--no-banded-columns",
    "banded_columns",
    default=None,
    help="Alternating column stripes.",
)
@click.pass_context
def table_set_banding(
    ctx: click.Context,
    table_index: int,
    first_row: bool | None,
    last_row: bool | None,
    first_column: bool | None,
    last_column: bool | None,
    banded_rows: bool | None,
    banded_columns: bool | None,
) -> None:
    """Toggle a table's style options / banding (needs a real table style applied; atomic-undo)."""
    flags = {
        "first_row": first_row,
        "last_row": last_row,
        "first_column": first_column,
        "last_column": last_column,
        "banded_rows": banded_rows,
        "banded_columns": banded_columns,
    }
    applied = {k: v for k, v in flags.items() if v is not None}

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            t = doc.tables[table_index]
            with doc.edit(f"CLI: set-banding table {table_index}"):
                t.set_banding(**applied)
            emit(
                {"ok": True, "table": table_index, "applied": applied},
                as_text=not ctx.obj["as_json"],
                text=f"banding table:{table_index}: {applied or '(no flags)'}",
            )

    _run(ctx, go)


@click.command(name="cell-valign")
@click.option("--anchor-id", "anchor_id", required=True, help="Cell anchor (table:N:R:C) to align.")
@click.option(
    "--align",
    "align",
    type=click.Choice(["top", "center", "bottom"]),
    required=True,
    help="Where the cell's content sits vertically.",
)
@click.pass_context
def cell_valign_cmd(ctx: click.Context, anchor_id: str, align: str) -> None:
    """Set a table cell's vertical alignment — top, center, or bottom (atomic-undo)."""

    def go() -> None:
        with attach() as word:
            from ..._tables import Cell

            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            if not isinstance(anchor, Cell):
                raise OpError(
                    f"cell-valign needs a cell anchor (table:N:R:C); "
                    f"{anchor_id!r} resolved to {anchor.kind}"
                )
            with doc.edit(f"CLI: cell-valign {anchor_id}"):
                anchor.set_vertical_alignment(align)
            emit(
                {"ok": True, "anchor_id": anchor_id, "align": align},
                as_text=not ctx.obj["as_json"],
                text=f"valigned {anchor_id}: {align}",
            )

    _run(ctx, go)
