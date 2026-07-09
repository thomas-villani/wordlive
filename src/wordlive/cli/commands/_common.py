"""Shared CLI helpers, formatters, and reusable option constants."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ...exceptions import OpError
from ..main import emit

_WITHIN_HELP = (
    "Scope to an anchor's range (e.g. 'heading:3', 'range:120-540'); default is the whole document."
)
_WRAP_CHOICES = ["inline", "auto", "square", "tight", "through", "top-bottom", "behind", "front"]
_WHICH_OPTION = click.option(
    "--which",
    "which",
    type=click.Choice(["primary", "first", "even"], case_sensitive=False),
    default="primary",
    show_default=True,
    help="Which header/footer: primary, first-page, or even-pages.",
)
_SECTION_OPTION = click.option(
    "--section",
    "section_index",
    type=int,
    default=1,
    show_default=True,
    help="1-based section index.",
)


def _fmt_status(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no documents open)"
    lines: list[str] = []
    width = max(len(str(r.get("name", ""))) for r in rows)
    for r in rows:
        marker = "*" if r.get("is_active") else " "
        lines.append(f"{marker} {str(r.get('name', '')):<{width}}  {r.get('path', '')}")
    return "\n".join(lines)


def _fmt_outline(items: list[dict[str, Any]]) -> str:
    if not items:
        return "(no headings)"
    lines: list[str] = []
    for it in items:
        level = int(it.get("level", 1))
        indent = "  " * max(level - 1, 0)
        lines.append(f"{indent}{it.get('text', '')}  [{it.get('anchor_id', '')}]")
    return "\n".join(lines)


def _fmt_paragraphs(items: list[dict[str, Any]]) -> str:
    if not items:
        return "(no paragraphs)"
    lines: list[str] = []
    for it in items:
        marker = f"H{it.get('level', 1)}" if it.get("is_heading") else "  "
        text = it.get("text", "")
        snippet = text if len(text) <= 60 else text[:57] + "…"
        lines.append(
            f"{marker} [{it.get('anchor_id', '')}] "
            f"{it.get('start', 0)}-{it.get('end', 0)}  {snippet}"
        )
    return "\n".join(lines)


def _fmt_find(matches: list[dict[str, Any]]) -> str:
    if not matches:
        return "(no matches)"
    return "\n".join(f"{m['start']:>6}–{m['end']:<6}  {m['text']!r}" for m in matches)


def _fmt_find_paragraphs(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no matches)"
    return "\n".join(f"{r['score']:.2f}  {r['anchor_id']:<10}  {r['text']!r}" for r in rows)


def _fmt_nearest_heading(anchor_id: str, direction: str, row: dict[str, Any] | None) -> str:
    if row is None:
        return f"(no heading {direction} {anchor_id})"
    return f"{row['anchor_id']}  (L{row['level']})  {row['text']!r}"


def _fmt_replace_summary(replacements: list[dict[str, Any]]) -> str:
    n = len(replacements)
    return f"replaced {n} occurrence{'s' if n != 1 else ''}"


def _fmt_format_info(info: dict[str, Any]) -> str:
    lines = [f"{info['anchor_id']}  style={info['style']!r}"]
    for group in ("paragraph", "font"):
        for field, cell in info[group].items():
            if field == "mixed":
                continue
            mark = " *override*" if cell.get("override") else ""
            lines.append(f"  {field}: {cell['value']} (style {cell['style']}){mark}")
    if info["font"].get("mixed"):
        lines.append(f"  mixed runs: {', '.join(info['font']['mixed'])}")
    return "\n".join(lines)


def _fmt_notes(rows: list[dict[str, Any]], scheme: str) -> str:
    if not rows:
        return f"(no {scheme}s)"
    lines = []
    for r in rows:
        where = f" @ {r['para']}" if r.get("para") else ""
        lines.append(f"{scheme}:{r['index']}{where}  {r.get('text', '')}")
    return "\n".join(lines)


def _fmt_revisions(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "no tracked changes"
    lines = []
    for r in rows:
        who = r.get("author") or "?"
        text = (r.get("text") or "").replace("\r", " ").replace("\n", " ")
        if len(text) > 60:
            text = text[:57] + "…"
        lines.append(f"{r['index']}. [{r.get('type')}] {who}: {text!r} ({r.get('anchor_id')})")
    return "\n".join(lines)


def _fmt_location(anchor_id: str, loc: dict[str, Any]) -> str:
    span = (
        f"page {loc['page']}"
        if loc["page"] == loc["end_page"]
        else f"pages {loc['page']}–{loc['end_page']}"
    )
    where = f"{anchor_id}: {span}, line {loc['line']}, col {loc['column']}"
    return where + (" (in table)" if loc["in_table"] else "")


def _fmt_stats(s: dict[str, Any]) -> str:
    order = [
        "pages",
        "words",
        "characters",
        "paragraphs",
        "lines",
        "sections",
        "headings",
        "tables",
        "images",
        "comments",
        "revisions",
    ]
    parts = [f"{k}: {s[k]}" for k in order if k in s]
    parts.append("saved" if s.get("saved") else "unsaved")
    return "  ".join(parts)


def _fmt_proofing(data: dict[str, Any]) -> str:
    sp, gr = data.get("spelling", {}), data.get("grammar", {})
    read = data.get("readability", {})
    lines = [
        f"spelling errors: {sp.get('count')}",
        f"grammar errors: {gr.get('count')}",
    ]
    for key in (
        "flesch_reading_ease",
        "flesch_kincaid_grade_level",
        "passive_sentences",
        "words_per_sentence",
    ):
        if key in read:
            lines.append(f"{key}: {read[key]}")
    return "\n".join(lines)


def _rules_selector(rule: tuple[str, ...], exclude: tuple[str, ...]) -> Any:
    """Build the `rules=` selector from the repeatable --rule / --exclude flags."""
    if rule and exclude:
        raise click.UsageError("pass either --rule or --exclude, not both")
    if rule:
        return list(rule)
    if exclude:
        return {"exclude": list(exclude)}
    return None


def _fmt_lint(findings: list[dict[str, Any]]) -> str:
    if not findings:
        return "(no findings)"
    lines = []
    for f in findings:
        fix = " [fixable]" if f.get("fixable") else ""
        lines.append(f"[{f['severity']}] {f['rule']} ({f['anchor_id']}): {f['message']}{fix}")
    return "\n".join(lines)


def _fmt_regularize(report: dict[str, Any]) -> str:
    applied, skipped = report.get("applied", []), report.get("skipped", [])
    deferred = report.get("deferred", [])
    verb = "would fix" if report.get("dry_run") else "fixed"
    summary = f"{verb} {len(applied)}; skipped {len(skipped)} (report-only / not fixable)"
    if deferred:
        summary += f"; deferred {len(deferred)} content fix(es) (pass --allow-content)"
    lines = [summary]
    for f in applied:
        lines.append(f"  {verb}: {f['rule']} ({f['anchor_id']})")
    for f in deferred:
        lines.append(f"  deferred: {f['rule']} ({f['anchor_id']})")
    return "\n".join(lines)


def _fmt_hyperlinks(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no hyperlinks)"
    lines: list[str] = []
    for r in rows:
        dest = r.get("address") or (f"#{r['sub_address']}" if r.get("sub_address") else "?")
        text = r.get("text") or ""
        para = r.get("para") or ""
        lines.append(f"[{r['anchor_id']}] {text!r} -> {dest}  {para}".rstrip())
    return "\n".join(lines)


def _fmt_fields(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no fields)"
    lines: list[str] = []
    for r in rows:
        result = r.get("result") or ""
        suffix = f" = {result!r}" if result else ""
        lines.append(f"[{r['anchor_id']}] {r['kind']}: {r['code']}{suffix}")
    return "\n".join(lines)


def _fmt_properties(data: dict[str, Any]) -> str:
    lines: list[str] = []
    for bag in ("builtin", "custom"):
        items = data.get(bag, {})
        if items:
            lines.append(f"[{bag}]")
            lines.extend(f"  {k}: {v}" for k, v in items.items())
    return "\n".join(lines) if lines else "(no properties)"


def _fmt_variables(data: dict[str, str]) -> str:
    if not data:
        return "(no variables)"
    return "\n".join(f"{k}: {v}" for k, v in data.items())


def _fmt_images(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no images)"
    lines: list[str] = []
    for r in rows:
        w, h = r.get("width"), r.get("height")
        size = f"  {w:.0f}×{h:.0f}pt" if w and h else ""
        mime = r.get("mime") or "?"
        para = r.get("para") or ""
        alt = r.get("alt_text") or ""
        crop = "  cropped" if r.get("crop") else ""
        suffix = f"  {alt!r}" if alt else ""
        lines.append(f"[{r['anchor_id']}] {mime}{size}{crop}  {para}{suffix}")
    return "\n".join(lines)


def _cc_anchor(word: Any, doc_name: str | None, anchor_id: str) -> Any:
    """Resolve `anchor_id` to a content control, raising a clean usage error otherwise."""
    from ..._anchors import ContentControl

    doc = _pick_doc(word, doc_name)
    anchor = doc.anchor_by_id(anchor_id)
    if not isinstance(anchor, ContentControl):
        raise click.UsageError(f"{anchor_id!r} is not a content control; pass a cc:NAME anchor")
    return doc, anchor


def _fmt_equations(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no equations)"
    lines: list[str] = []
    for r in rows:
        para = r.get("para") or ""
        linear = r.get("linear") or ""
        preview = f"  {linear}" if linear else ""
        lines.append(f"[{r['anchor_id']}] {r.get('type', '?')}  {para}{preview}")
    return "\n".join(lines)


def _fmt_charts(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no charts)"
    lines: list[str] = []
    for r in rows:
        para = r.get("para") or ""
        title = r.get("title")
        suffix = f"  {title!r}" if title else ""
        lines.append(f"[{r['anchor_id']}] {r.get('kind', '?')}  {para}{suffix}")
    return "\n".join(lines)


def _chart_anchor(word: Any, doc_name: str | None, anchor_id: str) -> Any:
    """Resolve `anchor_id` to a chart, raising a clean usage error otherwise."""
    from ..._anchors import ChartAnchor

    doc = _pick_doc(word, doc_name)
    anchor = doc.anchor_by_id(anchor_id)
    if not isinstance(anchor, ChartAnchor):
        raise click.UsageError(f"{anchor_id!r} is not a chart; pass a chart:N anchor")
    return doc, anchor


def _fmt_shapes(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no shapes)"
    lines: list[str] = []
    for r in rows:
        para = r.get("para") or ""
        w, h = r.get("width"), r.get("height")
        dims = f"{round(w)}x{round(h)}pt" if w is not None and h is not None else ""
        wrap = r.get("wrap")
        side = r.get("wrap_side")
        wrap_txt = f"wrap={wrap}" + (f"/{side}" if side and side != "both" else "")
        crop_txt = "  cropped" if r.get("crop") else ""
        lines.append(
            f"[{r['anchor_id']}] {r.get('shape_type', '?')}  {dims}  "
            f"{wrap_txt}{crop_txt}  {para}".rstrip()
        )
    return "\n".join(lines)


def _shape_anchor(word: Any, doc_name: str | None, anchor_id: str) -> Any:
    """Resolve `anchor_id` to a floating shape, raising a clean usage error otherwise."""
    from ..._anchors import ShapeAnchor

    doc = _pick_doc(word, doc_name)
    anchor = doc.anchor_by_id(anchor_id)
    if not isinstance(anchor, ShapeAnchor):
        raise click.UsageError(f"{anchor_id!r} is not a shape; pass a shape:N anchor")
    return doc, anchor


def _image_anchor(word: Any, doc_name: str | None, anchor_id: str) -> Any:
    """Resolve `anchor_id` to an inline image, raising a clean usage error otherwise."""
    from ..._anchors import ImageAnchor

    doc = _pick_doc(word, doc_name)
    anchor = doc.anchor_by_id(anchor_id)
    if not isinstance(anchor, ImageAnchor):
        raise click.UsageError(f"{anchor_id!r} is not an inline image; pass an image:N anchor")
    return doc, anchor


def _parse_pages_range(value: str) -> tuple[int, int]:
    """Parse a `--pages` value like `2-4` into an inclusive `(start, end)` span."""
    start_str, sep, end_str = value.partition("-")
    if not sep:
        raise click.UsageError("--pages must look like 'A-B' (inclusive), e.g. '2-4'")
    try:
        start, end = int(start_str), int(end_str)
    except ValueError as e:
        raise click.UsageError("--pages must look like 'A-B' (inclusive), e.g. '2-4'") from e
    if start < 1 or end < start:
        raise click.UsageError(f"invalid page span {value!r}: need 1 <= start <= end")
    return start, end


def _parse_rc(value: str) -> tuple[int, int]:
    """Parse a 1-based cell coordinate like `2:3` into `(row, col)`."""
    row_str, sep, col_str = value.partition(":")
    if not sep:
        raise click.UsageError("cell must look like 'R:C' (1-based), e.g. '2:3'")
    try:
        row, col = int(row_str), int(col_str)
    except ValueError as e:
        raise click.UsageError("cell must look like 'R:C' (1-based), e.g. '2:3'") from e
    if row < 1 or col < 1:
        raise click.UsageError(f"invalid cell {value!r}: row and column are 1-based")
    return row, col


def _parse_color(value: str | None) -> str | tuple[int, int, int] | None:
    """Turn a `--color` value into something `to_bgr` understands.

    A comma-separated `r,g,b` becomes an `(r, g, b)` tuple; anything else
    (a colour name or hex string) passes through unchanged for the helper to
    resolve. Returns `None` for `None` (option not given).
    """
    if value is None:
        return None
    if "," in value:
        try:
            r, g, b = (int(p.strip()) for p in value.split(","))
        except ValueError as e:
            raise click.UsageError(f"--color as r,g,b needs three integers; got {value!r}") from e
        return (r, g, b)
    return value


def _fmt_snapshot(images: list[dict[str, Any]], dpi: int) -> str:
    if not images:
        return "(no pages rendered)"
    lines: list[str] = []
    for im in images:
        size = f"{im['bytes']} bytes"
        where = im.get("path") or "base64"
        lines.append(f"page {im['page']}: {size} → {where}")
    head = f"rendered {len(images)} page(s) at {dpi} dpi"
    return head + "\n" + "\n".join(lines)


def _fmt_cursor(info: dict[str, Any]) -> str:
    para = info.get("paragraph")
    where = f"  in {para['anchor_id']}" if para else ""
    if info.get("collapsed"):
        return f"cursor at {info.get('start', 0)}{where}"
    sel = info.get("text") or ""
    return f"selection {info.get('start', 0)}-{info.get('end', 0)}: {sel!r}{where}"


def _load_checkpoint(path: Path) -> Any:
    """Read a checkpoint token from a file, mapping IO/parse errors to OpError
    (clean exit 1) rather than a traceback."""
    from ..._checkpoint import Checkpoint

    try:
        return Checkpoint.from_json(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise OpError(f"cannot read checkpoint file {str(path)!r}: {exc}") from exc
    except (ValueError, KeyError) as exc:
        raise OpError(f"invalid checkpoint file {str(path)!r}: {exc}") from exc


def _fmt_style_list(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no styles)"
    name_w = max(len(r["name"]) for r in rows)
    return "\n".join(
        f"{r['name']:<{name_w}}  {r['type']:<10}  "
        f"{'builtin' if r['builtin'] else 'custom':<8}  "
        f"{'in-use' if r['in_use'] else ''}"
        for r in rows
    )


def _fmt_table_list(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no tables)"
    return "\n".join(
        f"table:{r['index']}  {r['rows']}x{r['columns']}"
        + (f"  {r['title']!r}" if r.get("title") else "")
        for r in rows
    )


def _fmt_table_read(grid: dict[str, Any]) -> str:
    cells = grid.get("cells") or []
    if not cells:
        return f"table:{grid.get('index')} (empty)"
    # Rows can be ragged on a merged / split table, so size columns off the
    # widest row and guard each per-column scan against shorter rows.
    ncols = max((len(row) for row in cells), default=0)
    widths = [
        max((len(row[c]["text"]) for row in cells if c < len(row)), default=0) for c in range(ncols)
    ]
    lines = []
    for row in cells:
        lines.append("  ".join(cell["text"].ljust(widths[i]) for i, cell in enumerate(row)))
    return "\n".join(lines)


def _fmt_comment_list(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no comments)"
    lines: list[str] = []
    for r in rows:
        author = r.get("author") or "?"
        state = "resolved" if r.get("done") else "open"
        scope = r.get("scope") or ""
        on = f"  on {scope!r}" if scope else ""
        lines.append(f"[{r['index']}] {author} ({state}): {r.get('text', '')}{on}")
    return "\n".join(lines)


def _fmt_list_show(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no lists)"
    return "\n".join(
        f"list:{r['index']}  {r['type']}  "
        f"{r['count']} item{'s' if r['count'] != 1 else ''}  [{r['anchor_id']}]"
        for r in rows
    )


def _fmt_list_info(info: dict[str, Any]) -> str:
    if info.get("type") == "none":
        return "not in a list"
    return (
        f"{info['type']} (level {info['level']}, number {info['number']}, "
        f"marker {info['string']!r})"
    )


def _fmt_list_levels(levels: list[dict[str, Any]]) -> str:
    if not levels:
        return "not in a list"
    return "\n".join(
        f"L{lv['level']}: {lv['kind']} {lv['format']!r}"
        + (f" ({lv['style']})" if lv["kind"] == "number" else f" font {lv['font']!r}")
        + f"  trailing={lv['trailing']}  num@{lv['number_position']:g}pt text@{lv['text_position']:g}pt"
        for lv in levels
    )


def _fmt_section_list(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no sections)"
    lines: list[str] = []
    for r in rows:
        ps = r.get("page_setup", {})
        lines.append(
            f"section:{r['index']}  {ps.get('orientation', '?')}  "
            f"{ps.get('page_width', 0):.0f}x{ps.get('page_height', 0):.0f}pt"
        )
    return "\n".join(lines)


def _emit_section_list(ctx: click.Context) -> None:
    with attach() as word:
        doc = _pick_doc(word, ctx.obj["doc_name"])
        rows = doc.sections.list()
        emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_section_list(rows))


def _claude_desktop_config_path() -> Path:
    """Where Claude Desktop keeps `claude_desktop_config.json` on this OS."""
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA")
        base = Path(appdata) if appdata else Path.home() / "AppData" / "Roaming"
        return base / "Claude" / "claude_desktop_config.json"
    if sys.platform == "darwin":
        return (
            Path.home()
            / "Library"
            / "Application Support"
            / "Claude"
            / "claude_desktop_config.json"
        )
    return Path.home() / ".config" / "Claude" / "claude_desktop_config.json"


def _mcp_server_entry(directory: str | None) -> dict[str, Any]:
    """The `mcpServers` entry that launches the wordlive stdio server.

    Default (repo-less) form runs the published package straight from PyPI with
    `uvx` — `wordlive-mcp` is a console script *inside* `wordlive`, so it needs
    `--from "wordlive[mcp,snapshot]"` to tell uv which package provides it (and
    the `snapshot` extra enables the vision tool). With `--directory` (a local
    checkout) wordlive *is* the project, so a plain `uv run wordlive-mcp`
    resolves it without `--from`.
    """
    if directory:
        return {"command": "uv", "args": ["run", "--directory", directory, "wordlive-mcp"]}
    return {"command": "uvx", "args": ["--from", "wordlive[mcp,snapshot]", "wordlive-mcp"]}
