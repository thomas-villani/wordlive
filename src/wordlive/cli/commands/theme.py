"""Document-theme commands."""

from __future__ import annotations

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit


@click.command(name="theme")
@click.pass_context
def theme_cmd(ctx: click.Context) -> None:
    """Show the document's current theme (colours + major/minor fonts). Non-mutating."""

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data = doc.theme.to_dict()
            text = "\n".join(
                [
                    f"major font: {data['major_font']}",
                    f"minor font: {data['minor_font']}",
                    *(f"{k}: {v}" for k, v in data["colors"].items()),
                ]
            )
            emit(data, as_text=not ctx.obj["as_json"], text=text)

    _run(ctx, go)


@click.command(name="list-themes")
@click.pass_context
def list_themes_cmd(ctx: click.Context) -> None:
    """List the built-in themes, colour schemes, and font schemes Office ships.

    These names feed `apply-theme --theme`, `set-theme-colors --scheme`, and
    `set-theme-fonts --scheme`. Non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            data = doc.theme.list_available()
            text = "\n".join(
                [
                    f"themes: {', '.join(data['themes'])}",
                    f"color schemes: {', '.join(data['color_schemes'])}",
                    f"font schemes: {', '.join(data['font_schemes'])}",
                ]
            )
            emit(data, as_text=not ctx.obj["as_json"], text=text)

    _run(ctx, go)


@click.command(name="apply-theme")
@click.option(
    "--theme",
    "theme",
    required=True,
    help="Built-in theme name (e.g. Facet, Ion) or a .thmx file path.",
)
@click.pass_context
def apply_theme_cmd(ctx: click.Context, theme: str) -> None:
    """Apply a whole document theme — colours, fonts, and effects (atomic-undo).

    See `list-themes` for the built-in names. Brand colours/fonts can then be
    overridden with `set-theme-colors` / `set-theme-fonts`.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit(f"CLI: apply theme {theme}"):
                applied = doc.theme.apply(theme)
            emit(
                {"ok": True, "applied": {"theme": applied}},
                as_text=not ctx.obj["as_json"],
                text=f"applied theme {applied!r}",
            )

    _run(ctx, go)


@click.command(name="set-theme-colors")
@click.option("--scheme", "scheme", default=None, help="Built-in colour scheme name or .xml path.")
@click.option("--text1", "text1", default=None, help="Text 1 (dark 1) colour.")
@click.option("--background1", "background1", default=None, help="Background 1 (light 1) colour.")
@click.option("--text2", "text2", default=None, help="Text 2 (dark 2) colour.")
@click.option("--background2", "background2", default=None, help="Background 2 (light 2) colour.")
@click.option("--accent1", "accent1", default=None, help="Accent 1 colour (name/hex).")
@click.option("--accent2", "accent2", default=None, help="Accent 2 colour.")
@click.option("--accent3", "accent3", default=None, help="Accent 3 colour.")
@click.option("--accent4", "accent4", default=None, help="Accent 4 colour.")
@click.option("--accent5", "accent5", default=None, help="Accent 5 colour.")
@click.option("--accent6", "accent6", default=None, help="Accent 6 colour.")
@click.option("--hyperlink", "hyperlink", default=None, help="Hyperlink colour.")
@click.option(
    "--followed-hyperlink", "followed_hyperlink", default=None, help="Followed-hyperlink colour."
)
@click.pass_context
def set_theme_colors_cmd(
    ctx: click.Context,
    scheme: str | None,
    text1: str | None,
    background1: str | None,
    text2: str | None,
    background2: str | None,
    accent1: str | None,
    accent2: str | None,
    accent3: str | None,
    accent4: str | None,
    accent5: str | None,
    accent6: str | None,
    hyperlink: str | None,
    followed_hyperlink: str | None,
) -> None:
    """Set the theme's colour scheme and/or individual brand colours (atomic-undo).

    Pass `--scheme` for a named built-in scheme, and/or any `--accentN`/`--text*`
    flag to override a single colour (a name like `navy` or hex like `#1A73E8`).
    """
    overrides = {
        k: v
        for k, v in {
            "text1": text1,
            "background1": background1,
            "text2": text2,
            "background2": background2,
            "accent1": accent1,
            "accent2": accent2,
            "accent3": accent3,
            "accent4": accent4,
            "accent5": accent5,
            "accent6": accent6,
            "hyperlink": hyperlink,
            "followed_hyperlink": followed_hyperlink,
        }.items()
        if v is not None
    }

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: set theme colors"):
                colors = doc.theme.set_colors(scheme=scheme, **overrides)
            emit(
                {"ok": True, "colors": colors, "applied": {"scheme": scheme, **overrides}},
                as_text=not ctx.obj["as_json"],
                text="set theme colours",
            )

    _run(ctx, go)


@click.command(name="set-theme-fonts")
@click.option("--scheme", "scheme", default=None, help="Built-in font scheme name or .xml path.")
@click.option("--major", "major", default=None, help="Major (heading) font name.")
@click.option("--minor", "minor", default=None, help="Minor (body) font name.")
@click.pass_context
def set_theme_fonts_cmd(
    ctx: click.Context, scheme: str | None, major: str | None, minor: str | None
) -> None:
    """Set the theme's fonts via a named scheme and/or explicit names (atomic-undo).

    `--scheme` loads a named built-in font scheme; `--major`/`--minor` override
    the heading/body font names.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            with doc.edit("CLI: set theme fonts"):
                fonts = doc.theme.set_fonts(scheme=scheme, major=major, minor=minor)
            emit(
                {
                    "ok": True,
                    **fonts,
                    "applied": {"scheme": scheme, "major": major, "minor": minor},
                },
                as_text=not ctx.obj["as_json"],
                text=f"set theme fonts (major={fonts['major_font']}, minor={fonts['minor_font']})",
            )

    _run(ctx, go)
