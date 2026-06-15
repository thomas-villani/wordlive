"""Document theme — the document-wide brand primitive behind `doc.theme`.

A Word *document theme* (the Design tab's `OfficeTheme`) bundles a **colour
scheme** (12 theme colours: text/background pairs, six accents, two hyperlink
colours), a **font scheme** (a major heading font + a minor body font), and an
effect scheme. It's the natural target for "make this document match our brand":
set the accent colours and the fonts once and every theme-aware style follows.

wordlive surfaces it as `doc.theme` — a view mirroring `doc.styles`:

- `theme.apply("Facet")` swaps the whole theme (a built-in name or a ``.thmx``
  path);
- `theme.set_colors(scheme="Blue")` loads a named colour scheme, and
  `theme.set_colors(accent1="#1A73E8", text1="navy")` overrides individual brand
  colours;
- `theme.set_fonts(scheme="Garamond")` / `theme.set_fonts(major="Arial",
  minor="Calibri")` set the heading/body fonts;
- `theme.colors` / `theme.major_font` / `theme.minor_font` / `theme.to_dict()`
  read the current theme back.

Built-in themes, colour schemes, and font schemes ship with Office on disk
(``<office>/Document Themes <ver>``); `set_*`/`apply` resolve a friendly name
against that library or accept an absolute path to a brand file. All facts here
(the 12-colour ``ThemeColorScheme``, the BGR ``.RGB`` ints, ``Load``/``Save`` on
each scheme, ``ApplyDocumentTheme``) were confirmed against live Word.
"""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, Any

from . import _com
from ._format import bgr_to_hex, to_bgr
from .exceptions import OpError

if TYPE_CHECKING:
    from ._document import Document

# Friendly colour name (+ aliases) -> the 1-based ThemeColorScheme index Word
# uses. Indices confirmed live: 1 dk1, 2 lt1, 3 dk2, 4 lt2, 5-10 accent1-6,
# 11 hyperlink, 12 followedHyperlink. The text/background aliases are the names
# Word's own UI shows; dark/light mirror the OOXML dk/lt element names.
_COLOR_INDEX: dict[str, int] = {
    "text1": 1,
    "dark1": 1,
    "background1": 2,
    "light1": 2,
    "text2": 3,
    "dark2": 3,
    "background2": 4,
    "light2": 4,
    "accent1": 5,
    "accent2": 6,
    "accent3": 7,
    "accent4": 8,
    "accent5": 9,
    "accent6": 10,
    "hyperlink": 11,
    "followed_hyperlink": 12,
}

# 1-based index -> the canonical friendly name used when *reading* colours back.
_COLOR_NAMES: dict[int, str] = {
    1: "text1",
    2: "background1",
    3: "text2",
    4: "background2",
    5: "accent1",
    6: "accent2",
    7: "accent3",
    8: "accent4",
    9: "accent5",
    10: "accent6",
    11: "hyperlink",
    12: "followed_hyperlink",
}

# kind -> (subdirectory, file extension) within the Office theme library.
_KIND_DIR: dict[str, tuple[str, str]] = {
    "theme": ("", ".thmx"),
    "colors": ("Theme Colors", ".xml"),
    "fonts": ("Theme Fonts", ".xml"),
}


def _themes_dir(app_com: Any) -> str:
    """The Office built-in theme library directory for the running install.

    ``<dirname(app.Path)>/Document Themes <major>`` — e.g. ``app.Path`` is
    ``…/Root/Office16`` and ``app.Version`` is ``"16.0"``, giving
    ``…/Root/Document Themes 16``. Derived at runtime so it tracks the actual
    Office version rather than hard-coding ``16``.
    """
    root = os.path.dirname(str(app_com.Path))
    major = str(app_com.Version).split(".")[0]
    return os.path.join(root, f"Document Themes {major}")


def _resolve_theme_file(app_com: Any, name: str, kind: str) -> str:
    """Resolve a theme/colour-scheme/font-scheme `name` to a file path.

    `kind` is ``"theme"`` (``.thmx``), ``"colors"``, or ``"fonts"``. An existing
    file path is returned as-is (the brand-file escape hatch); otherwise `name`
    is matched case-insensitively against the built-in library. A miss raises
    `OpError` listing a few available names.
    """
    if not str(name).strip():
        raise OpError(f"{kind} name must be a non-empty string")
    if os.path.isfile(name):
        return name
    subdir, ext = _KIND_DIR[kind]
    base = os.path.join(_themes_dir(app_com), subdir)
    target = name if name.lower().endswith(ext) else name + ext
    if os.path.isdir(base):
        for entry in os.listdir(base):
            if entry.lower() == target.lower():
                return os.path.join(base, entry)
    available = _list_builtin(app_com)[
        {"theme": "themes", "colors": "color_schemes", "fonts": "font_schemes"}[kind]
    ]
    hint = ", ".join(available[:8]) + ("…" if len(available) > 8 else "")
    raise OpError(f"unknown {kind} {name!r}; pass a file path or one of the built-ins ({hint})")


def _list_builtin(app_com: Any) -> dict[str, list[str]]:
    """List the built-in themes, colour schemes, and font schemes (no extension)."""
    base = _themes_dir(app_com)

    def names(subdir: str, ext: str) -> list[str]:
        d = os.path.join(base, subdir) if subdir else base
        if not os.path.isdir(d):
            return []
        out = [
            os.path.splitext(f)[0]
            for f in os.listdir(d)
            if f.lower().endswith(ext) and os.path.isfile(os.path.join(d, f))
        ]
        return sorted(out)

    return {
        "themes": names("", ".thmx"),
        "color_schemes": names("Theme Colors", ".xml"),
        "font_schemes": names("Theme Fonts", ".xml"),
    }


class DocumentTheme:
    """The document's theme — `doc.theme`.

    A read/mutate view over Word's `OfficeTheme`: read the current colours/fonts,
    `apply(...)` a whole theme, or `set_colors(...)`/`set_fonts(...)` to brand a
    document. Mutations should be wrapped in `doc.edit(...)` for atomic undo and
    to preserve the user's selection/scroll.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def _theme(self) -> Any:
        return self._doc.com.DocumentTheme

    def _app(self) -> Any:
        return self._doc.com.Application

    @property
    def com(self) -> Any:
        """Raw COM `OfficeTheme` object — the escape hatch."""
        return self._theme()

    @property
    def colors(self) -> dict[str, str]:
        """The 12 theme colours as ``{friendly_name: "#RRGGBB"}``.

        Keys are `text1`, `background1`, `text2`, `background2`, `accent1`–
        `accent6`, `hyperlink`, `followed_hyperlink`.
        """
        with _com.translate_com_errors():
            scheme = self._theme().ThemeColorScheme
            return {
                _COLOR_NAMES[i]: bgr_to_hex(int(scheme.Colors(i).RGB))
                for i in range(1, int(scheme.Count) + 1)
                if i in _COLOR_NAMES
            }

    @property
    def major_font(self) -> str:
        """The theme's major (heading) font name."""
        with _com.translate_com_errors():
            return str(self._theme().ThemeFontScheme.MajorFont.Item(1).Name)

    @property
    def minor_font(self) -> str:
        """The theme's minor (body) font name."""
        with _com.translate_com_errors():
            return str(self._theme().ThemeFontScheme.MinorFont.Item(1).Name)

    def set_colors(self, scheme: str | None = None, **colors: Any) -> dict[str, str]:
        """Set the theme's colour scheme and/or individual brand colours.

        `scheme` loads a named built-in colour scheme (``"Blue"``, ``"Orange"``,
        …) or a Theme-Colors ``.xml`` path; then each keyword override
        (`accent1="#1A73E8"`, `text1="navy"`, … — friendly names from
        `theme.colors`) is applied. Colour values take any form `to_bgr` accepts
        (a colour name, a hex string, or an ``(r, g, b)`` tuple).
        Wrap in `doc.edit(...)`. Unknown scheme name or colour key/value raises
        `OpError`. Returns the resulting `colors` dict.
        """
        try:
            with _com.translate_com_errors():
                tcs = self._theme().ThemeColorScheme
                if scheme is not None:
                    tcs.Load(_resolve_theme_file(self._app(), scheme, "colors"))
                for key, value in colors.items():
                    idx = _COLOR_INDEX.get(key.lower())
                    if idx is None:
                        raise ValueError(
                            f"unknown theme colour {key!r}; expected one of "
                            f"{sorted(set(_COLOR_NAMES.values()))}"
                        )
                    tcs.Colors(idx).RGB = to_bgr(value)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        return self.colors

    def set_fonts(
        self, scheme: str | None = None, *, major: str | None = None, minor: str | None = None
    ) -> dict[str, str]:
        """Set the theme's fonts via a named scheme and/or explicit names.

        `scheme` loads a named built-in font scheme (``"Garamond"``, ``"Arial"``,
        …) or a Theme-Fonts ``.xml`` path; then `major` (heading font) and
        `minor` (body font) override individual names. Wrap in `doc.edit(...)`.
        An unknown scheme name raises `OpError`. Returns
        ``{"major_font", "minor_font"}``.
        """
        with _com.translate_com_errors():
            tfs = self._theme().ThemeFontScheme
            if scheme is not None:
                tfs.Load(_resolve_theme_file(self._app(), scheme, "fonts"))
            if major is not None:
                tfs.MajorFont.Item(1).Name = str(major)
            if minor is not None:
                tfs.MinorFont.Item(1).Name = str(minor)
        return {"major_font": self.major_font, "minor_font": self.minor_font}

    def apply(self, theme: str) -> str:
        """Apply a whole document theme (colours + fonts + effects).

        `theme` is a built-in name (``"Facet"``, ``"Ion"``, … — see
        `doc.theme` discovery via the ``list-themes`` surfaces) or a ``.thmx``
        file path. Wrap in `doc.edit(...)` for atomic undo. An unknown name
        raises `OpError`. Returns the resolved theme display name.
        """
        path = _resolve_theme_file(self._app(), theme, "theme")
        with _com.translate_com_errors():
            self._doc.com.ApplyDocumentTheme(path)
        return os.path.splitext(os.path.basename(path))[0]

    def list_available(self) -> dict[str, list[str]]:
        """The built-in themes, colour schemes, and font schemes Office ships.

        Returns ``{"themes": [...], "color_schemes": [...], "font_schemes":
        [...]}`` (names without extension) — the values `apply`, `set_colors`,
        and `set_fonts` accept. Empty lists if the library directory is absent.
        """
        with _com.translate_com_errors():
            return _list_builtin(self._app())

    def to_dict(self) -> dict[str, Any]:
        """The current theme as ``{colors, major_font, minor_font}``."""
        return {
            "colors": self.colors,
            "major_font": self.major_font,
            "minor_font": self.minor_font,
        }

    def __repr__(self) -> str:
        return f"<DocumentTheme major={self.major_font!r} minor={self.minor_font!r}>"
