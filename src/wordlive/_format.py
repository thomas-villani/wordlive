"""Colour and length coercion — the shared plumbing under character/run, shading,
border, and tab-stop formatting.

Two pure helpers, no COM:

- [`to_bgr`][wordlive._format.to_bgr] turns a colour (named / hex / RGB tuple)
  into the **BGR long** integer Word stores. Word's colour properties
  (`Font.Color`, `Shading.BackgroundPatternColor`, `Borders(...).Color`) are an
  `RGB(r, g, b)` value *byte-swapped* — red is `0x0000FF`, not `0xFF0000` — so a
  naive `0xRRGGBB` int comes out with red and blue transposed. Always route a
  colour through here.
- [`to_points`][wordlive._format.to_points] coerces a length to points (Word's
  native unit for sizes, indents, spacing, tab positions). Bare numbers are
  already points; strings carry an explicit unit (`pt`/`in`/`cm`/`mm`).

Internal — not exported through `__all__`, mirroring how `constants` enums stay
private. Bad input raises `ValueError`/`TypeError`; the public anchor/style
methods translate those into `OpError` (exit 1, bad-input) so a batch reports
cleanly instead of crashing the op loop.
"""

from __future__ import annotations

import re

# The 16 HTML/CSS basic colours (+ the two common aliases). Kept deliberately
# small — extend only as a real need appears, mirroring the narrow-enum policy in
# `constants.py`. Values are (r, g, b).
_NAMED_COLORS: dict[str, tuple[int, int, int]] = {
    "black": (0, 0, 0),
    "silver": (192, 192, 192),
    "gray": (128, 128, 128),
    "grey": (128, 128, 128),
    "white": (255, 255, 255),
    "maroon": (128, 0, 0),
    "red": (255, 0, 0),
    "purple": (128, 0, 128),
    "fuchsia": (255, 0, 255),
    "magenta": (255, 0, 255),
    "green": (0, 128, 0),
    "lime": (0, 255, 0),
    "olive": (128, 128, 0),
    "yellow": (255, 255, 0),
    "navy": (0, 0, 128),
    "blue": (0, 0, 255),
    "teal": (0, 128, 128),
    "aqua": (0, 255, 255),
    "cyan": (0, 255, 255),
}

_HEX_RE = re.compile(r"^#?[0-9a-fA-F]{6}$")

# value + optional unit; bare number => points.
_LENGTH_RE = re.compile(r"^\s*([+-]?\d+(?:\.\d+)?)\s*(pt|in|cm|mm)?\s*$", re.IGNORECASE)

# Conversion factors to points.
_TO_POINTS: dict[str, float] = {
    "pt": 1.0,
    "in": 72.0,
    "cm": 72.0 / 2.54,
    "mm": 72.0 / 25.4,
}


def rgb_to_bgr(r: int, g: int, b: int) -> int:
    """Pack an (r, g, b) triple into Word's BGR long integer."""
    for name, value in (("r", r), ("g", g), ("b", b)):
        if isinstance(value, bool) or not isinstance(value, int):
            raise TypeError(f"colour component {name} must be an int 0-255; got {value!r}")
        if not 0 <= value <= 255:
            raise ValueError(f"colour component {name} out of range 0-255: {value}")
    return (b << 16) | (g << 8) | r


def to_bgr(value: str | tuple[int, int, int] | list[int]) -> int:
    """Coerce a colour to the **BGR long** integer Word stores.

    Accepts a named colour (`"red"`, `"navy"`, …; the 16 HTML/CSS basics), a hex
    string (`"#FF0000"` or `"FF0000"`), or an `(r, g, b)` tuple/list. Returns
    `(b << 16) | (g << 8) | r` — e.g. red `#FF0000` -> `255`, blue -> `16711680`.

    Raises `ValueError` for an unknown name, malformed hex, or an out-of-range
    component; `TypeError` for an unsupported type (a bare `int` or `bool` is
    rejected — pass an explicit form so the byte-order is unambiguous).
    """
    if isinstance(value, str):
        key = value.strip().lower()
        if key in _NAMED_COLORS:
            return rgb_to_bgr(*_NAMED_COLORS[key])
        if _HEX_RE.match(value.strip()):
            h = value.strip().lstrip("#")
            return rgb_to_bgr(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
        raise ValueError(
            f"unknown colour {value!r}; expected a named colour "
            f"({', '.join(sorted(set(_NAMED_COLORS)))}), a hex string like '#FF0000', "
            "or an (r, g, b) tuple"
        )
    if isinstance(value, (tuple, list)):
        if len(value) != 3:
            raise ValueError(f"colour tuple must have 3 components (r, g, b); got {len(value)}")
        return rgb_to_bgr(*value)
    raise TypeError(
        f"colour must be a name, hex string, or (r, g, b) tuple; got {type(value).__name__}"
    )


def to_points(value: int | float | str) -> float:
    """Coerce a length to points (Word's native unit).

    A bare number is taken as points already. A string carries an explicit unit:
    `"36pt"`, `"1in"`, `"2cm"`, `"10mm"` (a unitless string like `"36"` is also
    points). 1in = 72pt, 1cm = 72/2.54pt, 1mm = 72/25.4pt.

    Raises `ValueError` for an unparseable string or unknown unit; `TypeError`
    for an unsupported type (a `bool` is rejected).
    """
    if isinstance(value, bool):
        raise TypeError("length must be a number or unit string, not bool")
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        m = _LENGTH_RE.match(value)
        if m is None:
            raise ValueError(
                f"cannot parse length {value!r}; expected a number optionally suffixed "
                "with pt/in/cm/mm (e.g. '36', '1in', '2.5cm')"
            )
        magnitude = float(m.group(1))
        unit = (m.group(2) or "pt").lower()
        return magnitude * _TO_POINTS[unit]
    raise TypeError(f"length must be a number or unit string; got {type(value).__name__}")
