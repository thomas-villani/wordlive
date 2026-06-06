"""Unit tests for the colour/units coercion helper (`wordlive._format`).

Pure — no COM, runs everywhere.
"""

from __future__ import annotations

import math

import pytest

from wordlive._format import to_bgr, to_points


class TestToBgr:
    def test_hex_with_hash(self) -> None:
        # #FF0000 is pure red -> BGR long has red in the low byte.
        assert to_bgr("#FF0000") == 255

    def test_hex_without_hash(self) -> None:
        # 0000FF is pure blue -> red/blue swapped vs. the hex digits.
        assert to_bgr("0000FF") == 16711680

    def test_hex_green(self) -> None:
        assert to_bgr("00FF00") == 65280

    def test_rgb_tuple_white(self) -> None:
        assert to_bgr((255, 255, 255)) == 16777215

    def test_rgb_list(self) -> None:
        assert to_bgr([255, 0, 0]) == 255

    def test_named_red(self) -> None:
        assert to_bgr("red") == 255

    def test_named_yellow(self) -> None:
        assert to_bgr("yellow") == 65535

    def test_named_case_insensitive(self) -> None:
        assert to_bgr("  Navy  ") == to_bgr("navy")

    def test_named_aliases(self) -> None:
        assert to_bgr("cyan") == to_bgr("aqua")
        assert to_bgr("magenta") == to_bgr("fuchsia")

    def test_unknown_name_raises_value_error(self) -> None:
        with pytest.raises(ValueError):
            to_bgr("burnt-sienna")

    def test_bad_hex_length_raises_value_error(self) -> None:
        with pytest.raises(ValueError):
            to_bgr("#FFF")

    def test_component_out_of_range_raises_value_error(self) -> None:
        with pytest.raises(ValueError):
            to_bgr((256, 0, 0))

    def test_wrong_tuple_arity_raises_value_error(self) -> None:
        with pytest.raises(ValueError):
            to_bgr((255, 0))

    def test_bare_int_raises_type_error(self) -> None:
        # A bare int is ambiguous (RGB vs BGR) — reject it.
        with pytest.raises(TypeError):
            to_bgr(255)  # type: ignore[arg-type]

    def test_bool_component_raises_type_error(self) -> None:
        with pytest.raises(TypeError):
            to_bgr((True, 0, 0))  # type: ignore[arg-type]


class TestToPoints:
    def test_bare_number_is_points(self) -> None:
        assert to_points(36) == 36.0

    def test_float_passthrough(self) -> None:
        assert to_points(36.5) == 36.5

    def test_inches(self) -> None:
        assert to_points("1in") == 72.0

    def test_centimetres(self) -> None:
        assert math.isclose(to_points("2.54cm"), 72.0)

    def test_millimetres(self) -> None:
        assert math.isclose(to_points("10mm"), 28.3464567, rel_tol=1e-6)

    def test_explicit_points(self) -> None:
        assert to_points("18pt") == 18.0

    def test_unitless_string_is_points(self) -> None:
        assert to_points("36") == 36.0

    def test_case_insensitive_unit(self) -> None:
        assert to_points("1IN") == 72.0

    def test_unknown_unit_raises_value_error(self) -> None:
        with pytest.raises(ValueError):
            to_points("18furlongs")

    def test_unparseable_raises_value_error(self) -> None:
        with pytest.raises(ValueError):
            to_points("wide")

    def test_bool_raises_type_error(self) -> None:
        with pytest.raises(TypeError):
            to_points(True)
