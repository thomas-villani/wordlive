"""Inline runs — the markdown-sugar parser and block-item normalisation.

Pure logic (`wordlive._runs`), no COM, so these run anywhere.
"""

from __future__ import annotations

import pytest

from wordlive._runs import (
    Run,
    normalize_block_items,
    parse_markup,
    runs_from_payload,
    runs_to_text,
)
from wordlive.exceptions import OpError


def _triples(text: str) -> list[tuple[str, bool | None, bool | None]]:
    return [(r.text, r.bold, r.italic) for r in parse_markup(text)]


class TestParseMarkup:
    def test_plain_text_is_one_unformatted_run(self):
        assert _triples("just plain text") == [("just plain text", None, None)]

    def test_bold_lead_in(self):
        assert _triples("**Fast** — sub-ms reads") == [
            ("Fast", True, None),
            (" — sub-ms reads", None, None),
        ]

    def test_italic(self):
        assert _triples("an *emphatic* word") == [
            ("an ", None, None),
            ("emphatic", None, True),
            (" word", None, None),
        ]

    def test_bold_and_italic_combined(self):
        assert _triples("***both*** ends") == [("both", True, True), (" ends", None, None)]

    def test_mixed_bold_and_italic(self):
        assert _triples("*i* and **b**") == [
            ("i", None, True),
            (" and ", None, None),
            ("b", True, None),
        ]

    def test_triple_star_wins_over_double_at_same_start(self):
        # *** must be tried before ** / * so it isn't mis-parsed as italic("*…").
        runs = parse_markup("***x***")
        assert runs == [Run(text="x", bold=True, italic=True)]

    def test_unmatched_single_star_stays_literal(self):
        assert _triples("a * b") == [("a * b", None, None)]
        assert _triples("see *.txt files") == [("see *.txt files", None, None)]

    def test_escaped_asterisk_is_literal(self):
        assert _triples(r"5 \* 3 = 15") == [("5 * 3 = 15", None, None)]

    def test_escaped_backslash_is_literal(self):
        assert _triples(r"path\\to") == [(r"path\to", None, None)]

    def test_escaped_star_inside_emphasis(self):
        # The escaped star is literal; the surrounding ** still bolds.
        assert _triples(r"**a\*b**") == [("a*b", True, None)]

    def test_empty_string_yields_one_empty_run(self):
        assert parse_markup("") == [Run(text="")]


class TestParseCodeSpans:
    """`` `code` `` — the span agents reach for on every identifier and path.

    Before this existed, backticks landed in the document as literal characters
    and `to_markdown` escaped them back out as ``\\```.
    """

    def test_code_span_becomes_a_code_run(self):
        assert parse_markup("use `attach()` now") == [
            Run(text="use "),
            Run(text="attach()", code=True),
            Run(text=" now"),
        ]

    def test_code_content_is_literal_not_emphasis(self):
        # Code binds tighter than `*`, so the stars inside stay text.
        assert parse_markup("`b*c`") == [Run(text="b*c", code=True)]

    def test_emphasis_does_not_reach_across_a_code_span(self):
        # Documented limitation of the flat run model: the asterisks land in
        # different segments and stay literal.
        assert parse_markup("*a `b` c*") == [
            Run(text="*a "),
            Run(text="b", code=True),
            Run(text=" c*"),
        ]

    def test_code_coexists_with_emphasis_in_separate_segments(self):
        assert parse_markup("`x` and **b**") == [
            Run(text="x", code=True),
            Run(text=" and "),
            Run(text="b", bold=True),
        ]

    def test_longer_fence_embeds_a_backtick(self):
        assert parse_markup("`` `x` ``") == [Run(text="`x`", code=True)]

    def test_escaped_backtick_does_not_open_a_span(self):
        assert parse_markup(r"escaped \`not code\`") == [Run(text="escaped `not code`")]

    def test_unmatched_backtick_stays_literal(self):
        assert parse_markup("a ` b") == [Run(text="a ` b")]

    def test_code_run_is_formatted(self):
        # Drives whether insert_block does a COM round-trip for the span at all.
        assert Run(text="x", code=True).formatted() is True
        assert Run(text="x").formatted() is False


class TestRunsFromPayload:
    def test_text_parses_markdown(self):
        assert runs_from_payload(text="**hi**") == [Run(text="hi", bold=True)]

    def test_structured_runs_pass_through(self):
        runs = runs_from_payload(runs=[{"text": "x", "bold": True, "style": "Strong"}])
        assert runs == [Run(text="x", bold=True, style="Strong")]

    def test_both_text_and_runs_is_error(self):
        with pytest.raises(OpError, match="exactly one of 'text' or 'runs'"):
            runs_from_payload(text="a", runs=[{"text": "b"}])

    def test_neither_text_nor_runs_is_error(self):
        with pytest.raises(OpError, match="exactly one of 'text' or 'runs'"):
            runs_from_payload()

    def test_empty_runs_list_is_error(self):
        with pytest.raises(OpError, match="non-empty list"):
            runs_from_payload(runs=[])

    def test_run_missing_text_is_error(self):
        with pytest.raises(OpError, match="string 'text'"):
            runs_from_payload(runs=[{"bold": True}])

    def test_run_unknown_field_is_error(self):
        with pytest.raises(OpError, match="unknown field"):
            runs_from_payload(runs=[{"text": "x", "colour": "red"}])

    def test_run_non_bool_flag_is_error(self):
        with pytest.raises(OpError, match="'bold' must be a boolean"):
            runs_from_payload(runs=[{"text": "x", "bold": "yes"}])


class TestNormalizeBlockItems:
    def test_string_item_is_plain_paragraph(self):
        assert normalize_block_items(["plain"]) == [([Run(text="plain")], None)]

    def test_dict_with_text_and_style(self):
        out = normalize_block_items([{"text": "**A** b", "style": "List Bullet"}])
        assert out == [([Run(text="A", bold=True), Run(text=" b")], "List Bullet")]

    def test_dict_with_runs(self):
        out = normalize_block_items([{"runs": [{"text": "x", "italic": True}]}])
        assert out == [([Run(text="x", italic=True)], None)]

    def test_mixed_item_shapes(self):
        out = normalize_block_items(["a", {"text": "b"}, {"runs": [{"text": "c"}]}])
        assert [s for _, s in out] == [None, None, None]
        assert [runs_to_text(r) for r, _ in out] == ["a", "b", "c"]

    def test_empty_items_is_error(self):
        with pytest.raises(OpError, match="non-empty 'items' list"):
            normalize_block_items([])

    def test_non_list_items_is_error(self):
        with pytest.raises(OpError, match="non-empty 'items' list"):
            normalize_block_items({"text": "x"})

    def test_item_unknown_field_is_error(self):
        with pytest.raises(OpError, match="unknown field"):
            normalize_block_items([{"text": "x", "list": "bulleted"}])

    def test_item_both_text_and_runs_is_error(self):
        with pytest.raises(OpError, match="exactly one of 'text' or 'runs'"):
            normalize_block_items([{"text": "a", "runs": [{"text": "b"}]}])

    def test_bad_style_type_is_error(self):
        with pytest.raises(OpError, match="'style' must be a string"):
            normalize_block_items([{"text": "x", "style": 7}])


def test_runs_to_text_concatenates():
    assert runs_to_text(parse_markup("**Fast** — go")) == "Fast — go"
