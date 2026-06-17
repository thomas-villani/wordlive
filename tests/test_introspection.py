"""Non-visual layout introspection — anchor.location() and doc.stats().

Both repaginate first, then read computed values off COM (`Range.Information`,
`Document.ComputeStatistics`). The fake range's `Information` defaults to 1 for
every selector, and `ComputeStatistics` is an unset MagicMock, so each test
seeds the selectors/counters it asserts on.
"""

from __future__ import annotations

from typing import Any

from click.testing import CliRunner

import wordlive
from wordlive._ops import run_batch
from wordlive.cli.main import EXIT_ANCHOR_NOT_FOUND, EXIT_OK, main
from wordlive.constants import WdInformation as _Info
from wordlive.constants import WdStatistic as _Stat
from wordlive.exceptions import OpError
from wordlive.mcp._worker import InlineWorker
from wordlive.mcp.server import _read_impl, _write_impl

W = InlineWorker()


def _invoke(args: list[str]) -> tuple[int, str, str]:
    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout, result.stderr


def _info(page: int, *, line: int = 1, col: int = 1, in_table: int = 0) -> Any:
    """A `Range.Information(selector)` side-effect keyed by the Wd* selector."""
    table = {
        int(_Info.ACTIVE_END_PAGE_NUMBER): page,
        int(_Info.FIRST_CHARACTER_LINE_NUMBER): line,
        int(_Info.FIRST_CHARACTER_COLUMN_NUMBER): col,
        int(_Info.WITH_IN_TABLE): in_table,
    }
    return lambda selector: table[int(selector)]


def _seed_location(dc: Any, start: int, end: int, *, start_page: int, end_page: int) -> None:
    """Seed the three ranges location() reads: [start,start], [end,end], [start,end]."""
    dc.Range(start, start).Information.side_effect = _info(start_page, line=4, col=7)
    dc.Range(end, end).Information.side_effect = _info(end_page)
    dc.Range(start, end).Information.side_effect = _info(start_page, in_table=0)


def _seed_stats(dc: Any, *, saved: bool = False) -> None:
    counts = {
        int(_Stat.WORDS): 120,
        int(_Stat.LINES): 30,
        int(_Stat.PAGES): 3,
        int(_Stat.CHARACTERS): 600,
        int(_Stat.PARAGRAPHS): 18,
    }
    dc.ComputeStatistics.side_effect = lambda s: counts[int(s)]
    dc.Saved = saved


# ---------------------------------------------------------------------------
# anchor.location()
# ---------------------------------------------------------------------------


def test_location_reports_page_span_line_column(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        _seed_location(doc.com, 0, 10, start_page=1, end_page=2)
        loc = doc.range(0, 10).location()
    assert loc == {"page": 1, "end_page": 2, "line": 4, "column": 7, "in_table": False}


def test_location_repaginates_first(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        _seed_location(doc.com, 0, 10, start_page=1, end_page=1)
        doc.range(0, 10).location()
    fake_word.ActiveDocument.Repaginate.assert_called()


def test_location_preserves_saved_dirty_bit(fake_word):
    # Repaginate dirties the document in real Word; a pure read must not. Emulate
    # the flip and assert location() puts the Saved flag back.
    with wordlive.attach() as word:
        doc = word.documents.active
        dc = doc.com
        _seed_location(dc, 0, 10, start_page=1, end_page=1)
        dc.Saved = True
        dc.Repaginate.side_effect = lambda: setattr(dc, "Saved", False)
        doc.range(0, 10).location()
    assert dc.Saved is True


def test_location_in_table_flag(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        dc = doc.com
        dc.Range(0, 0).Information.side_effect = _info(1)
        dc.Range(10, 10).Information.side_effect = _info(1)
        dc.Range(0, 10).Information.side_effect = _info(1, in_table=1)
        loc = doc.range(0, 10).location()
    assert loc["in_table"] is True


# ---------------------------------------------------------------------------
# doc.stats()
# ---------------------------------------------------------------------------


def test_stats_text_counts_and_structure(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        _seed_stats(doc.com, saved=False)
        s = doc.stats()
    assert s == {
        "pages": 3,
        "words": 120,
        "characters": 600,
        "paragraphs": 18,
        "lines": 30,
        # Structural counts come from wordlive's own collections (fixture seed):
        "sections": 1,
        "headings": 2,
        "tables": 1,
        "images": 1,
        "equations": 1,
        "charts": 0,
        "comments": 0,
        "revisions": 1,
        "saved": False,
    }


def test_stats_repaginates_first(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        _seed_stats(doc.com)
        doc.stats()
    fake_word.ActiveDocument.Repaginate.assert_called()


def test_stats_preserves_saved_dirty_bit(fake_word):
    # `stats` repaginates, which flips Word's dirty bit; a read of a saved
    # document must report (and leave) saved=True, not a spurious unsaved star.
    with wordlive.attach() as word:
        doc = word.documents.active
        dc = doc.com
        _seed_stats(dc, saved=True)
        dc.Repaginate.side_effect = lambda: setattr(dc, "Saved", False)
        s = doc.stats()
    assert dc.Saved is True
    assert s["saved"] is True


# ---------------------------------------------------------------------------
# CLI — locate / stats
# ---------------------------------------------------------------------------


def test_cli_locate(fake_word):
    import json

    _seed_location(fake_word.ActiveDocument, 0, 10, start_page=1, end_page=2)
    code, out, _ = _invoke(["locate", "--anchor-id", "range:0-10"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["anchor_id"] == "range:0-10"
    assert data["page"] == 1 and data["end_page"] == 2


def test_cli_locate_missing_anchor_exit_2(fake_word):
    code, _out, _err = _invoke(["locate", "--anchor-id", "heading:99"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_cli_stats(fake_word):
    import json

    _seed_stats(fake_word.ActiveDocument, saved=False)
    code, out, _ = _invoke(["stats"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["pages"] == 3 and data["tables"] == 1 and data["saved"] is False


# ---------------------------------------------------------------------------
# MCP — word_read location / stats / table_records, word_write table records
# ---------------------------------------------------------------------------


def test_mcp_read_location(fake_word):
    _seed_location(fake_word.ActiveDocument, 0, 10, start_page=2, end_page=2)
    out = _read_impl(W, "location", {"anchor_id": "range:0-10"})
    assert out["anchor_id"] == "range:0-10" and out["page"] == 2


def test_mcp_read_stats(fake_word):
    _seed_stats(fake_word.ActiveDocument)
    out = _read_impl(W, "stats", {})
    assert out["pages"] == 3 and out["revisions"] == 1


def test_mcp_read_table_records(fake_word):
    out = _read_impl(W, "table_records", {"table": 1})
    assert out == [{"A1": "A2", "B1": "B2"}]


def test_mcp_table_append_record_and_update_row(fake_word):
    r1 = _write_impl(
        W, "table", {"action": "append_record", "table": 1, "record": {"A1": "X", "B1": "Y"}}
    )
    assert r1["ok"] is True
    r2 = _write_impl(
        W,
        "table",
        {"action": "update_row", "table": 1, "key": "A2", "values": {"B1": "Z"}},
    )
    assert r2["ok"] is True
    with wordlive.attach() as word:
        t = word.documents.active.tables[1]
        assert t.cell(3, 1).text == "X" and t.cell(2, 2).text == "Z"


# ---------------------------------------------------------------------------
# exec op error semantics carry through the batch layer
# ---------------------------------------------------------------------------


def test_exec_update_row_unknown_header_fails_batch(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "update_row", "table": 1, "key": "A2", "values": {"Nope": "Z"}}],
            label="t",
        )
    assert isinstance(exc, OpError)
    assert result["ok"] is False


# ---------------------------------------------------------------------------
# Structural query helpers — doc.between / nearest_heading / find_paragraphs
#
# The default fake doc has three paragraphs (see conftest.fake_word):
#   heading:1 "Introduction" (L1)  range 0–13
#   para:2    "Body text here." (body) range 13–29
#   heading:3 "Risks" (L2)        range 29–35
# ---------------------------------------------------------------------------


def test_between_excludes_bounding_headings(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.com.Range(13, 29).Text = "Body text here.\r"
        span = doc.between("heading:1", "heading:3")
    assert (span.start, span.end) == (13, 29)
    assert span.anchor_id == "range:13-29"
    assert span.text == "Body text here.\r"


def test_between_inclusive_covers_both_headings(fake_word):
    with wordlive.attach() as word:
        span = word.documents.active.between("heading:1", "heading:3", inclusive=True)
    assert (span.start, span.end) == (0, 35)


def test_between_out_of_order_raises(fake_word):
    import pytest

    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError, match="begins before"):
            doc.between("heading:3", "heading:1")


def test_nearest_heading_before_is_enclosing(fake_word):
    with wordlive.attach() as word:
        row = word.documents.active.nearest_heading("para:2", direction="before")
    assert row == {"level": 1, "text": "Introduction", "anchor_id": "heading:1"}


def test_nearest_heading_after_is_next(fake_word):
    with wordlive.attach() as word:
        row = word.documents.active.nearest_heading("para:2", direction="after")
    assert row["anchor_id"] == "heading:3" and row["level"] == 2


def test_nearest_heading_none_past_last(fake_word):
    with wordlive.attach() as word:
        row = word.documents.active.nearest_heading("heading:3", direction="after")
    assert row is None


def test_nearest_heading_bad_direction_raises(fake_word):
    import pytest

    with wordlive.attach() as word:
        with pytest.raises(OpError, match="before.*after"):
            word.documents.active.nearest_heading("para:2", direction="sideways")


def test_find_paragraphs_ranks_best_first(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.find_paragraphs("Body text here", min_score=0.5)
    assert rows[0]["anchor_id"] == "para:2"
    assert rows[0]["index"] == 2 and rows[0]["is_heading"] is False
    assert rows == sorted(rows, key=lambda r: r["score"], reverse=True)


def test_find_paragraphs_min_score_filters(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.find_paragraphs("Risks", min_score=0.99)
    # find_paragraphs always addresses by para:N; level/is_heading flag headings.
    assert [r["anchor_id"] for r in rows] == ["para:3"]
    assert rows[0]["score"] == 1.0 and rows[0]["level"] == 2 and rows[0]["is_heading"] is True


def test_find_paragraphs_limit_caps(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.find_paragraphs("Risks", limit=1, min_score=0.0)
    assert len(rows) == 1


def test_find_paragraphs_empty_query_returns_empty(fake_word):
    with wordlive.attach() as word:
        assert word.documents.active.find_paragraphs("   ") == []


def test_find_paragraphs_normalizes_query(fake_word):
    # Smart quotes + em-dash in the paragraph fold onto the straight-ASCII query.
    dc = fake_word.ActiveDocument
    dc.Paragraphs._items[1].Range.Text = "“Smart” quotes—dash\r"
    with wordlive.attach() as word:
        rows = word.documents.active.find_paragraphs('"Smart" quotes-dash')
    assert rows[0]["anchor_id"] == "para:2" and rows[0]["score"] == 1.0


def test_find_paragraphs_bad_args_raise(fake_word):
    import pytest

    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.find_paragraphs("x", limit=0)
        with pytest.raises(OpError):
            doc.find_paragraphs("x", min_score=2.0)


# CLI ----------------------------------------------------------------------


def test_cli_read_between(fake_word):
    import json

    fake_word.ActiveDocument.Range(13, 29).Text = "Body text here.\r"
    code, out, _ = _invoke(["read", "between", "--start", "heading:1", "--end", "heading:3"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["anchor_id"] == "range:13-29" and data["inclusive"] is False


def test_cli_read_nearest_heading(fake_word):
    import json

    code, out, _ = _invoke(
        ["read", "nearest-heading", "--anchor-id", "para:2", "--direction", "before"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["heading"]["anchor_id"] == "heading:1" and data["direction"] == "before"


def test_cli_find_paragraph(fake_word):
    import json

    code, out, _ = _invoke(["find-paragraph", "--text", "Risks", "--min-score", "0.99"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert [r["anchor_id"] for r in data] == ["para:3"]


# MCP ----------------------------------------------------------------------


def test_mcp_read_between(fake_word):
    out = _read_impl(W, "between", {"start_anchor": "heading:1", "end_anchor": "heading:3"})
    assert out["anchor_id"] == "range:13-29" and out["start"] == "heading:1"


def test_mcp_read_nearest_heading(fake_word):
    out = _read_impl(W, "nearest_heading", {"anchor_id": "para:2", "direction": "after"})
    assert out["heading"]["anchor_id"] == "heading:3"


def test_mcp_read_nearest_heading_default_direction(fake_word):
    out = _read_impl(W, "nearest_heading", {"anchor_id": "para:2"})
    assert out["direction"] == "before" and out["heading"]["anchor_id"] == "heading:1"


def test_mcp_find_paragraphs(fake_word):
    out = _read_impl(W, "find_paragraphs", {"text": "Risks", "min_score": 0.99})
    assert [r["anchor_id"] for r in out] == ["para:3"]
