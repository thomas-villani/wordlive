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
