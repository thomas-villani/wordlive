"""CLI shape: JSON output and exit codes."""

from __future__ import annotations

import json
from pathlib import Path

from click.testing import CliRunner

from wordlive.cli.main import (
    EXIT_AMBIGUOUS_MATCH,
    EXIT_ANCHOR_NOT_FOUND,
    EXIT_OK,
    EXIT_OTHER,
    EXIT_WORD_NOT_RUNNING,
    main,
)


def _invoke(args: list[str], *, input: str | None = None) -> tuple[int, str, str]:
    runner = CliRunner()
    result = runner.invoke(main, args, input=input, catch_exceptions=False)
    return result.exit_code, result.stdout, result.stderr


def test_status_lists_active_doc(fake_word):
    code, out, _ = _invoke(["status"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert isinstance(data, list)
    assert data[0]["name"] == "Test.docx"
    assert data[0]["is_active"] is True


def test_outline_returns_headings(fake_word):
    code, out, _ = _invoke(["outline"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert [item["text"] for item in data] == ["Introduction", "Risks"]


def test_read_bookmark_success(fake_word):
    code, out, _ = _invoke(["read", "bookmark", "Address"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert "text" in data


def test_read_bookmark_missing(fake_word):
    code, _, err = _invoke(["read", "bookmark", "DoesNotExist"])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "bookmark" in err.lower()


def test_read_cc_success(fake_word):
    code, out, _ = _invoke(["read", "cc", "Signatory"])
    assert code == EXIT_OK
    assert json.loads(out) == {"text": "Jane Doe"}


def test_write_bookmark(fake_word):
    code, out, _ = _invoke(["write", "bookmark", "Address", "--text", "123 Main"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor"]["kind"] == "bookmark"


def test_insert_after_anchor_id(fake_word):
    code, out, _ = _invoke(["insert", "--anchor-id", "heading:1", "--text", "new para"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor_id"] == "heading:1"
    assert data["where"] == "after"


def test_insert_before_anchor_id(fake_word):
    code, out, _ = _invoke(["insert", "--anchor-id", "para:2", "--text", "intro", "--before"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["where"] == "before"
    # --before inserts at the paragraph's start offset (para:2 -> 13).
    assert fake_word.ActiveDocument.Range(13, 13).Text == "intro\r"


def test_insert_missing_anchor(fake_word):
    code, _, err = _invoke(["insert", "--anchor-id", "heading:99", "--text", "x"])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "heading" in err.lower()


def test_insert_break_page_after(fake_word):
    code, out, _ = _invoke(["insert-break", "--anchor-id", "bookmark:Address"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data == {"ok": True, "anchor_id": "bookmark:Address", "kind": "page", "where": "after"}
    # Address ends at 24; the page break is inserted there (wdPageBreak = 7).
    fake_word.ActiveDocument.Range(24, 24).InsertBreak.assert_called_once_with(Type=7)


def test_insert_break_section_before(fake_word):
    code, out, _ = _invoke(
        ["insert-break", "--anchor-id", "bookmark:Address", "--kind", "section_next", "--before"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["kind"] == "section_next"
    assert data["where"] == "before"
    # Address starts at 13; wdSectionBreakNextPage = 2.
    fake_word.ActiveDocument.Range(13, 13).InsertBreak.assert_called_once_with(Type=2)


def test_insert_break_bad_kind_is_usage_error(fake_word):
    code, _, _ = _invoke(["insert-break", "--anchor-id", "bookmark:Address", "--kind", "nope"])
    assert code != EXIT_OK


def test_insert_break_missing_anchor_returns_exit_2(fake_word):
    code, _, err = _invoke(["insert-break", "--anchor-id", "bookmark:Nope"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_format_paragraph_page_break_before(fake_word):
    code, out, _ = _invoke(
        ["format-paragraph", "--anchor-id", "bookmark:Address", "--page-break-before"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["applied"]["page_break_before"] is True
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert pf.PageBreakBefore is True


def test_append_paragraph_default(fake_word):
    code, out, _ = _invoke(["append", "--text", "Closing note."])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data == {"ok": True, "mode": "paragraph", "style": None}
    # New paragraph written as "<break><text>" just before the final mark (34).
    assert fake_word.ActiveDocument.Range(34, 34).Text == "\rClosing note."


def test_append_inline(fake_word):
    code, out, _ = _invoke(["append", "--text", " (verified)", "--inline"])
    assert code == EXIT_OK
    assert json.loads(out)["mode"] == "inline"
    fake_word.ActiveDocument.Content.InsertAfter.assert_called_once_with(" (verified)")


def test_append_with_style(fake_word):
    code, out, _ = _invoke(["append", "--text", "Body", "--style", "Body Text"])
    assert code == EXIT_OK
    assert json.loads(out)["style"] == "Body Text"


def test_append_inline_with_style_is_usage_error(fake_word):
    code, _, err = _invoke(["append", "--text", "x", "--inline", "--style", "Body Text"])
    assert code != EXIT_OK
    assert "paragraph mode" in err.lower()


def test_insert_anchor_id_end_appends(fake_word):
    """The `end` anchor resolves through --anchor-id like any other."""
    code, out, _ = _invoke(["insert", "--anchor-id", "end", "--text", "Tail."])
    assert code == EXIT_OK
    assert json.loads(out)["anchor_id"] == "end"
    assert fake_word.ActiveDocument.Range(34, 34).Text == "\rTail."


def test_prepend_paragraph_default(fake_word):
    code, out, _ = _invoke(["prepend", "--text", "DRAFT"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data == {"ok": True, "mode": "paragraph", "style": None}
    # New first paragraph written as "<text><break>" at offset 0.
    assert fake_word.ActiveDocument.Range(0, 0).Text == "DRAFT\r"


def test_prepend_inline(fake_word):
    code, out, _ = _invoke(["prepend", "--text", "Note: ", "--inline"])
    assert code == EXIT_OK
    assert json.loads(out)["mode"] == "inline"
    fake_word.ActiveDocument.Content.InsertBefore.assert_called_once_with("Note: ")


def test_prepend_inline_with_style_is_usage_error(fake_word):
    code, _, err = _invoke(["prepend", "--text", "x", "--inline", "--style", "Heading 1"])
    assert code != EXIT_OK
    assert "paragraph mode" in err.lower()


def test_insert_anchor_id_start_prepends(fake_word):
    code, out, _ = _invoke(["insert", "--anchor-id", "start", "--text", "Head."])
    assert code == EXIT_OK
    assert json.loads(out)["anchor_id"] == "start"
    assert fake_word.ActiveDocument.Range(0, 0).Text == "Head.\r"


def test_word_not_running_returns_exit_4(no_word):
    code, _, err = _invoke(["status"])
    assert code == EXIT_WORD_NOT_RUNNING
    assert "not running" in err.lower() or "word" in err.lower()


# ---------------------------------------------------------------------------
# v0.1: replace / go-to / exec
# ---------------------------------------------------------------------------


def test_replace_by_heading_id(fake_word):
    code, out, _ = _invoke(["replace", "--anchor-id", "heading:1", "--text", "New intro"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor"]["kind"] == "heading"


def test_replace_by_bookmark_id(fake_word):
    code, out, _ = _invoke(["replace", "--anchor-id", "bookmark:Address", "--text", "123 Main"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor"]["kind"] == "bookmark"
    assert data["anchor"]["name"] == "Address"


def test_replace_missing_anchor(fake_word):
    code, _, err = _invoke(["replace", "--anchor-id", "bookmark:Nope", "--text", "..."])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "bookmark" in err.lower()


def test_replace_bad_scheme(fake_word):
    code, _, _ = _invoke(["replace", "--anchor-id", "table:1", "--text", "..."])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_go_to_heading(fake_word):
    code, out, _ = _invoke(["go-to", "--anchor-id", "heading:1"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor_id"] == "heading:1"


def test_go_to_missing_anchor(fake_word):
    code, _, _ = _invoke(["go-to", "--anchor-id", "bookmark:Nope"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_exec_all_ops_succeed(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "label": "Batch update",
                "ops": [
                    {"op": "write_bookmark", "name": "Address", "text": "123 Main"},
                    {"op": "write_cc", "name": "Signatory", "text": "Jane Doe"},
                    {"op": "insert_paragraph", "anchor_id": "heading:1", "text": "New para"},
                    {"op": "replace", "anchor_id": "heading:3", "text": "Updated Risks"},
                ],
            }
        ),
        encoding="utf-8",
    )

    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["ops_run"] == 4
    assert data["label"] == "Batch update"


def test_exec_supports_append_paragraph_op(fake_word):
    payload = json.dumps(
        {"ops": [{"op": "append_paragraph", "text": "Tail.", "style": "Body Text"}]}
    )
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    assert fake_word.ActiveDocument.Range(34, 34).Text == "\rTail."


def test_exec_append_op_creates_a_new_paragraph(fake_word):
    # `append` now means a new final paragraph (matching its description and
    # `append_paragraph`), not inline concatenation.
    payload = json.dumps({"ops": [{"op": "append", "text": "Tail.", "style": "Body Text"}]})
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    assert fake_word.ActiveDocument.Range(34, 34).Text == "\rTail."


def test_exec_supports_append_inline_op(fake_word):
    payload = json.dumps({"ops": [{"op": "append_inline", "text": " more"}]})
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    fake_word.ActiveDocument.Content.InsertAfter.assert_called_once_with(" more")


def test_exec_append_inline_with_style_warns(fake_word):
    # An inline append takes no style; the ignored field surfaces as a warning
    # rather than silently vanishing.
    payload = json.dumps({"ops": [{"op": "append_inline", "text": "x", "style": "Heading 1"}]})
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ops_run"] == 1
    assert any(w["field"] == "style" for w in data.get("warnings", []))


def test_exec_append_paragraph_missing_text_reports_cleanly(fake_word):
    code, _, err = _invoke(["exec", "--ops", json.dumps({"ops": [{"op": "append_paragraph"}]})])
    assert code != EXIT_OK
    assert "text" in err.lower()


def test_exec_supports_prepend_paragraph_op(fake_word):
    payload = json.dumps(
        {"ops": [{"op": "prepend_paragraph", "text": "Title", "style": "Heading 1"}]}
    )
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    assert fake_word.ActiveDocument.Range(0, 0).Text == "Title\r"


def test_exec_prepend_op_creates_a_new_paragraph(fake_word):
    payload = json.dumps({"ops": [{"op": "prepend", "text": "Title", "style": "Heading 1"}]})
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    assert fake_word.ActiveDocument.Range(0, 0).Text == "Title\r"


def test_exec_supports_prepend_inline_op(fake_word):
    payload = json.dumps({"ops": [{"op": "prepend_inline", "text": "Note: "}]})
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    fake_word.ActiveDocument.Content.InsertBefore.assert_called_once_with("Note: ")


def test_exec_unknown_field_warns_but_succeeds(fake_word):
    # A typo'd / inapplicable field is reported, not silently dropped — the op
    # still applies.
    payload = json.dumps(
        {"ops": [{"op": "write_bookmark", "name": "Address", "text": "X", "anchorid": "oops"}]}
    )
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ops_run"] == 1
    assert any(w["field"] == "anchorid" for w in data.get("warnings", []))


def test_exec_inline_ops_json(fake_word):
    """`--ops '{...}'` applies a batch without needing a file on disk."""
    payload = json.dumps(
        {
            "label": "Inline",
            "ops": [{"op": "write_bookmark", "name": "Address", "text": "Inline St"}],
        }
    )
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["ops_run"] == 1
    assert data["label"] == "Inline"
    assert fake_word.ActiveDocument.Bookmarks._items["Address"][0] == 13


def test_exec_inline_ops_accepts_bare_array(fake_word):
    """A bare `[...]` array is shorthand for `{"ops": [...]}`."""
    payload = json.dumps([{"op": "write_bookmark", "name": "Address", "text": "X"}])
    code, out, _ = _invoke(["exec", "--ops", payload])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1


def test_exec_ops_from_stdin(fake_word):
    """`--ops -` reads the batch JSON from stdin."""
    payload = json.dumps({"ops": [{"op": "write_bookmark", "name": "Address", "text": "Piped"}]})
    code, out, _ = _invoke(["exec", "--ops", "-"], input=payload)
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1


def test_exec_requires_exactly_one_source(fake_word, tmp_path: Path):
    """Neither --script nor --ops, or both, is a usage error."""
    code, _, err = _invoke(["exec"])
    assert code != EXIT_OK
    assert "--script" in err and "--ops" in err

    script = tmp_path / "ops.json"
    script.write_text(json.dumps({"ops": []}), encoding="utf-8")
    code, _, err = _invoke(["exec", "--script", str(script), "--ops", "{}"])
    assert code != EXIT_OK
    assert "exactly one" in err


def test_exec_inline_malformed_json_is_clean_error(fake_word):
    """Malformed inline JSON surfaces as a clean error, not a traceback."""
    code, _, err = _invoke(["exec", "--ops", "{not json"])
    assert code != EXIT_OK
    assert "malformed" in err.lower()
    assert "Traceback" not in err


def test_exec_wraps_ops_in_single_undo_record(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "write_bookmark", "name": "Address", "text": "A"},
                    {"op": "write_bookmark", "name": "Address", "text": "B"},
                ]
            }
        ),
        encoding="utf-8",
    )

    code, _, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    # All ops should ride a single UndoRecord — Start/End called exactly once each.
    fake_word.UndoRecord.StartCustomRecord.assert_called_once()
    fake_word.UndoRecord.EndCustomRecord.assert_called_once()


def test_exec_stops_at_first_failure_and_reports_partial(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "write_bookmark", "name": "Address", "text": "ok-1"},
                    {"op": "write_bookmark", "name": "Nope", "text": "boom"},
                    {"op": "write_bookmark", "name": "Address", "text": "never-runs"},
                ]
            }
        ),
        encoding="utf-8",
    )

    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_ANCHOR_NOT_FOUND
    data = json.loads(out)
    assert data["ok"] is False
    assert data["ops_run"] == 1
    assert data["failure"]["index"] == 1
    assert data["failure"]["type"] == "AnchorNotFoundError"


def test_exec_insert_paragraph_honours_boolean_before(fake_word, tmp_path: Path):
    """Regression: the `insert_paragraph` op must honour `"before": true`, not
    just the verbose `"where": "before"`. The boolean mirrors the CLI's
    `--before/--after` flags, so an LLM that encodes the batch op the same way
    it would type the command gets a before-insert — not a silent after-insert.
    """
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {
                        "op": "insert_paragraph",
                        "anchor_id": "para:2",
                        "text": "intro",
                        "before": True,
                    }
                ]
            }
        ),
        encoding="utf-8",
    )

    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True
    # --before inserts at para:2's start offset (13), as a new paragraph.
    assert fake_word.ActiveDocument.Range(13, 13).Text == "intro\r"


def test_exec_insert_paragraph_where_before_still_works(fake_word, tmp_path: Path):
    """The original `"where": "before"` form must keep working alongside the
    boolean."""
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {
                        "op": "insert_paragraph",
                        "anchor_id": "para:2",
                        "text": "intro",
                        "where": "before",
                    }
                ]
            }
        ),
        encoding="utf-8",
    )

    code, _, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    assert fake_word.ActiveDocument.Range(13, 13).Text == "intro\r"


# ---------------------------------------------------------------------------
# --text human-readable output
# ---------------------------------------------------------------------------


def test_outline_text_mode_indents_by_level(fake_word):
    code, out, _ = _invoke(["--text", "outline"])
    assert code == EXIT_OK
    lines = out.splitlines()
    # Level 1 has no indent; level 2 has two spaces.
    assert lines[0].startswith("Introduction")
    assert "[heading:" in lines[0]
    assert lines[1].startswith("  Risks")


def test_status_text_mode_marks_active(fake_word):
    code, out, _ = _invoke(["--text", "status"])
    assert code == EXIT_OK
    # Active doc is prefixed with `*`.
    assert out.lstrip().startswith("* Test.docx")


def test_read_bookmark_text_mode_emits_only_text(fake_word):
    code, out, _ = _invoke(["--text", "read", "bookmark", "Address"])
    assert code == EXIT_OK
    # Text mode emits the raw bookmark text — not JSON.
    assert not out.strip().startswith("{")


def test_write_bookmark_text_mode_is_one_line(fake_word):
    code, out, _ = _invoke(["--text", "write", "bookmark", "Address", "--text", "X"])
    assert code == EXIT_OK
    assert out.strip() == "wrote bookmark:Address"


# ---------------------------------------------------------------------------
# read section
# ---------------------------------------------------------------------------


def test_read_section_by_heading(fake_word):
    code, out, _ = _invoke(["read", "section", "Introduction"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["heading"] == "Introduction"
    assert data["anchor_id"] == "heading:1"
    assert data["level"] == 1
    assert "text" in data


def test_read_section_by_anchor_id(fake_word):
    code, out, _ = _invoke(["read", "section", "--anchor-id", "heading:1"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["anchor_id"] == "heading:1"


def test_read_section_missing_heading(fake_word):
    code, _, err = _invoke(["read", "section", "Nope"])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "heading" in err.lower()


def test_read_section_rejects_non_heading_anchor(fake_word):
    code, _, err = _invoke(["read", "section", "--anchor-id", "bookmark:Address"])
    assert code != EXIT_OK
    assert "heading" in err.lower()


def test_read_section_requires_one_of(fake_word):
    code, _, _ = _invoke(["read", "section"])
    assert code != EXIT_OK
    code, _, _ = _invoke(["read", "section", "Introduction", "--anchor-id", "heading:1"])
    assert code != EXIT_OK


# ---------------------------------------------------------------------------
# read markdown / html
# ---------------------------------------------------------------------------


def test_read_markdown_json_envelope(fake_word):
    code, out, _ = _invoke(["read", "markdown"])
    assert code == EXIT_OK
    md = json.loads(out)["markdown"]
    assert "# Introduction" in md
    assert "## Risks" in md


def test_read_markdown_text_mode_emits_markdown(fake_word):
    code, out, _ = _invoke(["--text", "read", "markdown"])
    assert code == EXIT_OK
    assert "# Introduction" in out


def test_read_markdown_within_scope(fake_word):
    code, out, _ = _invoke(["read", "markdown", "--within", "heading:1"])
    assert code == EXIT_OK
    md = json.loads(out)["markdown"]
    assert "Introduction" in md
    assert "Risks" not in md


def test_read_html_json_envelope(fake_word):
    code, out, _ = _invoke(["read", "html"])
    assert code == EXIT_OK
    html = json.loads(out)["html"]
    assert "<h1>Introduction</h1>" in html
    assert "<h2>Risks</h2>" in html


def test_read_digest_json_envelope(fake_word):
    code, out, _ = _invoke(["read", "digest", "--budget", "200"])
    assert code == EXIT_OK
    digest = json.loads(out)["digest"]
    # Headings are the verbatim navigation spine, each tagged with its anchor.
    assert "# Introduction  <!-- heading:1 -->" in digest
    assert "## Risks  <!-- heading:3 -->" in digest


def test_read_digest_text_mode(fake_word):
    code, out, _ = _invoke(["--text", "read", "digest"])
    assert code == EXIT_OK
    assert "# Introduction" in out


# ---------------------------------------------------------------------------
# find / fuzzy replace
# ---------------------------------------------------------------------------


def test_find_locates_match(fake_word):
    code, out, _ = _invoke(["find", "--text", "Body text here"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert len(data) == 1
    assert data[0]["start"] == 13


def test_find_no_match_returns_empty(fake_word):
    code, out, _ = _invoke(["find", "--text", "nope"])
    assert code == EXIT_OK
    assert json.loads(out) == []


def test_replace_fuzzy_single_match(fake_word):
    code, out, _ = _invoke(["replace", "--find", "Body text here", "--text", "Replaced"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert len(data["replacements"]) == 1


def test_replace_fuzzy_zero_matches_is_anchor_not_found(fake_word):
    code, _, _ = _invoke(["replace", "--find", "no such phrase", "--text", "x"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_replace_fuzzy_ambiguous_returns_exit_5(fake_word):
    fake_word.ActiveDocument.Content.Text = "alpha beta alpha beta"
    fake_word.ActiveDocument.Content.End = len("alpha beta alpha beta")

    code, out, _ = _invoke(["replace", "--find", "alpha", "--text", "X"])
    assert code == EXIT_AMBIGUOUS_MATCH
    data = json.loads(out)
    assert data["error"] == "ambiguous_match"
    assert len(data["matches"]) == 2


def test_replace_fuzzy_all_succeeds(fake_word):
    fake_word.ActiveDocument.Content.Text = "alpha beta alpha beta"
    fake_word.ActiveDocument.Content.End = len("alpha beta alpha beta")

    code, out, _ = _invoke(["replace", "--find", "alpha", "--text", "X", "--all"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert len(data["replacements"]) == 2


def test_replace_fuzzy_occurrence_picks_second(fake_word):
    fake_word.ActiveDocument.Content.Text = "alpha beta alpha beta"
    fake_word.ActiveDocument.Content.End = len("alpha beta alpha beta")

    code, out, _ = _invoke(["replace", "--find", "alpha", "--text", "X", "--occurrence", "2"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert len(data["replacements"]) == 1
    assert data["replacements"][0]["start"] == 11


def test_replace_rejects_both_anchor_and_find(fake_word):
    code, _, _ = _invoke(
        [
            "replace",
            "--anchor-id",
            "heading:1",
            "--find",
            "alpha",
            "--text",
            "x",
        ]
    )
    assert code != EXIT_OK


def test_replace_rejects_all_with_anchor_id(fake_word):
    code, _, _ = _invoke(["replace", "--anchor-id", "heading:1", "--text", "x", "--all"])
    assert code != EXIT_OK


def test_exec_supports_find_replace_op(fake_word, tmp_path: Path):
    fake_word.ActiveDocument.Content.Text = "alpha beta alpha beta"
    fake_word.ActiveDocument.Content.End = len("alpha beta alpha beta")

    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "find_replace", "find": "alpha", "text": "X", "all": True},
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["ops_run"] == 1


def test_exec_unknown_op_is_click_error(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps({"ops": [{"op": "drop_table", "name": "anything"}]}),
        encoding="utf-8",
    )

    # Unknown op is a ClickException — non-zero exit, but not our taxonomy.
    code, _, err = _invoke(["exec", "--script", str(script)])
    assert code != EXIT_OK
    assert "drop_table" in err


# ---------------------------------------------------------------------------
# Malformed exec ops (regression for B-4 in issues.md)
# ---------------------------------------------------------------------------


def test_exec_op_missing_required_field_reports_cleanly(fake_word, tmp_path: Path):
    """A typo'd op payload (missing `name`) must surface as a clean Click
    error naming the missing field — not as a Python KeyError traceback.
    """
    script = tmp_path / "ops.json"
    script.write_text(
        # Missing 'name' on a write_bookmark op.
        json.dumps({"ops": [{"op": "write_bookmark", "text": "..."}]}),
        encoding="utf-8",
    )
    code, _, err = _invoke(["exec", "--script", str(script)])
    assert code != EXIT_OK
    # The error message names both the op kind and the missing field; no traceback.
    assert "write_bookmark" in err
    assert "name" in err
    assert "Traceback" not in err


def test_exec_op_missing_op_field_reports_cleanly(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(json.dumps({"ops": [{"name": "Address", "text": "X"}]}), encoding="utf-8")
    code, _, err = _invoke(["exec", "--script", str(script)])
    assert code != EXIT_OK
    assert "'op'" in err or "op " in err.lower()
    assert "Traceback" not in err


def test_exec_op_non_object_reports_cleanly(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(json.dumps({"ops": ["not-a-dict"]}), encoding="utf-8")
    code, _, err = _invoke(["exec", "--script", str(script)])
    assert code != EXIT_OK
    assert "Traceback" not in err


def test_exec_format_paragraph_missing_anchor_id_reports_cleanly(fake_word, tmp_path: Path):
    """format_paragraph has only one required field — make sure that's enforced too."""
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps({"ops": [{"op": "format_paragraph", "alignment": "center"}]}),
        encoding="utf-8",
    )
    code, _, err = _invoke(["exec", "--script", str(script)])
    assert code != EXIT_OK
    assert "format_paragraph" in err
    assert "anchor_id" in err
    assert "Traceback" not in err


# ---------------------------------------------------------------------------
# style list / style apply / format-paragraph / insert --style validation
# ---------------------------------------------------------------------------


def test_insert_with_valid_style_passes(fake_word):
    code, out, _ = _invoke(
        ["insert", "--anchor-id", "heading:1", "--text", "x", "--style", "Body Text"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["style"] == "Body Text"


def test_insert_with_bad_style_returns_exit_2(fake_word):
    code, _, err = _invoke(
        ["insert", "--anchor-id", "heading:1", "--text", "x", "--style", "NoSuchStyle"]
    )
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "style" in err.lower()


def test_style_list_returns_known_styles(fake_word):
    code, out, _ = _invoke(["style", "list"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert isinstance(data, list)
    names = [r["name"] for r in data]
    assert "Heading 1" in names
    assert "Body Text" in names
    assert all({"name", "type", "builtin", "in_use"} <= set(r) for r in data)


def test_style_list_text_mode(fake_word):
    code, out, _ = _invoke(["--text", "style", "list"])
    assert code == EXIT_OK
    assert "Heading 1" in out


def test_style_apply_happy_path(fake_word):
    code, out, _ = _invoke(
        ["style", "apply", "--anchor-id", "bookmark:Address", "--name", "Heading 2"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["style"] == "Heading 2"
    assert data["anchor"]["kind"] == "bookmark"


def test_style_apply_bad_anchor_returns_exit_2(fake_word):
    code, _, _ = _invoke(["style", "apply", "--anchor-id", "bookmark:Nope", "--name", "Heading 2"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_style_apply_bad_style_returns_exit_2(fake_word):
    code, _, err = _invoke(
        ["style", "apply", "--anchor-id", "bookmark:Address", "--name", "NoSuchStyle"]
    )
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "style" in err.lower()


def test_format_paragraph_alignment(fake_word):
    code, out, _ = _invoke(
        ["format-paragraph", "--anchor-id", "bookmark:Address", "--alignment", "center"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["applied"]["alignment"] == "center"


def test_format_paragraph_indent_and_spacing(fake_word):
    code, out, _ = _invoke(
        [
            "format-paragraph",
            "--anchor-id",
            "bookmark:Address",
            "--left-indent",
            "36",
            "--space-before",
            "6",
        ]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["applied"]["left_indent"] == 36.0
    assert data["applied"]["space_before"] == 6.0


def test_format_paragraph_requires_some_arg(fake_word):
    code, _, _ = _invoke(["format-paragraph", "--anchor-id", "bookmark:Address"])
    assert code != EXIT_OK


def test_format_paragraph_bad_anchor_returns_exit_2(fake_word):
    code, _, _ = _invoke(
        ["format-paragraph", "--anchor-id", "bookmark:Nope", "--alignment", "left"]
    )
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_exec_supports_apply_style_op(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "apply_style", "anchor_id": "bookmark:Address", "name": "Heading 2"},
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["ops_run"] == 1


def test_exec_supports_format_paragraph_op(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {
                        "op": "format_paragraph",
                        "anchor_id": "bookmark:Address",
                        "alignment": "center",
                        "space_before": 6,
                    },
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True


def test_exec_apply_style_with_bad_name_fails_to_exit_2(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "apply_style", "anchor_id": "bookmark:Address", "name": "NoSuchStyle"},
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_ANCHOR_NOT_FOUND
    data = json.loads(out)
    assert data["ok"] is False
    assert data["failure"]["type"] == "StyleNotFoundError"


# ---------------------------------------------------------------------------
# `kind` field consistency (regression test for B-1 in issues.md)
# ---------------------------------------------------------------------------


def test_kind_field_is_consistent_across_cc_commands(fake_word):
    """Every CLI command that returns an `anchor` for a content control must
    emit the same `kind` string. Otherwise an LLM that branches on `kind`
    breaks the moment it sees a different verb's output.
    """
    cmds = [
        ["write", "cc", "Signatory", "--text", "X"],
        ["replace", "--anchor-id", "cc:Signatory", "--text", "X"],
        ["go-to", "--anchor-id", "cc:Signatory"],
        ["style", "apply", "--anchor-id", "cc:Signatory", "--name", "Heading 2"],
        ["format-paragraph", "--anchor-id", "cc:Signatory", "--alignment", "left"],
    ]
    kinds: dict[str, str] = {}
    for cmd in cmds:
        code, out, _ = _invoke(cmd)
        assert code == EXIT_OK, f"{cmd} exited {code}"
        kinds[" ".join(cmd[:2])] = json.loads(out)["anchor"]["kind"]
    assert len(set(kinds.values())) == 1, f"kind drift: {kinds}"
    assert next(iter(kinds.values())) == "content_control"


def test_kind_field_is_consistent_across_bookmark_commands(fake_word):
    cmds = [
        ["write", "bookmark", "Address", "--text", "X"],
        ["replace", "--anchor-id", "bookmark:Address", "--text", "X"],
        ["go-to", "--anchor-id", "bookmark:Address"],
        ["style", "apply", "--anchor-id", "bookmark:Address", "--name", "Heading 2"],
        ["format-paragraph", "--anchor-id", "bookmark:Address", "--alignment", "left"],
    ]
    kinds: dict[str, str] = {}
    for cmd in cmds:
        code, out, _ = _invoke(cmd)
        assert code == EXIT_OK, f"{cmd} exited {code}"
        kinds[" ".join(cmd[:2])] = json.loads(out)["anchor"]["kind"]
    assert len(set(kinds.values())) == 1, f"kind drift: {kinds}"
    assert next(iter(kinds.values())) == "bookmark"


# ---------------------------------------------------------------------------
# tables: list / read / add-row / delete-row, cell anchors, exec ops
# ---------------------------------------------------------------------------


def test_table_list(fake_word):
    code, out, _ = _invoke(["table", "list"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data == [{"index": 1, "title": "Grid", "rows": 2, "columns": 2}]


def test_table_list_text_mode(fake_word):
    code, out, _ = _invoke(["--text", "table", "list"])
    assert code == EXIT_OK
    assert "table:1" in out
    assert "2x2" in out


def test_table_read(fake_word):
    code, out, _ = _invoke(["table", "read", "1"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["index"] == 1
    assert data["cells"][0][0]["text"] == "A1"
    assert data["cells"][1][1]["anchor_id"] == "table:1:2:2"


def test_table_read_missing_returns_exit_2(fake_word):
    code, _, err = _invoke(["table", "read", "5"])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "table" in err.lower()


def test_table_add_row(fake_word):
    code, out, _ = _invoke(["table", "add-row", "--table", "1"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["rows"] == 3


def test_table_add_row_with_values(fake_word):
    code, out, _ = _invoke(["table", "add-row", "--table", "1", "--values", '["X", "Y"]'])
    assert code == EXIT_OK
    # The new row's cells should be addressable and hold the values.
    code, out, _ = _invoke(["table", "read", "1"])
    data = json.loads(out)
    assert data["cells"][2][0]["text"] == "X"
    assert data["cells"][2][1]["text"] == "Y"


def test_table_add_row_bad_values_is_usage_error(fake_word):
    code, _, _ = _invoke(["table", "add-row", "--table", "1", "--values", "not-json"])
    assert code != EXIT_OK


def test_table_delete_row(fake_word):
    code, out, _ = _invoke(["table", "delete-row", "--table", "1", "--row", "1"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["rows"] == 1


def test_table_delete_row_out_of_range_returns_exit_2(fake_word):
    code, _, _ = _invoke(["table", "delete-row", "--table", "1", "--row", "9"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_replace_cell_via_anchor_id(fake_word):
    code, out, _ = _invoke(["replace", "--anchor-id", "table:1:1:1", "--text", "Z"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor"]["kind"] == "cell"


def test_style_apply_to_cell(fake_word):
    code, out, _ = _invoke(["style", "apply", "--anchor-id", "table:1:1:1", "--name", "Heading 2"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["anchor"]["kind"] == "cell"


def test_exec_supports_set_cell_op(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps({"ops": [{"op": "set_cell", "table": 1, "row": 1, "col": 2, "text": "new"}]}),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    # Confirm the write landed.
    code, out, _ = _invoke(["table", "read", "1"])
    assert json.loads(out)["cells"][0][1]["text"] == "new"


def test_exec_supports_add_and_delete_row_ops(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "add_row", "table": 1, "values": ["c", "d"]},
                    {"op": "delete_row", "table": 1, "row": 1},
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 2


def test_exec_set_cell_missing_field_reports_cleanly(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps({"ops": [{"op": "set_cell", "table": 1, "row": 1, "col": 2}]}),
        encoding="utf-8",
    )
    code, _, err = _invoke(["exec", "--script", str(script)])
    assert code != EXIT_OK
    assert "set_cell" in err
    assert "text" in err
    assert "Traceback" not in err


# ---------------------------------------------------------------------------
# v0.5: range anchors via the `range:` id
# ---------------------------------------------------------------------------


def test_replace_via_range_anchor_id(fake_word):
    code, out, _ = _invoke(["replace", "--anchor-id", "range:3-8", "--text", "Z"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor"]["kind"] == "range"


def test_replace_bad_range_id_returns_exit_2(fake_word):
    code, _, _ = _invoke(["replace", "--anchor-id", "range:nope", "--text", "Z"])
    assert code == EXIT_ANCHOR_NOT_FOUND


# ---------------------------------------------------------------------------
# v0.5: comments
# ---------------------------------------------------------------------------


def test_comment_list_empty(fake_word):
    code, out, _ = _invoke(["comment", "list"])
    assert code == EXIT_OK
    assert json.loads(out) == []


def test_comment_add(fake_word):
    code, out, _ = _invoke(
        ["comment", "add", "--anchor-id", "heading:1", "--text", "Review this", "--author", "Bot"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["comment"]["index"] == 1
    assert data["comment"]["author"] == "Bot"


def test_comment_add_then_list(fake_word):
    _invoke(["comment", "add", "--anchor-id", "heading:1", "--text", "note A"])
    code, out, _ = _invoke(["comment", "list"])
    assert code == EXIT_OK
    rows = json.loads(out)
    assert len(rows) == 1
    assert rows[0]["text"] == "note A"


def test_comment_add_bad_anchor_returns_exit_2(fake_word):
    code, _, _ = _invoke(["comment", "add", "--anchor-id", "bookmark:Nope", "--text", "x"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_comment_resolve(fake_word):
    _invoke(["comment", "add", "--anchor-id", "heading:1", "--text", "x"])
    code, out, _ = _invoke(["comment", "resolve", "--index", "1"])
    assert code == EXIT_OK
    assert json.loads(out)["done"] is True


def test_comment_resolve_out_of_range_returns_exit_2(fake_word):
    code, _, _ = _invoke(["comment", "resolve", "--index", "9"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_comment_delete(fake_word):
    _invoke(["comment", "add", "--anchor-id", "heading:1", "--text", "x"])
    code, out, _ = _invoke(["comment", "delete", "--index", "1"])
    assert code == EXIT_OK
    assert json.loads(out)["deleted"] is True
    code, out, _ = _invoke(["comment", "list"])
    assert json.loads(out) == []


def test_comment_list_text_mode(fake_word):
    _invoke(["comment", "add", "--anchor-id", "heading:1", "--text", "look", "--author", "Bot"])
    code, out, _ = _invoke(["--text", "comment", "list"])
    assert code == EXIT_OK
    assert "Bot" in out
    assert "look" in out


# ---------------------------------------------------------------------------
# v0.5: track changes
# ---------------------------------------------------------------------------


def test_track_status_default_off(fake_word):
    code, out, _ = _invoke(["track", "status"])
    assert code == EXIT_OK
    assert json.loads(out) == {"tracked": False}


def test_track_on_then_status(fake_word):
    code, out, _ = _invoke(["track", "on"])
    assert code == EXIT_OK
    assert json.loads(out)["tracked"] is True
    assert fake_word.ActiveDocument.TrackRevisions is True
    code, out, _ = _invoke(["track", "status"])
    assert json.loads(out) == {"tracked": True}


def test_track_off(fake_word):
    _invoke(["track", "on"])
    code, out, _ = _invoke(["track", "off"])
    assert code == EXIT_OK
    assert json.loads(out)["tracked"] is False
    assert fake_word.ActiveDocument.TrackRevisions is False


# ---------------------------------------------------------------------------
# v0.5: exec comment ops + tracked batch
# ---------------------------------------------------------------------------


def test_exec_supports_add_comment_op(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {
                        "op": "add_comment",
                        "anchor_id": "heading:1",
                        "text": "please review",
                        "author": "Bot",
                    },
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    code, out, _ = _invoke(["comment", "list"])
    assert json.loads(out)[0]["text"] == "please review"


def test_exec_supports_resolve_and_delete_comment_ops(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "add_comment", "anchor_id": "heading:1", "text": "a"},
                    {"op": "add_comment", "anchor_id": "heading:1", "text": "b"},
                    {"op": "resolve_comment", "index": 1},
                    {"op": "delete_comment", "index": 2},
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 4
    code, out, _ = _invoke(["comment", "list"])
    rows = json.loads(out)
    assert len(rows) == 1
    assert rows[0]["done"] is True


def test_exec_add_comment_missing_field_reports_cleanly(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps({"ops": [{"op": "add_comment", "anchor_id": "heading:1"}]}),
        encoding="utf-8",
    )
    code, _, err = _invoke(["exec", "--script", str(script)])
    assert code != EXIT_OK
    assert "add_comment" in err
    assert "text" in err
    assert "Traceback" not in err


def test_exec_tracked_payload_turns_on_then_restores(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "tracked": True,
                "ops": [
                    {"op": "write_bookmark", "name": "Address", "text": "123 Main"},
                ],
            }
        ),
        encoding="utf-8",
    )
    code, _, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    # Track Changes was flipped on for the batch, then restored to off.
    assert fake_word.ActiveDocument.TrackRevisions is False


# ---------------------------------------------------------------------------
# list (v0.6): show / apply / remove / info / restart / indent / outdent
# ---------------------------------------------------------------------------


def test_list_show(fake_word):
    code, out, _ = _invoke(["list", "show"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data[0]["index"] == 1
    assert data[0]["type"] == "numbered"
    assert data[0]["anchor_id"] == "range:13-29"


def test_list_apply_numbered(fake_word):
    code, out, _ = _invoke(["list", "apply", "--anchor-id", "range:0-12", "--type", "numbered"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["type"] == "numbered"
    assert data["continue_previous"] is False


def test_list_apply_with_continue_flag(fake_word):
    code, _, _ = _invoke(
        ["list", "apply", "--anchor-id", "range:0-12", "--type", "numbered", "--continue"]
    )
    assert code == EXIT_OK
    assert fake_word.ActiveDocument.Range(0, 12).ListFormat._continue is True


def test_list_apply_bad_anchor_returns_exit_2(fake_word):
    code, _, _ = _invoke(["list", "apply", "--anchor-id", "heading:99", "--type", "bulleted"])
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_list_info(fake_word):
    code, out, _ = _invoke(["list", "info", "--anchor-id", "range:13-29"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["type"] == "numbered"
    assert data["string"] == "1."


def test_list_format_applies_custom_levels(fake_word):
    levels = '[{"kind":"number","format":"%1)","style":"lower-letter"}]'
    code, out, _ = _invoke(["list", "format", "--anchor-id", "range:0-12", "--levels", levels])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["levels"] == 1
    # Read the per-level format back.
    code, out, _ = _invoke(["list", "levels", "--anchor-id", "range:0-12"])
    assert code == EXIT_OK
    lv = json.loads(out)["levels"]
    assert lv[0]["format"] == "%1)"
    assert lv[0]["style"] == "lower-letter"


def test_list_format_bad_levels_is_usage_error(fake_word):
    code, _, _ = _invoke(["list", "format", "--anchor-id", "range:0-12", "--levels", "not-json"])
    assert code != EXIT_OK


def test_list_format_bad_anchor_returns_exit_2(fake_word):
    code, _, _ = _invoke(
        ["list", "format", "--anchor-id", "heading:99", "--levels", '[{"kind":"number"}]']
    )
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_list_remove(fake_word):
    code, out, _ = _invoke(["list", "remove", "--anchor-id", "range:13-29"])
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True


def test_list_restart(fake_word):
    code, out, _ = _invoke(["list", "restart", "--anchor-id", "range:13-29"])
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True


def test_list_indent_then_outdent(fake_word):
    _invoke(["list", "apply", "--anchor-id", "range:0-12", "--type", "numbered"])
    code, _, _ = _invoke(["list", "indent", "--anchor-id", "range:0-12"])
    assert code == EXIT_OK
    code, _, _ = _invoke(["list", "outdent", "--anchor-id", "range:0-12"])
    assert code == EXIT_OK


# ---------------------------------------------------------------------------
# section / header / footer (v0.6)
# ---------------------------------------------------------------------------


def test_section_list(fake_word):
    code, out, _ = _invoke(["section", "list"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data[0]["index"] == 1
    assert data[0]["page_setup"]["orientation"] == "portrait"


def test_header_read_default_section(fake_word):
    code, out, _ = _invoke(["header", "read"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["text"] == "Confidential Draft"
    assert data["anchor_id"] == "header:1:primary"


def test_header_write_then_read(fake_word):
    code, out, _ = _invoke(["header", "write", "--section", "1", "--text", "ACME Corp"])
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True
    code, out, _ = _invoke(["header", "read", "--section", "1"])
    assert json.loads(out)["text"] == "ACME Corp"


def test_header_read_unseeded_which_is_empty(fake_word):
    code, out, _ = _invoke(["header", "read", "--which", "first"])
    assert code == EXIT_OK
    assert json.loads(out)["text"] == ""


def test_footer_read_and_write(fake_word):
    code, out, _ = _invoke(["footer", "read"])
    assert json.loads(out)["text"] == "Page 1"
    code, out, _ = _invoke(["footer", "write", "--text", "Page 2 of 5"])
    assert code == EXIT_OK
    code, out, _ = _invoke(["footer", "read"])
    assert json.loads(out)["text"] == "Page 2 of 5"


def test_header_text_mode(fake_word):
    code, out, _ = _invoke(["--text", "header", "read"])
    assert code == EXIT_OK
    assert out.strip() == "Confidential Draft"


# ---------------------------------------------------------------------------
# exec ops (v0.6): apply_list / restart_numbering / write_header / write_footer
# ---------------------------------------------------------------------------


def test_exec_supports_apply_list_op(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps({"ops": [{"op": "apply_list", "anchor_id": "range:0-12", "type": "numbered"}]}),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1
    assert fake_word.ActiveDocument.Range(0, 12).ListFormat.ListType == 3


def test_exec_supports_header_and_footer_ops(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "write_header", "section": 1, "text": "Hdr"},
                    {"op": "write_footer", "section": 1, "which": "primary", "text": "Ftr"},
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 2
    doc = fake_word.ActiveDocument
    assert doc.Sections(1).Headers(1).Range.Text == "Hdr"
    assert doc.Sections(1).Footers(1).Range.Text == "Ftr"


def test_exec_apply_list_missing_field_reports_cleanly(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(json.dumps({"ops": [{"op": "apply_list"}]}), encoding="utf-8")
    code, _, err = _invoke(["exec", "--script", str(script)])
    assert code != EXIT_OK
    assert "anchor_id" in err


# ---------------------------------------------------------------------------
# v0.7: paragraphs / outline --all / cursor
# ---------------------------------------------------------------------------


def test_paragraphs_lists_every_paragraph(fake_word):
    code, out, _ = _invoke(["paragraphs"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert [p["anchor_id"] for p in data] == ["para:1", "para:2", "para:3"]
    # Headings and body paragraphs both appear, flagged by is_heading.
    assert [p["is_heading"] for p in data] == [True, False, True]
    assert [p["text"] for p in data] == ["Introduction", "Body text here.", "Risks"]
    # Offsets are emitted so they can feed a range:START-END insertion.
    assert data[1]["start"] == 13 and data[1]["end"] == 29


def test_outline_all_matches_paragraphs(fake_word):
    code_a, out_a, _ = _invoke(["outline", "--all"])
    code_b, out_b, _ = _invoke(["paragraphs"])
    assert code_a == EXIT_OK
    assert json.loads(out_a) == json.loads(out_b)


def test_outline_default_is_headings_only(fake_word):
    code, out, _ = _invoke(["outline"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert [item["anchor_id"] for item in data] == ["heading:1", "heading:3"]


def test_replace_via_para_anchor(fake_word):
    code, out, _ = _invoke(["replace", "--anchor-id", "para:2", "--text", "Rewritten body"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    # para:2 spans 13-29; set_text preserves the trailing mark (writes 13-28).
    assert fake_word.ActiveDocument.Range(13, 28).Text == "Rewritten body"


def test_para_anchor_out_of_range_exit_2(fake_word):
    code, _, err = _invoke(["replace", "--anchor-id", "para:99", "--text", "x"])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "paragraph" in err.lower()


def test_cursor_read_reports_position_and_paragraph(fake_word):
    fake_word.Selection.Start = 15
    fake_word.Selection.End = 15
    fake_word.Selection.Text = ""
    code, out, _ = _invoke(["cursor", "read"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["start"] == 15
    assert data["collapsed"] is True
    # Offset 15 falls inside the body paragraph (13-29) -> para:2.
    assert data["paragraph"] == {"anchor_id": "para:2"}


def test_cursor_write_inserts_and_moves_cursor(fake_word):
    fake_word.Selection.Start = 0
    fake_word.Selection.End = 0
    code, out, _ = _invoke(["cursor", "write", "--text", "Hi"])
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True
    assert fake_word.ActiveDocument.Range(0, 0).Text == "Hi"


def test_cursor_write_replace_overwrites_selection(fake_word):
    fake_word.Selection.Start = 0
    fake_word.Selection.End = 12
    code, _, _ = _invoke(["cursor", "write", "--text", "New"])
    assert code == EXIT_OK
    # With a spanning selection, replace (default) overwrites 0-12.
    assert fake_word.ActiveDocument.Range(0, 12).Text == "New"


def test_cursor_write_no_replace_inserts_at_start(fake_word):
    fake_word.Selection.Start = 5
    fake_word.Selection.End = 12
    code, _, _ = _invoke(["cursor", "write", "--text", "X", "--no-replace"])
    assert code == EXIT_OK
    # --no-replace collapses to the selection start (5) before inserting.
    assert fake_word.ActiveDocument.Range(5, 5).Text == "X"


# ---------------------------------------------------------------------------
# insert-image
# ---------------------------------------------------------------------------

# A 1x1 transparent PNG, shared by the image CLI tests.
_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYPhfDwAChwGA"
    "60e6kgAAAABJRU5ErkJggg=="
)


def _png_path(tmp_path: Path) -> Path:
    import base64

    p = tmp_path / "pic.png"
    p.write_bytes(base64.b64decode(_PNG_B64))
    return p


def test_insert_image_from_path(fake_word, tmp_path: Path):
    img = _png_path(tmp_path)
    code, out, _ = _invoke(
        ["insert-image", "--anchor-id", "bookmark:Address", "--path", str(img), "--wrap", "inline"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["anchor"]["kind"] == "bookmark"
    assert data["wrap"] == "inline"
    assert data["where"] == "after"


def test_insert_image_square_before(fake_word, tmp_path: Path):
    img = _png_path(tmp_path)
    code, out, _ = _invoke(
        [
            "insert-image",
            "--anchor-id",
            "bookmark:Address",
            "--path",
            str(img),
            "--wrap",
            "square",
            "--before",
            "--width",
            "120",
            "--alt-text",
            "A diagram",
        ]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["where"] == "before"
    assert data["wrap"] == "square"


def test_insert_image_from_base64_arg(fake_word):
    code, out, _ = _invoke(
        ["insert-image", "--anchor-id", "bookmark:Address", "--base64", _PNG_B64, "--wrap", "auto"]
    )
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True


def test_insert_image_from_base64_stdin(fake_word):
    code, out, _ = _invoke(
        ["insert-image", "--anchor-id", "bookmark:Address", "--base64", "-", "--wrap", "inline"],
        input=_PNG_B64,
    )
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True


def test_insert_image_missing_file_is_exit_other(fake_word, tmp_path: Path):
    code, _, err = _invoke(
        [
            "insert-image",
            "--anchor-id",
            "bookmark:Address",
            "--path",
            str(tmp_path / "nope.png"),
            "--wrap",
            "inline",
        ]
    )
    assert code == EXIT_OTHER
    assert "image" in err.lower()


def test_insert_image_bad_anchor_is_exit_anchor_not_found(fake_word, tmp_path: Path):
    img = _png_path(tmp_path)
    code, _, _ = _invoke(
        ["insert-image", "--anchor-id", "bookmark:Nope", "--path", str(img), "--wrap", "inline"]
    )
    assert code == EXIT_ANCHOR_NOT_FOUND


def test_insert_image_requires_exactly_one_source(fake_word, tmp_path: Path):
    img = _png_path(tmp_path)
    # neither
    code, _, _ = _invoke(["insert-image", "--anchor-id", "bookmark:Address", "--wrap", "inline"])
    assert code == 2  # click usage error
    # both
    code, _, _ = _invoke(
        [
            "insert-image",
            "--anchor-id",
            "bookmark:Address",
            "--path",
            str(img),
            "--base64",
            _PNG_B64,
            "--wrap",
            "inline",
        ]
    )
    assert code == 2


def test_insert_image_bad_wrap_is_usage_error(fake_word, tmp_path: Path):
    img = _png_path(tmp_path)
    code, _, _ = _invoke(
        [
            "insert-image",
            "--anchor-id",
            "bookmark:Address",
            "--path",
            str(img),
            "--wrap",
            "diagonal",
        ]
    )
    assert code == 2  # click.Choice rejects it


def test_exec_insert_image_with_path(fake_word, tmp_path: Path):
    img = _png_path(tmp_path)
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {"op": "write_bookmark", "name": "Address", "text": "123 Main"},
                    {
                        "op": "insert_image",
                        "anchor_id": "bookmark:Address",
                        "path": str(img),
                        "wrap": "square",
                    },
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["ops_run"] == 2
    # Both ops ride a single undo record.
    fake_word.UndoRecord.StartCustomRecord.assert_called_once()


def test_exec_insert_image_with_base64(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {
                "ops": [
                    {
                        "op": "insert_image",
                        "anchor_id": "bookmark:Address",
                        "base64": _PNG_B64,
                        "wrap": "auto",
                    },
                ]
            }
        ),
        encoding="utf-8",
    )
    code, out, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1


def test_exec_insert_image_without_source_fails(fake_word, tmp_path: Path):
    script = tmp_path / "ops.json"
    script.write_text(
        json.dumps(
            {"ops": [{"op": "insert_image", "anchor_id": "bookmark:Address", "wrap": "inline"}]}
        ),
        encoding="utf-8",
    )
    code, _, _ = _invoke(["exec", "--script", str(script)])
    assert code == EXIT_OTHER  # ClickException for malformed op


# ---------------------------------------------------------------------------
# --version / --about (offline — no Word needed)
# ---------------------------------------------------------------------------


def test_version_flag_prints_version():
    """`--version`/`-v` prints the package version and exits cleanly."""
    from wordlive import __version__

    for flag in ("--version", "-v"):
        code, out, _ = _invoke([flag])
        assert code == EXIT_OK
        assert out.strip() == f"wordlive {__version__}"


def test_about_flag_shows_banner_and_metadata():
    """`--about`/`-A` renders the banner plus version, author, license, and repo."""
    from wordlive import __version__

    for flag in ("--about", "-A"):
        code, out, _ = _invoke([flag])
        assert code == EXIT_OK
        # CliRunner is not a tty, so the ANSI colour is stripped to clean ASCII.
        assert "\x1b[" not in out
        assert r"\_/\_/ \___/" in out  # a recognisable banner row
        assert __version__ in out
        assert "Tom Villani, Ph.D." in out
        assert "MIT" in out
        assert "github.com/thomas-villani/wordlive" in out


# ---------------------------------------------------------------------------
# llm-help (offline — no Word needed)
# ---------------------------------------------------------------------------


def test_llm_help_prints_cli_skill_body():
    """`llm-help` dumps the CLI guide as raw Markdown — no Word, no JSON."""
    code, out, _ = _invoke(["llm-help"])
    assert code == EXIT_OK
    # Raw Markdown, not a JSON object, and frontmatter is stripped.
    assert not out.lstrip().startswith("{")
    assert out.lstrip().startswith("# wordlive (CLI)")
    assert "name: wordlive" not in out  # YAML frontmatter dropped
    # Content sanity: the anchor model, a verb, and the exit-code contract.
    assert "--anchor-id" in out
    assert "insert-image" in out
    assert "Exit codes" in out


def test_llm_help_python_prints_python_guide():
    """`llm-help --python` dumps the Python-API guide instead of the CLI one."""
    code, out, _ = _invoke(["llm-help", "--python"])
    assert code == EXIT_OK
    assert out.lstrip().startswith("# wordlive (Python API)")
    assert "import wordlive as wl" in out
    assert 'doc.edit("' in out  # the atomic-undo idiom


def test_llm_help_ignores_json_flag():
    """It's documentation, like --help: raw Markdown even under the JSON default."""
    code_default, out_default, _ = _invoke(["llm-help"])
    code_json, out_json, _ = _invoke(["--json", "llm-help"])
    assert code_default == EXIT_OK and code_json == EXIT_OK
    assert out_default == out_json
    assert not out_json.lstrip().startswith("{")


def test_llm_help_matches_installed_skill_body(tmp_path: Path, monkeypatch):
    """`llm-help` output is the body of the same CLI skill `install-skill` writes."""
    monkeypatch.chdir(tmp_path)
    _invoke(["install-skill", "--cli"])
    installed = (tmp_path / ".agents" / "skills" / "wordlive-cli" / "SKILL.md").read_text(
        encoding="utf-8"
    )
    _, out, _ = _invoke(["llm-help"])
    # The printed guide is the installed skill minus its YAML frontmatter.
    assert installed.rstrip().endswith(out.rstrip())


# ---------------------------------------------------------------------------
# install-skill (offline — no Word needed)
# ---------------------------------------------------------------------------


def test_install_skill_installs_cli_by_default(tmp_path: Path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    code, out, _ = _invoke(["install-skill"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["scope"] == "local"
    assert [r["name"] for r in data["installed"]] == ["wordlive-cli"]
    cli = tmp_path / ".agents" / "skills" / "wordlive-cli" / "SKILL.md"
    assert cli.exists()
    assert "name: wordlive-cli" in cli.read_text(encoding="utf-8")
    # Python skill is NOT installed unless asked for.
    assert not (tmp_path / ".agents" / "skills" / "wordlive-python").exists()


def test_install_skill_both(tmp_path: Path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    code, out, _ = _invoke(["install-skill", "--both"])
    assert code == EXIT_OK
    data = json.loads(out)
    names = {r["name"] for r in data["installed"]}
    assert names == {"wordlive-cli", "wordlive-python"}
    cli = tmp_path / ".agents" / "skills" / "wordlive-cli" / "SKILL.md"
    py = tmp_path / ".agents" / "skills" / "wordlive-python" / "SKILL.md"
    assert cli.exists() and py.exists()
    assert "name: wordlive-python" in py.read_text(encoding="utf-8")


def test_install_skill_cli_only(tmp_path: Path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    code, out, _ = _invoke(["install-skill", "--cli"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert [r["name"] for r in data["installed"]] == ["wordlive-cli"]
    assert not (tmp_path / ".agents" / "skills" / "wordlive-python").exists()


def test_install_skill_python_only(tmp_path: Path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    code, out, _ = _invoke(["install-skill", "--python"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert [r["name"] for r in data["installed"]] == ["wordlive-python"]
    assert (tmp_path / ".agents" / "skills" / "wordlive-python" / "SKILL.md").exists()


def test_install_skill_system(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: tmp_path))
    code, out, _ = _invoke(["install-skill", "--system"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["scope"] == "system"
    assert (tmp_path / ".agents" / "skills" / "wordlive-cli" / "SKILL.md").exists()


def test_install_skill_refuses_overwrite_without_force(tmp_path: Path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    assert _invoke(["install-skill"])[0] == EXIT_OK
    code, _, err = _invoke(["install-skill"])
    assert code == EXIT_OTHER
    assert "force" in err.lower()


def test_install_skill_force_overwrites(tmp_path: Path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    _invoke(["install-skill"])
    code, out, _ = _invoke(["install-skill", "--force"])
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True


# ---------------------------------------------------------------------------
# install-mcp (offline — no Word needed)
# ---------------------------------------------------------------------------


def test_install_mcp_print_emits_pypi_snippet():
    """`--print` returns the uvx-from-PyPI entry without writing any file."""
    code, out, _ = _invoke(["install-mcp", "--print"])
    assert code == EXIT_OK
    data = json.loads(out)
    entry = data["mcpServers"]["wordlive"]
    assert entry["command"] == "uvx"
    assert entry["args"] == ["--from", "wordlive[mcp,snapshot]", "wordlive-mcp"]


def test_install_mcp_print_directory_uses_local_checkout():
    """`--directory` switches to `uv run --directory DIR wordlive-mcp` (dev)."""
    code, out, _ = _invoke(["install-mcp", "--print", "--directory", "C:/checkout"])
    assert code == EXIT_OK
    entry = json.loads(out)["mcpServers"]["wordlive"]
    assert entry["command"] == "uv"
    assert entry["args"] == ["run", "--directory", "C:/checkout", "wordlive-mcp"]


def test_install_mcp_writes_config(tmp_path: Path):
    cfg = tmp_path / "claude_desktop_config.json"
    code, out, _ = _invoke(["install-mcp", "--config", str(cfg)])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["action"] == "created"
    written = json.loads(cfg.read_text(encoding="utf-8"))
    assert written["mcpServers"]["wordlive"]["command"] == "uvx"


def test_install_mcp_merges_into_existing_config(tmp_path: Path):
    cfg = tmp_path / "cfg.json"
    cfg.write_text(json.dumps({"mcpServers": {"other": {"command": "x"}}}), encoding="utf-8")
    code, _, _ = _invoke(["install-mcp", "--config", str(cfg)])
    assert code == EXIT_OK
    written = json.loads(cfg.read_text(encoding="utf-8"))
    assert set(written["mcpServers"]) == {"other", "wordlive"}  # existing entry preserved


def test_install_mcp_claude_code_writes_local_mcp_json(tmp_path: Path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    code, _, _ = _invoke(["install-mcp", "--client", "claude-code"])
    assert code == EXIT_OK
    assert (tmp_path / ".mcp.json").exists()


def test_install_mcp_refuses_overwrite_without_force(tmp_path: Path):
    cfg = tmp_path / "cfg.json"
    assert _invoke(["install-mcp", "--config", str(cfg)])[0] == EXIT_OK
    code, _, err = _invoke(["install-mcp", "--config", str(cfg)])
    assert code == EXIT_OTHER
    assert "force" in err.lower()


def test_install_mcp_force_updates(tmp_path: Path):
    cfg = tmp_path / "cfg.json"
    _invoke(["install-mcp", "--config", str(cfg)])
    code, out, _ = _invoke(["install-mcp", "--config", str(cfg), "--force"])
    assert code == EXIT_OK
    assert json.loads(out)["action"] == "updated"
