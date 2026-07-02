"""Linter + regularizer — `doc.lint()` / `doc.regularize()`.

Detection runs over `fake_word` (whose ranges now carry real Font / Paragraph-
Format and an applied ParagraphStyle baseline). The structural `table-repeat-
header` rule needs real page geometry, so its multi-page case is smoke-only here
(the unit test asserts no false positive on the single-page fixture table); the
others round-trip in the fake.
"""

from __future__ import annotations

import wordlive
from wordlive import _com


def _attach(monkeypatch, app):
    monkeypatch.setattr(_com, "get_active_word", lambda: app)
    monkeypatch.setattr(_com, "launch_word", lambda visible=True: app)


def _make_app(**kwargs):
    from tests.conftest import _make_application, _make_document

    return _make_application([_make_document(**kwargs)])


def _rules_seen(findings):
    return {f["rule"] for f in findings}


# --- structural: heading-keep-with-next -------------------------------------


def test_lint_flags_headings_without_keep_with_next(fake_word):
    # fake_word's headings (Introduction, Risks) default to KeepWithNext off.
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["heading-keep-with-next"])
    assert {f["anchor_id"] for f in findings} == {"heading:1", "heading:3"}
    assert all(f["fixable"] and f["kind"] == "structural" for f in findings)
    assert findings[0]["fix"]["op"] == "format_paragraph"
    assert findings[0]["fix"]["keep_with_next"] is True


def test_regularize_fixes_keep_with_next_and_is_idempotent(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        first = doc.regularize(rules=["heading-keep-with-next"])
        assert len(first["applied"]) == 2
        # The fix actually landed: re-lint is clean and a second pass is a no-op.
        assert doc.lint(rules=["heading-keep-with-next"]) == []
        second = doc.regularize(rules=["heading-keep-with-next"])
    assert second["applied"] == []


def test_regularize_dry_run_writes_nothing(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        report = doc.regularize(rules=["heading-keep-with-next"], dry_run=True)
        assert report["dry_run"] is True
        assert report["applied"] == []
        # Still flagged afterwards — nothing was applied.
        assert len(doc.lint(rules=["heading-keep-with-next"])) == 2


# --- structural: list-numbering-continuity ----------------------------------


def test_lint_flags_split_numbered_lists(monkeypatch):
    # Two abutting numbered lists (end of A == start of B) = one list Word split.
    app = _make_app(lists=[{"start": 0, "end": 10, "type": 3}, {"start": 10, "end": 20, "type": 3}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["list-numbering-continuity"])
    assert len(findings) == 1
    f = findings[0]
    assert f["anchor_id"] == "range:0-20"
    assert [op["op"] for op in f["fix"]] == ["remove_list", "apply_list"]


def test_lint_flags_split_number_only_lists(monkeypatch):
    # number-only (WdListType.LIST_NUM_ONLY = 1) and mixed (5) numbered lists
    # suffer the same split footgun and must be flagged, not just simple/outline.
    app = _make_app(lists=[{"start": 0, "end": 10, "type": 1}, {"start": 10, "end": 20, "type": 5}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["list-numbering-continuity"])
    assert len(findings) == 1
    assert findings[0]["anchor_id"] == "range:0-20"


def test_lint_ignores_separated_numbered_lists(monkeypatch):
    # A gap between the two lists -> intentionally separate, not a split.
    app = _make_app(lists=[{"start": 0, "end": 10, "type": 3}, {"start": 15, "end": 25, "type": 3}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["list-numbering-continuity"])
    assert findings == []


# --- structural: table-repeat-header (no false positive on single-page) ------


def test_lint_single_page_table_not_flagged(fake_word):
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["table-repeat-header"])
    assert findings == []


# --- consistency: heading-font-consistent + idempotent targeted fix ----------


def test_lint_flags_heading_font_size_override(fake_word):
    # Bump heading:1's effective size above its style baseline (12pt).
    fake_word.ActiveDocument.Paragraphs._items[0].Range.Font.Size = 15.0
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["heading-font-consistent"])
    assert len(findings) == 1
    f = findings[0]
    # Consistency rules walk paragraphs, so a heading is addressed as para:N (the
    # same index space as heading:N — both resolve to the paragraph).
    assert f["anchor_id"] == "para:1"
    assert f["kind"] == "consistency"
    assert f["fix"] == {"op": "format_run", "anchor_id": "para:1", "size": 12.0}


def test_regularize_heading_font_targeted_fix_is_idempotent(fake_word):
    fake_word.ActiveDocument.Paragraphs._items[0].Range.Font.Size = 15.0
    with wordlive.attach() as word:
        doc = word.documents.active
        first = doc.regularize(rules=["heading-font-consistent"])
        assert len(first["applied"]) == 1
        assert doc.lint(rules=["heading-font-consistent"]) == []
        second = doc.regularize(rules=["heading-font-consistent"])
    assert second["applied"] == []


def test_regularize_attaches_run_batch_failure_detail(fake_word, monkeypatch):
    # When a fix op fails, run_batch's structured failure detail must ride on the
    # raised error rather than being dropped (so the caller sees which fix failed).
    import pytest

    import wordlive._ops as ops_mod
    from wordlive.exceptions import OpError

    failure = {"index": 0, "op": {"op": "format_run"}, "error": "boom", "type": "OpError"}

    def fake_run_batch(doc, ops, *, label, tracked=False):
        return {"ok": False, "ops_run": 0, "label": label, "failure": failure}, OpError("boom")

    monkeypatch.setattr(ops_mod, "run_batch", fake_run_batch)
    fake_word.ActiveDocument.Paragraphs._items[0].Range.Font.Size = 15.0
    with wordlive.attach() as word:
        with pytest.raises(OpError) as ei:
            word.documents.active.regularize(rules=["heading-font-consistent"])
    assert getattr(ei.value, "failure", None) == failure
    assert getattr(ei.value, "ops_run", None) == 0


# --- consistency: mixed-run-format is report-only ----------------------------


def test_lint_mixed_run_heading_is_report_only(fake_word):
    from wordlive.constants import WD_UNDEFINED

    fake_word.ActiveDocument.Paragraphs._items[0].Range.Font.Size = WD_UNDEFINED
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["mixed-run-format"])
    assert len(findings) == 1
    assert findings[0]["fixable"] is False
    assert findings[0]["fix"] is None


# --- rule selection ----------------------------------------------------------


def test_lint_default_runs_consistency_and_structural(fake_word):
    with wordlive.attach() as word:
        findings = word.documents.active.lint()
    # The default set includes the keep-with-next structural rule; policy rules
    # (none yet) would stay off.
    assert "heading-keep-with-next" in _rules_seen(findings)


def test_lint_rules_exclude(fake_word):
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules={"exclude": ["heading-keep-with-next"]})
    assert "heading-keep-with-next" not in _rules_seen(findings)


def test_lint_within_scopes_to_anchor(fake_word):
    # Scope to heading:3 (Risks, offsets 29-35); heading:1 falls outside.
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["heading-keep-with-next"], within="heading:3")
    assert {f["anchor_id"] for f in findings} == {"heading:3"}


def test_lint_within_scopes_consistency_rule(fake_word):
    # LINT-6: a consistency rule (which uses the unified `_in_span`/`_overlaps`
    # span test) must also honour `within`. heading:1 has a size override but
    # falls outside heading:3's span, so it isn't flagged.
    fake_word.ActiveDocument.Paragraphs._items[0].Range.Font.Size = 15.0
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["heading-font-consistent"], within="heading:3")
    assert findings == []


def test_lint_body_font_consistent_positive(fake_word):
    # LINT-6: a body paragraph whose font face drifts from its style is flagged,
    # with a targeted style-value fix.
    fake_word.ActiveDocument.Paragraphs._items[1].Range.Font.Name = "Comic Sans MS"
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["body-font-consistent"])
    assert len(findings) == 1
    f = findings[0]
    assert f["anchor_id"] == "para:2" and f["rule"] == "body-font-consistent"
    assert f["fix"]["op"] == "format_run" and f["fix"]["font"] == "Aptos"


def test_regularize_multiple_rules_ride_one_undo(fake_word):
    # LINT-6: a multi-rule regularize applies every fix under a single UndoRecord
    # (one Ctrl-Z reverts the whole pass).
    fake_word.ActiveDocument.Paragraphs._items[0].Range.Font.Size = 15.0
    fake_word.UndoRecord.StartCustomRecord.reset_mock()
    fake_word.UndoRecord.EndCustomRecord.reset_mock()
    with wordlive.attach() as word:
        report = word.documents.active.regularize(
            rules=["heading-keep-with-next", "heading-font-consistent"]
        )
    applied_rules = {a["rule"] for a in report["applied"]}
    assert {"heading-keep-with-next", "heading-font-consistent"} <= applied_rules
    fake_word.UndoRecord.StartCustomRecord.assert_called_once()
    fake_word.UndoRecord.EndCustomRecord.assert_called_once()


# --- exec op -----------------------------------------------------------------


def test_regularize_exec_op_returns_report(fake_word):
    from wordlive._ops import run_batch

    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "regularize", "rules": ["heading-keep-with-next"]}], label="t"
        )
    assert exc is None
    report = result["outputs"][0]
    assert report["op"] == "regularize"
    assert len(report["applied"]) == 2


# --- CLI ---------------------------------------------------------------------


def _invoke(args):
    from click.testing import CliRunner

    from wordlive.cli.main import main

    res = CliRunner().invoke(main, args, catch_exceptions=False)
    return res.exit_code, res.stdout


def test_cli_lint_json(fake_word):
    import json

    code, out = _invoke(["--json", "lint", "--rule", "heading-keep-with-next"])
    assert code == 0
    findings = json.loads(out)
    assert {f["anchor_id"] for f in findings} == {"heading:1", "heading:3"}


def test_cli_regularize_dry_run(fake_word):
    import json

    code, out = _invoke(["--json", "regularize", "--rule", "heading-keep-with-next", "--dry-run"])
    assert code == 0
    report = json.loads(out)
    assert report["dry_run"] is True
    assert report["applied"] == []


def test_cli_read_format(fake_word):
    import json

    code, out = _invoke(["--json", "read", "format", "--anchor-id", "heading:1"])
    assert code == 0
    info = json.loads(out)
    assert info["anchor_id"] == "heading:1"
    assert "keep_with_next" in info["paragraph"]


# --- typography (P2 text-scan) batch -----------------------------------------


def _typo_app(monkeypatch, paras):
    app = _make_app(paragraphs=paras)
    _attach(monkeypatch, app)
    return app


def test_lint_trailing_whitespace(monkeypatch):
    _typo_app(monkeypatch, [{"text": "foo  ", "start": 0, "end": 6}])
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["trailing-whitespace"])
    assert len(findings) == 1
    f = findings[0]
    assert f["anchor_id"] == "para:1" and f["kind"] == "structural"
    assert f["fix"]["op"] == "find_replace" and f["fix"]["mode"] == "regex"
    assert f["fix"]["in"] == "para:1" and f["fix"]["required"] is False


def test_lint_leading_whitespace(monkeypatch):
    _typo_app(monkeypatch, [{"text": "   indented", "start": 0, "end": 12}])
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["leading-whitespace"])
    assert {f["anchor_id"] for f in findings} == {"para:1"}
    assert findings[0]["fix"]["find"] == r"^[ \t]+"


def test_lint_double_space_skips_verbatim(monkeypatch):
    _typo_app(
        monkeypatch,
        [
            {"text": "a  b", "start": 0, "end": 5, "style": "Normal"},
            {"text": "x  y", "start": 5, "end": 10, "style": "HTML Preformatted"},
        ],
    )
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["double-space"])
    assert {f["anchor_id"] for f in findings} == {"para:1"}  # verbatim para skipped


def test_lint_space_before_punctuation(monkeypatch):
    _typo_app(monkeypatch, [{"text": "hello ,world", "start": 0, "end": 13}])
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["space-before-punctuation"])
    assert len(findings) == 1
    assert findings[0]["fix"]["text"] == r"\1"


def test_lint_hyphen_as_range_off_by_default(monkeypatch):
    _typo_app(monkeypatch, [{"text": "spanning 1990-1995", "start": 0, "end": 19}])
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "hyphen-as-range" not in _rules_seen(doc.lint())  # off by default
        findings = doc.lint(rules=["hyphen-as-range"])  # opt-in by id
    assert len(findings) == 1
    assert "–" in findings[0]["fix"]["text"]  # en-dash backreference replacement


def test_lint_em_dash_usage_report_only(monkeypatch):
    _typo_app(monkeypatch, [{"text": "a — b", "start": 0, "end": 6}])
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["em-dash-usage"])
    assert len(findings) == 1
    assert findings[0]["fixable"] is False and findings[0]["fix"] is None


def test_lint_tabs_for_layout_report_only(monkeypatch):
    _typo_app(monkeypatch, [{"text": "name\tvalue", "start": 0, "end": 11}])
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["tabs-for-layout"])
    assert len(findings) == 1 and findings[0]["fixable"] is False


def test_lint_manual_line_break_report_only(monkeypatch):
    _typo_app(monkeypatch, [{"text": "line one\x0bline two", "start": 0, "end": 18}])
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["manual-line-break"])
    assert len(findings) == 1 and findings[0]["rule"] == "manual-line-break"


def test_lint_manual_heading_formatting_report_only(monkeypatch):
    app = _typo_app(
        monkeypatch, [{"text": "Overview", "start": 0, "end": 9, "level": 10, "style": "Normal"}]
    )
    app.ActiveDocument.Paragraphs._items[0].Range.Font.Bold = True
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["manual-heading-formatting"])
    assert len(findings) == 1
    assert findings[0]["fixable"] is False and findings[0]["anchor_id"] == "para:1"


def test_lint_manual_heading_skips_sentences(monkeypatch):
    # A normal sentence (ends in a period) isn't a faux heading even if bold.
    app = _typo_app(
        monkeypatch, [{"text": "This is a sentence.", "start": 0, "end": 20, "style": "Normal"}]
    )
    app.ActiveDocument.Paragraphs._items[0].Range.Font.Bold = True
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["manual-heading-formatting"])
    assert findings == []


def test_lint_table_style_consistent_flags_minority(monkeypatch):
    app = _make_app(
        tables=[
            {"grid": [["A"]], "style": "Grid Table 4"},
            {"grid": [["B"]], "style": "Grid Table 4"},
            {"grid": [["C"]], "style": "Plain Table 1"},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["table-style-consistent"])
    assert len(findings) == 1
    f = findings[0]
    assert f["anchor_id"] == "table:3:1:1"
    assert f["fix"] == {"op": "set_table_style", "table": 3, "style": "Grid Table 4"}


def test_lint_typography_tag_includes_off_by_default_rules(monkeypatch):
    _typo_app(monkeypatch, [{"text": "a  b 1990-1995", "start": 0, "end": 15}])
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint(rules=["typography"]))
    # The tag pulls in both on-by-default (double-space) and off-by-default
    # (hyphen-as-range) typography rules.
    assert {"double-space", "hyphen-as-range"} <= seen


def test_lint_default_excludes_off_by_default_typography(monkeypatch):
    _typo_app(monkeypatch, [{"text": "a — b", "start": 0, "end": 6}])
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint())
    assert "em-dash-usage" not in seen  # default_on=False, not named/tagged


# --- finalization (P3) batch -------------------------------------------------


def test_lint_comments_present_report_only(monkeypatch):
    app = _make_app()
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.comments.add(doc.range(0, 4), "please revise")
        findings = doc.lint(rules=["comments-present"])
    assert len(findings) == 1
    f = findings[0]
    assert f["rule"] == "comments-present"
    assert f["fixable"] is False and f["fix"] is None
    assert f["anchor_id"] == "range:0-4"


def test_lint_unaccepted_revisions_report_only(monkeypatch):
    app = _make_app(revisions=[{"type": 1, "author": "Ann", "start": 0, "end": 4, "text": "new"}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["unaccepted-revisions"])
    assert len(findings) == 1
    f = findings[0]
    assert f["rule"] == "unaccepted-revisions" and f["severity"] == "warning"
    assert f["fixable"] is False


def test_lint_track_changes_on_report_only(monkeypatch):
    app = _make_app()
    app.ActiveDocument.TrackRevisions = True
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["track-changes-on"])
    assert len(findings) == 1
    assert findings[0]["anchor_id"] == "start" and findings[0]["fixable"] is False


def test_lint_track_changes_off_is_clean(monkeypatch):
    app = _make_app()  # TrackRevisions defaults False on the fake
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["track-changes-on"])
    assert findings == []


def test_lint_stale_fields_report_only(monkeypatch):
    app = _make_app(
        fields=[
            {"code": "TOC \\o", "start": 0, "end": 5},
            {"code": "PAGE", "start": 10, "end": 12},
            {"code": "HYPERLINK http://x", "start": 20, "end": 25},  # not updatable
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["stale-fields"])
    assert len(findings) == 1
    f = findings[0]
    assert f["rule"] == "stale-fields" and f["fixable"] is False
    assert "TOC" in f["message"] and "PAGE" in f["message"]


def test_lint_hidden_text_present_report_only(monkeypatch):
    app = _typo_app(monkeypatch, [{"text": "secret", "start": 0, "end": 7}])
    app.ActiveDocument.Paragraphs._items[0].Range.Font.Hidden = -1  # wdToggle on
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["hidden-text-present"])
    assert len(findings) == 1
    assert findings[0]["anchor_id"] == "para:1" and findings[0]["fixable"] is False


def test_lint_leftover_highlight_fix_and_idempotent(monkeypatch):
    from wordlive.constants import WdColorIndex

    app = _typo_app(monkeypatch, [{"text": "flagged", "start": 0, "end": 8}])
    app.ActiveDocument.Paragraphs._items[0].Range.HighlightColorIndex = int(WdColorIndex.YELLOW)
    with wordlive.attach() as word:
        doc = word.documents.active
        findings = doc.lint(rules=["leftover-highlight"])
        assert len(findings) == 1
        assert findings[0]["fix"] == {
            "op": "format_run",
            "anchor_id": "para:1",
            "highlight": "none",
        }
        first = doc.regularize(rules=["leftover-highlight"])
        assert len(first["applied"]) == 1
        # The clear landed: re-lint is clean and a second pass is a no-op.
        assert doc.lint(rules=["leftover-highlight"]) == []
        second = doc.regularize(rules=["leftover-highlight"])
    assert second["applied"] == []


def test_finalization_off_by_default(monkeypatch):
    app = _make_app(revisions=[{"type": 1, "start": 0, "end": 4}])
    app.ActiveDocument.TrackRevisions = True
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint())  # default set, no rules named
    assert not (seen & {"unaccepted-revisions", "track-changes-on"})


def test_finalization_tag_selects_cluster(monkeypatch):
    app = _make_app(
        revisions=[{"type": 1, "start": 0, "end": 4}],
        fields=[{"code": "PAGE", "start": 0, "end": 2}],
    )
    app.ActiveDocument.TrackRevisions = True
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint(rules=["finalization"]))
    assert {"unaccepted-revisions", "track-changes-on", "stale-fields"} <= seen


# --- field-code (P1) batch ---------------------------------------------------

_BROKEN_REF = "Error! Reference source not found."


def test_lint_broken_cross_reference(monkeypatch):
    app = _make_app(
        fields=[{"code": "REF _Ref1 \\h", "result": _BROKEN_REF, "start": 10, "end": 20}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["broken-cross-reference"])
    assert len(findings) == 1
    f = findings[0]
    assert f["anchor_id"] == "range:10-20" and f["kind"] == "structural"
    assert f["severity"] == "warning" and f["fixable"] is False and f["fix"] is None


def test_lint_broken_cross_reference_healthy_is_clean(monkeypatch):
    app = _make_app(
        fields=[{"code": "REF _Ref1 \\h", "result": "Figure 1", "start": 10, "end": 20}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["broken-cross-reference"]) == []


def test_lint_broken_cross_reference_within_scoping(monkeypatch):
    app = _make_app(fields=[{"code": "REF x", "result": _BROKEN_REF, "start": 100, "end": 110}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert len(doc.lint(rules=["broken-cross-reference"])) == 1
        # A within-span that misses the field excludes it.
        assert doc.lint(rules=["broken-cross-reference"], within="range:0-5") == []


def test_lint_caption_manual_numbering(monkeypatch):
    app = _make_app(
        paragraphs=[{"text": "Figure 3: Results", "style": "Caption", "start": 0, "end": 18}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["caption-manual-numbering"])
    assert len(findings) == 1
    assert findings[0]["anchor_id"] == "para:1" and findings[0]["fixable"] is False


def test_lint_caption_with_seq_field_is_clean(monkeypatch):
    app = _make_app(
        paragraphs=[{"text": "Figure 3: Results", "style": "Caption", "start": 0, "end": 18}],
        fields=[{"code": "SEQ Figure \\* ARABIC", "result": "3", "start": 6, "end": 7}],
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["caption-manual-numbering"]) == []


def test_lint_caption_without_number_is_clean(monkeypatch):
    # A Caption paragraph with no figure/table number isn't a manual-numbering defect.
    app = _make_app(
        paragraphs=[{"text": "A descriptive caption", "style": "Caption", "start": 0, "end": 22}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["caption-manual-numbering"]) == []


def test_lint_page_numbers_present_off_by_default(monkeypatch):
    app = _make_app()  # one section, empty headers/footers, no PAGE field
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "page-numbers-present" not in _rules_seen(doc.lint())  # policy, off
        assert {f["rule"] for f in doc.lint(rules=["page-numbers-present"])} == {
            "page-numbers-present"
        }
        assert {f["rule"] for f in doc.lint(rules=["layout"])} == {"page-numbers-present"}


def test_lint_page_numbers_present_clean_with_footer_field(monkeypatch):
    from tests.conftest import _FakeFields

    app = _make_app()
    footer = app.ActiveDocument.Sections(1).Footers(1)  # primary footer
    footer.Range.Fields = _FakeFields([{"code": "PAGE", "type": 33, "start": 0, "end": 2}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["layout"]) == []


def test_fields_batch_on_by_default(monkeypatch):
    app = _make_app(
        paragraphs=[{"text": "Figure 3", "style": "Caption", "start": 0, "end": 9}],
        fields=[{"code": "REF x", "result": _BROKEN_REF, "start": 20, "end": 30}],
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint())  # default set
    assert {"broken-cross-reference", "caption-manual-numbering"} <= seen
    assert "page-numbers-present" not in seen  # policy stays off


def test_fields_academia_tag_selects_cluster(monkeypatch):
    app = _make_app(
        paragraphs=[{"text": "Table 2", "style": "Caption", "start": 0, "end": 8}],
        fields=[{"code": "PAGEREF y", "result": _BROKEN_REF, "start": 20, "end": 30}],
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint(rules=["academia"]))
    assert {"broken-cross-reference", "caption-manual-numbering"} <= seen
    assert "page-numbers-present" not in seen  # layout, not academia


# --- MCP ---------------------------------------------------------------------


def test_mcp_read_lint_and_format_info(fake_word):
    import pytest

    pytest.importorskip("mcp")
    from wordlive.mcp._worker import InlineWorker
    from wordlive.mcp.server import _read_impl

    w = InlineWorker()
    findings = _read_impl(w, "lint", {"rules": ["heading-keep-with-next"]})
    assert {f["anchor_id"] for f in findings} == {"heading:1", "heading:3"}
    info = _read_impl(w, "format_info", {"anchor_id": "heading:1"})
    assert info["style"] == "Normal"


def test_mcp_write_regularize(fake_word):
    import pytest

    pytest.importorskip("mcp")
    from wordlive.mcp._worker import InlineWorker
    from wordlive.mcp.server import _write_impl

    out = _write_impl(InlineWorker(), "regularize", {"rules": ["heading-keep-with-next"]})
    assert out["ok"] is True
    assert len(out["result"]["applied"]) == 2
