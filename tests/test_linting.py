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


# --- content gate: adds_content + allow_content ------------------------------


def _gate_findings():
    """A formatting fix (applies by default) and an adds_content fix (gated)."""
    from wordlive._linting import Finding

    fmt = Finding(
        rule="fake-format",
        kind="consistency",
        severity="info",
        anchor_id="heading:1",
        message="drifted",
        fixable=True,
        fix={"op": "format_paragraph", "anchor_id": "heading:1", "keep_with_next": True},
    )
    content = Finding(
        rule="fake-content",
        kind="structural",
        severity="warning",
        anchor_id="heading:3",
        message="stray",
        fixable=True,
        adds_content=True,
        fix={"op": "format_paragraph", "anchor_id": "heading:3", "keep_with_next": True},
    )
    return [fmt, content]


def test_lint_findings_carry_adds_content_flag(fake_word):
    # Every finding dict now exposes adds_content; ordinary rules leave it False.
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["heading-keep-with-next"])
    assert findings and all(f["adds_content"] is False for f in findings)


def test_regularize_withholds_content_fixes_by_default(fake_word, monkeypatch):
    from wordlive import _linting

    monkeypatch.setattr(_linting, "run_lint", lambda *a, **k: _gate_findings())
    with wordlive.attach() as word:
        report = word.documents.active.regularize()
    assert {f["rule"] for f in report["applied"]} == {"fake-format"}
    assert {f["rule"] for f in report["deferred"]} == {"fake-content"}
    assert report["skipped"] == []


def test_regularize_applies_content_fixes_when_allowed(fake_word, monkeypatch):
    from wordlive import _linting

    monkeypatch.setattr(_linting, "run_lint", lambda *a, **k: _gate_findings())
    with wordlive.attach() as word:
        report = word.documents.active.regularize(allow_content=True)
    assert {f["rule"] for f in report["applied"]} == {"fake-format", "fake-content"}
    assert report["deferred"] == []


def test_regularize_dry_run_lists_deferred_content_fixes(fake_word, monkeypatch):
    from wordlive import _linting

    monkeypatch.setattr(_linting, "run_lint", lambda *a, **k: _gate_findings())
    with wordlive.attach() as word:
        report = word.documents.active.regularize(dry_run=True)
    assert report["dry_run"] is True
    assert report["applied"] == []
    assert {f["rule"] for f in report["deferred"]} == {"fake-content"}


def test_regularize_exec_op_honors_allow_content(fake_word, monkeypatch):
    from wordlive import _linting
    from wordlive._ops import apply_op

    monkeypatch.setattr(_linting, "run_lint", lambda *a, **k: _gate_findings())
    with wordlive.attach() as word:
        doc = word.documents.active
        gated = apply_op(doc, {"op": "regularize"})
        assert {f["rule"] for f in gated["deferred"]} == {"fake-content"}
        opened = apply_op(doc, {"op": "regularize", "allow_content": True})
        assert {f["rule"] for f in opened["applied"]} == {"fake-format", "fake-content"}


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
        # The `layout` tag pulls the whole §H cluster; page-numbers-present is in it.
        assert "page-numbers-present" in {f["rule"] for f in doc.lint(rules=["layout"])}


def test_lint_page_numbers_present_clean_with_footer_field(monkeypatch):
    from tests.conftest import _FakeFields

    app = _make_app()
    footer = app.ActiveDocument.Sections(1).Footers(1)  # primary footer
    footer.Range.Fields = _FakeFields([{"code": "PAGE", "type": 33, "start": 0, "end": 2}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["page-numbers-present"]) == []


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


def test_lint_xref_as_literal_text(monkeypatch):
    app = _make_app(
        paragraphs=[{"text": "As shown in Figure 3, the trend holds.", "start": 0, "end": 38}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["xref-as-literal-text"])
    assert len(findings) == 1
    f = findings[0]
    assert f["anchor_id"] == "para:1" and f["kind"] == "structural"
    assert f["severity"] == "info" and f["fixable"] is False and f["fix"] is None


def test_lint_xref_with_ref_field_is_clean(monkeypatch):
    # A REF field covering the paragraph means the mention is a real cross-reference.
    app = _make_app(
        paragraphs=[{"text": "As shown in Figure 3, the trend holds.", "start": 0, "end": 38}],
        fields=[{"code": "REF _Ref1 \\h", "result": "Figure 3", "start": 12, "end": 20}],
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["xref-as-literal-text"]) == []


def test_lint_xref_in_caption_is_clean(monkeypatch):
    # A Caption paragraph reading "Figure 3" is caption-manual-numbering's concern, not
    # an unlinked cross-reference.
    app = _make_app(
        paragraphs=[{"text": "Figure 3: Results", "style": "Caption", "start": 0, "end": 18}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["xref-as-literal-text"]) == []


def test_lint_xref_off_by_default(monkeypatch):
    app = _make_app(
        paragraphs=[{"text": "As shown in Figure 3, the trend holds.", "start": 0, "end": 38}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "xref-as-literal-text" not in _rules_seen(doc.lint())  # heuristic, off
        assert _rules_seen(doc.lint(rules=["crossref"])) >= {"xref-as-literal-text"}
        assert _rules_seen(doc.lint(rules=["academia"])) >= {"xref-as-literal-text"}


def test_lint_xref_within_scoping(monkeypatch):
    app = _make_app(
        paragraphs=[{"text": "As shown in Figure 3, the trend holds.", "start": 100, "end": 138}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert len(doc.lint(rules=["xref-as-literal-text"])) == 1
        assert doc.lint(rules=["xref-as-literal-text"], within="range:0-5") == []


# --- hyperlinks (§I) batch ---------------------------------------------------


def test_lint_hyperlink_broken_internal(monkeypatch):
    # An in-document jump (sub_address set, address empty) to a bookmark that
    # doesn't exist — a dead link.
    app = _make_app(
        hyperlinks=[{"text": "jump", "sub_address": "Ghost", "start": 10, "end": 14}],
        bookmarks={"Target": (0, 5)},
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["hyperlink-broken-internal"])
    assert len(findings) == 1
    f = findings[0]
    assert f["anchor_id"] == "range:10-14" and f["kind"] == "structural"
    assert f["severity"] == "warning" and f["fixable"] is False and f["fix"] is None
    assert "Ghost" in f["message"]


def test_lint_hyperlink_internal_valid_is_clean(monkeypatch):
    # Jump to a bookmark that exists (Bookmarks.Exists → True) — not flagged.
    app = _make_app(
        hyperlinks=[{"text": "jump", "sub_address": "Target", "start": 10, "end": 14}],
        bookmarks={"Target": (0, 5)},
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["hyperlink-broken-internal"]) == []


def test_lint_hyperlink_external_not_broken_internal(monkeypatch):
    # An external link (address set) with a sub_address is a fragment on a remote
    # URL, not an in-document jump — never a broken-internal defect.
    app = _make_app(
        hyperlinks=[
            {
                "text": "docs",
                "address": "https://x.example",
                "sub_address": "sec",
                "start": 5,
                "end": 9,
            }
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["hyperlink-broken-internal"]) == []


def test_lint_hyperlink_broken_internal_within_scoping(monkeypatch):
    app = _make_app(hyperlinks=[{"text": "jump", "sub_address": "Ghost", "start": 100, "end": 110}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert len(doc.lint(rules=["hyperlink-broken-internal"])) == 1
        assert doc.lint(rules=["hyperlink-broken-internal"], within="range:0-5") == []


def test_lint_hyperlink_bare_for_print(monkeypatch):
    # External link whose visible text doesn't contain its URL — hidden on paper.
    app = _make_app(
        hyperlinks=[{"text": "Acme", "address": "https://acme.example/page", "start": 5, "end": 9}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "hyperlink-bare-for-print" not in _rules_seen(doc.lint())  # off by default
        findings = doc.lint(rules=["hyperlink-bare-for-print"])
    assert len(findings) == 1
    assert findings[0]["kind"] == "policy" and findings[0]["fixable"] is False


def test_lint_hyperlink_bare_for_print_url_visible_is_clean(monkeypatch):
    # The label already contains the URL, so the destination is visible in print.
    app = _make_app(
        hyperlinks=[
            {
                "text": "Acme (https://acme.example/page)",
                "address": "https://acme.example/page",
                "start": 5,
                "end": 37,
            }
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["hyperlink-bare-for-print"]) == []


def test_lint_hyperlink_display_is_raw_url(monkeypatch):
    # The whole visible text is a bare URL where a label was wanted.
    app = _make_app(
        hyperlinks=[
            {
                "text": "https://raw.example/very/long/path",
                "address": "https://raw.example/very/long/path",
                "start": 5,
                "end": 39,
            }
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "hyperlink-display-is-raw-url" not in _rules_seen(doc.lint())  # off by default
        findings = doc.lint(rules=["hyperlink-display-is-raw-url"])
    assert len(findings) == 1
    assert findings[0]["kind"] == "consistency" and findings[0]["fixable"] is False
    # A raw-URL display is NOT also flagged bare-for-print (the URL is visible).
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["hyperlink-bare-for-print"]) == []


def test_lint_hyperlink_labelled_is_not_raw_url(monkeypatch):
    app = _make_app(
        hyperlinks=[{"text": "Acme", "address": "https://acme.example", "start": 5, "end": 9}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["hyperlink-display-is-raw-url"]) == []


def test_hyperlink_broken_internal_on_by_default(monkeypatch):
    app = _make_app(hyperlinks=[{"text": "jump", "sub_address": "Ghost", "start": 10, "end": 14}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint())  # default set, no rules named
    assert "hyperlink-broken-internal" in seen
    assert not (seen & {"hyperlink-bare-for-print", "hyperlink-display-is-raw-url"})


def test_hyperlinks_tag_selects_cluster(monkeypatch):
    app = _make_app(
        hyperlinks=[
            {"text": "jump", "sub_address": "Ghost", "start": 10, "end": 14},
            {"text": "Acme", "address": "https://acme.example/page", "start": 20, "end": 24},
            {
                "text": "https://raw.example/x",
                "address": "https://raw.example/x",
                "start": 30,
                "end": 51,
            },
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert _rules_seen(doc.lint(rules=["hyperlinks"])) == {
            "hyperlink-broken-internal",
            "hyperlink-bare-for-print",
            "hyperlink-display-is-raw-url",
        }
        # The `print` tag selects only the two print/sharing rules.
        assert _rules_seen(doc.lint(rules=["print"])) == {
            "hyperlink-bare-for-print",
            "hyperlink-display-is-raw-url",
        }


# --- layout / document-level (§H) batch --------------------------------------


def test_lint_document_properties_filled(monkeypatch):
    # Title is set, Author is missing from the bag (unset) — Author is flagged.
    app = _make_app(builtin_properties={"Title": "My Report"})
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "document-properties-filled" not in _rules_seen(doc.lint())  # policy, off
        findings = doc.lint(rules=["document-properties-filled"])
    assert len(findings) == 1
    f = findings[0]
    assert f["kind"] == "policy" and f["fixable"] is False and f["anchor_id"] == "start"
    assert "Author" in f["message"]


def test_lint_document_properties_filled_clean_when_set(monkeypatch):
    app = _make_app(builtin_properties={"Title": "My Report", "Author": "Jane Roe"})
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["document-properties-filled"]) == []


def test_lint_document_properties_filled_profile_required_override(monkeypatch):
    # A profile can name the required set; here only Company (which is unset) matters.
    app = _make_app(builtin_properties={"Title": "My Report", "Author": "Jane Roe"})
    _attach(monkeypatch, app)
    profile = {"rules": {"document-properties-filled": {"required": ["Company"]}}}
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["document-properties-filled"], profile=profile)
    assert len(findings) == 1 and "Company" in findings[0]["message"]


def test_document_properties_filled_opts_in_via_profile(monkeypatch):
    # A policy rule the profile enables joins the default set.
    app = _make_app(builtin_properties={"Title": "My Report"})
    _attach(monkeypatch, app)
    profile = {"rules": {"document-properties-filled": {"enabled": True}}}
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint(profile=profile))
    assert "document-properties-filled" in seen


def test_lint_confidentiality_notice_absent(monkeypatch):
    # Profile demands a confidentiality notice; it's nowhere in the document.
    app = _make_app(content="Just some body text.")
    _attach(monkeypatch, app)
    profile = {"rules": {"confidentiality-notice": {"text": "CONFIDENTIAL"}}}
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["confidentiality-notice"], profile=profile)
    assert len(findings) == 1
    f = findings[0]
    assert f["kind"] == "policy" and f["severity"] == "warning" and f["fixable"] is False
    assert "CONFIDENTIAL" in f["message"]


def test_lint_confidentiality_notice_present_in_footer_is_clean(monkeypatch):
    app = _make_app(sections=[{"footers": {"primary": "CONFIDENTIAL — Acme Corp"}}])
    _attach(monkeypatch, app)
    profile = {"rules": {"confidentiality-notice": {"text": "CONFIDENTIAL"}}}
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["confidentiality-notice"], profile=profile) == []


def test_confidentiality_notice_silent_without_configured_text(monkeypatch):
    # No `text` configured → the rule polices nothing, even when selected.
    app = _make_app(content="Body.")
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["confidentiality-notice"]) == []


def test_lint_copyright_notice_defaults_to_symbol(monkeypatch):
    # No configured text → the © symbol; absent here, so it fires.
    app = _make_app(content="Body text with no copyright line.")
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "copyright-notice" not in _rules_seen(doc.lint())  # policy, off by default
        findings = doc.lint(rules=["copyright-notice"])
    assert len(findings) == 1 and findings[0]["kind"] == "policy"


def test_lint_copyright_notice_present_in_body_is_clean(monkeypatch):
    app = _make_app(content="Body text. © 2026 Acme Corp. More text.")
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["copyright-notice"]) == []


def test_lint_header_footer_consistent_flags_divergent_text(monkeypatch):
    app = _make_app(
        sections=[
            {"headers": {"primary": "Report A"}},
            {"headers": {"primary": "Report B"}},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "header-footer-consistent" not in _rules_seen(doc.lint())  # off by default
        findings = doc.lint(rules=["header-footer-consistent"])
    assert len(findings) == 1
    f = findings[0]
    assert f["kind"] == "consistency" and f["anchor_id"] == "header:2:primary"
    assert "Report A" in f["message"] and "Report B" in f["message"]


def test_lint_header_footer_consistent_clean_when_matching(monkeypatch):
    app = _make_app(
        sections=[
            {"headers": {"primary": "Report"}},
            {"headers": {"primary": "Report"}},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["header-footer-consistent"]) == []


def test_lint_header_footer_consistent_single_section_is_clean(monkeypatch):
    app = _make_app(sections=[{"headers": {"primary": "Report"}}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["header-footer-consistent"]) == []


def test_lint_draft_watermark_present(monkeypatch):
    app = _make_app()
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.set_watermark("DRAFT")
        # The new read surface round-trips the watermark text.
        mark = doc.watermark()
        assert mark is not None and mark.text == "DRAFT" and mark.sections == [1]
        assert "draft-watermark-present" not in _rules_seen(doc.lint())  # off by default
        findings = doc.lint(rules=["draft-watermark-present"])
    assert len(findings) == 1
    f = findings[0]
    assert f["kind"] == "structural" and f["severity"] == "warning" and f["fixable"] is False
    assert "DRAFT" in f["message"]


def test_watermark_read_none_and_after_remove(monkeypatch):
    app = _make_app()
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.watermark() is None
        doc.set_watermark("CONFIDENTIAL")
        assert doc.watermark().text == "CONFIDENTIAL"
        doc.remove_watermark()
        assert doc.watermark() is None
        assert doc.lint(rules=["draft-watermark-present"]) == []


def test_layout_and_notices_tags_select_clusters(monkeypatch):
    # A bare doc: no props, no PAGE field, no © — the layout tag pulls the whole
    # cluster (incl. the shipped page-numbers-present), so the fired subset shows
    # the new rules alongside it.
    app = _make_app()
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        layout_seen = _rules_seen(doc.lint(rules=["layout"]))
        assert {
            "document-properties-filled",
            "copyright-notice",
            "page-numbers-present",
        } <= layout_seen
        # The `notices` tag selects only the two notice rules.
        profile = {"rules": {"confidentiality-notice": {"text": "SECRET"}}}
        assert _rules_seen(doc.lint(rules=["notices"], profile=profile)) == {
            "confidentiality-notice",
            "copyright-notice",
        }


# --- heading & document structure (§B) batch --------------------------------


def test_lint_heading_level_skip(monkeypatch):
    # H1 then H3 with no H2 between them: the outline skips level 2.
    app = _make_app(
        paragraphs=[
            {"level": 1, "text": "Intro", "start": 0, "end": 6},
            {"level": 10, "text": "Body", "start": 6, "end": 11},
            {"level": 3, "text": "Deep", "start": 11, "end": 16},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["heading-level-skip"])
    assert len(findings) == 1
    f = findings[0]
    assert f["anchor_id"] == "heading:3" and f["kind"] == "structural"
    assert f["fixable"] is False and "skips level 2" in f["message"]


def test_lint_heading_level_skip_allows_deep_start(monkeypatch):
    # A document that simply starts deep (H2, H3) and nests consistently is fine —
    # the skip is measured against the *previous* heading, not an assumed H1.
    app = _make_app(
        paragraphs=[
            {"level": 2, "text": "Sub", "start": 0, "end": 4},
            {"level": 3, "text": "Subsub", "start": 4, "end": 11},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["heading-level-skip"]) == []


def test_lint_empty_heading(monkeypatch):
    app = _make_app(
        paragraphs=[
            {"level": 1, "text": "Real", "start": 0, "end": 5},
            {"level": 2, "text": "", "start": 5, "end": 6},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(rules=["empty-heading"])
    assert len(findings) == 1
    assert findings[0]["anchor_id"] == "heading:2" and findings[0]["fixable"] is False


def test_lint_adjacent_headings_off_by_default(monkeypatch):
    # heading:1 and heading:2 are consecutive paragraphs — no body between.
    app = _make_app(
        paragraphs=[
            {"level": 1, "text": "Title", "start": 0, "end": 6},
            {"level": 2, "text": "Subtitle", "start": 6, "end": 15},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "adjacent-headings" not in _rules_seen(doc.lint())  # off by default
        findings = doc.lint(rules=["adjacent-headings"])
    assert len(findings) == 1
    assert findings[0]["anchor_id"] == "heading:1" and "Subtitle" in findings[0]["message"]


def test_lint_adjacent_headings_not_flagged_with_body_between(monkeypatch):
    app = _make_app(
        paragraphs=[
            {"level": 1, "text": "Title", "start": 0, "end": 6},
            {"level": 10, "text": "Some body", "start": 6, "end": 16},
            {"level": 2, "text": "Section", "start": 16, "end": 24},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["adjacent-headings"]) == []


def test_lint_heading_numbering_manual(monkeypatch):
    app = _make_app(paragraphs=[{"level": 1, "text": "3.1 Methods", "start": 0, "end": 12}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "heading-numbering-manual" not in _rules_seen(doc.lint())  # off by default
        findings = doc.lint(rules=["heading-numbering-manual"])
    assert len(findings) == 1
    assert findings[0]["anchor_id"] == "heading:1" and "'3.1'" in findings[0]["message"]


def test_lint_heading_numbering_manual_skips_auto_numbered(monkeypatch):
    # A heading Word auto-numbers (a live ListString) is not double-flagged, even if
    # its text happens to open with digits.
    app = _make_app(paragraphs=[{"level": 1, "text": "3.1 Methods", "start": 0, "end": 12}])
    app.ActiveDocument.Paragraphs._items[0].Range.ListFormat.ListString = "3.1"
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["heading-numbering-manual"]) == []


def test_lint_heading_trailing_period_fix_shape(monkeypatch):
    app = _make_app(paragraphs=[{"level": 1, "text": "Summary.", "start": 0, "end": 9}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "heading-trailing-period" not in _rules_seen(doc.lint())  # off by default
        findings = doc.lint(rules=["heading-trailing-period"])
    assert len(findings) == 1
    f = findings[0]
    assert f["anchor_id"] == "heading:1" and f["fixable"] is True
    assert f["fix"]["op"] == "find_replace" and f["fix"]["mode"] == "regex"
    # The fix scopes to the *paragraph* (para:N), not the heading anchor, whose
    # find_replace scope would expand to the body under the heading.
    assert f["fix"]["in"] == "para:1" and f["fix"]["required"] is False


def test_lint_heading_trailing_period_ignores_ellipsis(monkeypatch):
    app = _make_app(paragraphs=[{"level": 1, "text": "To be continued...", "start": 0, "end": 19}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["heading-trailing-period"]) == []


def test_lint_toc_present_and_current(monkeypatch):
    # A level-1 heading and no TOC field → the document lacks a table of contents.
    app = _make_app(paragraphs=[{"level": 1, "text": "Chapter", "start": 0, "end": 8}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "toc-present-and-current" not in _rules_seen(doc.lint())  # off by default
        findings = doc.lint(rules=["toc-present-and-current"])
    assert len(findings) == 1
    assert findings[0]["anchor_id"] == "start" and findings[0]["fixable"] is False


def test_lint_toc_not_flagged_when_present(monkeypatch):
    app = _make_app(
        paragraphs=[{"level": 1, "text": "Chapter", "start": 0, "end": 8}],
        fields=[{"code": 'TOC \\o "1-3" \\h', "type": 13, "start": 0, "end": 8}],
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["toc-present-and-current"]) == []


def test_lint_toc_not_flagged_without_top_level_headings(monkeypatch):
    # Only H2s — a TOC isn't expected, so the absence of one isn't reported.
    app = _make_app(paragraphs=[{"level": 2, "text": "Sub", "start": 0, "end": 4}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert word.documents.active.lint(rules=["toc-present-and-current"]) == []


def test_heading_structure_tags_select_cluster(monkeypatch):
    app = _make_app(
        paragraphs=[
            {"level": 1, "text": "1. Intro.", "start": 0, "end": 10},
            {"level": 3, "text": "Deep.", "start": 10, "end": 16},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        # `structure` pulls the whole §B cluster (on- and off-by-default alike).
        structure_seen = _rules_seen(doc.lint(rules=["structure"]))
        assert {
            "heading-level-skip",
            "heading-numbering-manual",
            "heading-trailing-period",
            "toc-present-and-current",
        } <= structure_seen
        # `headings` selects the §B cluster too, alongside the v1 heading rules.
        headings_seen = _rules_seen(doc.lint(rules=["headings"]))
        assert {"heading-level-skip", "heading-keep-with-next"} <= headings_seen


def test_lint_default_includes_on_by_default_heading_rules(monkeypatch):
    app = _make_app(
        paragraphs=[
            {"level": 1, "text": "Intro", "start": 0, "end": 6},
            {"level": 3, "text": "", "start": 6, "end": 7},
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint())
    # The two unambiguous outline defects run in the default set...
    assert {"heading-level-skip", "empty-heading"} <= seen
    # ...but the opinionated ones stay off until named/tagged.
    assert "adjacent-headings" not in seen and "heading-trailing-period" not in seen


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
