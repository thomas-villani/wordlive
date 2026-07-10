"""Linter profile loader + the policy-rule cluster it unlocks (Batch 4a).

Covers `Profile.load` (path / dict / None / malformed), the three policy rules
(`body-justified`, `body-line-spacing`, `table-numeric-right-align`) firing/clean
under a profile, severity override, disabling a default rule, the idempotency
contract for the policy fixes, and the `profile=` round-trip through CLI / MCP / exec.
"""

from __future__ import annotations

import json

import pytest

import wordlive

# Reuse the tiny attach/build helpers from the main linter test module.
from tests.test_linting import _attach, _invoke, _make_app, _rules_seen
from wordlive._lint_profile import Profile
from wordlive.constants import WdLineSpacing, WdParagraphAlignment
from wordlive.exceptions import OpError

_JUSTIFY = int(WdParagraphAlignment.JUSTIFY)
_RIGHT = int(WdParagraphAlignment.RIGHT)
_DOUBLE = int(WdLineSpacing.DOUBLE)

_POLICY_IDS = {"body-justified", "body-line-spacing", "table-numeric-right-align"}


# --- Profile.load ------------------------------------------------------------


def test_profile_load_none_is_empty():
    p = Profile.load(None)
    assert p.rules == {}
    assert p.is_enabled("body-justified") is None
    assert p.config_for("body-justified") == {}


def test_profile_load_dict():
    p = Profile.load(
        {
            "rules": {
                "body-justified": {"enabled": True, "severity": "warning"},
                "double-space": {"enabled": False},
            }
        }
    )
    assert p.is_enabled("body-justified") is True
    assert p.severity_for("body-justified") == "warning"
    assert p.is_enabled("double-space") is False
    assert p.is_enabled("never-mentioned") is None


def test_profile_bare_mention_enables():
    # A rule listed with only a target (no explicit `enabled`) still counts as on.
    p = Profile.load({"rules": {"body-line-spacing": {"target": "1.5"}}})
    assert p.is_enabled("body-line-spacing") is True
    assert p.config_for("body-line-spacing")["target"] == "1.5"


def test_profile_load_file(tmp_path):
    path = tmp_path / "wordlive.lint.json"
    path.write_text(json.dumps({"rules": {"body-justified": {"enabled": True}}}), encoding="utf-8")
    p = Profile.load(str(path))
    assert p.is_enabled("body-justified") is True


def test_profile_load_passthrough():
    original = Profile.load({"rules": {"body-justified": {"enabled": True}}})
    assert Profile.load(original) is original


def test_profile_load_empty_file_is_empty(tmp_path):
    path = tmp_path / "empty.json"
    path.write_text("", encoding="utf-8")
    assert Profile.load(str(path)).rules == {}


def test_profile_load_malformed_json_raises(tmp_path):
    path = tmp_path / "bad.json"
    path.write_text("{not json", encoding="utf-8")
    with pytest.raises(OpError):
        Profile.load(str(path))


def test_profile_load_non_object_raises(tmp_path):
    path = tmp_path / "list.json"
    path.write_text("[1, 2, 3]", encoding="utf-8")
    with pytest.raises(OpError):
        Profile.load(str(path))


def test_profile_missing_file_raises():
    with pytest.raises(OpError):
        Profile.load("no/such/profile.json")


# --- body-justified ----------------------------------------------------------


def test_body_justified_fires(monkeypatch):
    app = _make_app(paragraphs=[{"text": "Left-aligned body.", "start": 0, "end": 18}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(
            profile={"rules": {"body-justified": {"enabled": True}}}
        )
    body = [f for f in findings if f["rule"] == "body-justified"]
    assert len(body) == 1
    f = body[0]
    assert f["anchor_id"] == "para:1" and f["kind"] == "policy"
    assert f["fixable"] is True and f["fix"]["op"] == "format_paragraph"
    assert f["fix"]["alignment"] == "justify"


def test_body_justified_clean_when_justified(monkeypatch):
    app = _make_app(
        paragraphs=[{"text": "Already justified.", "start": 0, "end": 18, "alignment": _JUSTIFY}]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(
            profile={"rules": {"body-justified": {"enabled": True}}}
        )
    assert [f for f in findings if f["rule"] == "body-justified"] == []


def test_body_justified_skips_headings_and_empty(monkeypatch):
    app = _make_app(
        paragraphs=[
            {"text": "A heading", "start": 0, "end": 9, "level": 1},
            {"text": "", "start": 9, "end": 10},  # empty body paragraph
        ]
    )
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(
            profile={"rules": {"body-justified": {"enabled": True}}}
        )
    assert [f for f in findings if f["rule"] == "body-justified"] == []


def test_body_justified_off_without_profile(monkeypatch):
    app = _make_app(paragraphs=[{"text": "Left-aligned body.", "start": 0, "end": 18}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        assert "body-justified" not in _rules_seen(word.documents.active.lint())


# --- body-line-spacing -------------------------------------------------------


def test_body_line_spacing_fires(monkeypatch):
    app = _make_app(paragraphs=[{"text": "Single-spaced body.", "start": 0, "end": 19}])
    _attach(monkeypatch, app)
    profile = {"rules": {"body-line-spacing": {"enabled": True, "target": "double"}}}
    with wordlive.attach() as word:
        findings = word.documents.active.lint(profile=profile)
    body = [f for f in findings if f["rule"] == "body-line-spacing"]
    assert len(body) == 1
    assert body[0]["fix"]["line_spacing"] == "double" and body[0]["fixable"] is True


def test_body_line_spacing_noop_without_target(monkeypatch):
    # Enabled but no target → the rule polices nothing.
    app = _make_app(paragraphs=[{"text": "Single-spaced body.", "start": 0, "end": 19}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        findings = word.documents.active.lint(
            profile={"rules": {"body-line-spacing": {"enabled": True}}}
        )
    assert [f for f in findings if f["rule"] == "body-line-spacing"] == []


def test_body_line_spacing_clean_when_matches(monkeypatch):
    app = _make_app(
        paragraphs=[
            {"text": "Double-spaced body.", "start": 0, "end": 19, "line_spacing_rule": _DOUBLE}
        ]
    )
    _attach(monkeypatch, app)
    profile = {"rules": {"body-line-spacing": {"enabled": True, "target": "double"}}}
    with wordlive.attach() as word:
        findings = word.documents.active.lint(profile=profile)
    assert [f for f in findings if f["rule"] == "body-line-spacing"] == []


# --- table-numeric-right-align -----------------------------------------------


def _numeric_table_app():
    # Column 2 is numeric (skips the header row 1); column 1 is text.
    return _make_app(
        tables=[{"grid": [["Item", "Amount"], ["Widget", "1,200"], ["Gadget", "3.50"]]}]
    )


def test_table_numeric_right_align_fires(monkeypatch):
    app = _numeric_table_app()
    _attach(monkeypatch, app)
    profile = {"rules": {"table-numeric-right-align": {"enabled": True}}}
    with wordlive.attach() as word:
        findings = word.documents.active.lint(profile=profile)
    hits = {f["anchor_id"] for f in findings if f["rule"] == "table-numeric-right-align"}
    assert hits == {"table:1:2:2", "table:1:3:2"}  # both numeric body cells, not the text column


def test_table_numeric_text_column_not_flagged(monkeypatch):
    app = _make_app(tables=[{"grid": [["City", "Region"], ["Paris", "EU"], ["Tokyo", "APAC"]]}])
    _attach(monkeypatch, app)
    profile = {"rules": {"table-numeric-right-align": {"enabled": True}}}
    with wordlive.attach() as word:
        findings = word.documents.active.lint(profile=profile)
    assert [f for f in findings if f["rule"] == "table-numeric-right-align"] == []


def test_table_numeric_clean_when_right_aligned(monkeypatch):
    app = _numeric_table_app()
    table = app.ActiveDocument.Tables(1)
    for r in (2, 3):
        table.Cell(r, 2).Range.ParagraphFormat.Alignment = _RIGHT
    _attach(monkeypatch, app)
    profile = {"rules": {"table-numeric-right-align": {"enabled": True}}}
    with wordlive.attach() as word:
        findings = word.documents.active.lint(profile=profile)
    assert [f for f in findings if f["rule"] == "table-numeric-right-align"] == []


def test_table_numeric_threshold_respected(monkeypatch):
    # A mixed column (1 numeric of 2 non-empty = 0.5) with a 0.8 threshold is not numeric.
    app = _make_app(tables=[{"grid": [["H"], ["100"], ["N/A"]]}])
    _attach(monkeypatch, app)
    profile = {"rules": {"table-numeric-right-align": {"enabled": True, "threshold": 0.8}}}
    with wordlive.attach() as word:
        findings = word.documents.active.lint(profile=profile)
    assert [f for f in findings if f["rule"] == "table-numeric-right-align"] == []


# --- severity override, disable, default-set exclusion -----------------------


def test_profile_severity_override(monkeypatch):
    app = _make_app(paragraphs=[{"text": "Left-aligned body.", "start": 0, "end": 18}])
    _attach(monkeypatch, app)
    profile = {"rules": {"body-justified": {"enabled": True, "severity": "warning"}}}
    with wordlive.attach() as word:
        findings = word.documents.active.lint(profile=profile)
    body = [f for f in findings if f["rule"] == "body-justified"]
    assert body and body[0]["severity"] == "warning"  # overridden from the rule's "info"


def test_profile_disables_a_default_rule(fake_word):
    # heading-keep-with-next fires by default on the seeded headings; a profile can turn it off.
    with wordlive.attach() as word:
        doc = word.documents.active
        assert "heading-keep-with-next" in _rules_seen(doc.lint())
        off = doc.lint(profile={"rules": {"heading-keep-with-next": {"enabled": False}}})
        assert "heading-keep-with-next" not in _rules_seen(off)


def test_policy_rules_absent_from_default_set(fake_word):
    with wordlive.attach() as word:
        seen = _rules_seen(word.documents.active.lint())
    assert not (_POLICY_IDS & seen)


# --- idempotency of the policy fixes -----------------------------------------


def test_regularize_body_justified_idempotent(monkeypatch):
    app = _make_app(paragraphs=[{"text": "Left-aligned body.", "start": 0, "end": 18}])
    _attach(monkeypatch, app)
    profile = {"rules": {"body-justified": {"enabled": True}}}
    with wordlive.attach() as word:
        doc = word.documents.active
        first = doc.regularize(profile=profile)
        assert any(f["rule"] == "body-justified" for f in first["applied"])
        second = doc.regularize(profile=profile)
        assert second["applied"] == []  # the contract: a second pass changes nothing


def test_regularize_table_numeric_idempotent(monkeypatch):
    app = _numeric_table_app()
    _attach(monkeypatch, app)
    profile = {"rules": {"table-numeric-right-align": {"enabled": True}}}
    with wordlive.attach() as word:
        doc = word.documents.active
        first = doc.regularize(profile=profile)
        assert any(f["rule"] == "table-numeric-right-align" for f in first["applied"])
        second = doc.regularize(profile=profile)
        assert second["applied"] == []


# --- surface round-trips: CLI / MCP / exec op --------------------------------


def test_cli_lint_profile(fake_word, tmp_path):
    # fake_word's body paragraph (para:2) is left-aligned Normal → body-justified fires.
    path = tmp_path / "profile.json"
    path.write_text(json.dumps({"rules": {"body-justified": {"enabled": True}}}), encoding="utf-8")
    code, out = _invoke(["--json", "lint", "--profile", str(path)])
    assert code == 0
    findings = json.loads(out)
    assert any(f["rule"] == "body-justified" for f in findings)


def test_mcp_lint_profile(fake_word):
    pytest.importorskip("mcp")
    from wordlive.mcp._worker import InlineWorker
    from wordlive.mcp.server import _read_impl

    w = InlineWorker()
    findings = _read_impl(w, "lint", {"profile": {"rules": {"body-justified": {"enabled": True}}}})
    assert any(f["rule"] == "body-justified" for f in findings)


def test_exec_regularize_profile(monkeypatch):
    from wordlive._ops import run_batch

    app = _make_app(paragraphs=[{"text": "Left-aligned body.", "start": 0, "end": 18}])
    _attach(monkeypatch, app)
    profile = {"rules": {"body-justified": {"enabled": True}}}
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "regularize", "profile": profile}], label="t")
        assert exc is None and result.get("ops_run", 0) >= 1
        # The profile threaded through the exec op → the fix applied → re-lint is clean.
        assert [f for f in doc.lint(profile=profile) if f["rule"] == "body-justified"] == []


def test_exec_regularize_allow_content_is_not_warned_as_unused(monkeypatch):
    # `regularize` honours `allow_content`, so the batch must not warn that it was
    # ignored — a phantom warning on a field that demonstrably works undermines
    # the whole `warnings` channel.
    from wordlive._ops import run_batch, unexpected_fields

    assert unexpected_fields({"op": "regularize", "allow_content": True}, "regularize") == []

    app = _make_app(paragraphs=[{"text": "Body.", "start": 0, "end": 5}])
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "regularize", "allow_content": True}], label="t")
    assert exc is None
    assert "warnings" not in result, result.get("warnings")
