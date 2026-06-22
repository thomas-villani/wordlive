"""Checkpoint + diff — `doc.checkpoint()` / `doc.changes_since()` / `doc.diff()`.

The diff classifier is exercised by checkpointing two distinct `fake_word`
document states (re-attaching the monkeypatched app between them) and diffing the
tokens — so the alignment/classification logic round-trips off-Windows. The
`text+format` *reformat* path needs real per-paragraph format divergence, which
the fake doesn't vary, so that one is asserted structurally here and validated
live in the smoke pass.
"""

from __future__ import annotations

import pytest

import wordlive
from wordlive import Checkpoint, _com, _findreplace
from wordlive._checkpoint import _change_hash, _sha1, diff_checkpoints
from wordlive.exceptions import OpError


def _mk_cp(paras, tables=None, *, include="text+style", styles=None, fmts=None):
    """Build a `Checkpoint` directly (no COM) so the diff classifier — especially
    the table-diff and restyle/reformat gating — can be unit-tested precisely."""
    tables = tables or []
    pgs = []
    for i, text in enumerate(paras):
        style = (styles or {}).get(i, "Normal")
        fmt = (fmts or {}).get(i)
        norm = _findreplace._normalize(text).text
        pgs.append(
            {
                "i": i,
                "text": text,
                "style": style,
                "level": 10,
                "list": None,
                "fmt": fmt,
                "key": _sha1(norm),
                "hash": _change_hash(include, norm, style, fmt),
            }
        )
    doc_hash = _sha1(
        "\n".join(p["hash"] for p in pgs)
        + "\n\x1etables\x1e\n"
        + "\n".join(t["cells_hash"] for t in tables)
    )
    return Checkpoint(
        version=1,
        include=include,
        scope=None,
        paragraphs=pgs,
        tables=tables,
        doc_hash=doc_hash,
    )


def _attach(monkeypatch, app):
    monkeypatch.setattr(_com, "get_active_word", lambda: app)
    monkeypatch.setattr(_com, "launch_word", lambda visible=True: app)


def _make_app(**kwargs):
    from tests.conftest import _make_application, _make_document

    return _make_application([_make_document(**kwargs)])


def _paras(specs):
    """`specs` is a list of (text, style) — build self-consistent paragraph dicts.
    style defaults to 'Normal' when a bare string is given."""
    out = []
    pos = 0
    for spec in specs:
        text, style = (spec, "Normal") if isinstance(spec, str) else spec
        out.append(
            {"level": 10, "text": text, "style": style, "start": pos, "end": pos + len(text) + 1}
        )
        pos += len(text) + 1
    return out


def _checkpoint(monkeypatch, paras, **kw):
    app = _make_app(paragraphs=_paras(paras))
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        return word.documents.active.checkpoint(**kw)


# --- checkpoint shape & determinism -----------------------------------------


def test_checkpoint_is_deterministic(monkeypatch):
    a = _checkpoint(monkeypatch, ["Alpha", "Beta", "Gamma"])
    b = _checkpoint(monkeypatch, ["Alpha", "Beta", "Gamma"])
    assert a.doc_hash == b.doc_hash
    assert [p["key"] for p in a.paragraphs] == [p["key"] for p in b.paragraphs]
    assert a.version == 1 and a.include == "text+style" and a.scope is None
    # i is 0-based, anchor maps to 1-based para:N
    assert [p["i"] for p in a.paragraphs] == [0, 1, 2]


def test_normalization_ignores_cosmetic_whitespace(monkeypatch):
    # Same words, extra spacing — normalised text (and so key/hash) must match.
    a = _checkpoint(monkeypatch, ["Costs fell 4%."])
    b = _checkpoint(monkeypatch, ["Costs   fell 4%."])
    assert a.paragraphs[0]["key"] == b.paragraphs[0]["key"]
    assert a.doc_hash == b.doc_hash


# --- diff: fast path & the four core opcodes --------------------------------


def test_unchanged_returns_empty_via_fast_path(monkeypatch):
    a = _checkpoint(monkeypatch, ["One", "Two", "Three"])
    b = _checkpoint(monkeypatch, ["One", "Two", "Three"])
    assert diff_checkpoints(a, b) == []


def test_replace_carries_text_and_current_anchor(monkeypatch):
    a = _checkpoint(monkeypatch, ["Intro", "Costs fell 4%.", "End"])
    b = _checkpoint(monkeypatch, ["Intro", "Costs fell 9%.", "End"])
    changes = diff_checkpoints(a, b)
    assert len(changes) == 1
    c = changes[0]
    assert c["op"] == "replace"
    assert c["text_before"] == "Costs fell 4%." and c["text_after"] == "Costs fell 9%."
    assert c["anchor_id"] == "para:2" and c["index_after"] == 1 and c["index_before"] == 1


def test_insert_does_not_misreport_renumbered_tail(monkeypatch):
    a = _checkpoint(monkeypatch, ["P1", "P2", "P3"])
    b = _checkpoint(monkeypatch, ["P1", "NEW", "P2", "P3"])
    changes = diff_checkpoints(a, b)
    assert len(changes) == 1
    c = changes[0]
    assert c["op"] == "insert" and c["text_after"] == "NEW"
    assert c["anchor_id"] == "para:2" and c["index_after"] == 1
    assert "index_before" not in c


def test_replace_beside_insert_pairs_by_similarity(monkeypatch):
    # An edit ("4%"→"9%") next to a brand-new paragraph: the close pair must be
    # the replace and the unrelated paragraph the insert — not positional pairing.
    a = _checkpoint(monkeypatch, ["Intro", "Costs fell 4%.", "End"])
    b = _checkpoint(monkeypatch, ["Intro", "A brand new line.", "Costs fell 9%.", "End"])
    changes = diff_checkpoints(a, b)
    by_op = {c["op"]: c for c in changes}
    assert set(by_op) == {"replace", "insert"}
    assert by_op["replace"]["text_before"] == "Costs fell 4%."
    assert by_op["replace"]["text_after"] == "Costs fell 9%."
    assert by_op["insert"]["text_after"] == "A brand new line."


def test_delete_references_old_index_no_anchor(monkeypatch):
    a = _checkpoint(monkeypatch, ["P1", "P2", "P3"])
    b = _checkpoint(monkeypatch, ["P1", "P3"])
    changes = diff_checkpoints(a, b)
    assert len(changes) == 1
    c = changes[0]
    assert c["op"] == "delete" and c["text_before"] == "P2" and c["index_before"] == 1
    assert "anchor_id" not in c


# --- diff: restyle / reformat & include depth -------------------------------


def test_restyle_detected_in_text_plus_style(monkeypatch):
    a = _checkpoint(monkeypatch, [("Heading", "Normal"), ("Body", "Normal")])
    b = _checkpoint(monkeypatch, [("Heading", "Heading 2"), ("Body", "Normal")])
    changes = diff_checkpoints(a, b)
    assert len(changes) == 1
    c = changes[0]
    assert c["op"] == "restyle"
    assert c["style_before"] == "Normal" and c["style_after"] == "Heading 2"
    assert c["anchor_id"] == "para:1"


def test_text_mode_ignores_restyle(monkeypatch):
    a = _checkpoint(monkeypatch, [("Heading", "Normal")], include="text")
    b = _checkpoint(monkeypatch, [("Heading", "Heading 2")], include="text")
    # style is outside the 'text' fingerprint, so doc_hash matches → no change
    assert a.doc_hash == b.doc_hash
    assert diff_checkpoints(a, b) == []


def test_reformat_only_in_text_plus_format_with_real_fmt_diff(monkeypatch):
    # Same text, same style, different format fingerprint → reformat (text+format).
    a = _mk_cp(["Body"], include="text+format", fmts={0: "fmtA"})
    b = _mk_cp(["Body"], include="text+format", fmts={0: "fmtB"})
    changes = diff_checkpoints(a, b)
    assert [c["op"] for c in changes] == ["reformat"]


def test_no_reformat_verdict_without_format_fingerprint(monkeypatch):
    # In text+style mode there is no `fmt`, so a same-text hash difference can only
    # be a restyle — never a `reformat` (which would have no format data behind it).
    a = _mk_cp(["Body"], include="text+style", styles={0: "Normal"})
    b = _mk_cp(["Body"], include="text+style", styles={0: "Quote"})
    changes = diff_checkpoints(a, b)
    assert [c["op"] for c in changes] == ["restyle"]


def test_table_cell_change_surfaces_as_table_change(monkeypatch):
    # A table-only edit must not slip through (doc_hash breaks the fast path, then
    # the paragraph diff finds nothing) — it surfaces as a coarse table_change.
    a = _mk_cp(["Body"], tables=[{"index": 1, "shape": [2, 2], "cells_hash": _sha1("old")}])
    b = _mk_cp(["Body"], tables=[{"index": 1, "shape": [2, 2], "cells_hash": _sha1("new")}])
    changes = diff_checkpoints(a, b)
    assert [c["op"] for c in changes] == ["table_change"]
    assert changes[0]["anchor_id"] == "table:1"


def test_table_insert_and_delete(monkeypatch):
    none = _mk_cp(["Body"], tables=[])
    one = _mk_cp(["Body"], tables=[{"index": 1, "shape": [2, 2], "cells_hash": _sha1("x")}])
    assert [c["op"] for c in diff_checkpoints(none, one)] == ["table_insert"]
    assert [c["op"] for c in diff_checkpoints(one, none)] == ["table_delete"]


def test_edit_amid_blank_lines_detected_with_correct_anchor(monkeypatch):
    # Blank paragraphs share an alignment key, but the real content edit must still
    # be classified correctly with the right anchor (the documented limitation is
    # only spurious blank-line churn, not a misclassified real change).
    a = _checkpoint(monkeypatch, ["Alpha", "", "Beta", "", "Gamma"])
    b = _checkpoint(monkeypatch, ["Alpha", "", "Beta edited", "", "Gamma"])
    changes = diff_checkpoints(a, b)
    assert any(
        c["op"] == "replace" and c["text_after"] == "Beta edited" and c["anchor_id"] == "para:3"
        for c in changes
    )


def test_changes_since_rejects_offset_range_scope(monkeypatch):
    base = _checkpoint(monkeypatch, ["X"])
    forged = Checkpoint(
        version=1,
        include=base.include,
        scope="range:0-5",
        paragraphs=base.paragraphs,
        tables=base.tables,
        doc_hash=base.doc_hash,
    )
    app = _make_app(paragraphs=_paras(["X"]))
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        with pytest.raises(OpError, match="offset-based scope"):
            word.documents.active.changes_since(forged)


def test_text_plus_format_populates_fmt(monkeypatch):
    cp = _checkpoint(monkeypatch, ["Body"], include="text+format")
    assert cp.include == "text+format"
    assert cp.paragraphs[0]["fmt"] is not None
    # an identical doc still diffs clean
    cp2 = _checkpoint(monkeypatch, ["Body"], include="text+format")
    assert diff_checkpoints(cp, cp2) == []


def test_diff_rejects_mismatched_include(monkeypatch):
    a = _checkpoint(monkeypatch, ["X"], include="text")
    b = _checkpoint(monkeypatch, ["X"], include="text+style")
    with pytest.raises(OpError):
        diff_checkpoints(a, b)


# --- serialisation & token round-trip ---------------------------------------


def test_to_json_from_json_round_trips(monkeypatch):
    a = _checkpoint(monkeypatch, ["One", "Two"])
    token = a.to_json()
    restored = Checkpoint.from_json(token)
    assert restored == a
    # diff accepts the JSON string and the parsed dict directly
    import json

    assert diff_checkpoints(token, a) == []
    assert diff_checkpoints(json.loads(token), a) == []


def test_changes_since_rebuilds_now(monkeypatch):
    a = _checkpoint(monkeypatch, ["Intro", "old line", "End"])
    # Re-attach an edited document state; changes_since fingerprints "now".
    app_b = _make_app(paragraphs=_paras(["Intro", "new line", "End"]))
    _attach(monkeypatch, app_b)
    with wordlive.attach() as word:
        changes = word.documents.active.changes_since(a)
    assert [c["op"] for c in changes] == ["replace"]
    assert changes[0]["anchor_id"] == "para:2"


# --- include validation & scope ---------------------------------------------


def test_bad_include_raises(monkeypatch):
    app = _make_app(paragraphs=_paras(["X"]))
    _attach(monkeypatch, app)
    with wordlive.attach() as word:
        with pytest.raises(OpError):
            word.documents.active.checkpoint(include="text+everything")


def test_within_scopes_and_records_anchor(fake_word):
    # fake_word: heading:1 spans 0–13 (just "Introduction"); within clips the walk.
    with wordlive.attach() as word:
        cp = word.documents.active.checkpoint(within="heading:1")
    assert cp.scope == "heading:1"
    assert cp.tables == []  # tables skipped for a scoped checkpoint
    assert all(p["text"] == "Introduction" for p in cp.paragraphs)
