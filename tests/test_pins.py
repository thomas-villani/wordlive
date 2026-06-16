"""Durable handles (`pin:` / `pin_outline` / `bind`), `$ops` refs, stale hints.

These exercise the *plumbing* against the fake COM fixture: minting `_wl_`
bookmarks, the `pin:`<->`_wl_` mapping, idempotency-by-start, batch output
references, and the recovery hints on a stale positional anchor. Durability of a
handle across a real insert/delete is the smoke suite's job — the fake stores
fixed offsets and never shifts them.
"""

from __future__ import annotations

import pytest

import wordlive
from wordlive._ops import _resolve_op_refs, run_batch
from wordlive.exceptions import AnchorNotFoundError, OpError


@pytest.fixture
def codes(monkeypatch: pytest.MonkeyPatch):
    """Deterministic pin codes: code1, code2, … (patches `_new_pin_code`)."""
    counter = iter(f"code{i}" for i in range(1, 1000))
    monkeypatch.setattr("wordlive._document._new_pin_code", lambda: next(counter))


# ---------------------------------------------------------------------------
# Phase 1 — pin / stamp
# ---------------------------------------------------------------------------


def test_pin_mints_and_returns_handle(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("pin"):
            result = doc.pin("para:2")
    assert result == {"anchor_id": "pin:code1", "pin": "pin:code1", "target": "para:2"}
    # The hidden bookmark was minted under the `_wl_` namespace.
    assert doc.com.Bookmarks.Exists("_wl_code1")


def test_stamp_is_an_alias(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("stamp"):
            result = doc.stamp("para:2")
    assert result["pin"] == "pin:code1"


def test_pin_with_slug(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("pin"):
            result = doc.pin("para:2", name="budget-intro")
    assert result["pin"] == "pin:budget-intro"
    # A hyphenated slug stores as `_wl_budget_intro` (hyphens -> underscores).
    assert doc.com.Bookmarks.Exists("_wl_budget_intro")


@pytest.mark.parametrize("bad", ["Bad Slug", "x_y", "UPPER", "trailing-", "a--b"])
def test_pin_slug_validation(fake_word, bad):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.pin("para:2", name=bad)


def test_pin_slug_too_long(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError, match="40-character"):
            doc.pin("para:2", name="a" * 40)  # `_wl_` + 40 = 44 > 40


def test_anchor_by_id_resolves_pin(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("pin"):
            doc.pin("para:2")
        anchor = doc.anchor_by_id("pin:code1")
    assert anchor.anchor_id == "pin:code1"
    # Resolves to the pinned paragraph's range (para:2 spans 13–29 in the fixture).
    assert int(anchor.com.Start) == 13


def test_missing_pin_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc:
            doc.anchor_by_id("pin:deadbeef")
    assert exc.value.kind == "pin"


def test_pin_bookmarks_stay_hidden(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("pin"):
            doc.pin("para:2")
        names = doc.bookmarks.list()
        hidden = doc.bookmarks.list(include_hidden=True)
    assert "_wl_code1" not in names  # filtered from the user-facing list
    assert "_wl_code1" in hidden


def test_pin_op_in_batch_outputs(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "pin", "anchor_id": "para:2"}], label="t")
    assert exc is None
    assert result["outputs"] == [
        {"index": 0, "op": "pin", "anchor_id": "pin:code1", "pin": "pin:code1", "target": "para:2"}
    ]


# ---------------------------------------------------------------------------
# Phase 4 — pin_outline / outline(pin=True)
# ---------------------------------------------------------------------------


def test_pin_outline_maps_every_heading(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("pin-outline"):
            pins = doc.pin_outline()
    # Fixture headings: heading:1 ("Introduction"), heading:3 ("Risks").
    assert pins == {"heading:1": "pin:code1", "heading:3": "pin:code2"}


def test_pin_outline_is_idempotent(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("pin-outline"):
            first = doc.pin_outline()
        with doc.edit("pin-outline"):
            second = doc.pin_outline()
    assert first == second  # reuse-by-start mints nothing new the second time
    # Two headings -> two `_wl_` handles, even after a second pass.
    assert sum(n.startswith("_wl_") for n in doc.bookmarks.list(include_hidden=True)) == 2


def test_pin_outline_levels_filter(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("pin-outline"):
            pins = doc.pin_outline(levels=(1, 1))  # only level-1 headings
    assert pins == {"heading:1": "pin:code1"}


def test_outline_pin_true_adds_pin_field(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("outline-pin"):
            rows = doc.outline(pin=True)
    assert [r["anchor_id"] for r in rows] == ["heading:1", "heading:3"]
    assert all("pin" in r for r in rows)


def test_outline_default_has_no_pin(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.outline()
    assert all("pin" not in r for r in rows)


# ---------------------------------------------------------------------------
# Phase 2 — bind on insert ops
# ---------------------------------------------------------------------------


def test_insert_block_bind_returns_pin(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_block", "anchor_id": "end", "items": ["Hello"], "bind": "intro"}],
            label="t",
        )
    assert exc is None
    out = result["outputs"][0]
    assert out["pin"] == "pin:intro"
    assert doc.com.Bookmarks.Exists("_wl_intro")


def test_insert_block_bind_true_auto_code(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_block", "anchor_id": "end", "items": ["Hi"], "bind": True}],
            label="t",
        )
    assert result["outputs"][0]["pin"] == "pin:code1"


def test_insert_paragraph_bind(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_paragraph", "anchor_id": "end", "text": "Note", "bind": "note"}],
            label="t",
        )
    assert exc is None
    assert result["outputs"][0]["pin"] == "pin:note"


def test_insert_block_no_bind_no_pin(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "insert_block", "anchor_id": "end", "items": ["x"]}], label="t"
        )
    assert "pin" not in result["outputs"][0]


def test_bind_is_not_an_unexpected_field(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, _ = run_batch(
            doc,
            [{"op": "insert_block", "anchor_id": "end", "items": ["x"], "bind": "h"}],
            label="t",
        )
    assert "warnings" not in result  # bind is a known optional field, not flagged


# ---------------------------------------------------------------------------
# Phase 3 — $ops[N].field references
# ---------------------------------------------------------------------------


def test_resolve_op_refs_substitutes():
    op = {"op": "set_cell", "table": "$ops[0].table", "row": 1}
    resolved = _resolve_op_refs(op, {0: {"table": 7}})
    assert resolved == {"op": "set_cell", "table": 7, "row": 1}
    # Original is untouched.
    assert op["table"] == "$ops[0].table"


def test_resolve_op_refs_nested_in_list():
    op = {"op": "x", "items": ["literal", "$ops[0].pin"]}
    resolved = _resolve_op_refs(op, {0: {"pin": "pin:abc"}})
    assert resolved["items"] == ["literal", "pin:abc"]


def test_resolve_op_refs_partial_string_untouched():
    op = {"op": "x", "text": "see $ops[0].table for details"}
    resolved = _resolve_op_refs(op, {0: {"table": 1}})
    assert resolved["text"] == "see $ops[0].table for details"  # not a whole-string match


def test_resolve_op_refs_forward_reference_errors():
    with pytest.raises(OpError, match="has not produced output"):
        _resolve_op_refs({"op": "x", "a": "$ops[5].y"}, {})


def test_resolve_op_refs_unknown_field_errors():
    with pytest.raises(OpError, match="no output field"):
        _resolve_op_refs({"op": "x", "a": "$ops[0].nope"}, {0: {"table": 1}})


def test_ops_reference_end_to_end(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {"op": "pin", "anchor_id": "para:2"},
                {"op": "replace", "anchor_id": "$ops[0].anchor_id", "text": "Z"},
            ],
            label="t",
        )
    assert exc is None
    assert result["ops_run"] == 2


def test_ops_reference_failure_recorded(fake_word, codes):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {"op": "pin", "anchor_id": "para:2"},
                {"op": "replace", "anchor_id": "$ops[0].missing", "text": "Z"},
            ],
            label="t",
        )
    assert isinstance(exc, OpError)
    assert result["ok"] is False
    assert result["failure"]["index"] == 1


# ---------------------------------------------------------------------------
# Phase 5 — stale-anchor diagnostics
# ---------------------------------------------------------------------------


def test_para_out_of_range_hint(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc:
            doc.anchor_by_id("para:99").com
    hint = exc.value.hint or ""
    assert "out of range" in hint
    assert "3 paragraph" in hint
    assert "pin" in hint.lower()


def test_paragraph_collection_out_of_range_hint(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc:
            _ = doc.paragraphs[99]
    assert "out of range" in (exc.value.hint or "")


def test_heading_not_a_heading_hint(fake_word):
    # Para 2 is body text, so heading:2 misses with a "not a heading" hint.
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc:
            doc.anchor_by_id("heading:2").com
    hint = exc.value.hint or ""
    assert "not a heading" in hint
    assert "nearest heading is heading:" in hint


def test_stale_hint_in_batch_failure(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "replace", "anchor_id": "para:99", "text": "x"}], label="t"
        )
    assert isinstance(exc, AnchorNotFoundError)
    assert "out of range" in result["failure"]["error"]
