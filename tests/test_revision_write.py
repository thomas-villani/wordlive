"""Revision write surface — accept/reject, bulk accept/reject, and the
revision-aware read model (`text_final` / `text_original` / `revision_segments`).

The pure reconstruction logic (`segment_runs`) is tested directly; the COM-backed
accept/reject paths round-trip against the `fake_word` MagicMock, whose default
document carries one tracked insertion ("Body", range 13–18, inside the body
paragraph para:2 = "Body text here.").
"""

from __future__ import annotations

import json
from unittest.mock import MagicMock

import wordlive
from wordlive._revisions import segment_runs
from wordlive.cli.main import EXIT_OK, main


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


# --- segment_runs (pure) -------------------------------------------------------
#
# `final_text` is the range's Range.Text (the FINAL view — inserted runs present,
# deleted runs absent). Delete runs carry their text on the run itself, since the
# deleted characters are gone from final_text. Offsets are markup-space, where the
# deleted text still occupies positions (validated against live Word).


def _ins(start, end, text):
    return {"change": "insert", "start": start, "end": end, "text": text}


def _del(start, end, text):
    return {"change": "delete", "start": start, "end": end, "text": text}


def test_segment_runs_no_runs_is_one_unchanged_segment():
    segs = segment_runs("hello world", 0, [])
    assert segs == [{"text": "hello world", "change": None}]


def test_segment_runs_pure_delete_restores_text_not_in_final():
    # Nothing is in final_text, but the deleted run carries its own text.
    segs = segment_runs("", 0, [_del(0, 5, "gone")])
    assert segs == [{"text": "gone", "change": "delete"}]


def test_segment_runs_insert_then_unchanged():
    # "Body " (markup 13–18) inserted, "text here." unchanged.
    segs = segment_runs("Body text here.", 13, [_ins(13, 18, "Body ")])
    assert segs == [
        {"text": "Body ", "change": "insert"},
        {"text": "text here.", "change": None},
    ]


def test_segment_runs_replace_interleaves_insert_and_deleted_text():
    # The canonical tracked find_replace("quick" -> "slow"), as live Word reports
    # it: final_text has "slow"; "quick" survives only on the delete run.
    segs = segment_runs(
        "the slow brown fox",
        0,
        [_ins(4, 8, "slow"), _del(8, 13, "quick")],
    )
    assert segs == [
        {"text": "the ", "change": None},
        {"text": "slow", "change": "insert"},
        {"text": "quick", "change": "delete"},
        {"text": " brown fox", "change": None},
    ]
    final = "".join(s["text"] for s in segs if s["change"] != "delete")
    original = "".join(s["text"] for s in segs if s["change"] != "insert")
    assert final == "the slow brown fox"
    assert original == "the quick brown fox"


def test_segment_runs_skips_runs_before_the_cursor():
    # A run starting before base_start (already-consumed space) is ignored.
    segs = segment_runs("xyz", 10, [_ins(5, 8, "??")])
    assert segs == [{"text": "xyz", "change": None}]


def test_segment_runs_ignores_non_text_changes():
    segs = segment_runs("abc", 0, [{"change": "format", "start": 0, "end": 3, "text": ""}])
    assert segs == [{"text": "abc", "change": None}]


# --- revision-aware reads on an anchor -----------------------------------------


def test_text_final_keeps_inserted_runs(fake_word):
    with wordlive.attach() as word:
        para = word.documents.active.paragraphs[2]  # "Body text here." + insert 13–18
        # final = as if accepted: the inserted "Body " stays.
        assert para.text_final.startswith("Body text here.")


def test_text_original_drops_inserted_runs(fake_word):
    with wordlive.attach() as word:
        para = word.documents.active.paragraphs[2]
        # original = as if rejected: the inserted "Body " is gone.
        assert para.text_original.startswith("text here.")
        assert "Body " not in para.text_original


def test_revision_segments_classifies_runs(fake_word):
    with wordlive.attach() as word:
        segs = word.documents.active.paragraphs[2].revision_segments()
    assert segs[0] == {"text": "Body ", "change": "insert"}
    assert segs[1]["change"] is None and segs[1]["text"].startswith("text here.")


# --- accept / reject one --------------------------------------------------------


def test_revision_accept_consumes_it(fake_word):
    with wordlive.attach() as word:
        revs = word.documents.active.revisions
        assert len(revs) == 1
        revs[1].accept()
        assert len(word.documents.active.revisions) == 0


def test_revision_reject_consumes_it(fake_word):
    with wordlive.attach() as word:
        word.documents.active.revisions[1].reject()
        assert len(word.documents.active.revisions) == 0


# --- accept-all / reject-all ----------------------------------------------------


def test_accept_all_returns_count_and_clears(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.revisions.accept_all() == 1
        assert len(doc.revisions) == 0


def test_reject_all_empty_is_zero(fake_word):
    fake_word.ActiveDocument.Revisions = _empty_revisions()
    with wordlive.attach() as word:
        assert word.documents.active.revisions.reject_all() == 0


def test_accept_all_within_anchor_scopes_to_its_range(fake_word):
    # Seed the body range (13–29) with its own two tracked changes; accept_all
    # scoped to that anchor resolves exactly those.
    fake_word.ActiveDocument.Range(13, 29).Revisions = _revisions_with(2)
    with wordlive.attach() as word:
        doc = word.documents.active
        within = doc.range(13, 29)
        assert doc.revisions.accept_all(within=within) == 2
        # The whole-document collection is untouched by the scoped accept.
        assert len(doc.revisions) == 1


# --- CLI ------------------------------------------------------------------------


def test_cli_revision_accept(fake_word):
    code, out = _invoke(["--json", "revision", "accept", "--index", "1"])
    assert code == EXIT_OK
    assert json.loads(out) == {"ok": True, "index": 1, "accepted": True}


def test_cli_revision_reject(fake_word):
    code, out = _invoke(["--json", "revision", "reject", "--index", "1"])
    assert code == EXIT_OK
    assert json.loads(out)["rejected"] is True


def test_cli_revision_accept_all(fake_word):
    code, out = _invoke(["--json", "revision", "accept-all"])
    assert code == EXIT_OK
    assert json.loads(out)["accepted"] == 1


def test_cli_revision_list_alias(fake_word):
    code, out = _invoke(["--json", "revision", "list"])
    assert code == EXIT_OK
    assert json.loads(out)[0]["type"] == "insert"


def test_cli_read_text_final(fake_word):
    code, out = _invoke(["--json", "read", "text", "--anchor-id", "para:2", "--view", "final"])
    assert code == EXIT_OK
    payload = json.loads(out)
    assert payload["view"] == "final"
    assert payload["text"].startswith("Body text here.")


def test_cli_read_text_segments(fake_word):
    code, out = _invoke(["--json", "read", "text", "--anchor-id", "para:2", "--view", "segments"])
    assert code == EXIT_OK
    segs = json.loads(out)["segments"]
    assert segs[0] == {"text": "Body ", "change": "insert"}


# --- exec ops -------------------------------------------------------------------


def test_exec_accept_revision(fake_word):
    code, out = _invoke(["--json", "exec", "--ops", '[{"op": "accept_revision", "index": 1}]'])
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True
    with wordlive.attach() as word:
        assert len(word.documents.active.revisions) == 0


def test_exec_accept_all_revisions(fake_word):
    code, out = _invoke(["--json", "exec", "--ops", '[{"op": "accept_all_revisions"}]'])
    assert code == EXIT_OK
    assert json.loads(out)["ops_run"] == 1


# --- helpers --------------------------------------------------------------------


def _empty_revisions():
    coll = MagicMock(name="Revisions")
    coll.Count = 0
    coll.__iter__ = MagicMock(side_effect=lambda: iter([]))
    return coll


def _revisions_with(n: int):
    """A range-scoped Revisions stub: Count + AcceptAll/RejectAll that clear it."""
    state = {"count": n}

    coll = MagicMock(name="RangeRevisions")
    type(coll).Count = property(lambda self: state["count"])

    def _clear() -> None:
        state["count"] = 0

    coll.AcceptAll = MagicMock(side_effect=_clear)
    coll.RejectAll = MagicMock(side_effect=_clear)
    return coll
