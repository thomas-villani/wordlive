"""Track Changes — the persistent toggle and the self-restoring scope."""

from __future__ import annotations

import wordlive


def test_track_changes_reads_flag(fake_word):
    with wordlive.attach() as word:
        assert word.documents.active.track_changes is False


def test_track_changes_setter(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.track_changes = True
        assert doc.track_changes is True
        assert fake_word.ActiveDocument.TrackRevisions is True
        doc.track_changes = False
        assert doc.track_changes is False


def test_tracked_changes_scope_turns_on_then_restores(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.track_changes is False
        with doc.tracked_changes():
            assert fake_word.ActiveDocument.TrackRevisions is True
        # Restored to the prior (off) state on exit.
        assert doc.track_changes is False


def test_tracked_changes_scope_preserves_prior_on_state(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.track_changes = True
        with doc.tracked_changes():
            assert fake_word.ActiveDocument.TrackRevisions is True
        # The user already had it on — leave it on.
        assert doc.track_changes is True


def test_tracked_changes_restores_even_on_exception(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        try:
            with doc.tracked_changes():
                raise RuntimeError("boom")
        except RuntimeError:
            pass
        assert doc.track_changes is False


def test_tracked_changes_composes_with_edit(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.tracked_changes(), doc.edit("tracked edit"):
            doc.bookmarks["Address"].set_text("123 Main")
            assert fake_word.ActiveDocument.TrackRevisions is True
        assert doc.track_changes is False
        # The edit still rode a single UndoRecord.
        fake_word.UndoRecord.StartCustomRecord.assert_called_once()
