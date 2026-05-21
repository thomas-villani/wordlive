"""Politeness invariants for EditScope: UndoRecord + Selection preservation."""

from __future__ import annotations

import pytest

import wordlive


def test_edit_opens_and_closes_undo_record(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("Update address"):
            pass

    fake_word.UndoRecord.StartCustomRecord.assert_called_once_with("Update address")
    fake_word.UndoRecord.EndCustomRecord.assert_called_once()


def test_edit_closes_undo_record_even_on_exception(fake_word):
    class Boom(Exception):
        pass

    with pytest.raises(Boom):
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("Doomed"):
                raise Boom()

    fake_word.UndoRecord.StartCustomRecord.assert_called_once_with("Doomed")
    fake_word.UndoRecord.EndCustomRecord.assert_called_once()


def test_edit_restores_selection_by_default(fake_word):
    fake_word.Selection.Start = 10
    fake_word.Selection.End = 15

    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("Move-the-cursor work"):
            # Simulate the inside-the-scope code moving the cursor.
            fake_word.Selection.Start = 100
            fake_word.Selection.End = 100

    # Restoration calls ActiveDocument.Range(10, 15).Select().
    restore_calls = [c for c in fake_word.ActiveDocument.Range.call_args_list if c.args == (10, 15)]
    assert restore_calls, "expected ActiveDocument.Range(10, 15) for Selection restoration"


def test_allow_cursor_move_suppresses_restoration(fake_word):
    fake_word.Selection.Start = 10
    fake_word.Selection.End = 15

    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("explicit move") as scope:
            scope.allow_cursor_move()

    # No Range(10, 15) call — we opted out of restoring.
    restore_calls = [c for c in fake_word.ActiveDocument.Range.call_args_list if c.args == (10, 15)]
    assert not restore_calls, "did not expect Selection restoration after allow_cursor_move()"


def test_edit_does_not_restore_after_exception(fake_word):
    fake_word.Selection.Start = 10
    fake_word.Selection.End = 15

    class Boom(Exception):
        pass

    with pytest.raises(Boom):
        with wordlive.attach() as word:
            doc = word.documents.active
            with doc.edit("doomed"):
                raise Boom()

    restore_calls = [c for c in fake_word.ActiveDocument.Range.call_args_list if c.args == (10, 15)]
    assert not restore_calls, "should not restore Selection after a failed edit"
