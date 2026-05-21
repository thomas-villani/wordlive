"""Cursor surface — Selection.info() / Selection.write()."""

from __future__ import annotations

import wordlive


def test_info_reports_collapsed_cursor(fake_word):
    fake_word.Selection.Start = 7
    fake_word.Selection.End = 7
    fake_word.Selection.Text = ""
    with wordlive.attach() as word:
        info = word.documents.active.selection.info()
    assert info == {"start": 7, "end": 7, "collapsed": True, "text": ""}


def test_info_reports_spanning_selection(fake_word):
    fake_word.Selection.Start = 0
    fake_word.Selection.End = 12
    fake_word.Selection.Text = "Introduction"
    with wordlive.attach() as word:
        info = word.documents.active.selection.info()
    assert info["collapsed"] is False
    assert info["text"] == "Introduction"


def test_write_inserts_at_collapsed_cursor(fake_word):
    fake_word.Selection.Start = 0
    fake_word.Selection.End = 0
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.selection.write("Hi")
        assert doc.com.Range(0, 0).Text == "Hi"


def test_write_replace_overwrites_selection(fake_word):
    fake_word.Selection.Start = 0
    fake_word.Selection.End = 12
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.selection.write("New", replace=True)
        assert doc.com.Range(0, 12).Text == "New"


def test_write_no_replace_inserts_at_start(fake_word):
    fake_word.Selection.Start = 5
    fake_word.Selection.End = 12
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.selection.write("X", replace=False)
        assert doc.com.Range(5, 5).Text == "X"


def test_write_moves_cursor_after_inserted_text(fake_word):
    fake_word.Selection.Start = 0
    fake_word.Selection.End = 0
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.selection.write("Hello")
        # Cursor collapses to start + len("Hello") == 5.
        doc.com.Range(5, 5).Select.assert_called()
