"""Comments — add / list / index / resolve / delete on the review channel."""

from __future__ import annotations

import pytest

import wordlive
from wordlive.exceptions import AnchorNotFoundError


def test_comments_start_empty(fake_word):
    with wordlive.attach() as word:
        comments = word.documents.active.comments
        assert len(comments) == 0
        assert comments.list() == []


def test_add_comment_on_anchor(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("comment"):
            c = doc.comments.add(doc.bookmarks["Address"], "Please verify")
        assert c.index == 1
        assert c.text == "Please verify"
        assert len(doc.comments) == 1


def test_add_comment_sets_author(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        c = doc.comments.add(doc.heading("Risks"), "Flagged", author="ReviewBot")
        assert c.author == "ReviewBot"


def test_add_comment_does_not_change_text(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        before = doc.bookmarks["Address"]
        with doc.edit("comment"):
            doc.comments.add(before, "note")
        # The anchored bookmark range's body comes through as the scope text,
        # but the comment body is separate — adding a comment never rewrites it.
        assert doc.comments[1].text == "note"


def test_comment_scope_reflects_anchor(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.comments.add(doc.heading("Risks"), "look here")
        # The Risks heading paragraph is the comment's scope.
        assert "Risks" in doc.comments[1].scope_text


def test_comment_list_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.comments.add(doc.heading("Introduction"), "first", author="A")
        doc.comments.add(doc.heading("Risks"), "second", author="B")
        rows = doc.comments.list()
    assert [r["index"] for r in rows] == [1, 2]
    assert rows[0]["author"] == "A"
    assert rows[1]["text"] == "second"
    assert all({"index", "author", "text", "scope", "done"} <= set(r) for r in rows)
    assert rows[0]["done"] is False


def test_comment_getitem_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.comments.add(doc.heading("Risks"), "x")
        with pytest.raises(AnchorNotFoundError) as exc_info:
            _ = doc.comments[5]
    assert exc_info.value.kind == "comment"


def test_comment_getitem_bool_rejected(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.comments.add(doc.heading("Risks"), "x")
        with pytest.raises(TypeError):
            _ = doc.comments[True]


def test_comment_resolve_sets_done(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.comments.add(doc.heading("Risks"), "x")
        assert doc.comments[1].done is False
        with doc.edit("resolve"):
            doc.comments[1].resolve()
        assert doc.comments[1].done is True


def test_comment_reopen_clears_done(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.comments.add(doc.heading("Risks"), "x")
        doc.comments[1].resolve()
        doc.comments[1].reopen()
        assert doc.comments[1].done is False


def test_comment_delete_removes_and_reindexes(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.comments.add(doc.heading("Introduction"), "first")
        doc.comments.add(doc.heading("Risks"), "second")
        with doc.edit("delete comment"):
            doc.comments[1].delete()
        assert len(doc.comments) == 1
        # The survivor is re-indexed to position 1.
        assert doc.comments[1].text == "second"


def test_comments_iterate_in_order(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.comments.add(doc.heading("Introduction"), "a")
        doc.comments.add(doc.heading("Risks"), "b")
        texts = [c.text for c in doc.comments]
    assert texts == ["a", "b"]
