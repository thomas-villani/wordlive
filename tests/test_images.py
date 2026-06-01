"""insert_image — input resolution, wrap mapping, placement, and errors.

These exercise the method on the base `Anchor`, so a bookmark anchor stands in
for all kinds (one heading case guards against a `_range()` regression). The
fake `Range.InlineShapes.AddPicture` is a MagicMock, so the image file's
contents never matter — only that an existing path passes the readability
check.
"""

from __future__ import annotations

import base64
import os

import pytest

import wordlive
from wordlive.constants import WdWrapType
from wordlive.exceptions import ImageSourceError

# A 1x1 transparent PNG.
_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYPhfDwAChwGA"
    "60e6kgAAAABJRU5ErkJggg=="
)

# The "Address" bookmark in the fake document spans (13, 24).
_BM_START, _BM_END = 13, 24


@pytest.fixture
def png_file(tmp_path):
    p = tmp_path / "pic.png"
    p.write_bytes(_PNG)
    return p


def _insert_rng(fake_word, *, where: str = "after"):
    """The collapsed range insert_image targets for the Address bookmark."""
    pos = _BM_END if where == "after" else _BM_START
    return fake_word.ActiveDocument.Range(pos, pos)


# ---------------------------------------------------------------------------
# Inline vs. floating
# ---------------------------------------------------------------------------


def test_insert_image_inline_embeds_and_does_not_convert(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            doc.bookmarks["Address"].insert_image(str(png_file), wrap="inline")
    rng = _insert_rng(fake_word)
    rng.InlineShapes.AddPicture.assert_called_once()
    _, kwargs = rng.InlineShapes.AddPicture.call_args
    assert kwargs["LinkToFile"] is False
    assert kwargs["SaveWithDocument"] is True
    assert kwargs["FileName"] == str(png_file)
    # inline => never converted to a floating Shape
    assert rng.InlineShapes.shape.converted is None


def test_insert_image_square_converts_to_floating_shape(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            doc.bookmarks["Address"].insert_image(str(png_file), wrap="square")
    shape = _insert_rng(fake_word).InlineShapes.shape.converted
    assert shape is not None
    assert shape.WrapFormat.Type == int(WdWrapType.SQUARE)


def test_insert_image_top_bottom_sets_wrap_type(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(str(png_file), wrap="top-bottom")
    shape = _insert_rng(fake_word).InlineShapes.shape.converted
    assert shape.WrapFormat.Type == int(WdWrapType.TOP_BOTTOM)


# ---------------------------------------------------------------------------
# wrap="auto" half-text-width heuristic (Letter: usable 468pt, half 234pt)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "width,expected",
    [
        (117.0, WdWrapType.SQUARE),  # <= half => Square
        (421.0, WdWrapType.TOP_BOTTOM),  # >  half => TopBottom
        (None, WdWrapType.SQUARE),  # default fake width 100 <= half
    ],
)
def test_insert_image_auto_heuristic(fake_word, png_file, width, expected):
    extra = {} if width is None else {"width": width}
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(str(png_file), wrap="auto", **extra)
    shape = _insert_rng(fake_word).InlineShapes.shape.converted
    assert shape.WrapFormat.Type == int(expected)


# ---------------------------------------------------------------------------
# Sizing + alt text + lock aspect
# ---------------------------------------------------------------------------


def test_insert_image_sizing_alt_text_and_unlock_aspect(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(
            str(png_file),
            wrap="square",
            width=120.0,
            height=90.0,
            alt_text="A diagram",
            lock_aspect=False,
        )
    ish = _insert_rng(fake_word).InlineShapes.shape
    assert ish.Width == 120.0
    assert ish.Height == 90.0
    assert ish.LockAspectRatio == 0  # MsoTriState.FALSE
    assert ish.AlternativeText == "A diagram"
    # alt text is re-applied to the converted shape too
    assert ish.converted.AlternativeText == "A diagram"


def test_insert_image_locks_aspect_by_default(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(str(png_file), wrap="inline")
    assert _insert_rng(fake_word).InlineShapes.shape.LockAspectRatio == -1  # TRUE


# ---------------------------------------------------------------------------
# Placement (where=before/after)
# ---------------------------------------------------------------------------


def test_insert_image_before_targets_range_start(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(str(png_file), wrap="inline", where="before")
    _insert_rng(fake_word, where="before").InlineShapes.AddPicture.assert_called_once()
    _insert_rng(fake_word, where="after").InlineShapes.AddPicture.assert_not_called()


# ---------------------------------------------------------------------------
# Input shapes: path / bytes / base64 / data URL
# ---------------------------------------------------------------------------


def test_insert_image_path_str_used_directly(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(str(png_file), wrap="inline")
    _, kwargs = _insert_rng(fake_word).InlineShapes.AddPicture.call_args
    assert kwargs["FileName"] == str(png_file)
    assert os.path.exists(kwargs["FileName"])  # a real file, not a temp


def test_insert_image_relative_path_resolved_to_absolute(fake_word, png_file, monkeypatch):
    """A relative path must reach COM absolute: AddPicture resolves a relative
    name against Word's working directory, not ours, so a bare relative path
    silently fails with 0x80020009.
    """
    monkeypatch.chdir(png_file.parent)
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(png_file.name, wrap="inline")
    _, kwargs = _insert_rng(fake_word).InlineShapes.AddPicture.call_args
    assert os.path.isabs(kwargs["FileName"])
    assert os.path.samefile(kwargs["FileName"], str(png_file))


def test_insert_image_from_bytes_tempfiles_then_cleans_up(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(_PNG, wrap="inline")
    _, kwargs = _insert_rng(fake_word).InlineShapes.AddPicture.call_args
    tmp_used = kwargs["FileName"]
    assert tmp_used.endswith(".png")  # extension sniffed from magic bytes
    assert not os.path.exists(tmp_used)  # removed after embedding


def test_insert_image_from_base64(fake_word):
    b64 = base64.b64encode(_PNG).decode("ascii")
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(b64, wrap="inline")
    _, kwargs = _insert_rng(fake_word).InlineShapes.AddPicture.call_args
    assert not os.path.exists(kwargs["FileName"])


def test_insert_image_from_data_url(fake_word):
    data_url = "data:image/png;base64," + base64.b64encode(_PNG).decode("ascii")
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_image(data_url, wrap="inline")
    _insert_rng(fake_word).InlineShapes.AddPicture.assert_called_once()


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


def test_insert_image_missing_path_raises(fake_word, tmp_path):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ImageSourceError):
            doc.bookmarks["Address"].insert_image(tmp_path / "nope.png", wrap="inline")


def test_insert_image_invalid_base64_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ImageSourceError):
            doc.bookmarks["Address"].insert_image("not-base64-or-a-path!!", wrap="inline")


def test_insert_image_unrecognized_format_raises(fake_word):
    junk_b64 = base64.b64encode(b"this is not an image").decode("ascii")
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ImageSourceError):
            doc.bookmarks["Address"].insert_image(junk_b64, wrap="inline")


def test_insert_image_unrecognized_bytes_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ImageSourceError):
            doc.bookmarks["Address"].insert_image(b"definitely not an image", wrap="inline")


def test_insert_image_unknown_wrap_raises(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError):
            doc.bookmarks["Address"].insert_image(str(png_file), wrap="diagonal")


def test_insert_image_unknown_where_raises(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError):
            doc.bookmarks["Address"].insert_image(str(png_file), wrap="inline", where="sideways")


# ---------------------------------------------------------------------------
# Cross-kind: same method on a heading anchor (guards _range())
# ---------------------------------------------------------------------------


def test_insert_image_on_heading_anchor(fake_word, png_file):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("img"):
            doc.heading("Risks").insert_image(str(png_file), wrap="inline")
    para = next(
        p for p in fake_word.ActiveDocument.Paragraphs if p.Range.Text.rstrip("\r\n\x07") == "Risks"
    )
    end = para.Range.End
    fake_word.ActiveDocument.Range(end, end).InlineShapes.AddPicture.assert_called_once()
