"""Image extraction — `Anchor.read_image`, the `image:N` anchor, `doc.images`,
the OPC parser, and the `images` / `read-image` CLI + MCP surfaces.

Round-trips against the `fake_word` MagicMock: `doc.InlineShapes` is a
`_FakeDocInlineShapes` whose one seeded picture's `Range.WordOpenXML` carries a
single base64 image part (a PNG anchored at offset 20, inside the body
paragraph, with alt text "logo").
"""

from __future__ import annotations

import base64
import json

import pytest

import wordlive
from wordlive import _images
from wordlive.cli.main import EXIT_OK, main
from wordlive.exceptions import AnchorNotFoundError, ImageSourceError, OpError

_SEED_BYTES = b"\x89PNG\r\n\x1a\nSEEDED"


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


# --- the OPC parser ------------------------------------------------------------


def test_image_parts_in_opc_picks_image_parts():
    xml = (
        '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
        '<pkg:part pkg:name="/word/document.xml" '
        'pkg:contentType="application/xml"><pkg:xmlData></pkg:xmlData></pkg:part>'
        '<pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">'
        f"<pkg:binaryData>{base64.b64encode(b'one').decode()}</pkg:binaryData></pkg:part>"
        '<pkg:part pkg:name="/word/media/image2.jpeg" pkg:contentType="image/jpeg">'
        f"<pkg:binaryData>{base64.b64encode(b'two').decode()}</pkg:binaryData></pkg:part>"
        "</pkg:package>"
    )
    parts = _images.image_parts_in_opc(xml)
    assert [ctype for ctype, _ in parts] == ["image/png", "image/jpeg"]


def test_read_image_from_range_single():
    class _Rng:
        WordOpenXML = (
            '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
            '<pkg:part pkg:name="/word/media/image1.png" pkg:contentType="image/png">'
            f"<pkg:binaryData>{base64.b64encode(_SEED_BYTES).decode()}</pkg:binaryData>"
            "</pkg:part></pkg:package>"
        )

    data, mime = _images.read_image_from_range(_Rng())
    assert data == _SEED_BYTES
    assert mime == "image/png"


def test_read_image_from_range_none_raises():
    class _Rng:
        WordOpenXML = (
            '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
            "</pkg:package>"
        )

    with pytest.raises(ImageSourceError, match="no embedded image"):
        _images.read_image_from_range(_Rng())


def test_read_image_from_range_multiple_raises():
    one, two = base64.b64encode(b"a").decode(), base64.b64encode(b"b").decode()

    class _Rng:
        WordOpenXML = (
            '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
            f'<pkg:part pkg:name="/m/1.png" pkg:contentType="image/png">'
            f"<pkg:binaryData>{one}</pkg:binaryData></pkg:part>"
            f'<pkg:part pkg:name="/m/2.png" pkg:contentType="image/png">'
            f"<pkg:binaryData>{two}</pkg:binaryData></pkg:part></pkg:package>"
        )

    with pytest.raises(ImageSourceError, match="2 images"):
        _images.read_image_from_range(_Rng())


# --- doc.images + image:N anchor -----------------------------------------------


def test_doc_images_list_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        rows = doc.images.list()
    assert rows == [
        {
            "index": 1,
            "anchor_id": "image:1",
            "mime": "image/png",
            "width": 100.0,
            "height": 80.0,
            "crop": None,
            "alt_text": "logo",
            "para": "para:2",  # image at offset 20 sits in the body paragraph
        }
    ]


def test_image_anchor_resolves_and_reads(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        anchor = doc.anchor_by_id("image:1")
        assert isinstance(anchor, wordlive.ImageAnchor)
        assert anchor.anchor_id == "image:1"
        assert anchor.alt_text == "logo"
        data, mime = anchor.read_image()
    assert data == _SEED_BYTES
    assert mime == "image/png"


def test_images_collection_index_and_len(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert len(doc.images) == 1
        assert doc.images[1].read_image()[0] == _SEED_BYTES


def test_image_anchor_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("image:9")
        with pytest.raises(AnchorNotFoundError):
            doc.images[9]


def test_image_anchor_bad_id_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("image:notanumber")


def test_read_image_on_non_image_anchor_raises(fake_word):
    # para:1 ("Introduction") has no embedded image — its range's WordOpenXML
    # carries no media part, so read_image reports a clean bad-input error.
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ImageSourceError, match="no embedded image"):
            doc.anchor_by_id("para:1").read_image()


def test_image_anchor_set_text_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError, match="no text"):
            doc.images[1].set_text("x")


# --- CLI -----------------------------------------------------------------------


def test_cli_images_list(fake_word):
    code, out = _invoke(["--json", "images"])
    assert code == EXIT_OK
    rows = json.loads(out)
    assert rows[0]["anchor_id"] == "image:1"
    assert rows[0]["mime"] == "image/png"
    assert rows[0]["para"] == "para:2"


def test_cli_read_image_base64(fake_word):
    code, out = _invoke(["--json", "read-image", "--anchor-id", "image:1"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["mime"] == "image/png"
    assert data["bytes"] == len(_SEED_BYTES)
    assert base64.b64decode(data["base64"]) == _SEED_BYTES


def test_cli_read_image_out_file(fake_word, tmp_path):
    dest = tmp_path / "extracted.png"
    code, out = _invoke(["--json", "read-image", "--anchor-id", "image:1", "--out", str(dest)])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["path"] == str(dest)
    assert "base64" not in data
    assert dest.read_bytes() == _SEED_BYTES


def test_cli_read_image_missing_image_exit_1(fake_word):
    # An anchor with no image is bad input (exit 1), not a missing anchor (2).
    code, _ = _invoke(["--json", "read-image", "--anchor-id", "para:1"])
    assert code == 1


def test_cli_read_image_unknown_anchor_exit_2(fake_word):
    code, _ = _invoke(["--json", "read-image", "--anchor-id", "image:9"])
    assert code == 2
