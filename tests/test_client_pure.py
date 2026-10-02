"""Pure-logic unit tests for outlook_mcp.outlook.client's module-level
helper functions — no COM/Outlook required. Mirrors outlook-mcp-rs's
com.rs pure-function tests (parse_categories/join_categories)."""

import base64

import pytest

from outlook_mcp.errors import ToolError
from outlook_mcp.outlook.client import (
    _join_categories,
    _normalize_content_id,
    _parse_categories,
    _validate_inline_images,
)


def test_categories_round_trip_and_trim():
    assert _parse_categories("Work, Receipts") == ["Work", "Receipts"]
    assert _parse_categories("  Work ,  Personal ") == ["Work", "Personal"]
    assert _parse_categories("") == []
    assert _join_categories(["Work", "Personal"]) == "Work, Personal"
    assert _join_categories([]) == ""


# ---- inline image validation (issue #9) ------------------------------

PNG_BYTES = b"\x89PNG\r\n\x1a\n" + b"\x00" * 8
PNG_B64 = base64.b64encode(PNG_BYTES).decode()


def test_normalize_content_id_strips_brackets_and_cid_prefix():
    assert _normalize_content_id("logo") == "logo"
    assert _normalize_content_id("cid:logo") == "logo"
    assert _normalize_content_id(" <cid:logo> ") == "logo"
    assert _normalize_content_id("<logo@x>") == "logo@x"
    assert _normalize_content_id(None) == ""


def test_validate_base64_sniffs_png_and_names_file():
    [spec] = _validate_inline_images(
        [{"content_id": "cid:logo", "data_base64": PNG_B64}], html=True)
    assert spec["content_id"] == "logo"
    assert spec["data"] == PNG_BYTES
    assert spec["path"] is None
    assert spec["mime_type"] == "image/png"
    assert spec["filename"] == "logo.png"


def test_validate_accepts_data_uri_prefix_and_whitespace():
    wrapped = "data:image/png;base64," + PNG_B64[:8] + "\n  " + PNG_B64[8:]
    [spec] = _validate_inline_images(
        [{"content_id": "a", "data_base64": wrapped}], html=True)
    assert spec["data"] == PNG_BYTES


def test_validate_path_guesses_mime_from_extension(tmp_path):
    p = tmp_path / "chart.jpg"
    p.write_bytes(b"jpeg")
    [spec] = _validate_inline_images(
        [{"content_id": "chart", "path": str(p)}], html=True)
    assert spec["path"] == str(p)
    assert spec["filename"] == "chart.jpg"
    assert spec["mime_type"] == "image/jpeg"


def test_validate_explicit_filename_and_mime_win():
    [spec] = _validate_inline_images(
        [{"content_id": "x", "data_base64": PNG_B64,
          "filename": "a/b:c.bin", "mime_type": "image/x-custom"}], html=True)
    assert spec["mime_type"] == "image/x-custom"
    assert spec["filename"] == "a_b_c.bin"


def test_validate_unknown_data_falls_back_to_octet_stream():
    [spec] = _validate_inline_images(
        [{"content_id": "x", "data_base64": base64.b64encode(b"blob").decode()}],
        html=True)
    assert spec["mime_type"] == "application/octet-stream"


@pytest.mark.parametrize("images, html, message", [
    ([{"content_id": "a", "data_base64": PNG_B64}], False, "requires html=true"),
    ([{"content_id": " ", "data_base64": PNG_B64}], True, "missing a content_id"),
    ([{"content_id": "cid:", "data_base64": PNG_B64}], True, "missing a content_id"),
    ([{"content_id": "a"}], True, "exactly one of"),
    ([{"content_id": "a", "data_base64": PNG_B64, "path": "x.png"}], True,
     "exactly one of"),
    ([{"content_id": "a", "path": "/definitely/not/here.png"}], True,
     "inline image not found"),
    ([{"content_id": "a", "data_base64": "not base64!!"}], True, "not valid base64"),
    ([{"content_id": "a", "data_base64": "data:image/png,rawdata"}], True,
     "must be base64"),
    ([{"content_id": "a", "data_base64": PNG_B64},
      {"content_id": "<cid:A>", "data_base64": PNG_B64}], True, "duplicate"),
])
def test_validate_rejects_bad_specs(images, html, message):
    with pytest.raises(ToolError, match=message):
        _validate_inline_images(images, html=html)
