"""Pure-logic unit tests for outlook_mcp.outlook.client's module-level
helper functions — no COM/Outlook required. Mirrors outlook-mcp-rs's
com.rs pure-function tests (parse_categories/join_categories)."""

from outlook_mcp.outlook.client import (
    _clean_content_id,
    _guess_mime,
    _join_categories,
    _normalize_cid_request,
    _parse_categories,
)


def test_categories_round_trip_and_trim():
    assert _parse_categories("Work, Receipts") == ["Work", "Receipts"]
    assert _parse_categories("  Work ,  Personal ") == ["Work", "Personal"]
    assert _parse_categories("") == []
    assert _join_categories(["Work", "Personal"]) == "Work, Personal"
    assert _join_categories([]) == ""


def test_clean_content_id_strips_angle_brackets_and_blanks():
    assert _clean_content_id("<image001.png@01D9ABCD>") == "image001.png@01D9ABCD"
    assert _clean_content_id("  <abc@x>  ") == "abc@x"
    assert _clean_content_id("abc@x") == "abc@x"
    assert _clean_content_id("<>") is None
    assert _clean_content_id("") is None
    assert _clean_content_id(None) is None


def test_normalize_cid_request_strips_cid_prefix_and_brackets():
    assert _normalize_cid_request("image001.png@01D9") == "image001.png@01D9"
    assert _normalize_cid_request("cid:image001.png@01D9") == "image001.png@01D9"
    assert _normalize_cid_request("  CID:<abc@x>  ") == "abc@x"
    assert _normalize_cid_request("<abc@x>") == "abc@x"
    assert _normalize_cid_request("cid:") is None
    assert _normalize_cid_request("  ") is None
    assert _normalize_cid_request(None) is None


def test_guess_mime_prefers_tag_then_extension():
    assert _guess_mime("Image/PNG", "x.bin") == "image/png"
    assert _guess_mime(None, "report.pdf") == "application/pdf"
    assert _guess_mime("", "photo.JPG") == "image/jpeg"
    assert _guess_mime(None, "noextension") is None
    assert _guess_mime(None, "") is None
