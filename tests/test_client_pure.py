"""Pure-logic unit tests for outlook_mcp.outlook.client's module-level
helper functions — no COM/Outlook required. Mirrors outlook-mcp-rs's
com.rs pure-function tests (parse_categories/join_categories)."""

from outlook_mcp.outlook.client import (
    _clean_content_id,
    _guess_mime,
    _is_inline,
    _join_categories,
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


def test_guess_mime_prefers_tag_then_extension():
    assert _guess_mime("Image/PNG", "x.bin") == "image/png"
    assert _guess_mime(None, "report.pdf") == "application/pdf"
    assert _guess_mime("", "photo.JPG") == "image/jpeg"
    assert _guess_mime(None, "noextension") is None
    assert _guess_mime(None, "") is None


def test_is_inline_needs_a_content_id():
    html = '<img src="cid:img1@x">'
    assert _is_inline(None, False, html) is False
    assert _is_inline(None, True, html) is False
    assert _is_inline("", True, html) is False


def test_is_inline_when_referenced_from_html():
    assert _is_inline("img1@x", False, '<img src="cid:img1@x">') is True
    assert _is_inline("img1@x", False, "<img src='cid:img1@x'>") is True
    assert _is_inline("img1@x", False, '<IMG SRC="CID:IMG1@X">') is True
    assert _is_inline("img1@x", False, "background: url(cid:img1@x);") is True
    assert _is_inline("img1@x", False, "cid:img1@x") is True


def test_is_inline_cid_not_referenced_is_regular():
    assert _is_inline("img1@x", False, "") is False
    assert _is_inline("img1@x", False, None) is False
    assert _is_inline("img1@x", False, '<img src="cid:other@x">') is False
    # Whole-id match only: a longer id sharing the prefix doesn't count.
    assert _is_inline("img1", False, '<img src="cid:img10@x">') is False
    # Regex metacharacters in the id are literal.
    assert _is_inline("a.b", False, '<img src="cid:aXb">') is False


def test_is_inline_hidden_with_content_id():
    assert _is_inline("img1@x", True, "") is True
