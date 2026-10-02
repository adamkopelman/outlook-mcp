"""Pure-logic unit tests for outlook_mcp.outlook.client's module-level
helper functions — no COM/Outlook required. Mirrors outlook-mcp-rs's
com.rs pure-function tests (parse_categories/join_categories)."""

from outlook_mcp.outlook.client import _join_categories, _parse_categories


def test_categories_round_trip_and_trim():
    assert _parse_categories("Work, Receipts") == ["Work", "Receipts"]
    assert _parse_categories("  Work ,  Personal ") == ["Work", "Personal"]
    assert _parse_categories("") == []
    assert _join_categories(["Work", "Personal"]) == "Work, Personal"
    assert _join_categories([]) == ""


def test_truncate_reports_whether_it_cut():
    from outlook_mcp.outlook.client import _truncate

    assert _truncate("") == ("", False)
    assert _truncate("short", 10) == ("short", False)
    assert _truncate("x" * 10, 10) == ("x" * 10, False)
    text, truncated = _truncate("x" * 11, 10)
    assert truncated is True
    assert text.startswith("x" * 10 + "\n\n[... truncated at 10 characters]")


def test_clamp_body_limit():
    from outlook_mcp.outlook.client import (
        MAX_BODY_CHARS, MAX_BODY_CHARS_LIMIT, MIN_BODY_CHARS_LIMIT,
        _clamp_body_limit)

    assert _clamp_body_limit(None) == MAX_BODY_CHARS
    assert _clamp_body_limit(0) == MIN_BODY_CHARS_LIMIT
    assert _clamp_body_limit(-5) == MIN_BODY_CHARS_LIMIT
    assert _clamp_body_limit(250_000) == 250_000
    assert _clamp_body_limit(10**9) == MAX_BODY_CHARS_LIMIT
