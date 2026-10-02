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


def test_text_matches_hebrew_and_mixed():
    from outlook_mcp.outlook.client import _text_matches

    assert _text_matches("מייל שיקוף", "Re: מייל שיקוף שבועי")
    assert _text_matches("סיכום עשייה", "", None, "גוף: סיכום עשייה Q3")
    assert _text_matches("Q3 סיכום", "Weekly Q3 סיכום עשייה")
    assert not _text_matches("מייל שיקוף", "Weekly report", "Dana")


def test_text_matches_is_caseless():
    from outlook_mcp.outlook.client import _text_matches

    assert _text_matches("weekly", "WEEKLY Report")
    assert _text_matches("STRASSE", "Hauptstraße 5")  # casefold, not lower


def test_text_matches_ignores_normalization_form():
    import unicodedata

    from outlook_mcp.outlook.client import _text_matches

    # Hebrew with niqqud and Latin with accents, composed vs decomposed.
    for word in ("שָׁלוֹם", "café"):
        nfc = unicodedata.normalize("NFC", word)
        nfd = unicodedata.normalize("NFD", word)
        assert _text_matches(nfd, f"subject {nfc}")
        assert _text_matches(nfc, f"subject {nfd}")
