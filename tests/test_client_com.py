"""WindowsOutlookClient tests against the fake COM objects in
tests/com_fakes.py (no Outlook required)."""

from outlook_mcp.outlook import client as client_mod
from tests import com_fakes
from tests.com_fakes import FakeFolder, FakeItems, FakeMail, FakeNamespace

INBOX = 6


def _is_text_query(flt):
    return "urn:schemas:httpmail:subject" in flt


def _dasl_finds_nothing(flt, items):
    """Model issue #4: the DASL LIKE text query returns no rows; every other
    Restrict keeps the items unchanged."""
    return [] if _is_text_query(flt) else None


def _dasl_ascii_only(flt, items):
    """Model a DASL that works for ASCII subjects only."""
    if not _is_text_query(flt):
        return None
    return [i for i in items if "Weekly" in i.Subject]


def _inbox(mails, hook):
    items = FakeItems(mails, restrict_hook=hook)
    inbox = FakeFolder("Inbox", items=items)
    return inbox, FakeNamespace({INBOX: inbox})


def test_hebrew_query_falls_back_to_client_side_scan(monkeypatch):
    hit = FakeMail(Subject="מייל שיקוף שבועי", SenderName="Dana")
    body_hit = FakeMail(Subject="Re: update", Body="מצורף סיכום עשייה לרבעון")
    miss = FakeMail(Subject="Weekly report", SenderName="Avi")
    inbox, ns = _inbox([hit, body_hit, miss], _dasl_finds_nothing)
    client = com_fakes.install(monkeypatch, ns)

    results = client.list_emails(query="מייל שיקוף")
    assert [r["subject"] for r in results] == ["מייל שיקוף שבועי"]

    results = client.list_emails(query="סיכום עשייה")
    assert [r["subject"] for r in results] == ["Re: update"]

    # The DASL query was still attempted first.
    assert any(op == "Restrict" and _is_text_query(arg)
               for op, arg, *_ in inbox.Items.log)


def test_fallback_matches_decomposed_query_against_composed_subject(monkeypatch):
    import unicodedata

    word = unicodedata.normalize("NFC", "שָׁלוֹם")
    hit = FakeMail(Subject=f"{word} לכולם")
    _, ns = _inbox([hit], _dasl_finds_nothing)
    client = com_fakes.install(monkeypatch, ns)

    results = client.list_emails(query=unicodedata.normalize("NFD", word))
    assert len(results) == 1


def test_text_query_is_applied_after_other_restricts(monkeypatch):
    mail = FakeMail(Subject="מייל שיקוף", UnRead=True)
    inbox, ns = _inbox([mail], _dasl_finds_nothing)
    client = com_fakes.install(monkeypatch, ns)

    client.list_emails(query="מייל שיקוף", sender="Dana", unread_only=True)
    restricts = [arg for op, arg, *_ in inbox.Items.log if op == "Restrict"]
    assert _is_text_query(restricts[-1])
    assert "[UnRead] = True" in restricts


def test_fallback_still_applies_client_side_filters_and_count(monkeypatch):
    mails = [FakeMail(Subject=f"מייל שיקוף {n}", Categories="Work" if n % 2 else "")
             for n in range(6)]
    _, ns = _inbox(mails, _dasl_finds_nothing)
    client = com_fakes.install(monkeypatch, ns)

    results = client.list_emails(query="מייל שיקוף", category="work", count=2)
    assert [r["subject"] for r in results] == ["מייל שיקוף 1", "מייל שיקוף 3"]


def test_fallback_scan_is_capped(monkeypatch):
    monkeypatch.setattr(client_mod, "MAX_SCAN_ITEMS", 3)
    mails = [FakeMail(Subject="other") for _ in range(3)]
    mails.append(FakeMail(Subject="מייל שיקוף"))
    _, ns = _inbox(mails, _dasl_finds_nothing)
    client = com_fakes.install(monkeypatch, ns)

    assert client.list_emails(query="מייל שיקוף") == []


def test_ascii_query_uses_dasl_only(monkeypatch):
    hit = FakeMail(Subject="Weekly report")
    other = FakeMail(Subject="weekly lowercase, DASL says no")
    _, ns = _inbox([hit, other], _dasl_ascii_only)
    client = com_fakes.install(monkeypatch, ns)

    results = client.list_emails(query="Weekly")
    assert [r["subject"] for r in results] == ["Weekly report"]

    # An empty ASCII DASL result is trusted: no client-side scan.
    _, ns = _inbox([other], _dasl_finds_nothing)
    client = com_fakes.install(monkeypatch, ns)
    assert client.list_emails(query="weekly") == []


def test_hebrew_query_uses_dasl_results_when_dasl_finds_them(monkeypatch):
    hit = FakeMail(Subject="מייל שיקוף")
    also = FakeMail(Subject="מייל שיקוף 2")

    def hook(flt, items):
        return [hit] if _is_text_query(flt) else None

    _, ns = _inbox([hit, also], hook)
    client = com_fakes.install(monkeypatch, ns)
    results = client.list_emails(query="מייל שיקוף")
    assert [r["subject"] for r in results] == ["מייל שיקוף"]


def test_query_reaches_restrict_nfc_normalized_and_escaped(monkeypatch):
    import unicodedata

    inbox, ns = _inbox([], None)
    client = com_fakes.install(monkeypatch, ns)
    client.list_emails(query=unicodedata.normalize("NFD", "café d'été"))
    flt = next(arg for op, arg, *_ in inbox.Items.log
               if op == "Restrict" and _is_text_query(arg))
    assert "%café d''été%" in flt
    assert unicodedata.is_normalized("NFC", flt)
