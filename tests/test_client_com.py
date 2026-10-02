"""WindowsOutlookClient against the fake COM objects in tests/com_fakes."""

from tests import com_fakes
from tests.com_fakes import FakeFolder, FakeMail, FakeNamespace

from outlook_mcp.outlook import client as client_mod


def _inbox_client(monkeypatch, mails):
    inbox = FakeFolder("Inbox", items=mails)
    return com_fakes.install(monkeypatch, FakeNamespace({6: inbox}))


def _subjects(results):
    return [r["subject"] for r in results]


# ---- list_emails pagination ----------------------------------------

def test_list_emails_pages_with_offset(monkeypatch):
    mails = [FakeMail(Subject=f"m{i}") for i in range(7)]
    client = _inbox_client(monkeypatch, mails)

    pages = [client.list_emails(count=3, offset=off) for off in (0, 3, 6)]

    assert [len(p) for p in pages] == [3, 3, 1]
    ids = [r["id"] for p in pages for r in p]
    assert len(ids) == len(set(ids)) == 7
    assert _subjects(pages[0] + pages[1] + pages[2]) == [f"m{i}" for i in range(7)]


def test_list_emails_offset_beyond_end_is_empty(monkeypatch):
    client = _inbox_client(monkeypatch, [FakeMail(Subject="only")])
    assert client.list_emails(count=10, offset=5) == []


def test_list_emails_negative_offset_clamps_to_zero(monkeypatch):
    client = _inbox_client(monkeypatch, [FakeMail(Subject=f"m{i}") for i in range(3)])
    assert _subjects(client.list_emails(count=2, offset=-4)) == ["m0", "m1"]


def test_list_emails_count_is_capped(monkeypatch):
    assert client_mod.MAX_EMAIL_COUNT == 200
    client = _inbox_client(monkeypatch, [FakeMail() for _ in range(250)])
    assert len(client.list_emails(count=1000)) == 200


def test_list_emails_offset_counts_only_category_matches(monkeypatch):
    mails = []
    for i in range(6):
        mails.append(FakeMail(Subject=f"work{i}", Categories="Work"))
        mails.append(FakeMail(Subject=f"other{i}", Categories="Personal"))
    client = _inbox_client(monkeypatch, mails)

    page1 = client.list_emails(count=2, category="work")
    page2 = client.list_emails(count=2, offset=2, category="work")

    assert _subjects(page1) == ["work0", "work1"]
    assert _subjects(page2) == ["work2", "work3"]
