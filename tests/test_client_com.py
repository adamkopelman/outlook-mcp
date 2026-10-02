"""WindowsOutlookClient tests against the fake COM objects in
tests/com_fakes.py (runs on any platform)."""

import pytest

from outlook_mcp import constants as c
from outlook_mcp.errors import ToolError
from tests import com_fakes
from tests.com_fakes import FakeFolder, FakeMail, FakeNamespace, FakeStore


class _BrokenStore(FakeStore):
    """A store whose default folders can't be resolved (forces fallback)."""

    def GetDefaultFolder(self, folder_id):
        raise com_fakes.com_error("Store does not support this folder.")


def _mailbox():
    """Inbox + Deleted Items sharing one store, wired like Outlook."""
    store = FakeStore()
    deleted = FakeFolder("Deleted Items", store=store)
    inbox = FakeFolder("Inbox", store=store)
    store.default_folders[c.OL_FOLDER_DELETED_ITEMS] = deleted
    ns = FakeNamespace({c.OL_FOLDER_DELETED_ITEMS: deleted,
                        c.OL_FOLDER_INBOX: inbox})
    return ns, inbox, deleted


# ---- delete_email ---------------------------------------------------

def test_delete_email_soft_moves_to_deleted_items(monkeypatch):
    ns, inbox, deleted = _mailbox()
    mail = inbox.add(FakeMail(Subject="hello"))
    item_id = ns.register(mail)
    client = com_fakes.install(monkeypatch, ns)

    result = client.delete_email(item_id)

    assert result == {"status": "deleted", "subject": "hello",
                      "permanent": False, "note": "Moved to Deleted Items."}
    assert mail.calls == ["Delete"]
    assert mail.deleted is False
    assert mail in list(deleted.Items)


def test_delete_email_permanent_moves_then_hard_deletes(monkeypatch):
    ns, inbox, deleted = _mailbox()
    mail = inbox.add(FakeMail(Subject="draft"))
    item_id = ns.register(mail)
    client = com_fakes.install(monkeypatch, ns)

    result = client.delete_email(item_id, permanent=True)

    assert result["status"] == "deleted"
    assert result["subject"] == "draft"
    assert result["permanent"] is True
    assert "Permanently" in result["note"]
    assert mail.calls == [("Move", "Deleted Items"), "Delete"]
    assert mail.deleted is True
    assert deleted.Items.Count == 0  # not left behind in Deleted Items
    assert inbox.Items.Count == 0


def test_delete_email_permanent_already_in_deleted_items(monkeypatch):
    ns, _inbox, deleted = _mailbox()
    mail = deleted.add(FakeMail(Subject="old"))
    item_id = ns.register(mail)
    client = com_fakes.install(monkeypatch, ns)

    result = client.delete_email(item_id, permanent=True)

    assert result["permanent"] is True
    assert mail.calls == ["Delete"]  # no Move: a single Delete is permanent
    assert mail.deleted is True
    assert deleted.Items.Count == 0


def test_delete_email_permanent_falls_back_to_default_store(monkeypatch):
    deleted = FakeFolder("Deleted Items")
    inbox = FakeFolder("Inbox", store=_BrokenStore())
    ns = FakeNamespace({c.OL_FOLDER_DELETED_ITEMS: deleted})
    mail = inbox.add(FakeMail(Subject="x"))
    item_id = ns.register(mail)
    client = com_fakes.install(monkeypatch, ns)

    result = client.delete_email(item_id, permanent=True)

    assert result["permanent"] is True
    assert mail.calls == [("Move", "Deleted Items"), "Delete"]
    assert mail.deleted is True
    assert deleted.Items.Count == 0


def test_delete_email_com_error_becomes_tool_error(monkeypatch):
    ns, inbox, _deleted = _mailbox()
    mail = inbox.add(FakeMail(Subject="stuck"))
    item_id = ns.register(mail)

    def boom(_folder):
        raise com_fakes.com_error("Move refused")

    mail.Move = boom
    client = com_fakes.install(monkeypatch, ns)

    with pytest.raises(ToolError, match="Move refused"):
        client.delete_email(item_id, permanent=True)


# ---- empty_deleted_items --------------------------------------------

def test_empty_deleted_items_requires_confirm(monkeypatch):
    ns, _inbox, deleted = _mailbox()
    mail = deleted.add(FakeMail())
    client = com_fakes.install(monkeypatch, ns)

    with pytest.raises(ToolError, match="confirm=true"):
        client.empty_deleted_items()
    assert mail.deleted is False
    assert deleted.Items.Count == 1


def test_empty_deleted_items_deletes_items_and_subfolders(monkeypatch):
    ns, _inbox, deleted = _mailbox()
    mails = [deleted.add(FakeMail(Subject=f"m{i}")) for i in range(3)]
    subs = [FakeFolder(f"sub{i}", parent=deleted) for i in range(2)]
    for sub in subs:
        deleted.Folders.append(sub)
    client = com_fakes.install(monkeypatch, ns)

    result = client.empty_deleted_items(confirm=True)

    assert result == {"status": "emptied", "items_deleted": 3,
                      "folders_deleted": 2, "failed": 0}
    assert all(m.deleted for m in mails)  # none skipped by re-indexing
    assert all(s.deleted for s in subs)
    assert deleted.Items.Count == 0
    assert deleted.Folders.Count == 0


def test_empty_deleted_items_counts_failures_and_continues(monkeypatch):
    ns, _inbox, deleted = _mailbox()
    good1 = deleted.add(FakeMail())
    stuck = deleted.add(FakeMail())
    good2 = deleted.add(FakeMail())

    def refuse():
        raise com_fakes.com_error("Item is locked")

    stuck.Delete = refuse
    client = com_fakes.install(monkeypatch, ns)

    result = client.empty_deleted_items(confirm=True)

    assert result == {"status": "emptied", "items_deleted": 2,
                      "folders_deleted": 0, "failed": 1}
    assert good1.deleted and good2.deleted
    assert list(deleted.Items) == [stuck]
