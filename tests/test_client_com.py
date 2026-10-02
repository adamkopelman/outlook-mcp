"""WindowsOutlookClient tests against the fake COM objects in
tests/com_fakes.py (no Outlook required)."""

import pytest

from outlook_mcp import constants as c
from outlook_mcp.errors import ToolError
from tests import com_fakes


def _draft_setup(monkeypatch, **attrs):
    """A fake namespace holding one mail item in Drafts; returns
    (client, mail, id)."""
    ns = com_fakes.FakeNamespace()
    drafts = com_fakes.FakeFolder("Drafts")
    mail = drafts.add(com_fakes.FakeMail(**attrs))
    item_id = ns.register(mail)
    client = com_fakes.install(monkeypatch, ns)
    return client, mail, item_id


# ---- update_draft ---------------------------------------------------

def test_update_draft_sets_fields_and_saves_once(monkeypatch):
    client, mail, item_id = _draft_setup(
        monkeypatch, Subject="Old", Body="old body", To="old@x.com",
        CC="keep@x.com")
    result = client.update_draft(item_id, subject="New", body="new body",
                                 to=["a@x.com", "b@x.com"], bcc=["z@x.com"])
    assert result == {"status": "draft_updated", "id": item_id,
                      "changed": ["subject", "body", "to", "bcc"]}
    assert mail.Subject == "New"
    assert mail.Body == "new body"
    assert mail.BodyFormat == c.OL_FORMAT_PLAIN
    assert mail.To == "a@x.com; b@x.com"
    assert mail.CC == "keep@x.com"  # untouched field is left alone
    assert mail.BCC == "z@x.com"
    assert mail.saved == 1
    assert mail.sent is False


def test_update_draft_html_body_switches_format(monkeypatch):
    client, mail, item_id = _draft_setup(monkeypatch, BodyFormat=c.OL_FORMAT_PLAIN)
    result = client.update_draft(item_id, html_body="<p>hi</p>")
    assert result["changed"] == ["html_body"]
    assert mail.BodyFormat == c.OL_FORMAT_HTML
    assert mail.HTMLBody == "<p>hi</p>"


def test_update_draft_body_switches_html_draft_to_plain(monkeypatch):
    client, mail, item_id = _draft_setup(monkeypatch, BodyFormat=c.OL_FORMAT_HTML)
    client.update_draft(item_id, body="plain now")
    assert mail.BodyFormat == c.OL_FORMAT_PLAIN
    assert mail.Body == "plain now"


def test_update_draft_empty_list_clears_recipients(monkeypatch):
    client, mail, item_id = _draft_setup(monkeypatch, CC="a@x.com; b@x.com")
    result = client.update_draft(item_id, cc=[])
    assert result["changed"] == ["cc"]
    assert mail.CC == ""
    assert mail.saved == 1


def test_update_draft_rejects_sent_item(monkeypatch):
    client, mail, item_id = _draft_setup(monkeypatch, Sent=True, Subject="Old")
    with pytest.raises(ToolError, match="Only unsent drafts"):
        client.update_draft(item_id, subject="New")
    assert mail.Subject == "Old"
    assert mail.saved == 0


def test_update_draft_rejects_body_and_html_body(monkeypatch):
    client, mail, item_id = _draft_setup(monkeypatch)
    with pytest.raises(ToolError, match="not both"):
        client.update_draft(item_id, body="a", html_body="<p>b</p>")
    assert mail.saved == 0


def test_update_draft_rejects_empty_update(monkeypatch):
    client, mail, item_id = _draft_setup(monkeypatch)
    with pytest.raises(ToolError, match="at least one"):
        client.update_draft(item_id)
    with pytest.raises(ToolError, match="at least one"):
        client.update_draft(item_id, attachments=[])
    assert mail.saved == 0


def test_update_draft_appends_attachments(monkeypatch, tmp_path):
    path = tmp_path / "notes.txt"
    path.write_text("hello")
    existing = com_fakes.FakeAttachment("old.pdf", b"x")
    client, mail, item_id = _draft_setup(monkeypatch, attachments=[existing])
    result = client.update_draft(item_id, attachments=[str(path)])
    assert result["changed"] == ["attachments"]
    assert [a.FileName for a in mail.Attachments] == ["old.pdf", "notes.txt"]
    assert mail.saved == 1


def test_update_draft_missing_attachment_fails_before_changes(monkeypatch, tmp_path):
    client, mail, item_id = _draft_setup(monkeypatch, Subject="Old")
    with pytest.raises(ToolError, match="attachment not found"):
        client.update_draft(item_id, subject="New",
                            attachments=[str(tmp_path / "nope.pdf")])
    assert mail.Subject == "Old"
    assert mail.Attachments.Count == 0
    assert mail.saved == 0


def test_update_draft_bad_id(monkeypatch):
    client, _, _ = _draft_setup(monkeypatch)
    with pytest.raises(ToolError):
        client.update_draft("missing|store-1", subject="x")
