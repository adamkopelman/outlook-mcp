"""WindowsOutlookClient tests against the fake COM objects in
tests/com_fakes.py — exercises the real client code without Outlook."""

import base64
import os

import pytest

from outlook_mcp import constants as c
from outlook_mcp.errors import ToolError
from tests import com_fakes

PNG_BYTES = b"\x89PNG\r\n\x1a\n" + b"\x00" * 8
PNG_B64 = base64.b64encode(PNG_BYTES).decode()


class DraftsApp(com_fakes.FakeApp):
    """FakeApp whose new items live in a Drafts folder, so the client can
    build an id (EntryID|StoreID) after Save()."""

    def __init__(self, ns):
        super().__init__(ns)
        self.drafts = com_fakes.FakeFolder("Drafts")

    def CreateItem(self, item_type):
        return self.drafts.add(super().CreateItem(item_type))


@pytest.fixture
def app():
    return DraftsApp(com_fakes.FakeNamespace())


@pytest.fixture
def client(monkeypatch, app):
    return com_fakes.install(monkeypatch, app.ns, app=app)


def _created_mail(app):
    assert len(app.created) == 1
    item_type, mail = app.created[0]
    assert item_type == c.OL_MAIL_ITEM
    return mail


# ---- inline images (issue #9) ----------------------------------------

def test_create_draft_attaches_base64_inline_image(client, app):
    client.create_draft(to=["a@x.com"], subject="s", html=True,
                        body='<img src="cid:logo">',
                        inline_images=[{"content_id": "<cid:logo>",
                                        "data_base64": "data:image/png;base64," + PNG_B64}])
    mail = _created_mail(app)
    assert mail.saved == 1
    assert mail.BodyFormat == c.OL_FORMAT_HTML
    [(source, att_type, position, display)] = mail.Attachments.add_calls
    assert (att_type, position, display) == (c.OL_BY_VALUE, 0, "logo.png")
    att = mail.Attachments.Item(1)
    assert att.data == PNG_BYTES
    assert att.PropertyAccessor.set_calls == [
        (c.PR_ATTACH_CONTENT_ID, "logo"),
        (c.PR_ATTACH_MIME_TAG, "image/png"),
        (c.PR_ATTACHMENT_HIDDEN, True),
    ]
    # The temp copy of the base64 data is gone once the draft is saved.
    assert not os.path.exists(source)
    assert not os.path.exists(os.path.dirname(os.path.dirname(source)))


def test_create_draft_attaches_path_inline_image_without_temp_copy(client, app, tmp_path):
    img = tmp_path / "chart.png"
    img.write_bytes(PNG_BYTES)
    client.create_draft(to=["a@x.com"], subject="s", body="<p/>", html=True,
                        inline_images=[{"content_id": "chart", "path": str(img),
                                        "filename": "Chart Q3.png"}])
    mail = _created_mail(app)
    [(source, att_type, position, display)] = mail.Attachments.add_calls
    assert source == str(img)
    assert display == "Chart Q3.png"
    assert img.exists()  # caller's file is never cleaned up
    assert (c.PR_ATTACH_CONTENT_ID, "chart") in \
        mail.Attachments.Item(1).PropertyAccessor.set_calls


def test_create_draft_inline_images_alongside_regular_attachments(client, app, tmp_path):
    doc = tmp_path / "report.pdf"
    doc.write_bytes(b"%PDF")
    client.create_draft(to=["a@x.com"], subject="s", body="<p/>", html=True,
                        attachments=[str(doc)],
                        inline_images=[{"content_id": "a", "data_base64": PNG_B64},
                                       {"content_id": "b", "data_base64": PNG_B64,
                                        "filename": "a.png"}])
    mail = _created_mail(app)
    assert mail.Attachments.Count == 3
    regular = mail.Attachments.Item(1)
    assert regular.PropertyAccessor.set_calls == []
    # Equal filenames (b explicitly reuses a's) don't collide on disk.
    assert mail.Attachments.Item(2).data == PNG_BYTES
    assert mail.Attachments.Item(3).data == PNG_BYTES


def test_send_email_attaches_inline_image_then_sends(client, app):
    client.send_email(to=["a@x.com"], subject="s", body="<p/>", html=True,
                      inline_images=[{"content_id": "logo", "data_base64": PNG_B64}])
    mail = _created_mail(app)
    assert mail.sent is True
    [(source, *_)] = mail.Attachments.add_calls
    assert not os.path.exists(source)


def test_send_email_cleans_temp_dir_when_send_fails(client, app, monkeypatch):
    def boom(self):
        raise com_fakes.com_error("send failed")

    monkeypatch.setattr(com_fakes.FakeMail, "Send", boom)
    with pytest.raises(ToolError, match="send failed"):
        client.send_email(to=["a@x.com"], subject="s", body="<p/>", html=True,
                          inline_images=[{"content_id": "logo", "data_base64": PNG_B64}])
    [(source, *_)] = _created_mail(app).Attachments.add_calls
    assert not os.path.exists(source)


def test_inline_images_require_html_and_fail_before_creating_item(client, app):
    with pytest.raises(ToolError, match="requires html=true"):
        client.create_draft(to=["a@x.com"], subject="s", body="plain",
                            inline_images=[{"content_id": "logo",
                                            "data_base64": PNG_B64}])
    assert app.created == []


def test_bad_inline_image_fails_before_sending(client, app):
    with pytest.raises(ToolError, match="not valid base64"):
        client.send_email(to=["a@x.com"], subject="s", body="<p/>", html=True,
                          inline_images=[{"content_id": "ok", "data_base64": PNG_B64},
                                         {"content_id": "bad", "data_base64": "%%%"}])
    assert app.created == []


def test_create_draft_without_inline_images_unchanged(client, app):
    result = client.create_draft(to=["a@x.com"], subject="s", body="hi")
    mail = _created_mail(app)
    assert result["status"] == "draft_saved"
    assert mail.Attachments.Count == 0
    assert mail.BodyFormat == c.OL_FORMAT_PLAIN
