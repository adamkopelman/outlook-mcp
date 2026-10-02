"""WindowsOutlookClient tests against the fake COM objects in
tests/com_fakes.py — exercises the real client code paths without
Outlook."""

from outlook_mcp import constants as c
from tests import com_fakes
from tests.com_fakes import FakeAttachment, FakeFolder, FakeMail, FakeNamespace


def _client_with_mail(monkeypatch, attachments, html_body=""):
    ns = FakeNamespace()
    mail = FakeMail(attachments=attachments, HTMLBody=html_body)
    FakeFolder("Inbox", items=[mail])
    email_id = ns.register(mail)
    return com_fakes.install(monkeypatch, ns), email_id


def _inline_png():
    return FakeAttachment("image001.png", b"\x89PNG....", props={
        c.PR_ATTACH_CONTENT_ID: "<image001.png@01D9ABCD>",
        c.PR_ATTACH_MIME_TAG: "image/png",
        c.PR_ATTACHMENT_HIDDEN: True,
    })


# ---- list_attachments -----------------------------------------------

def test_list_attachments_reads_mapi_metadata(monkeypatch):
    client, email_id = _client_with_mail(monkeypatch, [_inline_png()])
    [info] = client.list_attachments(email_id)
    assert info == {
        "index": 1,
        "filename": "image001.png",
        "size": 8,
        "type": "file",
        "content_id": "image001.png@01D9ABCD",
        "mime_type": "image/png",
        "hidden": True,
        "is_inline": True,
    }


def test_list_attachments_falls_back_when_props_missing(monkeypatch):
    plain = FakeAttachment("report.pdf", b"%PDF-1.4")
    linked = FakeAttachment("notes.unknownext", b"", att_type=c.OL_BY_REFERENCE)
    client, email_id = _client_with_mail(monkeypatch, [plain, linked])
    pdf, link = client.list_attachments(email_id)
    assert pdf["index"] == 1
    assert pdf["content_id"] is None
    assert pdf["mime_type"] == "application/pdf"
    assert pdf["hidden"] is False
    assert pdf["type"] == "file"
    assert link["index"] == 2
    assert link["type"] == "link"
    assert link["mime_type"] is None


def test_list_attachments_empty_content_id_is_none(monkeypatch):
    att = FakeAttachment("a.txt", b"x", props={c.PR_ATTACH_CONTENT_ID: ""})
    client, email_id = _client_with_mail(monkeypatch, [att])
    assert client.list_attachments(email_id)[0]["content_id"] is None


def _cid_png(cid="<img1@x>"):
    return FakeAttachment("a.png", b"png", props={c.PR_ATTACH_CONTENT_ID: cid})


def test_list_attachments_flags_cid_referenced_as_inline(monkeypatch):
    plain = FakeAttachment("report.pdf", b"%PDF-1.4")
    client, email_id = _client_with_mail(
        monkeypatch, [_cid_png(), plain], html_body='<img src="cid:img1@x">')
    png, pdf = client.list_attachments(email_id)
    assert png["is_inline"] is True
    assert pdf["is_inline"] is False


def test_list_attachments_unreferenced_cid_is_not_inline(monkeypatch):
    client, email_id = _client_with_mail(
        monkeypatch, [_cid_png()], html_body="<p>no images here</p>")
    [png] = client.list_attachments(email_id)
    assert png["content_id"] == "img1@x"
    assert png["is_inline"] is False


def test_list_attachments_hidden_cid_is_inline_without_html(monkeypatch):
    client, email_id = _client_with_mail(monkeypatch, [_inline_png()])
    assert client.list_attachments(email_id)[0]["is_inline"] is True


# ---- get_email ------------------------------------------------------

def test_get_email_lists_inline_images(monkeypatch):
    plain = FakeAttachment("report.pdf", b"%PDF-1.4")
    client, email_id = _client_with_mail(
        monkeypatch, [plain, _cid_png()], html_body='<img src="cid:IMG1@x">')
    detail = client.get_email(email_id)
    assert detail["attachments"] == ["report.pdf", "a.png"]
    assert detail["inline_images"] == [{
        "index": 2,
        "content_id": "img1@x",
        "filename": "a.png",
        "mime_type": "image/png",
    }]


def test_get_email_without_inline_images(monkeypatch):
    plain = FakeAttachment("report.pdf", b"%PDF-1.4")
    client, email_id = _client_with_mail(monkeypatch, [plain])
    detail = client.get_email(email_id)
    assert detail["attachments"] == ["report.pdf"]
    assert detail["inline_images"] == []

    client, email_id = _client_with_mail(monkeypatch, [])
    detail = client.get_email(email_id)
    assert detail["attachments"] == []
    assert detail["inline_images"] == []


# ---- save_attachments -----------------------------------------------

def test_save_attachments_returns_metadata(monkeypatch, tmp_path):
    plain = FakeAttachment("report.pdf", b"%PDF-1.4")
    client, email_id = _client_with_mail(monkeypatch, [_inline_png(), plain])
    png, pdf = client.save_attachments(email_id, str(tmp_path))
    assert png["status"] == "saved"
    assert png["index"] == 1
    assert png["content_id"] == "image001.png@01D9ABCD"
    assert png["mime_type"] == "image/png"
    assert png["size"] == 8
    assert (tmp_path / "image001.png").read_bytes() == b"\x89PNG...."
    assert pdf["index"] == 2
    assert pdf["content_id"] is None
    assert pdf["mime_type"] == "application/pdf"
    assert pdf["saved_to"] == str(tmp_path / "report.pdf")
    assert png["is_inline"] is True
    assert pdf["is_inline"] is False


def test_save_attachments_filtered_keeps_original_index(monkeypatch, tmp_path):
    plain = FakeAttachment("report.pdf", b"%PDF-1.4")
    client, email_id = _client_with_mail(monkeypatch, [_inline_png(), plain])
    [pdf] = client.save_attachments(email_id, str(tmp_path),
                                    attachment_names=["REPORT.pdf"])
    assert pdf["index"] == 2
    assert pdf["filename"] == "report.pdf"


def test_save_attachments_failure_still_carries_metadata(monkeypatch, tmp_path):
    class Broken(FakeAttachment):
        def SaveAsFile(self, path):
            raise com_fakes.com_error("Cannot save the attachment.")

    client, email_id = _client_with_mail(monkeypatch, [Broken("x.pdf", b"abc")])
    [res] = client.save_attachments(email_id, str(tmp_path))
    assert res["status"] == "failed"
    assert res["index"] == 1
    assert res["size"] == 3
    assert res["mime_type"] == "application/pdf"
    assert "error" in res
    assert "saved_to" not in res
