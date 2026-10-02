"""WindowsOutlookClient tests against the fake COM objects in
tests/com_fakes.py — exercises the real client code paths without
Outlook."""

import base64
import os

import pytest

from outlook_mcp import constants as c
from outlook_mcp.errors import ToolError
from outlook_mcp.outlook import client as client_mod
from tests import com_fakes
from tests.com_fakes import FakeAttachment, FakeFolder, FakeMail, FakeNamespace


def _client_with_mail(monkeypatch, attachments):
    ns = FakeNamespace()
    mail = FakeMail(attachments=attachments)
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


# ---- get_inline_image -----------------------------------------------

def _decode(data_uri):
    header, payload = data_uri.split(",", 1)
    return header, base64.b64decode(payload)


@pytest.mark.parametrize("requested", [
    "image001.png@01D9ABCD",
    "cid:image001.png@01D9ABCD",
    "<IMAGE001.PNG@01d9abcd>",
    "  CID:<image001.png@01D9ABCD> ",
])
def test_get_inline_image_matches_normalized_content_id(monkeypatch, requested):
    plain = FakeAttachment("report.pdf", b"%PDF-1.4")
    client, email_id = _client_with_mail(monkeypatch, [plain, _inline_png()])
    result = client.get_inline_image(email_id, requested)
    assert result["content_id"] == "image001.png@01D9ABCD"
    assert result["filename"] == "image001.png"
    assert result["mime_type"] == "image/png"
    assert result["size"] == 8
    header, data = _decode(result["data_uri"])
    assert header == "data:image/png;base64"
    assert data == b"\x89PNG...."


def test_get_inline_image_falls_back_to_octet_stream(monkeypatch):
    att = FakeAttachment("blob", b"\x00\x01", props={c.PR_ATTACH_CONTENT_ID: "blob@x"})
    client, email_id = _client_with_mail(monkeypatch, [att])
    result = client.get_inline_image(email_id, "blob@x")
    assert result["mime_type"] == "application/octet-stream"
    assert result["data_uri"].startswith("data:application/octet-stream;base64,")


def test_get_inline_image_not_found_lists_available_ids(monkeypatch):
    other = FakeAttachment("logo.gif", b"GIF89a", props={
        c.PR_ATTACH_CONTENT_ID: "<logo@01D9>"})
    plain = FakeAttachment("report.pdf", b"%PDF-1.4")
    client, email_id = _client_with_mail(monkeypatch, [_inline_png(), plain, other])
    with pytest.raises(ToolError) as exc:
        client.get_inline_image(email_id, "cid:missing@x")
    msg = str(exc.value)
    assert "missing@x" in msg
    assert "image001.png@01D9ABCD" in msg
    assert "logo@01D9" in msg


def test_get_inline_image_email_without_content_ids(monkeypatch):
    plain = FakeAttachment("report.pdf", b"%PDF-1.4")
    client, email_id = _client_with_mail(monkeypatch, [plain])
    with pytest.raises(ToolError, match="no attachments with a Content-ID"):
        client.get_inline_image(email_id, "image001.png@01D9ABCD")


def test_get_inline_image_email_without_attachments(monkeypatch):
    client, email_id = _client_with_mail(monkeypatch, [])
    with pytest.raises(ToolError, match="list_attachments"):
        client.get_inline_image(email_id, "image001.png@01D9ABCD")


def test_get_inline_image_rejects_empty_content_id(monkeypatch):
    client, email_id = _client_with_mail(monkeypatch, [_inline_png()])
    with pytest.raises(ToolError, match="non-empty"):
        client.get_inline_image(email_id, "cid:<>")


def test_get_inline_image_size_cap_checked_before_saving(monkeypatch):
    att = _inline_png()
    att.Size = client_mod.MAX_INLINE_IMAGE_BYTES + 1
    client, email_id = _client_with_mail(monkeypatch, [att])
    with pytest.raises(ToolError, match="save_attachments"):
        client.get_inline_image(email_id, "image001.png@01D9ABCD")
    assert att.saved_to == []


def test_get_inline_image_size_cap_checked_on_actual_bytes(monkeypatch, tmp_path):
    monkeypatch.setattr(client_mod, "MAX_INLINE_IMAGE_BYTES", 4)
    att = _inline_png()  # 8 bytes on save...
    att.Size = 2         # ...though MAPI claims 2
    work = tmp_path / "work"
    monkeypatch.setattr(client_mod.tempfile, "mkdtemp",
                        lambda **_kw: (work.mkdir(), str(work))[1])
    client, email_id = _client_with_mail(monkeypatch, [att])
    with pytest.raises(ToolError, match="save_attachments"):
        client.get_inline_image(email_id, "image001.png@01D9ABCD")
    assert att.saved_to
    assert not work.exists()


def test_get_inline_image_removes_temp_dir(monkeypatch, tmp_path):
    work = tmp_path / "work"
    monkeypatch.setattr(client_mod.tempfile, "mkdtemp",
                        lambda **_kw: (work.mkdir(), str(work))[1])
    att = _inline_png()
    client, email_id = _client_with_mail(monkeypatch, [att])
    client.get_inline_image(email_id, "image001.png@01D9ABCD")
    [saved] = att.saved_to
    assert os.path.dirname(saved) == str(work)
    assert not work.exists()


def test_get_inline_image_removes_temp_dir_when_save_fails(monkeypatch, tmp_path):
    class Broken(FakeAttachment):
        def SaveAsFile(self, path):
            raise com_fakes.com_error("Cannot save the attachment.")

    work = tmp_path / "work"
    monkeypatch.setattr(client_mod.tempfile, "mkdtemp",
                        lambda **_kw: (work.mkdir(), str(work))[1])
    att = Broken("x.png", b"abc", props={c.PR_ATTACH_CONTENT_ID: "x@y"})
    client, email_id = _client_with_mail(monkeypatch, [att])
    with pytest.raises(ToolError, match="Cannot save"):
        client.get_inline_image(email_id, "x@y")
    assert not work.exists()
