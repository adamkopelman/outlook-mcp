"""WindowsOutlookClient tests against the fake COM objects in
tests/com_fakes.py — no Outlook required."""

from tests import com_fakes
from tests.com_fakes import FakeFolder, FakeMail, FakeNamespace


def _client_with_mail(monkeypatch, **attrs):
    mail = FakeMail(**attrs)
    FakeFolder("Inbox", items=[mail])
    ns = FakeNamespace()
    email_id = ns.register(mail)
    return com_fakes.install(monkeypatch, ns), email_id


def test_get_email_reports_html_truncation_and_full_length(monkeypatch):
    html = "<html><body>" + "a" * 150_000 + "</body></html>"
    client, email_id = _client_with_mail(monkeypatch, Body="hello",
                                         HTMLBody=html)

    info = client.get_email(email_id, prefer_html=True)
    assert info["html_truncated"] is True
    assert info["html_length"] == len(html)
    assert info["html_body"].startswith(html[:100_000])
    assert "[... truncated at 100000 characters]" in info["html_body"]
    assert info["body"] == "hello"
    assert info["body_truncated"] is False
    assert info["body_length"] == 5

    info = client.get_email(email_id, prefer_html=True, max_body_chars=200_000)
    assert info["html_truncated"] is False
    assert info["html_body"] == html


def test_get_email_without_html_omits_html_fields(monkeypatch):
    client, email_id = _client_with_mail(monkeypatch, Body="b" * 2_000)

    info = client.get_email(email_id, max_body_chars=1_500)
    assert "html_body" not in info
    assert "html_truncated" not in info
    assert info["body_truncated"] is True
    assert info["body_length"] == 2_000


def test_get_email_clamps_tiny_limit(monkeypatch):
    client, email_id = _client_with_mail(monkeypatch, Body="c" * 1_000)

    info = client.get_email(email_id, max_body_chars=10)
    assert info["body_truncated"] is False  # clamped up to 1,000
    assert info["body"] == "c" * 1_000
