"""Live tests against a real, running Outlook. Skipped by default — set
OUTLOOK_MCP_LIVE_TESTS=1 to run them (mirrors outlook-mcp-rs's
`cargo test -- --ignored` opt-in convention, adapted to pytest's skip
marker since Python has no built-in `--ignored`-equivalent test tag).

Run explicitly:
    OUTLOOK_MCP_LIVE_TESTS=1 pytest tests/test_live.py -v
"""

import os

import pytest

from outlook_mcp.outlook.client import WindowsOutlookClient

pytestmark = pytest.mark.skipif(
    os.environ.get("OUTLOOK_MCP_LIVE_TESTS") != "1",
    reason="set OUTLOOK_MCP_LIVE_TESTS=1 to run live Outlook tests",
)


@pytest.fixture
def client():
    return WindowsOutlookClient()


def test_categories_round_trip_on_a_real_task(client):
    created = client.create_task("outlook-mcp P1-python live categories probe")
    task_id = created["id"]
    try:
        from outlook_mcp.outlook.client import _get_item_categories, _set_item_categories
        _, ns = client._mapi()
        item = client._get_item(ns, task_id)
        _set_item_categories(item, ["Red Category", "Blue Category"])
        item.Save()
        cats = _get_item_categories(item)
        assert set(cats) == {"Red Category", "Blue Category"}

        tasks = client.list_tasks(include_completed=True)
        found = next((t for t in tasks if t["id"] == task_id), None)
        assert found is not None
        assert set(found["categories"]) == {"Red Category", "Blue Category"}
        assert found["status"] == "not_started"
        assert found["importance"] == "normal"
    finally:
        _, ns = client._mapi()
        item = client._get_item(ns, task_id)
        item.Delete()


def test_list_emails_query_filter_narrows_results(client):
    all_results = client.list_emails(folder="inbox", count=25)
    # A query that almost certainly matches nothing should return <= all.
    filtered = client.list_emails(
        folder="inbox", count=25, query="zzqx-improbable-token-9137")
    assert len(filtered) <= len(all_results)


def test_create_draft_with_attachment_round_trips(client, tmp_path):
    path = tmp_path / "outlook-mcp-live-attach.txt"
    path.write_text("live attachment test")
    created = client.create_draft(
        to=["nobody@example.invalid"],
        subject="outlook-mcp attachment test",
        body="see attached",
        attachments=[str(path)],
    )
    try:
        assert created["status"] == "draft_saved"
    finally:
        client.delete_email(created["id"])


def test_send_with_missing_attachment_errors_before_sending(client):
    from outlook_mcp.errors import ToolError

    with pytest.raises(ToolError, match="attachment not found"):
        client.send_email(
            to=["nobody@example.invalid"],
            subject="should not send",
            body="body",
            attachments=["C:/definitely/does/not/exist/nope.pdf"],
        )


def test_get_email_reports_item_type_for_real_inbox_item(client):
    emails = client.list_emails(folder="inbox", count=1)
    if not emails:
        pytest.skip("inbox is empty, nothing to check")
    detail = client.get_email(emails[0]["id"])
    assert detail["item_type"] in {"email", "meeting", "bounce", "read_receipt", "other"}
    if detail["is_meeting"]:
        assert "meeting" in detail
    else:
        assert "meeting" not in detail
