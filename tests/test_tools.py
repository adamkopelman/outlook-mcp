"""Behavior tests: each tool delegates to the client with the right
arguments, and errors surface as MCP tool errors (isError: true)."""

import asyncio
import json
import sys

import pytest

from outlook_mcp import server
from tests.conftest import EMAIL_ID, EVENT_ID, NOTE_ID, TASK_ID


def call_tool(name, arguments):
    """Call through FastMCP (exercises schema validation + serialization)."""
    return asyncio.run(server.mcp.call_tool(name, arguments))


def result_json(content):
    assert content[0].type == "text"
    return json.loads(content[0].text)


# ---- Email ----------------------------------------------------------

def test_list_folders(fake_client):
    content = call_tool("list_folders", {})
    assert fake_client.calls == [("list_folders", {})]


def test_list_emails_passes_arguments(fake_client):
    call_tool("list_emails", {"folder": "sent", "count": 5,
                              "unread_only": True})
    assert fake_client.calls == [
        ("list_emails", {"folder": "sent", "count": 5, "unread_only": True,
                         "query": None, "sender": None, "category": None,
                         "received_after": None, "received_before": None,
                         "since_days": None, "has_attachments": None,
                         "flagged": False, "high_importance": False})
    ]


def test_list_emails_defaults(fake_client):
    call_tool("list_emails", {})
    assert fake_client.calls == [
        ("list_emails", {"folder": "inbox", "count": 10, "unread_only": False,
                         "query": None, "sender": None, "category": None,
                         "received_after": None, "received_before": None,
                         "since_days": None, "has_attachments": None,
                         "flagged": False, "high_importance": False})
    ]


def test_list_emails_returns_categories(fake_client):
    content = call_tool("list_emails", {})
    result = result_json(content)
    assert result["categories"] == ["Work"]


def test_list_emails_forwards_query_and_filters(fake_client):
    call_tool("list_emails", {
        "query": "invoice", "sender": "ada@x.com", "category": "Work",
        "since_days": 30, "has_attachments": True, "flagged": True,
        "high_importance": True,
    })
    name, kwargs = fake_client.calls[0]
    assert kwargs["query"] == "invoice"
    assert kwargs["sender"] == "ada@x.com"
    assert kwargs["category"] == "Work"
    assert kwargs["since_days"] == 30
    assert kwargs["has_attachments"] is True
    assert kwargs["flagged"] is True
    assert kwargs["high_importance"] is True


def test_get_email(fake_client):
    content = call_tool("get_email", {"email_id": EMAIL_ID})
    assert result_json(content)["body"] == "Hi there"


def test_get_email_includes_item_type(fake_client):
    content = call_tool("get_email", {"email_id": EMAIL_ID})
    result = result_json(content)
    assert result["item_type"] == "email"
    assert result["is_meeting"] is False


def test_send_email(fake_client):
    call_tool("send_email", {"to": ["a@example.com", "b@example.com"],
                             "subject": "Hi", "body": "Hello!"})
    name, kwargs = fake_client.calls[0]
    assert name == "send_email"
    assert kwargs["to"] == ["a@example.com", "b@example.com"]
    assert kwargs["html"] is False


def test_send_email_forwards_attachments(fake_client):
    call_tool("send_email", {
        "to": ["a@x.com"], "subject": "Hi", "body": "yo",
        "attachments": ["C:/tmp/a.pdf", "C:/tmp/b.png"],
    })
    name, kwargs = fake_client.calls[0]
    assert kwargs["attachments"] == ["C:/tmp/a.pdf", "C:/tmp/b.png"]


def test_create_draft(fake_client):
    content = call_tool("create_draft", {"to": ["a@example.com"],
                                         "subject": "Hi", "body": "Hello!"})
    assert result_json(content)["status"] == "draft_saved"


def test_reply_email(fake_client):
    call_tool("reply_email", {"email_id": EMAIL_ID, "body": "Thanks!",
                              "reply_all": True, "send": False})
    name, kwargs = fake_client.calls[0]
    assert kwargs["reply_all"] is True
    assert kwargs["send"] is False


def test_update_email_move_returns_new_id(fake_client):
    content = call_tool("update_email", {"email_id": EMAIL_ID,
                                         "move_to": "Archive"})
    result = result_json(content)
    assert result["id"] == "new-entry|store-1"
    assert result["status"] == "updated"
    assert result["changed"] == ["move_to"]


def test_update_email_state_only_keeps_same_id_and_lists_changes(fake_client):
    content = call_tool("update_email", {
        "email_id": EMAIL_ID, "mark_read": True, "flag": "follow_up",
        "add_categories": ["Work"], "importance": "high",
    })
    result = result_json(content)
    # No move -> id unchanged.
    assert result["id"] == EMAIL_ID
    assert result["changed"] == ["mark_read", "flag", "add_categories", "importance"]
    # The client saw the full update.
    name, kwargs = fake_client.calls[0]
    assert name == "update_email"
    assert kwargs["flag"] == "follow_up"
    assert kwargs["importance"] == "high"


def test_delete_email(fake_client):
    call_tool("delete_email", {"email_id": EMAIL_ID})
    assert fake_client.calls == [("delete_email", {"email_id": EMAIL_ID})]


# ---- Calendar -------------------------------------------------------

def test_list_events(fake_client):
    call_tool("list_events", {"start_date": "2026-06-10",
                              "end_date": "2026-06-17"})
    assert fake_client.calls == [
        ("list_events", {"start_date": "2026-06-10",
                         "end_date": "2026-06-17"})
    ]


def test_get_event(fake_client):
    content = call_tool("get_event", {"event_id": EVENT_ID})
    assert result_json(content)["subject"] == "Standup"


def test_create_event_with_attendees(fake_client):
    call_tool("create_event", {"subject": "Sync",
                               "start": "2026-06-12T14:00",
                               "end": "2026-06-12T15:00",
                               "attendees": ["a@example.com"]})
    name, kwargs = fake_client.calls[0]
    assert kwargs["attendees"] == ["a@example.com"]


def test_respond_to_meeting(fake_client):
    call_tool("respond_to_meeting", {"event_id": EVENT_ID,
                                     "response": "accept"})
    name, kwargs = fake_client.calls[0]
    assert kwargs["response"] == "accept"
    assert kwargs["send"] is True


# ---- Attachments ----------------------------------------------------

def test_list_attachments(fake_client):
    content = call_tool("list_attachments", {"email_id": EMAIL_ID})
    assert result_json(content)["filename"] == "report.pdf"


def test_save_attachments(fake_client):
    call_tool("save_attachments", {"email_id": EMAIL_ID,
                                   "save_dir": "/tmp/x",
                                   "attachment_names": ["report.pdf"]})
    name, kwargs = fake_client.calls[0]
    assert kwargs["save_dir"] == "/tmp/x"
    assert kwargs["attachment_names"] == ["report.pdf"]


# ---- Tasks ----------------------------------------------------------

def test_list_tasks(fake_client):
    call_tool("list_tasks", {"include_completed": True})
    assert fake_client.calls == [("list_tasks", {"include_completed": True})]


def test_create_task(fake_client):
    call_tool("create_task", {"subject": "Buy milk", "due_date": "2026-06-15",
                              "importance": "high"})
    name, kwargs = fake_client.calls[0]
    assert kwargs["importance"] == "high"


def test_complete_task(fake_client):
    call_tool("complete_task", {"task_id": TASK_ID})
    assert fake_client.calls == [("complete_task", {"task_id": TASK_ID})]


# ---- Notes ----------------------------------------------------------

def test_list_notes(fake_client):
    call_tool("list_notes", {})
    assert fake_client.calls == [("list_notes", {})]


def test_get_note(fake_client):
    content = call_tool("get_note", {"note_id": NOTE_ID})
    assert result_json(content)["body"].startswith("Ideas")


def test_create_note(fake_client):
    call_tool("create_note", {"body": "Ideas\n- one"})
    assert fake_client.calls == [("create_note", {"body": "Ideas\n- one"})]


# ---- Error paths ----------------------------------------------------

def test_client_error_propagates_as_tool_error(fake_client):
    fake_client.fail_with = "Outlook exploded"
    with pytest.raises(Exception, match="Outlook exploded"):
        call_tool("list_emails", {})


def test_unknown_tool_rejected(fake_client):
    with pytest.raises(Exception, match="(?i)unknown tool"):
        call_tool("format_hard_drive", {})


@pytest.mark.skipif(sys.platform == "win32",
                    reason="on Windows the default client is the real one")
def test_off_windows_default_client_errors_helpfully():
    server.set_client(None)
    try:
        with pytest.raises(Exception, match="requires Windows"):
            call_tool("list_emails", {})
    finally:
        server.set_client(None)
