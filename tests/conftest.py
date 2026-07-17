from typing import Optional

import pytest

from outlook_mcp import server
from outlook_mcp.errors import ToolError
from outlook_mcp.outlook.base import OutlookClientBase

EMAIL_ID = "entry-1|store-1"
EVENT_ID = "entry-2|store-1"
TASK_ID = "entry-3|store-1"
NOTE_ID = "entry-4|store-1"


class FakeOutlookClient(OutlookClientBase):
    """In-memory stand-in for COM Outlook; records every call."""

    def __init__(self):
        self.calls = []
        self.fail_with: Optional[str] = None

    def _record(self, name, **kwargs):
        if self.fail_with:
            raise ToolError(self.fail_with)
        self.calls.append((name, kwargs))

    # Email
    def list_folders(self):
        self._record("list_folders")
        return [{"name": "Inbox", "path": "Inbox", "items": 2, "unread": 1}]

    def list_emails(self, folder="inbox", count=10, unread_only=False,
                    query=None, sender=None, category=None,
                    received_after=None, received_before=None,
                    since_days=None, has_attachments=None,
                    flagged=False, high_importance=False):
        self._record("list_emails", folder=folder, count=count,
                     unread_only=unread_only, query=query, sender=sender,
                     category=category, received_after=received_after,
                     received_before=received_before, since_days=since_days,
                     has_attachments=has_attachments, flagged=flagged,
                     high_importance=high_importance)
        return [{"id": EMAIL_ID, "subject": "Hello", "sender": "Ada",
                 "unread": True, "categories": ["Work"]}]

    def get_email(self, email_id, prefer_html=False):
        self._record("get_email", email_id=email_id, prefer_html=prefer_html)
        return {"id": email_id, "subject": "Hello", "body": "Hi there",
                "categories": []}

    def send_email(self, to, subject, body, cc=None, bcc=None, html=False,
                   attachments=None):
        self._record("send_email", to=to, subject=subject, body=body, cc=cc,
                     bcc=bcc, html=html, attachments=attachments)
        return {"status": "sent", "to": "; ".join(to), "subject": subject}

    def create_draft(self, to, subject, body, cc=None, bcc=None, html=False,
                     attachments=None):
        self._record("create_draft", to=to, subject=subject, body=body,
                     cc=cc, bcc=bcc, html=html, attachments=attachments)
        return {"status": "draft_saved", "id": EMAIL_ID, "subject": subject}

    def reply_email(self, email_id, body, reply_all=False, html=False,
                    send=True, attachments=None):
        self._record("reply_email", email_id=email_id, body=body,
                     reply_all=reply_all, html=html, send=send,
                     attachments=attachments)
        return {"status": "sent" if send else "draft_saved"}

    def move_email(self, email_id, target_folder):
        self._record("move_email", email_id=email_id,
                     target_folder=target_folder)
        return {"status": "moved", "folder": target_folder,
                "id": "new-entry|store-1"}

    def delete_email(self, email_id):
        self._record("delete_email", email_id=email_id)
        return {"status": "deleted"}

    # Calendar
    def list_events(self, start_date=None, end_date=None):
        self._record("list_events", start_date=start_date, end_date=end_date)
        return [{"id": EVENT_ID, "subject": "Standup", "categories": []}]

    def get_event(self, event_id):
        self._record("get_event", event_id=event_id)
        return {"id": event_id, "subject": "Standup", "body": "",
                "response": "accepted", "categories": []}

    def create_event(self, subject, start, end, body=None, location=None,
                     attendees=None, all_day=False, reminder_minutes=None):
        self._record("create_event", subject=subject, start=start, end=end,
                     body=body, location=location, attendees=attendees,
                     all_day=all_day, reminder_minutes=reminder_minutes)
        return {"status": "saved", "id": EVENT_ID, "subject": subject}

    def respond_to_meeting(self, event_id, response, comment=None, send=True):
        self._record("respond_to_meeting", event_id=event_id,
                     response=response, comment=comment, send=send)
        return {"status": f"{response}_sent"}

    # Attachments
    def list_attachments(self, email_id):
        self._record("list_attachments", email_id=email_id)
        return [{"index": 1, "filename": "report.pdf", "size": 1234}]

    def save_attachments(self, email_id, save_dir, attachment_names=None):
        self._record("save_attachments", email_id=email_id, save_dir=save_dir,
                     attachment_names=attachment_names)
        return [{"filename": "report.pdf", "saved_to": save_dir,
                 "status": "saved"}]

    # Tasks
    def list_tasks(self, include_completed=False):
        self._record("list_tasks", include_completed=include_completed)
        return [{"id": TASK_ID, "subject": "Buy milk", "complete": False,
                 "status": "not_started", "importance": "normal",
                 "categories": []}]

    def create_task(self, subject, body=None, due_date=None,
                    importance="normal"):
        self._record("create_task", subject=subject, body=body,
                     due_date=due_date, importance=importance)
        return {"status": "created", "id": TASK_ID, "subject": subject}

    def complete_task(self, task_id):
        self._record("complete_task", task_id=task_id)
        return {"status": "completed"}

    # Notes
    def list_notes(self):
        self._record("list_notes")
        return [{"id": NOTE_ID, "subject": "Ideas", "categories": []}]

    def get_note(self, note_id):
        self._record("get_note", note_id=note_id)
        return {"id": note_id, "subject": "Ideas", "body": "Ideas\n- one",
                "categories": []}

    def create_note(self, body):
        self._record("create_note", body=body)
        return {"status": "created", "id": NOTE_ID}


@pytest.fixture
def fake_client():
    client = FakeOutlookClient()
    server.set_client(client)
    yield client
    server.set_client(None)
