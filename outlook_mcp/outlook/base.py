"""Abstract Outlook client interface.

This module is importable on any platform. The real implementation
(WindowsOutlookClient in outlook_mcp.outlook.client) is only imported on
Windows; tests inject a fake implementing this interface.

All methods accept and return plain JSON-able Python values (str, int,
bool, dict, list). Items are addressed by an opaque ``id`` string of the
form ``"{EntryID}|{StoreID}"`` returned from list/search calls.
"""

from typing import Optional

from outlook_mcp.errors import ToolError


class OutlookClientBase:
    # ---- Email -----------------------------------------------------
    def list_folders(self) -> list:
        raise NotImplementedError

    def list_emails(self, folder: str = "inbox", count: int = 10,
                    unread_only: bool = False) -> list:
        raise NotImplementedError

    def search_emails(self, query: str, folder: str = "inbox", count: int = 10,
                      since_days: Optional[int] = None) -> list:
        raise NotImplementedError

    def get_email(self, email_id: str, prefer_html: bool = False) -> dict:
        raise NotImplementedError

    def send_email(self, to: list, subject: str, body: str,
                   cc: Optional[list] = None, bcc: Optional[list] = None,
                   html: bool = False) -> dict:
        raise NotImplementedError

    def create_draft(self, to: list, subject: str, body: str,
                     cc: Optional[list] = None, bcc: Optional[list] = None,
                     html: bool = False) -> dict:
        raise NotImplementedError

    def reply_email(self, email_id: str, body: str, reply_all: bool = False,
                    html: bool = False, send: bool = True) -> dict:
        raise NotImplementedError

    def move_email(self, email_id: str, target_folder: str) -> dict:
        raise NotImplementedError

    def delete_email(self, email_id: str) -> dict:
        raise NotImplementedError

    # ---- Calendar --------------------------------------------------
    def list_events(self, start_date: Optional[str] = None,
                    end_date: Optional[str] = None) -> list:
        raise NotImplementedError

    def get_event(self, event_id: str) -> dict:
        raise NotImplementedError

    def create_event(self, subject: str, start: str, end: str,
                     body: Optional[str] = None, location: Optional[str] = None,
                     attendees: Optional[list] = None, all_day: bool = False,
                     reminder_minutes: Optional[int] = None) -> dict:
        raise NotImplementedError

    def respond_to_meeting(self, event_id: str, response: str,
                           comment: Optional[str] = None,
                           send: bool = True) -> dict:
        raise NotImplementedError

    # ---- Attachments -----------------------------------------------
    def list_attachments(self, email_id: str) -> list:
        raise NotImplementedError

    def save_attachments(self, email_id: str, save_dir: str,
                         attachment_names: Optional[list] = None) -> list:
        raise NotImplementedError

    # ---- Tasks -----------------------------------------------------
    def list_tasks(self, include_completed: bool = False) -> list:
        raise NotImplementedError

    def create_task(self, subject: str, body: Optional[str] = None,
                    due_date: Optional[str] = None,
                    importance: str = "normal") -> dict:
        raise NotImplementedError

    def complete_task(self, task_id: str) -> dict:
        raise NotImplementedError

    # ---- Notes -----------------------------------------------------
    def list_notes(self) -> list:
        raise NotImplementedError

    def get_note(self, note_id: str) -> dict:
        raise NotImplementedError

    def create_note(self, body: str) -> dict:
        raise NotImplementedError


class UnavailableClient(OutlookClientBase):
    """Stands in when not running on Windows: the server starts and lists
    tools anywhere, but every call explains the platform requirement."""

    _MESSAGE = ("Outlook is not available: this server requires Windows with "
                "classic Outlook desktop installed (Outlook COM automation).")

    def __getattribute__(self, name):
        base = object.__getattribute__(self, "__class__").__mro__[1]
        if name.startswith("_") or not callable(getattr(base, name, None)):
            return object.__getattribute__(self, name)

        def _unavailable(*args, **kwargs):
            raise ToolError(UnavailableClient._MESSAGE)

        return _unavailable
