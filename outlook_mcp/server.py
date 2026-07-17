"""FastMCP server exposing Microsoft Outlook desktop over MCP (stdio).

The MCP protocol layer is handled entirely by the official `mcp` SDK;
this module just registers tools that delegate to an Outlook client.
On Windows the client drives classic Outlook via COM (pywin32); on other
platforms an UnavailableClient makes every tool return a clear error so
the server can still start and list its tools.
"""

import logging
import sys
from typing import Optional

from mcp.server.fastmcp import FastMCP

from outlook_mcp.outlook.base import OutlookClientBase, UnavailableClient

logger = logging.getLogger("outlook-mcp")

mcp = FastMCP("outlook")

_client: Optional[OutlookClientBase] = None


def get_client() -> OutlookClientBase:
    global _client
    if _client is None:
        if sys.platform == "win32":
            from outlook_mcp.outlook.client import WindowsOutlookClient
            _client = WindowsOutlookClient()
        else:
            logger.warning(
                "Not running on Windows: Outlook COM is unavailable and all "
                "tools will return errors."
            )
            _client = UnavailableClient()
    return _client


def set_client(client: Optional[OutlookClientBase]) -> None:
    """Inject a client (used by tests)."""
    global _client
    _client = client


# ---- Email tools ----------------------------------------------------

@mcp.tool()
def list_folders():
    """List all Outlook mail folders with their paths, item counts and
    unread counts. Folder paths returned here can be passed to other
    tools' `folder`/`target_folder` arguments."""
    return get_client().list_folders()


@mcp.tool()
def list_emails(folder: str = "inbox", count: int = 10,
                unread_only: bool = False, query: Optional[str] = None,
                sender: Optional[str] = None, category: Optional[str] = None,
                received_after: Optional[str] = None,
                received_before: Optional[str] = None,
                since_days: Optional[int] = None,
                has_attachments: Optional[bool] = None,
                flagged: bool = False, high_importance: bool = False):
    """Find emails in a folder with optional text query and filters.
    `folder` accepts a well-known name (inbox, sent, drafts, deleted, outbox)
    or a path like 'Inbox/Receipts'. `count` is capped at 50. `query`
    searches subject/sender/body text. `sender` matches sender name or
    address. `category` filters by color category. `received_after`/
    `received_before` are ISO dates; `since_days` is a relative alternative.
    `has_attachments`, `flagged`, `high_importance` are optional boolean
    filters."""
    return get_client().list_emails(
        folder=folder, count=count, unread_only=unread_only, query=query,
        sender=sender, category=category, received_after=received_after,
        received_before=received_before, since_days=since_days,
        has_attachments=has_attachments, flagged=flagged,
        high_importance=high_importance)


@mcp.tool()
def get_email(email_id: str, prefer_html: bool = False):
    """Read a full email (headers, body, attachment names) by the id
    returned from list_emails. Set `prefer_html` to also get the HTML
    body."""
    return get_client().get_email(email_id=email_id, prefer_html=prefer_html)


@mcp.tool()
def send_email(to: list[str], subject: str, body: str,
               cc: Optional[list[str]] = None,
               bcc: Optional[list[str]] = None, html: bool = False,
               attachments: Optional[list[str]] = None):
    """Send an email immediately as the signed-in Outlook user. Use
    create_draft instead if the user should review before sending.
    Set `html` to true if `body` is HTML. `attachments` is a list of local
    file paths; a missing path fails before anything is sent."""
    return get_client().send_email(to=to, subject=subject, body=body, cc=cc,
                                   bcc=bcc, html=html, attachments=attachments)


@mcp.tool()
def create_draft(to: list[str], subject: str, body: str,
                 cc: Optional[list[str]] = None,
                 bcc: Optional[list[str]] = None, html: bool = False,
                 attachments: Optional[list[str]] = None):
    """Compose an email and save it to Drafts without sending. Returns the
    draft's id. `attachments` is a list of local file paths; a missing path
    fails before saving."""
    return get_client().create_draft(to=to, subject=subject, body=body, cc=cc,
                                     bcc=bcc, html=html, attachments=attachments)


@mcp.tool()
def reply_email(email_id: str, body: str, reply_all: bool = False,
                html: bool = False, send: bool = True,
                attachments: Optional[list[str]] = None):
    """Reply to an email (reply-all optional). Sends immediately by
    default; pass send=false to save the reply as a draft instead.
    `attachments` is a list of local file paths; a missing path fails
    before anything is sent or saved."""
    return get_client().reply_email(email_id=email_id, body=body,
                                    reply_all=reply_all, html=html, send=send,
                                    attachments=attachments)


@mcp.tool()
def move_email(email_id: str, target_folder: str):
    """Move an email to another folder. Returns the email's NEW id (ids
    change when an item moves)."""
    return get_client().move_email(email_id=email_id,
                                   target_folder=target_folder)


@mcp.tool()
def delete_email(email_id: str):
    """Delete an email (moves it to Deleted Items)."""
    return get_client().delete_email(email_id=email_id)


# ---- Calendar tools -------------------------------------------------

@mcp.tool()
def list_events(start_date: Optional[str] = None,
                end_date: Optional[str] = None):
    """List calendar events between two ISO dates (recurring events are
    expanded). Defaults to the next 7 days starting today."""
    return get_client().list_events(start_date=start_date, end_date=end_date)


@mcp.tool()
def get_event(event_id: str):
    """Get full details of a calendar event by id, including attendees
    and body."""
    return get_client().get_event(event_id=event_id)


@mcp.tool()
def create_event(subject: str, start: str, end: str,
                 body: Optional[str] = None, location: Optional[str] = None,
                 attendees: Optional[list[str]] = None, all_day: bool = False,
                 reminder_minutes: Optional[int] = None):
    """Create a calendar appointment. `start`/`end` are ISO datetimes like
    '2026-06-12T14:00'. If `attendees` is given, the event becomes a
    meeting and invitations are SENT immediately."""
    return get_client().create_event(subject=subject, start=start, end=end,
                                     body=body, location=location,
                                     attendees=attendees, all_day=all_day,
                                     reminder_minutes=reminder_minutes)


@mcp.tool()
def respond_to_meeting(event_id: str, response: str,
                       comment: Optional[str] = None, send: bool = True):
    """Respond to a meeting invitation: `response` is 'accept', 'decline'
    or 'tentative'. The response is sent to the organizer as the signed-in
    user (pass send=false to save without sending)."""
    return get_client().respond_to_meeting(event_id=event_id,
                                           response=response,
                                           comment=comment, send=send)


# ---- Attachment tools -----------------------------------------------

@mcp.tool()
def list_attachments(email_id: str):
    """List an email's attachments (file names and sizes)."""
    return get_client().list_attachments(email_id=email_id)


@mcp.tool()
def save_attachments(email_id: str, save_dir: str,
                     attachment_names: Optional[list[str]] = None):
    """Save an email's attachments to a local directory (created if
    missing). By default saves all attachments; pass `attachment_names`
    to save only specific files."""
    return get_client().save_attachments(email_id=email_id, save_dir=save_dir,
                                         attachment_names=attachment_names)


# ---- Task tools -----------------------------------------------------

@mcp.tool()
def list_tasks(include_completed: bool = False):
    """List Outlook tasks (open tasks only, unless include_completed)."""
    return get_client().list_tasks(include_completed=include_completed)


@mcp.tool()
def create_task(subject: str, body: Optional[str] = None,
                due_date: Optional[str] = None, importance: str = "normal"):
    """Create an Outlook task. `due_date` is an ISO date; `importance` is
    'low', 'normal' or 'high'."""
    return get_client().create_task(subject=subject, body=body,
                                    due_date=due_date, importance=importance)


@mcp.tool()
def complete_task(task_id: str):
    """Mark a task as complete."""
    return get_client().complete_task(task_id=task_id)


# ---- Note tools -----------------------------------------------------

@mcp.tool()
def list_notes():
    """List Outlook sticky notes (id, first line, creation time)."""
    return get_client().list_notes()


@mcp.tool()
def get_note(note_id: str):
    """Read the full body of a note by id."""
    return get_client().get_note(note_id=note_id)


@mcp.tool()
def create_note(body: str):
    """Create an Outlook sticky note. The first line of `body` becomes
    the note's title."""
    return get_client().create_note(body=body)


def main() -> None:
    logging.basicConfig(stream=sys.stderr, level=logging.INFO)
    mcp.run()  # stdio transport


if __name__ == "__main__":
    main()
