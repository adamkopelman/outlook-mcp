"""Win32 COM implementation of the Outlook client.

Only importable on Windows (requires pywin32 and classic Outlook desktop).
Every public method:
  * calls pythoncom.CoInitialize() — tool calls may run on different worker
    threads, and CoInitialize is idempotent per thread;
  * creates its own Dispatch/namespace so COM objects never cross threads;
  * converts pywintypes.com_error into ToolError with a readable message.
"""

import functools
import os
import re
from datetime import date, datetime, time, timedelta
from typing import Optional

import pythoncom
import pywintypes
import win32com.client

from outlook_mcp import constants as c
from outlook_mcp import friendly
from outlook_mcp.errors import ToolError, format_com_error
from outlook_mcp.outlook.base import OutlookClientBase

MAX_EMAIL_COUNT = 50
MAX_CALENDAR_ITEMS = 250
MAX_BODY_CHARS = 100_000

_UNSAFE_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def _com(method):
    """Decorator for public client methods: COM init + error mapping."""

    @functools.wraps(method)
    def wrapper(self, *args, **kwargs):
        pythoncom.CoInitialize()
        try:
            return method(self, *args, **kwargs)
        except pywintypes.com_error as exc:
            raise ToolError(format_com_error(exc)) from exc

    return wrapper


def _to_iso(value) -> Optional[str]:
    """pywintypes.datetime (or anything date-like) -> ISO-8601 string."""
    if value is None:
        return None
    try:
        return value.isoformat()
    except (AttributeError, ValueError, OSError):
        return str(value)


def _jet_dt(dt: datetime) -> str:
    """Format a datetime for JET Restrict filters: MM/DD/YYYY HH:MM AM/PM
    (US format, no seconds — anything else silently misfilters)."""
    return dt.strftime("%m/%d/%Y %I:%M %p")


def _parse_dt(value: str, field: str) -> datetime:
    try:
        return datetime.fromisoformat(value)
    except (TypeError, ValueError):
        raise ToolError(
            f"Invalid {field} {value!r}: expected ISO format like "
            f"'2026-06-10' or '2026-06-10T14:30'"
        )


def _truncate(text: str) -> str:
    if text and len(text) > MAX_BODY_CHARS:
        return text[:MAX_BODY_CHARS] + f"\n\n[... truncated at {MAX_BODY_CHARS} characters]"
    return text


def _safe_filename(name: str) -> str:
    cleaned = _UNSAFE_FILENAME_CHARS.sub("_", name or "").strip(". ")
    return cleaned or "attachment"


def _parse_categories(raw: str) -> list:
    """Outlook stores categories as one ", "-joined string. Split it into
    names, trimming whitespace and dropping empties."""
    return [s.strip() for s in (raw or "").split(",") if s.strip()]


def _join_categories(cats: list) -> str:
    """Join category names back into the ", "-separated string Outlook
    expects."""
    return ", ".join(cats)


def _get_item_categories(item) -> list:
    """Read an item's color categories (empty list if the property is
    missing or blank)."""
    return _parse_categories(getattr(item, "Categories", "") or "")


def _set_item_categories(item, cats: list) -> None:
    """Overwrite an item's categories with the given list."""
    item.Categories = _join_categories(cats)


def _attach_files(item, paths: list) -> None:
    """Attach local files to a mail/reply item. Validates every path exists
    FIRST (so a bad path fails before anything is sent), then adds each via
    item.Attachments.Add(path)."""
    for p in paths:
        if not os.path.isfile(p):
            raise ToolError(f"attachment not found: {p}")
    for p in paths:
        item.Attachments.Add(p)


class WindowsOutlookClient(OutlookClientBase):
    """Talks to a running (or auto-launched) classic Outlook via COM."""

    # ---- plumbing ---------------------------------------------------

    def _mapi(self):
        app = win32com.client.Dispatch("Outlook.Application")
        return app, app.GetNamespace("MAPI")

    @staticmethod
    def _make_id(item) -> str:
        return f"{item.EntryID}|{item.Parent.StoreID}"

    def _get_item(self, ns, item_id: str):
        entry_id, sep, store_id = (item_id or "").partition("|")
        if not sep or not entry_id or not store_id:
            raise ToolError(
                f"Invalid item id {item_id!r}: expected the opaque id returned "
                f"by a list/search tool."
            )
        try:
            return ns.GetItemFromID(entry_id, store_id)
        except pywintypes.com_error as exc:
            raise ToolError(
                "Item not found — it may have been moved or deleted "
                "(item ids change when an item moves to another folder). "
                + format_com_error(exc)
            ) from exc

    def _resolve_folder(self, ns, folder: Optional[str]):
        name = (folder or "inbox").strip()
        key = name.lower()
        if key in c.FOLDER_NAME_TO_ID:
            return ns.GetDefaultFolder(c.FOLDER_NAME_TO_ID[key])
        # Walk a path like "Inbox/Receipts" from the default store root.
        current = ns.GetDefaultFolder(c.OL_FOLDER_INBOX).Parent
        for part in [p for p in re.split(r"[/\\]", name) if p]:
            match = None
            for sub in current.Folders:
                if sub.Name.lower() == part.lower():
                    match = sub
                    break
            if match is None:
                raise ToolError(
                    f"Folder not found: {name!r} (no subfolder named {part!r} "
                    f"under {current.Name!r})"
                )
            current = match
        return current

    # ---- summaries --------------------------------------------------

    def _email_summary(self, item) -> dict:
        attachments = getattr(item, "Attachments", None)
        return {
            "id": self._make_id(item),
            "subject": getattr(item, "Subject", "") or "",
            "sender": getattr(item, "SenderName", "") or "",
            "sender_email": getattr(item, "SenderEmailAddress", "") or "",
            "to": getattr(item, "To", "") or "",
            "received": _to_iso(getattr(item, "ReceivedTime", None)),
            "unread": bool(getattr(item, "UnRead", False)),
            "has_attachments": bool(attachments and attachments.Count > 0),
            "categories": _get_item_categories(item),
        }

    def _event_summary(self, item) -> dict:
        return {
            "id": self._make_id(item),
            "subject": getattr(item, "Subject", "") or "",
            "start": _to_iso(getattr(item, "Start", None)),
            "end": _to_iso(getattr(item, "End", None)),
            "location": getattr(item, "Location", "") or "",
            "organizer": getattr(item, "Organizer", "") or "",
            "all_day": bool(getattr(item, "AllDayEvent", False)),
            "is_recurring": bool(getattr(item, "IsRecurring", False)),
            "is_meeting": getattr(item, "MeetingStatus", c.OL_NONMEETING) != c.OL_NONMEETING,
            "categories": _get_item_categories(item),
        }

    def _task_summary(self, item) -> dict:
        return {
            "id": self._make_id(item),
            "subject": getattr(item, "Subject", "") or "",
            "due_date": _to_iso(getattr(item, "DueDate", None)),
            "complete": bool(getattr(item, "Complete", False)),
            "status": friendly.task_status_word(
                getattr(item, "Status", c.OL_TASK_NOT_STARTED)
            ),
            "importance": friendly.importance_word(
                getattr(item, "Importance", c.OL_IMPORTANCE_NORMAL)
            ),
            "categories": _get_item_categories(item),
        }

    def _note_summary(self, item) -> dict:
        body = getattr(item, "Body", "") or ""
        first_line = body.strip().splitlines()[0] if body.strip() else ""
        return {
            "id": self._make_id(item),
            "subject": first_line[:120],
            "created": _to_iso(getattr(item, "CreationTime", None)),
            "categories": _get_item_categories(item),
        }

    # ---- Email ------------------------------------------------------

    @_com
    def list_folders(self) -> list:
        _, ns = self._mapi()
        root = ns.GetDefaultFolder(c.OL_FOLDER_INBOX).Parent

        results = []

        def walk(folder, path, depth):
            try:
                item_count = folder.Items.Count
            except pywintypes.com_error:
                item_count = 0
            results.append({
                "name": folder.Name,
                "path": path,
                "items": item_count,
                "unread": getattr(folder, "UnReadItemCount", 0),
            })
            if depth >= 3:
                return
            for sub in folder.Folders:
                walk(sub, f"{path}/{sub.Name}", depth + 1)

        for sub in root.Folders:
            walk(sub, sub.Name, 1)
        return results

    @_com
    def list_emails(self, folder: str = "inbox", count: int = 10,
                    unread_only: bool = False, query: Optional[str] = None,
                    sender: Optional[str] = None, category: Optional[str] = None,
                    received_after: Optional[str] = None,
                    received_before: Optional[str] = None,
                    since_days: Optional[int] = None,
                    has_attachments: Optional[bool] = None,
                    flagged: bool = False, high_importance: bool = False) -> list:
        _, ns = self._mapi()
        count = max(1, min(int(count), MAX_EMAIL_COUNT))
        items = self._resolve_folder(ns, folder).Items

        # Text query: DASL @SQL across subject/sender/body (escaped).
        if query:
            q = query.replace("'", "''")
            dasl = (
                '@SQL=("urn:schemas:httpmail:subject" LIKE \'%{0}%\' '
                'OR "urn:schemas:httpmail:fromname" LIKE \'%{0}%\' '
                'OR "urn:schemas:httpmail:textdescription" LIKE \'%{0}%\')'
            ).format(q)
            items = items.Restrict(dasl)
        # Sender: DASL @SQL against fromname + fromemail.
        if sender:
            s = sender.replace("'", "''")
            dasl = (
                '@SQL=("urn:schemas:httpmail:fromname" LIKE \'%{0}%\' '
                'OR "urn:schemas:httpmail:fromemail" LIKE \'%{0}%\')'
            ).format(s)
            items = items.Restrict(dasl)
        if unread_only:
            items = items.Restrict("[UnRead] = True")
        if flagged:
            # FlagStatus 2 = flagged/marked.
            items = items.Restrict("[FlagStatus] = 2")
        if high_importance:
            items = items.Restrict("[Importance] = 2")
        # Date filters: since_days (relative), received_after/before (absolute).
        if since_days:
            cutoff = datetime.now() - timedelta(days=int(since_days))
            items = items.Restrict(f"[ReceivedTime] >= '{_jet_dt(cutoff)}'")
        if received_after:
            dt = _parse_dt(received_after, "received_after")
            items = items.Restrict(f"[ReceivedTime] >= '{_jet_dt(dt)}'")
        if received_before:
            dt = _parse_dt(received_before, "received_before")
            items = items.Restrict(f"[ReceivedTime] <= '{_jet_dt(dt)}'")

        items.Sort("[ReceivedTime]", True)

        # Client-side fuzzy filters: category + has_attachments. Iterate,
        # build each summary, keep it only if it passes, stop at count.
        cat_want = category.lower() if category else None
        results = []
        for item in items:
            summary = self._email_summary(item)
            if cat_want is not None:
                if not any(c.lower() == cat_want for c in summary["categories"]):
                    continue
            if has_attachments is not None:
                if summary["has_attachments"] != has_attachments:
                    continue
            results.append(summary)
            if len(results) >= count:
                break
        return results

    @_com
    def get_email(self, email_id: str, prefer_html: bool = False) -> dict:
        _, ns = self._mapi()
        item = self._get_item(ns, email_id)
        info = self._email_summary(item)
        info["cc"] = getattr(item, "CC", "") or ""
        info["bcc"] = getattr(item, "BCC", "") or ""
        info["body"] = _truncate(getattr(item, "Body", "") or "")
        if prefer_html:
            info["html_body"] = _truncate(getattr(item, "HTMLBody", "") or "")
        attachments = getattr(item, "Attachments", None)
        info["attachments"] = (
            [attachments.Item(i).FileName for i in range(1, attachments.Count + 1)]
            if attachments and attachments.Count else []
        )

        message_class = getattr(item, "MessageClass", "") or ""
        item_type = friendly.item_type_from_class(message_class)
        is_meeting = item_type == "meeting"
        info["item_type"] = item_type
        info["is_meeting"] = is_meeting
        if is_meeting:
            # AddToCalendar=False: reading a meeting request must never
            # silently add it to the user's calendar as a side effect.
            try:
                appt = item.GetAssociatedAppointment(False)
                info["meeting"] = {
                    "meeting_type": friendly.meeting_type_from_class(message_class),
                    "start": _to_iso(getattr(appt, "Start", None)),
                    "end": _to_iso(getattr(appt, "End", None)),
                    "location": getattr(appt, "Location", "") or "",
                    "organizer": getattr(appt, "Organizer", "") or "",
                    "required_attendees": getattr(appt, "RequiredAttendees", "") or "",
                    "optional_attendees": getattr(appt, "OptionalAttendees", "") or "",
                    "is_recurring": bool(getattr(appt, "IsRecurring", False)),
                }
            except pywintypes.com_error:
                pass  # malformed meeting item: leave "meeting" absent

        return info

    def _compose(self, app, to, subject, body, cc, bcc, html):
        mail = app.CreateItem(c.OL_MAIL_ITEM)
        mail.To = "; ".join(to or [])
        if cc:
            mail.CC = "; ".join(cc)
        if bcc:
            mail.BCC = "; ".join(bcc)
        mail.Subject = subject or ""
        if html:
            mail.BodyFormat = c.OL_FORMAT_HTML
            mail.HTMLBody = body or ""
        else:
            mail.BodyFormat = c.OL_FORMAT_PLAIN
            mail.Body = body or ""
        return mail

    @_com
    def send_email(self, to: list, subject: str, body: str,
                   cc: Optional[list] = None, bcc: Optional[list] = None,
                   html: bool = False, attachments: Optional[list] = None) -> dict:
        if not to:
            raise ToolError("send_email requires at least one recipient in 'to'.")
        app, _ = self._mapi()
        mail = self._compose(app, to, subject, body, cc, bcc, html)
        if attachments:
            _attach_files(mail, attachments)
        mail.Send()
        return {"status": "sent", "to": "; ".join(to), "subject": subject}

    @_com
    def create_draft(self, to: list, subject: str, body: str,
                     cc: Optional[list] = None, bcc: Optional[list] = None,
                     html: bool = False, attachments: Optional[list] = None) -> dict:
        app, _ = self._mapi()
        mail = self._compose(app, to, subject, body, cc, bcc, html)
        if attachments:
            _attach_files(mail, attachments)
        mail.Save()  # Save first so EntryID exists
        return {"status": "draft_saved", "id": self._make_id(mail),
                "subject": subject}

    @_com
    def reply_email(self, email_id: str, body: str, reply_all: bool = False,
                    html: bool = False, send: bool = True,
                    attachments: Optional[list] = None) -> dict:
        _, ns = self._mapi()
        item = self._get_item(ns, email_id)
        reply = item.ReplyAll() if reply_all else item.Reply()
        if html:
            reply.HTMLBody = (body or "") + reply.HTMLBody
        else:
            reply.Body = (body or "") + "\n\n" + reply.Body
        if attachments:
            _attach_files(reply, attachments)
        if send:
            reply.Send()
            return {"status": "sent", "subject": reply.Subject}
        reply.Save()
        return {"status": "draft_saved", "id": self._make_id(reply),
                "subject": reply.Subject}

    @_com
    def update_email(self, email_id: str, move_to: Optional[str] = None,
                     mark_read: Optional[bool] = None,
                     flag: Optional[str] = None,
                     add_categories: Optional[list] = None,
                     remove_categories: Optional[list] = None,
                     importance: Optional[str] = None) -> dict:
        _, ns = self._mapi()
        item = self._get_item(ns, email_id)
        changed = []

        # ---- state changes first (they address the item by its current id) ----

        if mark_read is not None:
            # UnRead is the inverse of "read".
            item.UnRead = not mark_read
            item.Save()
            changed.append("mark_read")

        if flag is not None:
            flag_key = flag.strip().lower()
            if flag_key == "follow_up":
                # MarkAsTask flags for follow-up with no due date.
                item.MarkAsTask(c.OL_MARK_NO_DATE)
            elif flag_key == "complete":
                item.FlagStatus = c.OL_FLAG_COMPLETE
            elif flag_key == "clear":
                # ClearTaskFlag removes the follow-up flag entirely.
                item.ClearTaskFlag()
            else:
                raise ToolError(
                    f"Invalid flag {flag!r}: expected 'follow_up', 'complete', "
                    f"or 'clear'."
                )
            item.Save()
            changed.append("flag")

        # Categories: read the current set once, then add/remove against it,
        # so tagging never wipes existing categories.
        if add_categories is not None or remove_categories is not None:
            cats = _get_item_categories(item)
            if add_categories is not None:
                for a in add_categories:
                    if not any(existing.lower() == a.lower() for existing in cats):
                        cats.append(a)
                changed.append("add_categories")
            if remove_categories is not None:
                wanted = {r.lower() for r in remove_categories}
                cats = [c_ for c_ in cats if c_.lower() not in wanted]
                changed.append("remove_categories")
            _set_item_categories(item, cats)
            item.Save()

        if importance is not None:
            importance_key = importance.strip().lower()
            if importance_key not in c.IMPORTANCE_NAME_TO_ID:
                raise ToolError(
                    f"Invalid importance {importance!r}: use 'low', 'normal' "
                    f"or 'high'."
                )
            item.Importance = c.IMPORTANCE_NAME_TO_ID[importance_key]
            item.Save()
            changed.append("importance")

        # ---- move last (Move changes the EntryID) ----

        if move_to is not None:
            target = self._resolve_folder(ns, move_to)
            moved = item.Move(target)
            changed.append("move_to")
            new_id = self._make_id(moved)  # EntryID changed — return the new id.
        else:
            new_id = email_id

        return {"status": "updated", "id": new_id, "changed": changed}

    @_com
    def delete_email(self, email_id: str) -> dict:
        _, ns = self._mapi()
        item = self._get_item(ns, email_id)
        subject = getattr(item, "Subject", "") or ""
        item.Delete()
        return {"status": "deleted", "subject": subject,
                "note": "Moved to Deleted Items."}

    # ---- Calendar ---------------------------------------------------

    @_com
    def list_events(self, start_date: Optional[str] = None,
                    end_date: Optional[str] = None) -> list:
        _, ns = self._mapi()
        start = (_parse_dt(start_date, "start_date") if start_date
                 else datetime.combine(date.today(), time.min))
        end = (_parse_dt(end_date, "end_date") if end_date
               else start + timedelta(days=7))
        if end.time() == time.min and end_date and "T" not in end_date:
            end = datetime.combine(end.date(), time.max)  # whole end day
        items = ns.GetDefaultFolder(c.OL_FOLDER_CALENDAR).Items
        items.IncludeRecurrences = True  # must precede Sort/Restrict
        items.Sort("[Start]")
        flt = f"[Start] >= '{_jet_dt(start)}' AND [Start] <= '{_jet_dt(end)}'"
        results = []
        for item in items.Restrict(flt):
            results.append(self._event_summary(item))
            if len(results) >= MAX_CALENDAR_ITEMS:
                break  # recurrences without an end date expand forever
        return results

    @_com
    def get_event(self, event_id: str) -> dict:
        _, ns = self._mapi()
        item = self._get_item(ns, event_id)
        info = self._event_summary(item)
        info["body"] = _truncate(getattr(item, "Body", "") or "")
        info["required_attendees"] = getattr(item, "RequiredAttendees", "") or ""
        info["optional_attendees"] = getattr(item, "OptionalAttendees", "") or ""
        info["response"] = friendly.response_word(
            getattr(item, "ResponseStatus", c.OL_RESPONSE_NONE) or c.OL_RESPONSE_NONE
        )
        return info

    @_com
    def create_event(self, subject: str, start: str, end: str,
                     body: Optional[str] = None, location: Optional[str] = None,
                     attendees: Optional[list] = None, all_day: bool = False,
                     reminder_minutes: Optional[int] = None) -> dict:
        app, _ = self._mapi()
        appt = app.CreateItem(c.OL_APPOINTMENT_ITEM)
        appt.Subject = subject or ""
        appt.Start = _parse_dt(start, "start")
        appt.End = _parse_dt(end, "end")
        if all_day:
            appt.AllDayEvent = True
        if body:
            appt.Body = body
        if location:
            appt.Location = location
        if reminder_minutes is not None:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = int(reminder_minutes)
        if attendees:
            appt.MeetingStatus = c.OL_MEETING
            for address in attendees:
                appt.Recipients.Add(address)
            appt.Recipients.ResolveAll()
            appt.Send()
            status = "meeting_sent"
        else:
            appt.Save()
            status = "saved"
        return {"status": status, "id": self._make_id(appt), "subject": subject}

    @_com
    def respond_to_meeting(self, event_id: str, response: str,
                           comment: Optional[str] = None,
                           send: bool = True) -> dict:
        response_key = (response or "").strip().lower()
        if response_key not in c.MEETING_RESPONSE_TO_ID:
            raise ToolError(
                f"Invalid response {response!r}: use 'accept', 'decline' or "
                f"'tentative'."
            )
        _, ns = self._mapi()
        item = self._get_item(ns, event_id)
        # A meeting request from the inbox resolves to a MeetingItem;
        # get its appointment. Calendar ids resolve straight to appointments.
        if hasattr(item, "GetAssociatedAppointment"):
            item = item.GetAssociatedAppointment(True)
        resp = item.Respond(c.MEETING_RESPONSE_TO_ID[response_key], True)
        if comment and resp is not None:
            resp.Body = comment
        if resp is not None:
            if send:
                resp.Send()
            else:
                resp.Save()
        return {"status": f"{response_key}{'_sent' if send else '_saved'}",
                "subject": getattr(item, "Subject", "") or ""}

    # ---- Attachments ------------------------------------------------

    @_com
    def list_attachments(self, email_id: str) -> list:
        _, ns = self._mapi()
        item = self._get_item(ns, email_id)
        attachments = getattr(item, "Attachments", None)
        results = []
        if attachments:
            for i in range(1, attachments.Count + 1):  # COM is 1-based
                att = attachments.Item(i)
                results.append({
                    "index": i,
                    "filename": getattr(att, "FileName", "") or "",
                    "size": getattr(att, "Size", 0),
                })
        return results

    @_com
    def save_attachments(self, email_id: str, save_dir: str,
                         attachment_names: Optional[list] = None) -> list:
        _, ns = self._mapi()
        item = self._get_item(ns, email_id)
        attachments = getattr(item, "Attachments", None)
        if not attachments or attachments.Count == 0:
            raise ToolError("This email has no attachments.")
        save_dir = os.path.abspath(os.path.expanduser(save_dir))
        os.makedirs(save_dir, exist_ok=True)
        wanted = {n.lower() for n in attachment_names} if attachment_names else None
        results = []
        for i in range(1, attachments.Count + 1):
            att = attachments.Item(i)
            filename = getattr(att, "FileName", "") or f"attachment-{i}"
            if wanted is not None and filename.lower() not in wanted:
                continue
            target = os.path.join(save_dir, _safe_filename(filename))
            try:
                att.SaveAsFile(target)
                results.append({"filename": filename, "saved_to": target,
                                "status": "saved"})
            except pywintypes.com_error as exc:
                results.append({"filename": filename, "status": "failed",
                                "error": format_com_error(exc)})
        if not results:
            raise ToolError(
                "No attachments matched attachment_names; use list_attachments "
                "to see the exact file names."
            )
        return results

    # ---- Tasks ------------------------------------------------------

    @_com
    def list_tasks(self, include_completed: bool = False) -> list:
        _, ns = self._mapi()
        items = ns.GetDefaultFolder(c.OL_FOLDER_TASKS).Items
        if not include_completed:
            items = items.Restrict("[Complete] = False")
        return [self._task_summary(item) for item in items]

    @_com
    def create_task(self, subject: str, body: Optional[str] = None,
                    due_date: Optional[str] = None,
                    importance: str = "normal") -> dict:
        importance_key = (importance or "normal").strip().lower()
        if importance_key not in c.IMPORTANCE_NAME_TO_ID:
            raise ToolError(
                f"Invalid importance {importance!r}: use 'low', 'normal' or 'high'."
            )
        app, _ = self._mapi()
        task = app.CreateItem(c.OL_TASK_ITEM)
        task.Subject = subject or ""
        if body:
            task.Body = body
        if due_date:
            task.DueDate = _parse_dt(due_date, "due_date")
        task.Importance = c.IMPORTANCE_NAME_TO_ID[importance_key]
        task.Save()
        return {"status": "created", "id": self._make_id(task), "subject": subject}

    @_com
    def complete_task(self, task_id: str) -> dict:
        _, ns = self._mapi()
        task = self._get_item(ns, task_id)
        task.MarkComplete()
        return {"status": "completed",
                "subject": getattr(task, "Subject", "") or ""}

    # ---- Notes ------------------------------------------------------

    @_com
    def list_notes(self) -> list:
        _, ns = self._mapi()
        items = ns.GetDefaultFolder(c.OL_FOLDER_NOTES).Items
        return [self._note_summary(item) for item in items]

    @_com
    def get_note(self, note_id: str) -> dict:
        _, ns = self._mapi()
        note = self._get_item(ns, note_id)
        info = self._note_summary(note)
        info["body"] = _truncate(getattr(note, "Body", "") or "")
        return info

    @_com
    def create_note(self, body: str) -> dict:
        if not body:
            raise ToolError("create_note requires a non-empty body.")
        app, _ = self._mapi()
        note = app.CreateItem(c.OL_NOTE_ITEM)
        note.Body = body
        note.Save()
        return {"status": "created", "id": self._make_id(note)}
