"""Outlook object model constants, hardcoded so we never depend on
win32com.client.constants (which requires makepy/gencache type libraries).

Values come from the documented Outlook enumerations:
https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
"""

# OlDefaultFolders
OL_FOLDER_DELETED_ITEMS = 3
OL_FOLDER_OUTBOX = 4
OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_INBOX = 6
OL_FOLDER_CALENDAR = 9
OL_FOLDER_CONTACTS = 10
OL_FOLDER_JOURNAL = 11
OL_FOLDER_NOTES = 12
OL_FOLDER_TASKS = 13
OL_FOLDER_DRAFTS = 16

# OlItemType (Application.CreateItem)
OL_MAIL_ITEM = 0
OL_APPOINTMENT_ITEM = 1
OL_CONTACT_ITEM = 2
OL_TASK_ITEM = 3
OL_JOURNAL_ITEM = 4
OL_NOTE_ITEM = 5
OL_POST_ITEM = 6

# OlBodyFormat
OL_FORMAT_PLAIN = 1
OL_FORMAT_HTML = 2
OL_FORMAT_RICHTEXT = 3

# OlMeetingResponse (AppointmentItem.Respond)
OL_MEETING_TENTATIVE = 2
OL_MEETING_ACCEPTED = 3
OL_MEETING_DECLINED = 4

# OlMeetingStatus
OL_NONMEETING = 0
OL_MEETING = 1

# OlTaskStatus
OL_TASK_NOT_STARTED = 0
OL_TASK_IN_PROGRESS = 1
OL_TASK_COMPLETE = 2

# OlImportance
OL_IMPORTANCE_LOW = 0
OL_IMPORTANCE_NORMAL = 1
OL_IMPORTANCE_HIGH = 2

# OlBusyStatus (AppointmentItem.BusyStatus)
OL_FREE = 0
OL_TENTATIVE = 1
OL_BUSY = 2
OL_OUT_OF_OFFICE = 3
OL_WORKING_ELSEWHERE = 4

# OlTaskStatus (full set — OL_TASK_NOT_STARTED/IN_PROGRESS/COMPLETE already exist above)
OL_TASK_WAITING = 3
OL_TASK_DEFERRED = 4

# OlResponseStatus (AppointmentItem.ResponseStatus)
OL_RESPONSE_NONE = 0
OL_RESPONSE_ORGANIZED = 1
OL_RESPONSE_TENTATIVE = 2
OL_RESPONSE_ACCEPTED = 3
OL_RESPONSE_DECLINED = 4
OL_RESPONSE_NOT_RESPONDED = 5

# Friendly folder names accepted by tools, mapped to default folder ids.
FOLDER_NAME_TO_ID = {
    "inbox": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT_MAIL,
    "sent items": OL_FOLDER_SENT_MAIL,
    "drafts": OL_FOLDER_DRAFTS,
    "deleted": OL_FOLDER_DELETED_ITEMS,
    "deleted items": OL_FOLDER_DELETED_ITEMS,
    "trash": OL_FOLDER_DELETED_ITEMS,
    "outbox": OL_FOLDER_OUTBOX,
}

IMPORTANCE_NAME_TO_ID = {
    "low": OL_IMPORTANCE_LOW,
    "normal": OL_IMPORTANCE_NORMAL,
    "high": OL_IMPORTANCE_HIGH,
}

MEETING_RESPONSE_TO_ID = {
    "accept": OL_MEETING_ACCEPTED,
    "decline": OL_MEETING_DECLINED,
    "tentative": OL_MEETING_TENTATIVE,
}
