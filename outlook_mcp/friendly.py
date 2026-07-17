"""Convert Outlook enum integers to/from the lowercase friendly words the
MCP API exposes, so callers see "accepted" / "busy" rather than 3 / 2.
"""

from typing import Optional

from outlook_mcp import constants as c

_BUSY_STATUS_WORDS = {
    c.OL_FREE: "free",
    c.OL_TENTATIVE: "tentative",
    c.OL_OUT_OF_OFFICE: "out_of_office",
    c.OL_WORKING_ELSEWHERE: "working_elsewhere",
}

_TASK_STATUS_WORDS = {
    c.OL_TASK_IN_PROGRESS: "in_progress",
    c.OL_TASK_COMPLETE: "complete",
    c.OL_TASK_WAITING: "waiting",
    c.OL_TASK_DEFERRED: "deferred",
}

_RESPONSE_WORDS = {
    c.OL_RESPONSE_ORGANIZED: "organizer",
    c.OL_RESPONSE_TENTATIVE: "tentative",
    c.OL_RESPONSE_ACCEPTED: "accepted",
    c.OL_RESPONSE_DECLINED: "declined",
    c.OL_RESPONSE_NOT_RESPONDED: "not_responded",
}

_BUSY_STATUS_IDS = {word: v for v, word in _BUSY_STATUS_WORDS.items()}
_BUSY_STATUS_IDS["busy"] = c.OL_BUSY  # busy itself has no dedicated dict entry above (it's the default)
_TASK_STATUS_IDS = {word: v for v, word in _TASK_STATUS_WORDS.items()}
_TASK_STATUS_IDS["not_started"] = c.OL_TASK_NOT_STARTED  # not_started is the default


def importance_word(v: int) -> str:
    if v == c.OL_IMPORTANCE_LOW:
        return "low"
    if v == c.OL_IMPORTANCE_HIGH:
        return "high"
    return "normal"


def response_word(v: int) -> str:
    return _RESPONSE_WORDS.get(v, "none")


def busy_status_word(v: int) -> str:
    return _BUSY_STATUS_WORDS.get(v, "busy")


def task_status_word(v: int) -> str:
    return _TASK_STATUS_WORDS.get(v, "not_started")


def busy_status_to_id(name: str) -> Optional[int]:
    return _BUSY_STATUS_IDS.get((name or "").strip().lower())


def task_status_to_id(name: str) -> Optional[int]:
    return _TASK_STATUS_IDS.get((name or "").strip().lower())


def item_type_from_class(message_class: str) -> str:
    """Map an Outlook MessageClass to a coarse item type."""
    m = (message_class or "").upper()
    if m.startswith("IPM.SCHEDULE.MEETING"):
        return "meeting"
    if "NDR" in m:
        return "bounce"
    if m.startswith("REPORT.") and "RN" in m:
        return "read_receipt"
    if m.startswith("IPM.NOTE"):
        return "email"
    return "other"


def meeting_type_from_class(message_class: str) -> str:
    """Map a meeting-item MessageClass to a meeting type. Updates are
    delivered with the same class as requests, so they map to "request"."""
    m = (message_class or "").upper()
    if "CANCELED" in m or "CANCELLED" in m:
        return "cancellation"
    if "RESP" in m:
        return "response"
    return "request"
