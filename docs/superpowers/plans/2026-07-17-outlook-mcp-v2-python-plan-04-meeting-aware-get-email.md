# v2 Python Plan 4 — Meeting-aware get_email Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development or superpowers:executing-plans. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Make `get_email` label every item's type and, when the item is a meeting (invite/update/cancellation/response), surface the embedded meeting details. Ports Rust v2 Plan 4 (`docs/superpowers/plans/2026-07-07-outlook-mcp-v2-plan-04-meeting-aware-get-email.md`, shipped) into the Python/pywin32 implementation.

**Architecture:** Add two pure mapping functions to `outlook_mcp/friendly.py`: `item_type_from_class(message_class)` (derives a coarse type — `"email"`/`"meeting"`/`"bounce"`/`"read_receipt"`/`"other"` — from the item's `MessageClass` string) and `meeting_type_from_class(message_class)` (derives `"request"`/`"cancellation"`/`"response"` for meeting items specifically). `get_email` gains three new dict keys: `item_type: str`, `is_meeting: bool`, and `meeting: dict` (present only when `is_meeting` is true — Python has no `Option`/`serde(skip_serializing_if)`, so this is a plain conditional `dict.__setitem__`, matching this codebase's existing "add key only when it applies" convention, e.g. `html_body` in `get_email` today).

**One deliberate divergence from Rust:** Rust detects meeting-ness by probing whether the COM item exposes `GetAssociatedAppointment` at all (`has_member`, a reflection-style check against the raw `IDispatch` vtable/dispinterface). pywin32's dynamic `Dispatch` objects don't support a reliable equivalent — `hasattr()` on a `win32com.client.CDispatch` object does not reflect real COM member existence the way `has_member` does at the raw-`IDispatch` level, since pywin32's dynamic dispatch resolves member IDs lazily and inconsistently for `hasattr` purposes. This codebase instead derives `is_meeting` directly from `item_type_from_class(message_class) == "meeting"` (the same `MessageClass` string Rust also parses for `item_type`, just used for one more decision) and ONLY THEN calls `GetAssociatedAppointment`, wrapped in a `try/except pywintypes.com_error` so a malformed meeting item degrades to `is_meeting=True, meeting` absent rather than failing the whole `get_email` call. This is a stronger tolerance guarantee than Rust's per-property `.unwrap_or_default()` (which assumes the method call itself succeeds) and is a closer match to this codebase's own `@_com` error-mapping conventions.

**Tech Stack:** Python 3.13, pywin32 (`win32com.client`), FastMCP.

**Depends on:** Plan 1 (Foundations, for the `friendly.py` module and its established forward-mapping style). Plans 1–3 already shipped.

## Global Constraints

- Target repo: `C:\Users\adamk\projects\outlook-mcp`.
- `get_email` is on `OutlookClientBase` (abstract, `outlook_mcp/outlook/base.py`) with no NEW parameters in this plan — only its RETURN dict shape changes. `base.py` needs no signature edit; both implementors (`WindowsOutlookClient` in `client.py`, `FakeOutlookClient` in `tests/conftest.py`) change their `get_email` body/return value, plus `outlook_mcp/server.py`'s docstring (optional, for discoverability) and `tests/test_tools.py`.
- Tolerance: every appointment-property read inside the `meeting` block uses this codebase's existing `getattr(appt, "X", default) or default` pattern (never a bare attribute access) — a malformed meeting item's individual property gaps must not raise. The `GetAssociatedAppointment` CALL ITSELF is separately wrapped in `try/except pywintypes.com_error` (see divergence note above) so a total appointment-fetch failure also can't error the whole `get_email`.
- `meeting` key is present in the returned dict ONLY when `is_meeting` is `True` — absent for normal emails (mirrors Rust's `#[serde(skip_serializing_if = "Option::is_none")]` via plain Python dict-key-omission).
- `GetAssociatedAppointment(False)` — the boolean argument is `AddToCalendar`. Pass `False` so merely READING a meeting request does NOT silently add it to the user's calendar as a side effect. This is safety-critical, not a style choice — matches Rust's Task 2 note verbatim.
- Commit after each task; `pytest` green before commit. Push to `main` at plan end (confirmed workflow for this repo).

---

### Task 1: Pure mapping helpers + fake/test plumbing

**Files:**
- Modify: `outlook_mcp/friendly.py` (add `item_type_from_class` + `meeting_type_from_class`)
- Modify: `tests/test_friendly.py` (add tests for both new functions)
- Modify: `tests/conftest.py` (`FakeOutlookClient.get_email` returns `item_type`/`is_meeting`, no `meeting` key — a plain fake email)
- Modify: `tests/test_tools.py` (assert `get_email` now returns `item_type`/`is_meeting`)

**Interfaces:**
- Produces: `friendly.item_type_from_class(message_class: str) -> str`, `friendly.meeting_type_from_class(message_class: str) -> str`. Both pure, string-in/string-out, case-insensitive on the input.

- [ ] **Step 1: Add the pure mappings to `outlook_mcp/friendly.py`**

Add at the end of the file, after `task_status_to_id`:
```python
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
```
(Note: the Rust version's bounce check has a redundant `||` — `c.contains("NDR") || c.starts_with("REPORT.") && c.contains("NDR")` — which simplifies to just `c.contains("NDR")` since the second disjunct is a strict subset of the first. The Python version above implements the simplified, logically-equivalent form: `"NDR" in m`.)

- [ ] **Step 2: Write the mapping tests in `tests/test_friendly.py`**

Add to the existing file (match its current style — read the file first to match import/assertion conventions):
```python
def test_item_type_from_class():
    assert friendly.item_type_from_class("IPM.Note") == "email"
    assert friendly.item_type_from_class("IPM.Schedule.Meeting.Request") == "meeting"
    assert friendly.item_type_from_class("IPM.Schedule.Meeting.Canceled") == "meeting"
    assert friendly.item_type_from_class("REPORT.IPM.Note.NDR") == "bounce"
    assert friendly.item_type_from_class("REPORT.IPM.Note.IPNRN") == "read_receipt"
    assert friendly.item_type_from_class("IPM.Contact") == "other"
    assert friendly.item_type_from_class("") == "other"
    assert friendly.item_type_from_class(None) == "other"


def test_meeting_type_from_class():
    assert friendly.meeting_type_from_class("IPM.Schedule.Meeting.Request") == "request"
    assert friendly.meeting_type_from_class("IPM.Schedule.Meeting.Canceled") == "cancellation"
    assert friendly.meeting_type_from_class("IPM.Schedule.Meeting.Resp.Pos") == "response"
```

- [ ] **Step 3: Run the mapping tests**

Run: `pytest tests/test_friendly.py -v` → both new tests pass alongside the existing ones.

- [ ] **Step 4: Update the fake in `tests/conftest.py`**

Current `get_email` (conftest.py:46-49):
```python
    def get_email(self, email_id, prefer_html=False):
        self._record("get_email", email_id=email_id, prefer_html=prefer_html)
        return {"id": email_id, "subject": "Hello", "body": "Hi there",
                "categories": []}
```
New (a plain, non-meeting fake email — matches how the fake already models a normal case; no fake meeting-item test is needed since that logic is real-COM-only, deferred to Task 3's live test, matching this plan's own Task-2/Task-3 split and Plan 2's established precedent of not fake-testing COM-only branching logic):
```python
    def get_email(self, email_id, prefer_html=False):
        self._record("get_email", email_id=email_id, prefer_html=prefer_html)
        return {"id": email_id, "subject": "Hello", "body": "Hi there",
                "categories": [], "item_type": "email", "is_meeting": False}
```

- [ ] **Step 5: Add a fake-backed assertion in `tests/test_tools.py`**

Extend the existing `test_get_email` (test_tools.py:76-78) or add a sibling test — match this file's `call_tool`/`result_json` idiom:
```python
def test_get_email_includes_item_type(fake_client):
    content = call_tool("get_email", {"email_id": EMAIL_ID})
    result = result_json(content)
    assert result["item_type"] == "email"
    assert result["is_meeting"] is False
```

- [ ] **Step 6: Run tests + commit**

Run: `pytest` (all green, no regressions).
```bash
git add outlook_mcp/friendly.py tests/test_friendly.py tests/conftest.py tests/test_tools.py
git commit -m "Add item_type/meeting_type class-mapping helpers and fake plumbing"
```

---

### Task 2: Real meeting detection in the Windows client

**Files:**
- Modify: `outlook_mcp/outlook/client.py` (`get_email` — read `MessageClass`, detect + build the `meeting` block)

**Interfaces:**
- Consumes: `friendly.item_type_from_class`, `friendly.meeting_type_from_class` (Task 1), `_to_iso` (existing helper, client.py:47-54), `pywintypes.com_error` (already imported).

- [ ] **Step 1: Insert the detection + meeting-block logic in `client.py::get_email`**

Current `get_email` (client.py:308-323):
```python
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
        return info
```
New — insert the meeting-detection block after `attachments` is computed, before `return info`:
```python
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
```
Note: `friendly` is already imported in `client.py` (`from outlook_mcp import friendly`, used by `busy_status_word`/`task_status_word` elsewhere). `_to_iso` and `pywintypes` are already module-level in this file — no new imports needed.

- [ ] **Step 2: Run tests**

Run: `pytest` (all green — fake-backed tests unaffected, since this only changes the real-COM path in `client.py`).

- [ ] **Step 3: Commit**

```bash
git add outlook_mcp/outlook/client.py
git commit -m "Detect meeting items in get_email and surface meeting details"
```

---

### Task 3: Live test

**Files:**
- Modify: `tests/test_live.py`

**Interfaces:**
- Consumes: real `WindowsOutlookClient.list_emails` + `get_email`.

- [ ] **Step 1: Add a live test that get_email returns a valid item_type for a real inbox item**

Add to `tests/test_live.py`, after the existing tests, using the established `client` fixture (no new skip decorator — the module-level `pytestmark` already gates the file):
```python
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
```

- [ ] **Step 2: Confirm the plain (non-live) suite still skips it**

Run: `pytest` (no env var)
Expected: this new test shows as skipped along with the rest of `test_live.py`.

- [ ] **Step 3: Run it live**

Confirm Outlook is running, then:
Run: `OUTLOOK_MCP_LIVE_TESTS=1 pytest tests/test_live.py -v`
Expected: PASS (or a clean `pytest.skip` if the inbox happens to be empty). This test makes no COM writes (pure read), so a failure here would be a real classification/detection bug, not a cleanup/side-effect problem — do not weaken the assertion. If the real inbox item under test happens to be a meeting invite, this is a good opportunity to eyeball the `meeting` block's fields for plausibility (start/end/location/organizer) — note anything odd in the report even if the assertions technically pass.

- [ ] **Step 4: Commit**

```bash
git add tests/test_live.py
git commit -m "Add live get_email item_type test"
```

---

## Self-Review

- **Spec coverage:** `item_type` ✅ (from `MessageClass`), `is_meeting` ✅, `meeting{}` block with `meeting_type`/start/end/location/organizer/attendees/is_recurring ✅ (Task 2). Tolerance ✅ (per-property `getattr(...) or default` + a `try/except` around the `GetAssociatedAppointment` call itself, which is a stronger guarantee than Rust's per-property-only tolerance, documented above as a deliberate divergence).
- **Placeholder scan:** none — full code in every step.
- **Type consistency:** `item_type_from_class`/`meeting_type_from_class` signatures (Task 1) match their call sites in `client.py` (Task 2); the `meeting` dict's keys match exactly between the plan's Task 2 code and Task 3's live-test assertions.
- **Side-effect safety:** `GetAssociatedAppointment(False)` — reading a meeting request never adds it to the calendar. Explicitly called out as safety-critical, not style.
- **Divergence from Rust documented:** the `has_member`-vs-`MessageClass`-based detection strategy difference is explained in the Architecture section with concrete reasoning (pywin32 dynamic dispatch doesn't support a reliable member-existence probe the way raw `IDispatch` does), not silently substituted.

## Execution Handoff

Python Plan 4 of 12 (parallel track — see `V2-RESUME-PYTHON.md`). Once complete and green (including the live test), update `V2-RESUME-PYTHON.md`'s Progress checklist and proceed to Python Plan 5 (update_email, absorbs move_email). Execute with superpowers:subagent-driven-development (model per task: Task 1 sonnet, Task 2 opus [COM judgment — the detection-strategy divergence from Rust needs careful implementation, matches the Rust plan's own model choice for this equivalent task], Task 3 haiku).
