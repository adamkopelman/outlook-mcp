# v2 Python Plan 1 — Foundations (friendly words + categories) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Port the same foundation the Rust v2 build started with: a friendly-word enum module (outputs say `"accepted"`/`"busy"`, not `3`/`2`) and category read/write support, into the Python implementation.

**Architecture:** A new pure module `outlook_mcp/friendly.py` converts Outlook enum integers ↔ lowercase strings (forward) and back (reverse, for later plans' inputs) — same shape as `outlook-mcp-rs/src/friendly.rs`, written idiomatically (plain functions + dict lookups, matching this codebase's existing `IMPORTANCE_NAME_TO_ID`-style constants). Category parse/join + COM read/write helpers go in `outlook_mcp/outlook/client.py` as module-level functions (this codebase has no `com.rs`-style pure/impure split file — module-level free functions in `client.py`, next to the existing `_to_iso`/`_truncate`/`_safe_filename` helpers, is the established local pattern). `categories: list[str]` is added to every summary dict; the raw-integer output fields (`response_status`, task `status`/`importance`) become friendly strings. Both implementors (`WindowsOutlookClient`, `FakeOutlookClient` in `tests/conftest.py`) are updated together, matching this codebase's own established two-implementor discipline.

**Tech Stack:** Python 3.13, pywin32 (`win32com.client`), FastMCP (`mcp.server.fastmcp`); test with `pytest`.

**Reference:** `outlook-mcp-rs/src/friendly.rs`, `outlook-mcp-rs/src/outlook/com.rs` (categories helpers), `outlook-mcp-rs/src/outlook/types.rs`/`client.rs`/`fake.rs` (Task 3/4 equivalents) — the shipped, reviewed, live-tested Rust implementation of this exact plan. Port the *behavior* (exact friendly-word vocabulary, exact enum values, exact tolerance conventions), not Rust syntax.

**Spec:** `docs/superpowers/specs/2026-07-07-outlook-mcp-v2-features-design.md` (cross-cutting principles §1–§2).

## Global Constraints

- Target repo: `C:\Users\adamk\projects\outlook-mcp` (the Python package `outlook_mcp/`) only — do not touch `outlook-mcp-rs`.
- `OutlookClientBase` (`outlook_mcp/outlook/base.py`) has TWO implementors — `WindowsOutlookClient` (`outlook_mcp/outlook/client.py`) and the test-only `FakeOutlookClient` (`tests/conftest.py`) — any output-shape change touches both plus `tests/test_tools.py`. `UnavailableClient` (also in `base.py`) needs no changes here — it inherits `OutlookClientBase`'s abstract methods and raises for all of them regardless of shape.
- Friendly words (identical vocabulary to Rust, do not invent new spellings): response → `"organizer"`/`"accepted"`/`"declined"`/`"tentative"`/`"not_responded"`/`"none"`; busy → `"free"`/`"tentative"`/`"busy"`/`"out_of_office"`/`"working_elsewhere"`; importance → `"low"`/`"normal"`/`"high"`; task status → `"not_started"`/`"in_progress"`/`"complete"`/`"waiting"`/`"deferred"`.
- Missing-property tolerance: this codebase's existing convention is `getattr(item, "X", default) or default` (see every existing `_*_summary` method in `client.py`) — new property reads must follow the same pattern, never let a missing/None property raise.
- The Outlook `Categories` property is a single string of category names joined by `", "` (comma-space); empty string = no categories. Identical to Rust's convention.
- Run `pytest` (all existing tests must stay green) before each commit.

---

### Task 1: Friendly-word conversion module

**Files:**
- Create: `outlook_mcp/friendly.py`
- Modify: `outlook_mcp/constants.py` (add the OlBusyStatus + full OlTaskStatus + OlResponseStatus enum values)
- Create: `tests/test_friendly.py`

**Interfaces:**
- Produces:
  - `friendly.importance_word(v: int) -> str`
  - `friendly.response_word(v: int) -> str`
  - `friendly.busy_status_word(v: int) -> str`
  - `friendly.task_status_word(v: int) -> str`
  - `friendly.busy_status_to_id(name: str) -> Optional[int]`
  - `friendly.task_status_to_id(name: str) -> Optional[int]`
  - (importance/response name→id already exist in `constants.py` as `IMPORTANCE_NAME_TO_ID` / `MEETING_RESPONSE_TO_ID` dicts — reuse them, do not duplicate.)

- [ ] **Step 1: Add the enum constants to `outlook_mcp/constants.py`**

Append after the existing `OL_IMPORTANCE_HIGH = 2` block:

```python
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
```

- [ ] **Step 2: Write the failing test — create `tests/test_friendly.py`**

```python
from outlook_mcp import friendly


def test_words_map_known_and_unknown_values():
    assert friendly.importance_word(2) == "high"
    assert friendly.importance_word(99) == "normal"  # unknown -> default
    assert friendly.response_word(3) == "accepted"
    assert friendly.response_word(5) == "not_responded"
    assert friendly.busy_status_word(0) == "free"
    assert friendly.busy_status_word(3) == "out_of_office"
    assert friendly.busy_status_word(99) == "busy"  # unknown -> default
    assert friendly.task_status_word(1) == "in_progress"
    assert friendly.task_status_word(99) == "not_started"


def test_reverse_lookups_are_case_insensitive_and_reject_garbage():
    assert friendly.busy_status_to_id("Out_Of_Office") == 3
    assert friendly.busy_status_to_id("nope") is None
    assert friendly.task_status_to_id("COMPLETE") == 2
    assert friendly.task_status_to_id("nope") is None
```

- [ ] **Step 3: Run it to see it fail**

Run: `pytest tests/test_friendly.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'outlook_mcp.friendly'`.

- [ ] **Step 4: Create `outlook_mcp/friendly.py`**

```python
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
```

- [ ] **Step 5: Run the tests**

Run: `pytest tests/test_friendly.py -v`
Expected: both tests PASS.

- [ ] **Step 6: Commit**

```bash
git add outlook_mcp/friendly.py outlook_mcp/constants.py tests/test_friendly.py
git commit -m "Add friendly-word enum conversion module"
```

---

### Task 2: Category parse/join helpers + COM read/write wrappers

**Files:**
- Modify: `outlook_mcp/outlook/client.py` (add pure `_parse_categories`/`_join_categories` module-level functions with tests, and COM `_get_item_categories`/`_set_item_categories` module-level wrappers)
- Create: `tests/test_client_pure.py` (this repo has no existing file for client.py's pure-logic unit tests — `tests/test_tools.py` only covers fake-backed tool-layer behavior; the Rust equivalent of this test lives inside `com.rs`'s own `#[cfg(test)]` module, colocated with the code — Python's colocated-doctest option is weaker for this kind of table-driven case, so a small dedicated test file mirrors this codebase's existing `tests/test_*.py`-per-concern layout)

**Interfaces:**
- Consumes: nothing new (pywin32's `item.Categories` is a plain string property, read/written the same way every other string property already is in this file, e.g. `item.Subject`).
- Produces:
  - `client._parse_categories(raw: str) -> list[str]` (pure)
  - `client._join_categories(cats: list[str]) -> str` (pure)
  - `client._get_item_categories(item) -> list[str]` (reads `Categories`; empty list on missing/empty)
  - `client._set_item_categories(item, cats: list[str]) -> None` (writes joined string)

- [ ] **Step 1: Write the failing pure-logic test — create `tests/test_client_pure.py`**

```python
"""Pure-logic unit tests for outlook_mcp.outlook.client's module-level
helper functions — no COM/Outlook required. Mirrors outlook-mcp-rs's
com.rs pure-function tests (parse_categories/join_categories)."""

from outlook_mcp.outlook.client import _join_categories, _parse_categories


def test_categories_round_trip_and_trim():
    assert _parse_categories("Work, Receipts") == ["Work", "Receipts"]
    assert _parse_categories("  Work ,  Personal ") == ["Work", "Personal"]
    assert _parse_categories("") == []
    assert _join_categories(["Work", "Personal"]) == "Work, Personal"
    assert _join_categories([]) == ""
```

- [ ] **Step 2: Run it to see it fail**

Run: `pytest tests/test_client_pure.py -v`
Expected: FAIL — `_parse_categories`/`_join_categories` don't exist yet.

- [ ] **Step 3: Add the pure functions to `outlook_mcp/outlook/client.py`**

Add right after the existing `_safe_filename` function:

```python
def _parse_categories(raw: str) -> list:
    """Outlook stores categories as one ", "-joined string. Split it into
    names, trimming whitespace and dropping empties."""
    return [s.strip() for s in (raw or "").split(",") if s.strip()]


def _join_categories(cats: list) -> str:
    """Join category names back into the ", "-separated string Outlook
    expects."""
    return ", ".join(cats)
```

- [ ] **Step 4: Run the pure test**

Run: `pytest tests/test_client_pure.py -v`
Expected: PASS.

- [ ] **Step 5: Add the COM wrapper functions to `outlook_mcp/outlook/client.py`**

Add right after `_parse_categories`/`_join_categories` (still module-level, not methods — these take a raw COM `item`, matching the free-function style of `_to_iso`/`_truncate`, not the `self`-taking `_*_summary` methods, since they don't need any client instance state):

```python
def _get_item_categories(item) -> list:
    """Read an item's color categories (empty list if the property is
    missing or blank)."""
    return _parse_categories(getattr(item, "Categories", "") or "")


def _set_item_categories(item, cats: list) -> None:
    """Overwrite an item's categories with the given list."""
    item.Categories = _join_categories(cats)
```

- [ ] **Step 6: Run the full suite**

Run: `pytest`
Expected: all existing tests + the 2 new ones pass. `_get_item_categories`/`_set_item_categories` are compile/import-checked here only (no live COM in this test suite) — exercised for real in Task 3 below and live-verified in this plan's final live-test pass.

- [ ] **Step 7: Commit**

```bash
git add outlook_mcp/outlook/client.py tests/test_client_pure.py
git commit -m "Add category parse/join and COM read/write helpers"
```

---

### Task 3: Add `categories` to every summary dict and populate them

**Files:**
- Modify: `outlook_mcp/outlook/client.py` (`_email_summary`, `_event_summary`, `_task_summary`, `_note_summary` each gain a `"categories"` key via `_get_item_categories`)
- Modify: `outlook_mcp/outlook/base.py` (no signature change needed — return types are already untyped `dict`/`list`, matching this codebase's convention of not declaring per-field shapes in the abstract base; skip this file for this task, note it explicitly so no one wastes time looking for a change here)
- Modify: `tests/conftest.py` (`FakeOutlookClient` — add `"categories"` to every canned summary dict it returns)
- Modify: `tests/test_tools.py` (add an assertion that categories flow through)

**Interfaces:**
- Consumes: `_get_item_categories` (Task 2).
- Produces: every summary dict now has a `"categories": [...]` key.

- [ ] **Step 1: Add `"categories"` to the four summary builders in `outlook_mcp/outlook/client.py`**

In `_email_summary`, add as the last key:

```python
            "categories": _get_item_categories(item),
```

Do the exact same (add `"categories": _get_item_categories(item),` as the last dict key) in `_event_summary`, `_task_summary`, and `_note_summary`.

- [ ] **Step 2: Update `FakeOutlookClient` in `tests/conftest.py`**

Add a `"categories"` key to every dict literal returned by `list_folders` (skip — folders have no categories), `list_emails`, `search_emails`, `get_email`, `list_events`, `get_event`, `list_tasks`, `list_notes`, `get_note` — i.e., every dict that represents a mail/event/task/note item. Give `list_emails`'s canned item a non-empty value so a test can assert it flows through end-to-end; give every other one an empty list:

In `list_emails`, change:
```python
        return [{"id": EMAIL_ID, "subject": "Hello", "sender": "Ada",
                 "unread": True}]
```
to:
```python
        return [{"id": EMAIL_ID, "subject": "Hello", "sender": "Ada",
                 "unread": True, "categories": ["Work"]}]
```

For every other item-returning method listed above, add `, "categories": []` to the returned dict (e.g. `search_emails`'s `{"id": EMAIL_ID, "subject": "Hello", "categories": []}`, `get_email`'s `{"id": email_id, "subject": "Hello", "body": "Hi there", "categories": []}`, and so on for `list_events`, `get_event`, `list_tasks`, `list_notes`, `get_note`).

- [ ] **Step 3: Add a fake-backed test in `tests/test_tools.py`**

Add near the existing email tests:

```python
def test_list_emails_returns_categories(fake_client):
    content = call_tool("list_emails", {})
    result = result_json(content)
    assert result[0]["categories"] == ["Work"]
```

- [ ] **Step 4: Run the full suite**

Run: `pytest`
Expected: all existing tests + the new one pass.

- [ ] **Step 5: Commit**

```bash
git add outlook_mcp/outlook/client.py tests/conftest.py tests/test_tools.py
git commit -m "Add categories field to all summary dicts"
```

---

### Task 4: Convert raw-number output fields to friendly words

**Files:**
- Modify: `outlook_mcp/outlook/client.py` (`get_event` produces a friendly `response` string instead of raw `response_status`; `_task_summary` produces friendly `status`/`importance` strings instead of raw numbers)
- Modify: `tests/conftest.py` (`FakeOutlookClient`'s `get_event`/`list_tasks` return friendly-word strings)
- Modify: `tests/test_tools.py` (fix any test asserting the old numeric output fields)

**Interfaces:**
- Consumes: `friendly.response_word`, `friendly.task_status_word`, `friendly.importance_word` (Task 1).
- Produces: `get_event`'s output has a `response` string key (replacing `response_status`); `_task_summary`'s `status`/`importance` keys are now friendly strings instead of raw ints.

- [ ] **Step 1: Fix `get_event` in `outlook_mcp/outlook/client.py`**

Find:
```python
        info["response_status"] = getattr(item, "ResponseStatus", None)
```

Replace with:
```python
        info["response"] = friendly.response_word(
            getattr(item, "ResponseStatus", c.OL_RESPONSE_NONE) or c.OL_RESPONSE_NONE
        )
```

Add the import at the top of the file (alongside the existing `from outlook_mcp import constants as c`):
```python
from outlook_mcp import friendly
```

- [ ] **Step 2: Fix `_task_summary` in `outlook_mcp/outlook/client.py`**

Find:
```python
    def _task_summary(self, item) -> dict:
        return {
            "id": self._make_id(item),
            "subject": getattr(item, "Subject", "") or "",
            "due_date": _to_iso(getattr(item, "DueDate", None)),
            "complete": bool(getattr(item, "Complete", False)),
            "status": getattr(item, "Status", c.OL_TASK_NOT_STARTED),
            "importance": getattr(item, "Importance", c.OL_IMPORTANCE_NORMAL),
        }
```

Replace with:
```python
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
```

(This re-states the `categories` key added in Task 3 so the whole method reads correctly in one place — if Task 3 already left it in a different position, just confirm the final method has all 6 keys: `id`, `subject`, `due_date`, `complete`, `status`, `importance`, `categories` — 7 keys total.)

- [ ] **Step 3: Fix `FakeOutlookClient` in `tests/conftest.py`**

In `get_event`, change the returned dict to include `"response": "accepted"` instead of any `response_status` key (there wasn't one before Task 3, so this is a pure addition — check the current dict literal and add the key).

In `list_tasks`, change:
```python
        return [{"id": TASK_ID, "subject": "Buy milk", "complete": False}]
```
to:
```python
        return [{"id": TASK_ID, "subject": "Buy milk", "complete": False,
                 "status": "not_started", "importance": "normal",
                 "categories": []}]
```

- [ ] **Step 4: Run the full suite, fix anything that breaks**

Run: `pytest`
Expected: all pass. If any pre-existing test in `tests/test_tools.py` asserted a raw numeric `status`/`importance`/`response_status` value from a fake response, update it to the friendly-word equivalent — search first with `grep -n "response_status\|TASK_ID" tests/test_tools.py` to find every place `get_event`'s or `list_tasks`'s output shape is asserted.

- [ ] **Step 5: Commit**

```bash
git add outlook_mcp/outlook/client.py tests/conftest.py tests/test_tools.py
git commit -m "Return friendly words for response/status/importance instead of raw enum numbers"
```

---

### Task 5: Live-verify against the real mailbox

**Files:**
- Create: `tests/test_live.py` (this repo currently has no live-Outlook test file at all — Rust's `tests/live_outlook.rs` precedent needs a Python equivalent, marked to skip by default so plain `pytest` never touches real Outlook)

**Interfaces:**
- Consumes: everything from Tasks 1-4, exercised against real COM.

- [ ] **Step 1: Create the live test scaffold**

```python
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
```

- [ ] **Step 2: Run it for real**

Confirm Outlook is running (`tasklist` should show `OUTLOOK.EXE`), then:
Run: `OUTLOOK_MCP_LIVE_TESTS=1 pytest tests/test_live.py -v`
Expected: PASS against the real mailbox. If it fails, root-cause it — do not weaken the assertion. In particular, verify the category string written via `_set_item_categories` actually round-trips through Outlook's own `Categories` property exactly as expected; pywin32 may or may not need different handling than Rust's raw `windows`-crate `VARIANT` marshalling did for this specific string property (categories are a plain `BSTR`-typed property, so this is a low-risk one, but confirm rather than assume).

- [ ] **Step 3: Confirm the plain (non-live) suite is unaffected**

Run: `pytest` (without the env var)
Expected: `tests/test_live.py`'s test is skipped (shown as `s` in pytest output), everything else passes.

- [ ] **Step 4: Clean up and commit**

Confirm no leftover `outlook-mcp P1-python live categories probe` task remains in the real Tasks folder (the test's own `finally` block should have deleted it — verify via `client.list_tasks(include_completed=True)` or a quick manual Outlook check).

```bash
git add tests/test_live.py
git commit -m "Add live test verifying categories round-trip against real Outlook"
```

---

## Self-Review

- **Spec coverage:** Cross-cutting §1 (categories first-class — visible part) ✅ Tasks 2–3; filter/settable parts land in later plans, matching the Rust build's own sequencing. §2 (friendly words) ✅ Tasks 1 & 4. Tolerance principle ✅ (`getattr(item, "X", default) or default`, this codebase's existing convention, used throughout).
- **Placeholder scan:** none — every step has concrete code or an exact command.
- **Type consistency:** `friendly.*` function names/signatures in Task 1 match their call sites in Task 4; `_get_item_categories`/`_set_item_categories` in Task 2 match the call in Task 3; every summary dict gains exactly one `"categories"` key.
- **Divergence from the Rust plan, and why:** (1) categories/COM helpers are module-level functions in `client.py` rather than a separate `com.rs`-equivalent file, matching this codebase's existing flat-file convention; (2) a new `tests/test_live.py` file had to be created from scratch (Rust already had `tests/live_outlook.rs` as an established pattern to extend; Python has no live-test file at all yet) — this plan's Task 5 establishes that pattern for all future Python plans to extend, the same role Rust's Task 5-in-Plan-9 or dedicated live-test tasks played; (3) the reverse lookup dicts in `friendly.py` (`_BUSY_STATUS_IDS`/`_TASK_STATUS_IDS`) need one explicit entry each for the "default" word (`"busy"`, `"not_started"`) since Python dict comprehensions built from the forward-mapping dict don't include values that map to a fallback — called out explicitly in Task 1's code so an implementer doesn't drop it.

## Execution Handoff

This is Python Plan 1 of 12 (parallel track to the Rust build — see `V2-RESUME-PYTHON.md`). Once complete and green (including the live test), update `V2-RESUME-PYTHON.md`'s Progress checklist and proceed to Python Plan 2 (Email finder). Execute this plan with superpowers:subagent-driven-development.
