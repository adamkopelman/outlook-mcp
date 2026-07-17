# v2 Python Plan 5 — update_email (absorbs move_email) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development or superpowers:executing-plans. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Replace the single-purpose `move_email` tool with one `update_email` tool that performs every non-destructive change to an existing email — move, mark read/unread, flag, add/remove categories, set importance — in a single call. Ports Rust v2 Plan 5 (`docs/superpowers/plans/2026-07-07-outlook-mcp-v2-plan-05-update-email.md`, shipped) into the Python/pywin32 implementation.

**Architecture:** `move_email` is removed from all four layers (`base.py`, `client.py`, `conftest.py`'s fake, `server.py`) and `update_email(email_id, move_to=None, mark_read=None, flag=None, add_categories=None, remove_categories=None, importance=None)` takes its place — plain kwargs, matching this codebase's established convention (no `EmailUpdate`-struct equivalent; Python has no need for Rust's params-struct indirection). `FakeOutlookClient.update_email` records the call and returns `{"status": "updated", "id": <current-or-new>, "changed": [...]}`. `WindowsOutlookClient.update_email` applies state changes first (mark_read, flag, categories, importance) and `move_to` **last**, because `Move` changes the item's EntryID — moving first would invalidate the id every subsequent step needs. The MCP tool `move_email` is retired in the same plan that adds `update_email`, so there is never a broken intermediate tool set.

**Tech Stack:** Python 3.13, pywin32 (`win32com.client`), FastMCP.

**Depends on:** Plan 1 (Foundations, for `_get_item_categories`/`_set_item_categories`). Plans 1–4 already shipped.

## Global Constraints

- Target repo: `C:\Users\adamk\projects\outlook-mcp`.
- Four layers change together: `outlook_mcp/outlook/base.py` (abstract — `move_email` removed, `update_email` added), `outlook_mcp/outlook/client.py` (real impl), `tests/conftest.py` (`FakeOutlookClient`), `outlook_mcp/server.py` (tool). Plus `tests/test_tools.py`, `tests/test_registry.py` (its `EXPECTED_TOOLS` set names `move_email`), and `README.md` (its tool table lists `move_email`).
- Tolerance: any NEW read of an existing property inside `update_email` (e.g. re-reading categories before modifying them) uses this codebase's existing `getattr(item, "X", default) or default` pattern. State *writes* (the actual mutations) may raise naturally through `@_com`'s existing `pywintypes.com_error` → `ToolError` mapping — a write failure is a real error worth surfacing, not something to swallow.
- Return shape for `update_email`: `{"status": "updated", "id": <str>, "changed": [<str>, ...]}` — `changed` lists, in APPLICATION ORDER, the fields that were touched: any of `"mark_read"`, `"flag"`, `"add_categories"`, `"remove_categories"`, `"importance"`, `"move_to"`. This exact ordering (state changes first, `move_to` last) is both the application order AND the order fields appear in `changed` — mirrors Rust's fake/real ordering parity exactly.
- `flag` accepts `"follow_up" | "complete" | "clear"`; `importance` accepts `"low" | "normal" | "high"` (reuse the existing `c.IMPORTANCE_NAME_TO_ID` dict already used by `create_task`, `client.py:598-602` — same validate-then-raise-`ToolError` pattern). An invalid value for either raises `ToolError` with a message naming the valid options.
- `add_categories`/`remove_categories` are non-destructive: read the current category set once via `_get_item_categories` (Plan 1), apply additions (case-insensitive dedup — don't add a category that's already present under different casing) and removals (case-insensitive match) against that in-memory list, then write back once via `_set_item_categories`. Never wipe existing categories the caller didn't mention.
- Commit after each task; `pytest` green before commit. Push to `main` at plan end (confirmed workflow for this repo).

---

### Task 1: Interface + fake + tool layer + unit tests

Swap `move_email` for `update_email` across `base.py`/`fake` (`conftest.py`)/`server.py`, retire the `move_email` tool, and cover it with fake-client tool tests. The real COM impl is stubbed here (`raise NotImplementedError`) and filled in Task 2 — mirrors Rust's `todo!()` handoff pattern.

**Files:**
- Modify: `outlook_mcp/outlook/base.py` (remove abstract `move_email`, lines 50-51; add abstract `update_email`)
- Modify: `tests/conftest.py` (`FakeOutlookClient` — remove `move_email`, lines 70-74; add `update_email`)
- Modify: `outlook_mcp/outlook/client.py` (remove `move_email`, lines 379-387; add `update_email` stub — real impl in Task 2)
- Modify: `outlook_mcp/server.py` (remove `move_email` tool, lines 127-132; add `update_email` tool)
- Modify: `tests/test_tools.py` (replace `test_move_email_returns_new_id`, lines 120-123, with `update_email` tests)
- Modify: `tests/test_registry.py` (`EXPECTED_TOOLS` — replace `"move_email"` with `"update_email"`, line 8)
- Modify: `README.md` (replace the `move_email` table row and the "Ids change when an item moves" paragraph's tool reference)

**Interfaces:**
- Produces: `update_email(email_id: str, move_to: Optional[str] = None, mark_read: Optional[bool] = None, flag: Optional[str] = None, add_categories: Optional[list] = None, remove_categories: Optional[list] = None, importance: Optional[str] = None) -> dict`, identical kwarg name/default across all four layers.
- Produces (fake return / real return contract, Task 2): `{"status": "updated", "id": "<current-or-new>", "changed": [...]}`.

- [ ] **Step 1: Remove `move_email` and add `update_email` to `outlook_mcp/outlook/base.py`**

Remove (base.py:50-51):
```python
    def move_email(self, email_id: str, target_folder: str) -> dict:
        raise NotImplementedError
```
Add in its place:
```python
    def update_email(self, email_id: str, move_to: Optional[str] = None,
                     mark_read: Optional[bool] = None,
                     flag: Optional[str] = None,
                     add_categories: Optional[list] = None,
                     remove_categories: Optional[list] = None,
                     importance: Optional[str] = None) -> dict:
        raise NotImplementedError
```

- [ ] **Step 2: Replace `move_email` with `update_email` in `tests/conftest.py`**

Remove (conftest.py:70-74):
```python
    def move_email(self, email_id, target_folder):
        self._record("move_email", email_id=email_id,
                     target_folder=target_folder)
        return {"status": "moved", "folder": target_folder,
                "id": "new-entry|store-1"}
```
Add in its place:
```python
    def update_email(self, email_id, move_to=None, mark_read=None, flag=None,
                     add_categories=None, remove_categories=None,
                     importance=None):
        self._record("update_email", email_id=email_id, move_to=move_to,
                     mark_read=mark_read, flag=flag,
                     add_categories=add_categories,
                     remove_categories=remove_categories,
                     importance=importance)
        # Mirror the real client's `changed` ordering: state changes first, move last.
        changed = []
        if mark_read is not None:
            changed.append("mark_read")
        if flag is not None:
            changed.append("flag")
        if add_categories is not None:
            changed.append("add_categories")
        if remove_categories is not None:
            changed.append("remove_categories")
        if importance is not None:
            changed.append("importance")
        # Move changes the EntryID; simulate a new id only when we moved.
        if move_to is not None:
            changed.append("move_to")
            new_id = "new-entry|store-1"
        else:
            new_id = email_id
        return {"status": "updated", "id": new_id, "changed": changed}
```

- [ ] **Step 3: Replace `move_email` with an `update_email` stub in `outlook_mcp/outlook/client.py`**

Remove (client.py:379-387):
```python
    @_com
    def move_email(self, email_id: str, target_folder: str) -> dict:
        _, ns = self._mapi()
        item = self._get_item(ns, email_id)
        target = self._resolve_folder(ns, target_folder)
        moved = item.Move(target)
        # EntryID changes on Move — hand back the new id.
        return {"status": "moved", "folder": target.Name,
                "id": self._make_id(moved)}
```
Add in its place:
```python
    @_com
    def update_email(self, email_id: str, move_to: Optional[str] = None,
                     mark_read: Optional[bool] = None,
                     flag: Optional[str] = None,
                     add_categories: Optional[list] = None,
                     remove_categories: Optional[list] = None,
                     importance: Optional[str] = None) -> dict:
        # Real COM implementation added in Plan 5 Task 2.
        raise NotImplementedError("update_email real COM impl — Plan 5 Task 2")
```

- [ ] **Step 4: Replace the `move_email` tool with `update_email` in `outlook_mcp/server.py`**

Remove (server.py:127-132):
```python
@mcp.tool()
def move_email(email_id: str, target_folder: str):
    """Move an email to another folder. Returns the email's NEW id (ids
    change when an item moves)."""
    return get_client().move_email(email_id=email_id,
                                   target_folder=target_folder)
```
Add in its place:
```python
@mcp.tool()
def update_email(email_id: str, move_to: Optional[str] = None,
                 mark_read: Optional[bool] = None,
                 flag: Optional[str] = None,
                 add_categories: Optional[list[str]] = None,
                 remove_categories: Optional[list[str]] = None,
                 importance: Optional[str] = None):
    """Update an existing email: move to a folder, mark read/unread, flag
    ("follow_up"/"complete"/"clear"), add/remove categories, or set
    importance ("low"/"normal"/"high"). Combine any of these in one call.
    Returns the current-or-new id — moving changes the id."""
    return get_client().update_email(
        email_id=email_id, move_to=move_to, mark_read=mark_read, flag=flag,
        add_categories=add_categories, remove_categories=remove_categories,
        importance=importance)
```

- [ ] **Step 5: Replace the tool test in `tests/test_tools.py`**

Remove (test_tools.py:120-123):
```python
def test_move_email_returns_new_id(fake_client):
    content = call_tool("move_email", {"email_id": EMAIL_ID,
                                       "target_folder": "Archive"})
    assert result_json(content)["id"] == "new-entry|store-1"
```
Add in its place:
```python
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
```

- [ ] **Step 6: Update `tests/test_registry.py`**

In `EXPECTED_TOOLS` (test_registry.py:8), replace `"move_email"` with `"update_email"`:
```python
    "send_email", "create_draft", "reply_email", "update_email", "delete_email",
```

- [ ] **Step 7: Update `README.md`**

Replace the table row (README.md:98):
```
| `move_email` | Move an email to another folder (returns its new id) |
```
with:
```
| `update_email` | Move, mark read/unread, flag, add/remove categories, or set importance on an email |
```
Replace the sentence (README.md:114):
```
**Ids change when an item moves folders** — `move_email` returns the new id,
```
with:
```
**Ids change when an item moves folders** — `update_email` returns the new id,
```

- [ ] **Step 8: Run tests + commit**

Run: `pytest` (all green, including the two new `update_email_*` tests; no regressions). The `NotImplementedError` stub in `client.py` is never reached — fake-backed unit tests use `FakeOutlookClient`, and the real path is only exercised by Task 3's live test after Task 2 fills it in.
```bash
git add outlook_mcp/outlook/base.py outlook_mcp/outlook/client.py tests/conftest.py outlook_mcp/server.py tests/test_tools.py tests/test_registry.py README.md
git commit -m "Add update_email tool and retire move_email (interface + fake + tool layer)"
```

---

### Task 2: Real COM implementation in `WindowsOutlookClient`

Fill in the stub with real COM: apply state changes first, `move_to` last, building the `changed` list.

**Files:**
- Modify: `outlook_mcp/constants.py` (add flag constants)
- Modify: `outlook_mcp/outlook/client.py` (implement `update_email`)

**Interfaces:**
- Consumes: `c.IMPORTANCE_NAME_TO_ID` (existing, `constants.py`), `_get_item_categories`/`_set_item_categories` (Plan 1, `client.py`), `self._mapi`/`self._get_item`/`self._resolve_folder`/`self._make_id` (existing `client.py` plumbing).

- [ ] **Step 1: Add flag constants to `outlook_mcp/constants.py`**

Add after the `OlImportance` block (after `OL_IMPORTANCE_HIGH = 2`, constants.py line 51):
```python
# OlFlagStatus (MailItem.FlagStatus)
OL_NO_FLAG = 0
OL_FLAG_COMPLETE = 1
OL_FLAG_MARKED = 2

# OlMarkInterval (MailItem.MarkAsTask)
OL_MARK_NO_DATE = 0
```

- [ ] **Step 2: Implement `update_email` in `outlook_mcp/outlook/client.py`**

Replace the `NotImplementedError` stub from Task 1 with:
```python
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
```
Note: `c` (the `outlook_mcp.constants` module alias), `ToolError`, `_get_item_categories`, and `_set_item_categories` are all already imported/defined in `client.py` (the last two by Plan 1, right after `_safe_filename`) — no new imports needed beyond the two constants added in Step 1. The local loop variable `c_` (inside the `remove_categories` list comprehension) is deliberately NOT named `c`, since `c` is already the module-level `constants` alias in this file's scope — shadowing it inside a comprehension would still technically work (comprehension-local scope in Python 3), but avoid the shadow anyway for readability, matching this file's existing care around the `c` alias elsewhere.

- [ ] **Step 3: Run tests + commit**

Run: `pytest` (all green — fake-backed tests unaffected, since this only changes the real-COM path in `client.py`). No new fake-backed test here — the real COM path is covered by the live test in Task 3.
```bash
git add outlook_mcp/constants.py outlook_mcp/outlook/client.py
git commit -m "Implement update_email real COM (mark_read, flag, categories, importance, move-last)"
```

---

### Task 3: Live test

An end-to-end test against real Outlook: create a draft, update several fields, verify, move it, then delete for cleanup.

**Files:**
- Modify: `tests/test_live.py`

**Interfaces:**
- Consumes: real `WindowsOutlookClient.create_draft` + `update_email` + `get_email` + `delete_email`.

**Known Outlook constraint (discovered and already worked around in the Rust build — see `outlook-mcp-rs/tests/live_outlook.rs:333-397` and `outlook-mcp-rs/TESTING.md:59-67`):** `MarkAsTask` (the COM method backing `flag="follow_up"`) is only valid on items that have been SENT or RECEIVED — Outlook rejects it on a draft with `"Draft items cannot be marked. MarkAsTask is only valid on items that have been sent or received."` A draft is the only safe, disposable target this automated test can create, so `flag` is deliberately NOT exercised by the automated live test below (verified manually instead — see Step 1a). Do not attempt to "fix" this by having the implementation detect draft status and skip the flag operation — a real received/sent email SHOULD honor `flag="follow_up"`, so silently no-op'ing it would hide a real capability. This is a test-scope limitation, not an implementation bug.

- [ ] **Step 1: Add the live test**

Add to `tests/test_live.py`, after the existing tests, using the established `client` fixture (no new skip decorator — the module-level `pytestmark` already gates the file):
```python
def test_update_email_applies_state_then_moves(client):
    # A draft is a safe, disposable target (never sent).
    created = client.create_draft(
        to=["nobody@example.invalid"],
        subject="outlook-mcp update_email live test",
        body="body",
    )
    email_id = created["id"]

    # Apply state changes only (no move yet) so we can read them back by the same id.
    # NOTE: `flag` is deliberately NOT exercised here. `MarkAsTask` (follow_up)
    # is only valid on sent/received items — Outlook rejects it on a draft
    # ("MarkAsTask is only valid on items that have been sent or received").
    # A draft is the only safe disposable target we can create here, so flag
    # is verified manually instead (see Step 1a below).
    res = client.update_email(
        email_id=email_id, mark_read=True,
        add_categories=["Work"], importance="high",
    )
    assert res["status"] == "updated"
    assert res["id"] == email_id  # no move -> id unchanged
    assert "importance" in res["changed"]
    assert "add_categories" in res["changed"]

    # Verify the category landed via the public get_email path.
    detail = client.get_email(email_id)
    assert "Work" in detail.get("categories", [])
    # mark_read(True) -> the item must now read as read (unread == False).
    assert detail["unread"] is False

    # get_email's dict (built from _email_summary) does not surface
    # importance — verify that write landed via a direct COM read instead.
    _, ns = client._mapi()
    item = client._get_item(ns, email_id)
    assert item.Importance == 2  # OL_IMPORTANCE_HIGH

    # A standalone mark_read (no other field, so nothing else Saves afterward)
    # must still persist — set it back to unread and confirm it stuck.
    unread_res = client.update_email(email_id=email_id, mark_read=False)
    assert unread_res["changed"] == ["mark_read"]
    redetail = client.get_email(email_id)
    assert redetail["unread"] is True

    # Now move it; the id must change, then delete via the new id for cleanup.
    moved = client.update_email(email_id=email_id, move_to="Deleted Items")
    assert moved["changed"] == ["move_to"]
    new_id = moved["id"]
    client.delete_email(new_id)
```
(`get_email`'s returned dict, built from `_email_summary` — `client.py:170-182` — does not include an `importance` key, unlike `_task_summary`. This was confirmed by reading `client.py` directly before writing this test, not guessed. The importance assertion therefore reads `item.Importance` straight off the COM item via `client._mapi()`/`client._get_item()`, both already-existing private methods this test can call directly since it lives in the same package's test suite and other live tests already do this, e.g. the Plan 1 categories round-trip test.)

- [ ] **Step 1a: Manually verify `flag` against a real sent/received email (not automated — record the result in your task report)**

Pick a received email in the test mailbox (`adamkopelman2@gmail.com`'s Outlook) and call `client.update_email(email_id=<that email's id>, flag="follow_up")` — e.g. via a throwaway Python REPL snippet or an ad-hoc pytest invocation, not a permanent test file — confirm a follow-up flag appears on the item in Outlook. Repeat with `flag="complete"` (flag shows complete) and `flag="clear"` (flag removed). Report the outcome (pass/fail per value) in your task report; do not skip this step even though it produces no committed test code — it's the only coverage `flag` gets, since the automated test above cannot exercise it on a draft.

- [ ] **Step 2: Confirm the plain (non-live) suite still skips it**

Run: `pytest` (no env var)
Expected: this new test shows as skipped along with the rest of `test_live.py`.

- [ ] **Step 3: Run it live**

Confirm Outlook is running, then:
Run: `OUTLOOK_MCP_LIVE_TESTS=1 pytest tests/test_live.py -v`
Expected: PASS. If the `importance`/`categories` read-back assertions fail, the property write didn't stick — investigate against real Outlook (this is the point of the live test); do not weaken the assertion to route around a real problem. Confirm no stray draft is left in Drafts or Deleted Items after the test runs (the final `delete_email` should have cleaned it up — Deleted Items retention is fine, that's the intended destination).

- [ ] **Step 4: Commit**

```bash
git add tests/test_live.py
git commit -m "Add live update_email round-trip test"
```

---

## Self-Review

- **Spec coverage:** `email_id` (required) ✅; `move_to` ✅ applied last, returns new id; `mark_read` true/false ✅ via `UnRead` inverse; `flag` follow_up/complete/clear ✅ via `MarkAsTask`/`FlagStatus`/`ClearTaskFlag`; `add_categories`/`remove_categories` non-destructive add/remove ✅ (read-modify-write, case-insensitive); `importance` low/normal/high ✅ via `c.IMPORTANCE_NAME_TO_ID`. Return shape + `changed` ordering (state first, move last) ✅. Retire `move_email` ✅ across base/fake/server/client/tests/registry/README.
- **Placeholder scan:** the only placeholder is the deliberate Task 1 → Task 2 `NotImplementedError` handoff, replaced in Task 2 Step 2. Task 3's Step 1 flags one genuine unknown (whether `_email_summary` surfaces `importance`) and gives an explicit fallback rather than guessing — this is a verification instruction, not an unresolved placeholder in the shipped code.
- **Corrected after first execution attempt:** Task 3's automated live test originally included `flag="follow_up"` in the draft-based state-change call. The first implementer dispatch hit a real Outlook COM error (`"Draft items cannot be marked. MarkAsTask is only valid on items that have been sent or received."`) — confirmed as the same constraint Rust's build already discovered and documented (`outlook-mcp-rs/tests/live_outlook.rs:333-397`, `outlook-mcp-rs/TESTING.md:59-67`). The plan was corrected to match Rust's approach: `flag` removed from the automated draft-based test, verified manually against a real received email instead (Step 1a). The `client.py` implementation itself needs no change — this is a genuine Outlook-imposed constraint on drafts, not a bug in `MarkAsTask`'s usage.
- **Type consistency:** `update_email`'s kwarg names/defaults are identical across `base.py`/`client.py`/`conftest.py`/`server.py`. `changed` ordering is identical between the Task 1 fake and the Task 2 real client: `mark_read`, `flag`, `add_categories`, `remove_categories`, `importance`, `move_to`. `c.IMPORTANCE_NAME_TO_ID` matches the existing `constants.py` dict already used by `create_task`.
- **Retirement completeness:** `move_email` removed from `base.py`, `client.py`, `conftest.py`, `server.py`, `test_tools.py`, `test_registry.py`'s `EXPECTED_TOOLS`, and `README.md` — no dangling references. Grep for `move_email` across the whole repo at the end of Task 1 to confirm (matching Plan 2's precedent for `search_emails` retirement).

## Execution Handoff

Python Plan 5 of 12 (parallel track — see `V2-RESUME-PYTHON.md`). Once complete and green (including the live test), update `V2-RESUME-PYTHON.md`'s Progress checklist and proceed to Python Plan 6 (Calendar finder). Execute with superpowers:subagent-driven-development (model per task: Task 1 sonnet [interface + fake + tool ripple across 7 files], Task 2 opus [real COM state mutation — mark_read/flag/categories/importance/move-last ordering needs careful judgment, matches the Rust plan's own model choice], Task 3 haiku [live test, code given verbatim modulo one flagged verification step]).
