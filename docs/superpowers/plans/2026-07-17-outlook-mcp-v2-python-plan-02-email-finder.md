# v2 Python Plan 2 — Email finder (merge search + filters) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Merge `search_emails` into `list_emails` and give it rich composable filters (`query`, `sender`, `category`, `received_after`/`received_before`, `since_days`, `has_attachments`, `flagged`, `high_importance`), so there is one powerful email-finder tool — matching the already-shipped Rust `list_emails`/`EmailQuery`.

**Architecture:** Replace `OutlookClientBase`'s two methods (`list_emails(folder, count, unread_only)` + `search_emails(query, folder, count, since_days)`) with a single `list_emails(folder=..., count=..., unread_only=..., query=None, sender=None, category=None, received_after=None, received_before=None, since_days=None, has_attachments=None, flagged=False, high_importance=False)` — kwargs, not a struct (this codebase's own established style: every existing method already takes plain kwargs with defaults, e.g. `create_task`, `send_email`; there is no Rust-`struct`-equivalent convention here, and introducing one now would be inconsistent with everything else in this file). `WindowsOutlookClient.list_emails` applies cheap filters as sequential COM `Restrict` calls (they AND together, same JET/DASL syntax already used by the existing `search_emails`) and the fuzzy ones (`category`, `has_attachments`) client-side while collecting summaries — identical strategy to the shipped Rust `list_emails`. `search_emails` is removed entirely: from `OutlookClientBase`, `WindowsOutlookClient`, `FakeOutlookClient`, and the `@mcp.tool()` registration in `server.py`.

**Tech Stack:** Python 3.13, pywin32, FastMCP; test with `pytest`.

**Reference:** `outlook-mcp-rs`'s shipped, reviewed, live-tested Plan 2 (`docs/superpowers/plans/2026-07-07-outlook-mcp-v2-plan-02-email-finder.md`, `src/outlook/client.rs`'s `list_emails`) — port the exact filtering behavior and DASL/JET query syntax (already proven correct against a real mailbox), adapted to Python/pywin32 method-call syntax.

**Depends on:** Python Plan 1 (Foundations) — shipped.

## Global Constraints

- Target repo: `C:\Users\adamk\projects\outlook-mcp` (`outlook_mcp/`) only.
- **Naming divergence from Rust, required by the language:** Rust's filter field is named `from` (a valid Rust field name); `from` is a reserved keyword in Python. This plan uses `sender` as the parameter/kwarg name throughout (matching this codebase's own `_email_summary`'s existing `"sender"`/`"sender_email"` output keys, which already avoid the same clash) — do not attempt to name it `from_` or anything else; `sender` is the deliberate, final choice.
- `OutlookClientBase` has TWO implementors — `WindowsOutlookClient` (`outlook_mcp/outlook/client.py`) and `FakeOutlookClient` (`tests/conftest.py`) — every signature change touches both plus `tests/test_tools.py`. `UnavailableClient` needs no change (inherits the abstract signature).
- `count` is clamped `max(1, min(int(count), MAX_EMAIL_COUNT))` — this exact clamping already exists in both `list_emails` and `search_emails`; keep it, don't relax it.
- Default folder `"inbox"`, default count `10` — these already exist as `@mcp.tool()` parameter defaults in `server.py` for `list_emails`; keep them there (not in `OutlookClientBase`/`WindowsOutlookClient`, matching this file's existing convention where defaults live at the tool-registration layer for optional-with-a-default fields, and at the client-method layer only for genuinely-always-required-with-a-sensible-default fields like `folder`/`count` themselves, which already default at BOTH layers today — preserve that duplication, it's how this codebase already works, not a new inconsistency to introduce).
- DASL single quotes are escaped by doubling (`query.replace("'", "''")`) — this exact pattern already exists in `search_emails`; reuse it verbatim for the new `sender` filter's DASL clause too.
- Fuzzy filters done client-side: `category` (case-insensitive match against `_get_item_categories(item)`, from Plan 1), `has_attachments` (compare the built summary's `has_attachments` bool).
- Run `pytest` (all existing tests green) before each commit.

---

### Task 1: Merge `search_emails` into `list_emails`, keep everything working, add filter plumbing (no new filtering logic yet)

**Files:**
- Modify: `outlook_mcp/outlook/base.py` (replace `list_emails`'s signature; remove `search_emails`)
- Modify: `tests/conftest.py` (`FakeOutlookClient`: new `list_emails` signature, remove `search_emails`)
- Modify: `outlook_mcp/server.py` (update `list_emails`'s `@mcp.tool()` signature to the full filter set; remove the `search_emails` tool)
- Modify: `outlook_mcp/outlook/client.py` (temporarily: `list_emails` accepts the new kwargs but only implements today's behavior — folder/count/unread_only/query, since `search_emails`'s query logic already exists and can move over as-is; the rest of the new filters are no-ops for now, real filtering lands in Task 2. Remove `search_emails` entirely.)
- Modify: `tests/test_tools.py` (update `list_emails`-related tests to the new signature; remove/convert the `search_emails` test)

**Interfaces:**
- Produces: `OutlookClientBase.list_emails(self, folder="inbox", count=10, unread_only=False, query=None, sender=None, category=None, received_after=None, received_before=None, since_days=None, has_attachments=None, flagged=False, high_importance=False) -> list`.

- [ ] **Step 1: Replace `list_emails` and remove `search_emails` in `outlook_mcp/outlook/base.py`**

Find:
```python
    def list_emails(self, folder: str = "inbox", count: int = 10,
                    unread_only: bool = False) -> list:
        raise NotImplementedError

    def search_emails(self, query: str, folder: str = "inbox", count: int = 10,
                      since_days: Optional[int] = None) -> list:
        raise NotImplementedError
```

Replace with:
```python
    def list_emails(self, folder: str = "inbox", count: int = 10,
                    unread_only: bool = False, query: Optional[str] = None,
                    sender: Optional[str] = None, category: Optional[str] = None,
                    received_after: Optional[str] = None,
                    received_before: Optional[str] = None,
                    since_days: Optional[int] = None,
                    has_attachments: Optional[bool] = None,
                    flagged: bool = False, high_importance: bool = False) -> list:
        raise NotImplementedError
```

- [ ] **Step 2: Update `FakeOutlookClient` in `tests/conftest.py`**

Find:
```python
    def list_emails(self, folder="inbox", count=10, unread_only=False):
        self._record("list_emails", folder=folder, count=count,
                     unread_only=unread_only)
        return [{"id": EMAIL_ID, "subject": "Hello", "sender": "Ada",
                 "unread": True, "categories": ["Work"]}]

    def search_emails(self, query, folder="inbox", count=10, since_days=None):
        self._record("search_emails", query=query, folder=folder, count=count,
                     since_days=since_days)
        return [{"id": EMAIL_ID, "subject": "Hello", "categories": []}]
```

Replace with (delete `search_emails` entirely, `list_emails` records every new kwarg so tests can assert forwarding):
```python
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
```

- [ ] **Step 3: Update `outlook_mcp/server.py`**

Find:
```python
@mcp.tool()
def list_emails(folder: str = "inbox", count: int = 10,
                unread_only: bool = False):
    """List recent emails in a folder, newest first. `folder` accepts a
    well-known name (inbox, sent, drafts, deleted, outbox) or a path like
    'Inbox/Receipts'. `count` is capped at 50."""
    return get_client().list_emails(folder=folder, count=count,
                                    unread_only=unread_only)


@mcp.tool()
def search_emails(query: str, folder: str = "inbox", count: int = 10,
                  since_days: Optional[int] = None):
    """Search emails by text across subject, sender name and body.
    Optionally limit to messages received in the last `since_days` days."""
    return get_client().search_emails(query=query, folder=folder, count=count,
                                      since_days=since_days)
```

Replace with:
```python
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
```

Delete the entire `search_emails` `@mcp.tool()` function.

- [ ] **Step 4: Temporarily adapt `outlook_mcp/outlook/client.py` (real filtering comes in Task 2)**

Find the existing `list_emails` and `search_emails` methods:
```python
    @_com
    def list_emails(self, folder: str = "inbox", count: int = 10,
                    unread_only: bool = False) -> list:
        _, ns = self._mapi()
        count = max(1, min(int(count), MAX_EMAIL_COUNT))
        items = self._resolve_folder(ns, folder).Items
        if unread_only:
            items = items.Restrict("[UnRead] = True")
        items.Sort("[ReceivedTime]", True)
        results = []
        for item in items:
            results.append(self._email_summary(item))
            if len(results) >= count:
                break
        return results

    @_com
    def search_emails(self, query: str, folder: str = "inbox", count: int = 10,
                      since_days: Optional[int] = None) -> list:
        _, ns = self._mapi()
        count = max(1, min(int(count), MAX_EMAIL_COUNT))
        q = (query or "").replace("'", "''")
        dasl = (
            '@SQL=("urn:schemas:httpmail:subject" LIKE \'%{0}%\' '
            'OR "urn:schemas:httpmail:fromname" LIKE \'%{0}%\' '
            'OR "urn:schemas:httpmail:textdescription" LIKE \'%{0}%\')'
        ).format(q)
        items = self._resolve_folder(ns, folder).Items.Restrict(dasl)
        if since_days:
            cutoff = datetime.now() - timedelta(days=int(since_days))
            items = items.Restrict(f"[ReceivedTime] >= '{_jet_dt(cutoff)}'")
        items.Sort("[ReceivedTime]", True)
        results = []
        for item in items:
            results.append(self._email_summary(item))
            if len(results) >= count:
                break
        return results
```

Replace both with a single temporary `list_emails` that folds in `query`/`since_days` (the two pieces of `search_emails` logic that already work) but leaves `sender`/`category`/`received_after`/`received_before`/`has_attachments`/`flagged`/`high_importance` unimplemented (accepted but ignored) for now — Task 2 fills those in:

```python
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
        if query:
            q = query.replace("'", "''")
            dasl = (
                '@SQL=("urn:schemas:httpmail:subject" LIKE \'%{0}%\' '
                'OR "urn:schemas:httpmail:fromname" LIKE \'%{0}%\' '
                'OR "urn:schemas:httpmail:textdescription" LIKE \'%{0}%\')'
            ).format(q)
            items = items.Restrict(dasl)
        if unread_only:
            items = items.Restrict("[UnRead] = True")
        if since_days:
            cutoff = datetime.now() - timedelta(days=int(since_days))
            items = items.Restrict(f"[ReceivedTime] >= '{_jet_dt(cutoff)}'")
        items.Sort("[ReceivedTime]", True)
        results = []
        for item in items:
            results.append(self._email_summary(item))
            if len(results) >= count:
                break
        return results
```

- [ ] **Step 5: Update `tests/test_tools.py`**

Find `test_list_emails_passes_arguments`, `test_list_emails_defaults`, `test_search_emails`, and any test using `list_emails`/`search_emails`'s old shape. Update the assertions to the new kwarg set (extra kwargs will appear as `None`/`False` in `fake_client.calls`). Delete the standalone search test or convert it into a `list_emails`-with-`query` test. Example updated defaults test:

```python
def test_list_emails_defaults(fake_client):
    call_tool("list_emails", {})
    assert fake_client.calls == [
        ("list_emails", {"folder": "inbox", "count": 10, "unread_only": False,
                         "query": None, "sender": None, "category": None,
                         "received_after": None, "received_before": None,
                         "since_days": None, "has_attachments": None,
                         "flagged": False, "high_importance": False})
    ]
```

And a new forwarding test:
```python
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
```

Search the file for `search_emails` (`grep -n search_emails tests/test_tools.py`) and remove/convert every reference — do not leave a test calling a tool that no longer exists.

- [ ] **Step 6: Run the full suite**

Run: `pytest`
Expected: all pass (the new forwarding test + updated defaults test + no remaining `search_emails` references).

- [ ] **Step 7: Commit**

```bash
git add outlook_mcp/outlook/base.py outlook_mcp/outlook/client.py outlook_mcp/server.py tests/conftest.py tests/test_tools.py
git commit -m "Merge search_emails into list_emails with the full filter set (plumbing only)"
```

---

### Task 2: Implement the real filtering in `WindowsOutlookClient`

**Files:**
- Modify: `outlook_mcp/outlook/client.py` (the `list_emails` body — apply every filter)

**Interfaces:**
- Consumes: `_get_item_categories` (Plan 1), `_jet_dt`/`_parse_dt` (already exist in this file), `_resolve_folder`, `_email_summary`.
- Produces: `list_emails` that honors every filter.

- [ ] **Step 1: Replace `list_emails`'s body with full filtering**

Cheap filters become sequential `Restrict` calls (they AND); `category`/`has_attachments` are filtered client-side while collecting summaries, stopping once `count` results are found:

```python
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
```

Note: `_parse_dt`/`_jet_dt` already exist as module-level functions in this file (used by `create_event`/`list_events`/the old `search_emails`) — confirm they're in scope, no new import needed.

- [ ] **Step 2: Run the full suite**

Run: `pytest`
Expected: all pass (fake-backed tests unaffected — this task only changes the real COM path, no fake-client test needed since the fake returns canned data regardless of filters, matching the shipped Rust build's same precedent).

- [ ] **Step 3: Commit**

```bash
git add outlook_mcp/outlook/client.py
git commit -m "Implement list_emails filtering (query, sender, dates, flagged, importance, category, attachments)"
```

---

### Task 3: Live verification

**Files:**
- Modify: `tests/test_live.py` (add a filter smoke test)

**Interfaces:**
- Consumes: the real `WindowsOutlookClient.list_emails`.

- [ ] **Step 1: Add a live filter test**

Add to `tests/test_live.py` (after the existing categories test), matching this file's established fixture/skip pattern:

```python
def test_list_emails_query_filter_narrows_results(client):
    all_results = client.list_emails(folder="inbox", count=25)
    # A query that almost certainly matches nothing should return <= all.
    filtered = client.list_emails(
        folder="inbox", count=25, query="zzqx-improbable-token-9137")
    assert len(filtered) <= len(all_results)
```

- [ ] **Step 2: Confirm the plain (non-live) suite still skips it**

Run: `pytest` (no env var)
Expected: this new test shows as skipped along with the existing live test.

- [ ] **Step 3: Run it live**

Confirm Outlook is running, then:
Run: `OUTLOOK_MCP_LIVE_TESTS=1 pytest tests/test_live.py -v`
Expected: PASS. If it fails, root-cause it (this test makes no COM writes, so a failure here would be a real filtering bug or a `_resolve_folder`/`Restrict` issue, not a cleanup/side-effect problem) — do not weaken the assertion.

- [ ] **Step 4: Commit**

```bash
git add tests/test_live.py
git commit -m "Add live filter smoke test for list_emails"
```

---

## Self-Review

- **Spec coverage:** `search_emails` merged ✅ (Task 1 removes it, folds `query`/`since_days` in); filters `sender`/`category`/`received_after`/`received_before`/`since_days`/`has_attachments`/`flagged`/`high_importance` ✅ (Task 2); `categories` already in output from Plan 1.
- **Placeholder scan:** none — full code in every step.
- **Type consistency:** the kwarg list is identical across `base.py`, `client.py` (both Task 1's placeholder and Task 2's real version), `conftest.py`'s fake, and `server.py`'s tool signature — verify this by eye across all four when reviewing, since Python has no compiler to catch a mismatched kwarg name the way Rust's struct-field-based approach would.
- **Retirement completeness:** `search_emails` removed from `base.py` (Task 1.1), `conftest.py`'s fake (Task 1.2), `server.py`'s tool (Task 1.3), `client.py` (Task 1.4), and every test reference (Task 1.5) — no dangling references. Grep for `search_emails` across the whole repo at the end of Task 1 to confirm.
- **Naming divergence from Rust:** `from` → `sender`, required by Python's reserved-keyword rules — documented in Global Constraints, not an oversight.

## Execution Handoff

Python Plan 2 of 12 (parallel track — see `V2-RESUME-PYTHON.md`). Once complete and green (including the live test), update `V2-RESUME-PYTHON.md`'s Progress checklist and proceed to Python Plan 3 (Compose attachments). Execute with superpowers:subagent-driven-development (model per task: Task 1 sonnet, Task 2 opus [COM filtering logic, matches the Rust plan's own model choice for this equivalent task], Task 3 sonnet).
