# v2 Python Plan 3 — Compose attachments Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development or superpowers:executing-plans. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Let `send_email`, `create_draft`, and `reply_email` attach local files, given file paths, validated to exist before anything is sent. Ports Rust v2 Plan 3 (`docs/superpowers/plans/2026-07-07-outlook-mcp-v2-plan-03-compose-attachments.md`, shipped) into the Python/pywin32 implementation.

**Architecture:** Add an `attachments: Optional[list] = None` param (list of local file paths) as the LAST kwarg of `send_email`, `create_draft`, `reply_email` across all four layers (`base.py`, `client.py`, `conftest.py`'s fake, `server.py`). A shared module-level `_attach_files(item, paths)` helper in `client.py` validates every path exists (fail fast — nothing sends if a path is bad) then calls `item.Attachments.Add(path)` for each — pywin32 exposes `Attachments.Add` as a plain method taking a filesystem path string, no VARIANT marshalling needed (simpler than Rust's raw `IDispatch::Invoke` call in `attach_files`). `send_email`/`create_draft` attach after `_compose`, before `Send`/`Save`; `reply_email` attaches to the reply item, after the body is set and before the `if send:` branch.

**Tech Stack:** Python 3.13, pywin32 (`win32com.client`), FastMCP.

**Depends on:** nothing (independent of Plan 2); Plans 1–2 are already shipped.

## Global Constraints

- Target repo: `C:\Users\adamk\projects\outlook-mcp`.
- Four layers change together: `outlook_mcp/outlook/base.py` (abstract signature), `outlook_mcp/outlook/client.py` (`WindowsOutlookClient` — real impl), `tests/conftest.py` (`FakeOutlookClient`), `outlook_mcp/server.py` (`@mcp.tool()` signature). This mirrors the Rust build's "trait has two implementors, plus server + tests" discipline, just with four Python layers instead of trait+2+server+tests.
- Validate ALL paths exist BEFORE attaching/sending — a missing path raises `ToolError` and nothing is sent (send is irreversible). Use `os.path.isfile(p)`, matching this codebase's existing `os.path`-based file handling in `save_attachments` (`client.py:518-528`) rather than introducing `pathlib`.
- `send_email` keeps its existing empty-`to` guard (`client.py:345-346`, unchanged).
- `attachments` defaults to `None` everywhere (not `[]`) — matches every other `Optional[list]` param already in this codebase (`cc`, `bcc`, `attachment_names`).
- Commit after each task; `pytest` green before commit. Push to `main` at plan end (confirmed workflow for this repo, same as Plan 1/2).

---

### Task 1: Thread `attachments` through base, client, fake, server, tests

**Files:**
- Modify: `outlook_mcp/outlook/base.py` (add `attachments: Optional[list] = None` to the three abstract methods, lines 35-47)
- Modify: `outlook_mcp/outlook/client.py` (add `_attach_files` helper; wire into `send_email`/`create_draft`/`reply_email`, lines 325-377)
- Modify: `tests/conftest.py` (add the param + record it in `FakeOutlookClient`'s three methods, lines 51-65)
- Modify: `outlook_mcp/server.py` (add `attachments` param to the three `@mcp.tool()` functions, lines 90-116, pass through)
- Modify: `tests/test_tools.py` (add a forwarding assertion for at least `send_email`)

**Interfaces:**
- Consumes: nothing new — extends existing `send_email`/`create_draft`/`reply_email` signatures already present in all four files.
- Produces: `send_email(to, subject, body, cc=None, bcc=None, html=False, attachments=None)`, `create_draft(to, subject, body, cc=None, bcc=None, html=False, attachments=None)`, `reply_email(email_id, body, reply_all=False, html=False, send=True, attachments=None)` — identical kwarg name/default across all four layers.
- Produces (client.py module-level): `def _attach_files(item, paths: list) -> None`, raises `ToolError` on the first missing path, placed near the other module-level helpers (right after `_compose`, before `send_email`).

- [ ] **Step 1: Change the abstract signatures in `outlook_mcp/outlook/base.py`**

Current (lines 35-47):
```python
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
```
Add `attachments: Optional[list] = None` as the last parameter of each:
```python
    def send_email(self, to: list, subject: str, body: str,
                   cc: Optional[list] = None, bcc: Optional[list] = None,
                   html: bool = False, attachments: Optional[list] = None) -> dict:
        raise NotImplementedError

    def create_draft(self, to: list, subject: str, body: str,
                     cc: Optional[list] = None, bcc: Optional[list] = None,
                     html: bool = False, attachments: Optional[list] = None) -> dict:
        raise NotImplementedError

    def reply_email(self, email_id: str, body: str, reply_all: bool = False,
                    html: bool = False, send: bool = True,
                    attachments: Optional[list] = None) -> dict:
        raise NotImplementedError
```

- [ ] **Step 2: Add the `_attach_files` helper in `outlook_mcp/outlook/client.py`**

Place it right after `_compose` (client.py:325-339), before `send_email`:
```python
def _attach_files(item, paths: list) -> None:
    """Attach local files to a mail/reply item. Validates every path exists
    FIRST (so a bad path fails before anything is sent), then adds each via
    item.Attachments.Add(path)."""
    for p in paths:
        if not os.path.isfile(p):
            raise ToolError(f"attachment not found: {p}")
    for p in paths:
        item.Attachments.Add(p)
```
(`os` is already imported at the top of `client.py` — confirm before writing; if not, add `import os` alongside the existing imports.)

- [ ] **Step 3: Wire into the three client methods**

`send_email` (client.py:341-350) — add the param, attach after `_compose` and BEFORE `mail.Send()`:
```python
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
```

`create_draft` (client.py:352-360) — add the param, attach after `_compose` and BEFORE `mail.Save()`:
```python
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
```

`reply_email` (client.py:362-377) — add the param, attach after the body is set (after the `if html: ... else: ...` block) and BEFORE the `if send:` branch:
```python
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
```

- [ ] **Step 4: Update the fake in `tests/conftest.py`**

Add `attachments=None` to the three fake methods' signatures and include it in the recorded call. Current (conftest.py:51-65):
```python
    def send_email(self, to, subject, body, cc=None, bcc=None, html=False):
        self._record("send_email", to=to, subject=subject, body=body, cc=cc,
                     bcc=bcc, html=html)
        return {"status": "sent", "to": "; ".join(to), "subject": subject}

    def create_draft(self, to, subject, body, cc=None, bcc=None, html=False):
        self._record("create_draft", to=to, subject=subject, body=body,
                     cc=cc, bcc=bcc, html=html)
        return {"status": "draft_saved", "id": EMAIL_ID, "subject": subject}

    def reply_email(self, email_id, body, reply_all=False, html=False,
                    send=True):
        self._record("reply_email", email_id=email_id, body=body,
                     reply_all=reply_all, html=html, send=send)
        return {"status": "sent" if send else "draft_saved"}
```
New:
```python
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
```

- [ ] **Step 5: Update the server in `outlook_mcp/server.py`**

Current (server.py:90-116):
```python
@mcp.tool()
def send_email(to: list[str], subject: str, body: str,
               cc: Optional[list[str]] = None,
               bcc: Optional[list[str]] = None, html: bool = False):
    """Send an email immediately as the signed-in Outlook user. Use
    create_draft instead if the user should review before sending.
    Set `html` to true if `body` is HTML."""
    return get_client().send_email(to=to, subject=subject, body=body, cc=cc,
                                   bcc=bcc, html=html)


@mcp.tool()
def create_draft(to: list[str], subject: str, body: str,
                 cc: Optional[list[str]] = None,
                 bcc: Optional[list[str]] = None, html: bool = False):
    """Compose an email and save it to Drafts without sending. Returns the
    draft's id."""
    return get_client().create_draft(to=to, subject=subject, body=body, cc=cc,
                                     bcc=bcc, html=html)


@mcp.tool()
def reply_email(email_id: str, body: str, reply_all: bool = False,
                html: bool = False, send: bool = True):
    """Reply to an email (reply-all optional). Sends immediately by
    default; pass send=false to save the reply as a draft instead."""
    return get_client().reply_email(email_id=email_id, body=body,
                                    reply_all=reply_all, html=html, send=send)
```
New:
```python
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
```

- [ ] **Step 6: Update `tests/test_tools.py`**

This file's established idiom (see `test_list_emails_forwards_query_and_filters`, lines 60-73) is: call the module-level `call_tool(name, arguments)` helper (test_tools.py:14-16, wraps `asyncio.run(server.mcp.call_tool(...))`), then assert against `fake_client.calls[0]`. Add, near the other email-tool tests:
```python
def test_send_email_forwards_attachments(fake_client):
    call_tool("send_email", {
        "to": ["a@x.com"], "subject": "Hi", "body": "yo",
        "attachments": ["C:/tmp/a.pdf", "C:/tmp/b.png"],
    })
    name, kwargs = fake_client.calls[0]
    assert kwargs["attachments"] == ["C:/tmp/a.pdf", "C:/tmp/b.png"]
```

- [ ] **Step 7: Run tests**

Run: `pytest` (all green, including the new forwarding test; no regressions).

- [ ] **Step 8: Commit**

```bash
git add outlook_mcp/outlook/base.py outlook_mcp/outlook/client.py tests/conftest.py outlook_mcp/server.py tests/test_tools.py
git commit -m "Add attachments param to send_email/create_draft/reply_email"
```

---

### Task 2: Live attachment round-trip test

**Files:**
- Modify: `tests/test_live.py`

**Interfaces:**
- Consumes: real `WindowsOutlookClient.create_draft` + `delete_email` + `send_email`.

- [ ] **Step 1: Add two live tests — a round-trip and a fail-fast case**

Add to `tests/test_live.py`, after the existing tests, matching the file's established `client` fixture pattern (no new skip decorator needed — the module-level `pytestmark` already gates the whole file):
```python
def test_create_draft_with_attachment_round_trips(client, tmp_path):
    path = tmp_path / "outlook-mcp-live-attach.txt"
    path.write_text("live attachment test")
    created = client.create_draft(
        to=["nobody@example.invalid"],
        subject="outlook-mcp attachment test",
        body="see attached",
        attachments=[str(path)],
    )
    try:
        assert created["status"] == "draft_saved"
    finally:
        client.delete_email(created["id"])


def test_send_with_missing_attachment_errors_before_sending(client):
    from outlook_mcp.errors import ToolError

    with pytest.raises(ToolError, match="attachment not found"):
        client.send_email(
            to=["nobody@example.invalid"],
            subject="should not send",
            body="body",
            attachments=["C:/definitely/does/not/exist/nope.pdf"],
        )
```
(The second test proves the fail-fast path errors before `Send` — it never actually sends because the path check fails first; safe to run against the real mailbox. `tmp_path` is pytest's built-in temp-directory fixture — no manual cleanup needed for the file itself.)

- [ ] **Step 2: Confirm the plain (non-live) suite still skips both**

Run: `pytest` (no env var)
Expected: both new tests show as skipped along with the rest of `test_live.py`.

- [ ] **Step 3: Run it live**

Confirm Outlook is running, then:
Run: `OUTLOOK_MCP_LIVE_TESTS=1 pytest tests/test_live.py -v`
Expected: PASS. If the round-trip test fails, root-cause it (this is real-COM `Attachments.Add` behavior, not a plumbing issue) — do not weaken the assertion. Confirm no stray draft is left in Drafts after the round-trip test runs (the `finally` block should have deleted it).

- [ ] **Step 4: Commit**

```bash
git add tests/test_live.py
git commit -m "Add live attachment round-trip and fail-fast tests"
```

---

## Self-Review

- **Spec coverage:** `attachments` on send/draft/reply ✅ (Task 1); paths validated before send ✅ (`_attach_files` validates all first, before any `.Add()` call); shared logic (single `_attach_files` helper) ✅.
- **Placeholder scan:** none — full code throughout, except Task 1 Step 6's test which explicitly calls out matching the file's real existing idiom rather than a blind copy (this codebase's exact `mcp.call_tool` / result-parsing pattern wasn't fully re-derived here — the implementer must read one neighboring test first, as instructed).
- **Type consistency:** `_attach_files(item, paths: list) -> None` signature matches its three call sites; `attachments: Optional[list] = None` is consistent across `base.py`/`client.py`/`conftest.py`/`server.py`.
- **Safety:** validation happens before `Send`/`Save`; a bad path raises `ToolError` with nothing sent — matches Rust's `attach_files` fail-fast design exactly.
- **Retirement/naming:** no retirement in this plan (pure addition); no reserved-keyword clashes (`attachments` is not a Python keyword).

## Execution Handoff

Python Plan 3 of 12 (parallel track — see `V2-RESUME-PYTHON.md`). Once complete and green (including the live test), update `V2-RESUME-PYTHON.md`'s Progress checklist and proceed to Python Plan 4 (Meeting-aware get_email). Execute with superpowers:subagent-driven-development (model per task: Task 1 sonnet [4-file signature threading + one new test, moderate multi-file coordination], Task 2 haiku [complete code given verbatim in the brief, mechanical transcription + live run]).
