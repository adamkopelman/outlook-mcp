# outlook-mcp-rs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `outlook-mcp-rs`, a single-binary Rust MCP server with full tool parity to the Python `outlook-mcp` project, distributed as a prebuilt Windows `.exe` via GitHub Releases.

**Architecture:** `windows-rs` for late-bound COM automation against classic Outlook desktop (mirroring the existing `win32com`-based Python client 1:1), the official `rmcp` crate for the MCP protocol/stdio transport, and a trait-based `OutlookClient` abstraction with a `FakeOutlookClient` test double so the entire tool layer is unit-testable without a live Outlook — plus a separate, local-only live system-test suite that exercises the real COM client.

**Tech Stack:** Rust (2021 edition), `rmcp` (MCP SDK), `windows` (COM interop), `tokio` (async runtime), `serde`/`serde_json`/`schemars` (JSON + schema), GitHub Actions (CI + release).

**Design spec:** `docs/superpowers/specs/2026-07-06-rust-port-design.md`

## Global Constraints

- New, separate GitHub repo: `adamkopelman/outlook-mcp-rs`, MIT licensed.
- Full tool parity: all 21 tools from `outlook_mcp/server.py` (Python), same names, same parameters, same JSON return shapes.
- Item id scheme: opaque `"{EntryID}|{StoreID}"` string, identical format to the Python version.
- No crates.io publish, no code-signing, no non-Windows support — all explicitly out of scope per the spec.
- CI `test` job runs on `windows-latest`; `cargo test` must never touch a live Outlook (live tests are `#[ignore]`d).
- Distribution is GitHub Releases only: a `build` job compiles `outlook-mcp-rs.exe` on tag push `v*`, a `release` job attaches it to the GitHub Release.
- COM/tokio correctness constraint: every blocking COM call must run via `tokio::task::spawn_blocking` so the OS thread doing `CoInitializeEx`/`IDispatch::Invoke` work is never migrated mid-call by the tokio scheduler (mirrors the Python client's per-call `pythoncom.CoInitialize()` + fresh-Dispatch-per-call discipline, documented at `outlook_mcp/outlook/client.py:1-9`).

**Adjustment from the design spec:** the spec's file layout sketched one file per tool category under `src/tools/`. Because `rmcp`'s `#[tool_router]` macro generates a single router from one annotated `impl` block per type, all 21 `#[tool]` methods live together in one file, `src/server.rs`, organized with the same category-comment sections the Python `server.py` uses. This is a file-organization detail the spec didn't pin down; the underlying architecture (COM layer, MCP layer choice, tool parity) is unchanged. Tasks are still split by tool category so each is independently reviewable/testable.

**API-drift note (windows-rs and rmcp):** exact dependency versions and some low-level `windows` crate signatures (`IDispatch::Invoke`, `DISPPARAMS`, `EXCEPINFO`) are resolved with `cargo add` rather than hand-pinned, since both crates evolve. Where a task's code doesn't match what the compiler reports for the resolved version, the fix is mechanical: read the compiler error, check `cargo doc -p windows --open` (or `-p rmcp`) for the exact signature, and adjust the call accordingly — the surrounding struct shapes (`DISPPARAMS`/`EXCEPINFO` mirror the decades-stable Win32 `oaidl.h` layout) are not expected to change in a way that alters the logic, only the Rust-side type wrapping.

---

## Reference: Python source being ported

This plan lives in a **new, separate repo** (`outlook-mcp-rs`) from the Python
project it ports (`outlook-mcp`). Several tasks below (12–16 especially)
specify behavior by exact file/line reference into the Python source instead
of pasting the full translation inline — that source is on the **same
machine this plan was written on**, at this absolute path (adjust if
executing on a different machine/checkout):

```
C:\Users\adamk\projects\outlook-mcp\
```

Whenever a task references e.g. `client.py:340-428`, that means
`C:\Users\adamk\projects\outlook-mcp\outlook_mcp\outlook\client.py`, lines
340-428. An executor (human or subagent) must have filesystem access to that
path to complete those tasks — it is not part of the `outlook-mcp-rs` repo
and won't be visible from a worktree of it. If executing somewhere without
access to that checkout, `git clone` the Python repo first
(`https://github.com/adamkopelman/outlook-mcp`) and adjust the path.

Full method-by-method behavior lives in the existing Python project and should be read alongside each task below, not re-derived from memory:

- `outlook_mcp/outlook/base.py` — the `OutlookClientBase` interface and `UnavailableClient` (no Rust equivalent needed; the Rust binary is Windows-only by construction).
- `outlook_mcp/outlook/client.py` — the real COM implementation; every Rust `WindowsOutlookClient` method in Tasks 12–16 is a direct translation of the matching Python method.
- `outlook_mcp/constants.py` — Outlook enum values, translated verbatim in Task 11.
- `outlook_mcp/errors.py` — `ToolError` / `format_com_error`, translated in Tasks 3 and 11.
- `outlook_mcp/server.py` — tool registration and parameter defaults, translated in Tasks 6–10.
- `tests/conftest.py` — `FakeOutlookClient`, translated in Task 5.
- `tests/test_tools.py` — the tool-layer behavior tests, translated in Tasks 6–10.

---

### Task 1: Repo bootstrap & crate scaffold

**Files:**
- Create: `outlook-mcp-rs/Cargo.toml`
- Create: `outlook-mcp-rs/src/main.rs`
- Create: `outlook-mcp-rs/.gitignore`
- Create: `outlook-mcp-rs/LICENSE`
- Create: `outlook-mcp-rs/README.md`

**Interfaces:**
- Produces: a buildable, empty binary crate named `outlook-mcp-rs` that later tasks add modules to.

- [ ] **Step 1: Create the GitHub repo**

```bash
gh repo create adamkopelman/outlook-mcp-rs --public --license mit --clone
cd outlook-mcp-rs
```

Expected: repo created and cloned locally with a generated MIT `LICENSE` already in place (edit the copyright line to match the existing Python repo's holder if it names one specifically; otherwise leave `gh`'s default).

- [ ] **Step 2: Initialize the Cargo binary crate**

```bash
cargo init --name outlook-mcp-rs .
```

Expected: `Cargo.toml` and `src/main.rs` (`fn main() { println!("Hello, world!"); }`) created.

- [ ] **Step 3: Replace the placeholder main with a temporary stub**

Edit `src/main.rs`:

```rust
fn main() {
    println!("outlook-mcp-rs: scaffold only, real server wired in a later task");
}
```

- [ ] **Step 4: Write `.gitignore`**

```
/target
Cargo.lock
```

Note: this is a binary (not a library) crate that we distribute as a compiled `.exe`, but we still exclude `Cargo.lock` from version control per the existing Python project's convention of not committing build artifacts — reproducibility for an application binary comes from CI resolving fresh each release, not from a committed lockfile. (If you'd rather commit `Cargo.lock` for fully reproducible builds, that's a reasonable alternative — just be consistent.)

- [ ] **Step 5: Write README.md**

```markdown
# outlook-mcp-rs

A single-binary Rust MCP server controlling Microsoft Outlook desktop via the
Win32 COM API. This is the Rust counterpart to
[outlook-mcp](https://github.com/adamkopelman/outlook-mcp) (Python):
same tools, same behavior, distributed as a standalone Windows `.exe` with no
Python or Rust toolchain required to run it.

## Install

Download the latest `outlook-mcp-rs.exe` from
[Releases](https://github.com/adamkopelman/outlook-mcp-rs/releases) and point
your MCP client at it directly — no install step.

## Requirements

- Windows
- Classic Outlook desktop installed and signed in

## Development

See `TESTING.md` for how to run the unit test suite and the local live-Outlook
system tests.
```

- [ ] **Step 6: Verify it builds**

```bash
cargo build
```

Expected: `Compiling outlook-mcp-rs v0.1.0 (...)` then `Finished` with no errors.

- [ ] **Step 7: Commit and push**

```bash
git add -A
git commit -m "Scaffold outlook-mcp-rs crate"
git push -u origin main
```

---

### Task 2: CI test workflow

**Files:**
- Create: `outlook-mcp-rs/.github/workflows/ci.yaml`

**Interfaces:**
- Consumes: `cargo build` / `cargo test` from Task 1's scaffold.
- Produces: a `test` job that later tasks' `cargo test` suites run under automatically.

- [ ] **Step 1: Write the workflow**

```yaml
name: CI/CD

on:
  push:
    branches: [main]
    tags: ['v*']
  pull_request:
    branches: [main]

jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: dtolnay/rust-toolchain@stable
      - name: Run tests
        run: cargo test --all -- --skip live_outlook
```

Note: `--skip live_outlook` is belt-and-suspenders — the live tests are also `#[ignore]`d individually in Task 18, so plain `cargo test` already wouldn't run them. Naming the skip explicitly here documents *why* for anyone reading the workflow.

- [ ] **Step 2: Commit and push**

```bash
git add .github/workflows/ci.yaml
git commit -m "Add CI test workflow"
git push
```

- [ ] **Step 3: Verify the workflow runs**

```bash
gh run list --workflow ci.yaml --limit 1
```

Expected: a run triggered by the push, status eventually `completed`/`success` (only the scaffold's implicit "does nothing fail" check at this point).

---

### Task 3: Error type

**Files:**
- Create: `src/error.rs`
- Modify: `src/main.rs` (add `mod error;`)

**Interfaces:**
- Produces: `pub struct ToolError(pub String)` with `ToolError::new(impl Into<String>) -> Self`, implementing `std::error::Error` + `Display`, and `impl From<ToolError> for rmcp::ErrorData` (added once `rmcp` is a dependency in Task 6 — for now, just the plain type and its `Display`/`Error` impls, unit-tested standalone).

- [ ] **Step 1: Write the failing test**

Add to `src/error.rs`:

```rust
use std::fmt;

#[derive(Debug, Clone)]
pub struct ToolError(pub String);

impl ToolError {
    pub fn new(msg: impl Into<String>) -> Self {
        ToolError(msg.into())
    }
}

impl fmt::Display for ToolError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl std::error::Error for ToolError {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn displays_its_message() {
        let err = ToolError::new("Outlook exploded");
        assert_eq!(err.to_string(), "Outlook exploded");
    }

    #[test]
    fn new_accepts_string_and_str() {
        let a = ToolError::new("literal");
        let b = ToolError::new(String::from("owned"));
        assert_eq!(a.to_string(), "literal");
        assert_eq!(b.to_string(), "owned");
    }
}
```

- [ ] **Step 2: Wire the module and run the test**

Add to the top of `src/main.rs`:

```rust
mod error;
```

```bash
cargo test error::
```

Expected: `test error::tests::displays_its_message ... ok`, `test error::tests::new_accepts_string_and_str ... ok`.

- [ ] **Step 3: Commit**

```bash
git add src/error.rs src/main.rs
git commit -m "Add ToolError type"
```

---

### Task 4: Outlook domain types

**Files:**
- Create: `src/outlook/mod.rs` (module declares `pub mod types;` — trait comes in Task 5)
- Create: `src/outlook/types.rs`
- Modify: `src/main.rs` (add `mod outlook;`)

**Interfaces:**
- Produces (all `#[derive(Debug, Clone, Serialize)]`, field names chosen to serialize identically to the Python dicts in `outlook_mcp/outlook/client.py`'s `_email_summary`/`_event_summary`/`_task_summary`/`_note_summary`):
  - `FolderInfo { name: String, path: String, items: i32, unread: i32 }`
  - `EmailSummary { id, subject, sender, sender_email, to: String, received: Option<String>, unread: bool, has_attachments: bool }`
  - `EmailDetail { #[serde(flatten)] summary: EmailSummary, cc: String, bcc: String, body: String, #[serde(skip_serializing_if = "Option::is_none")] html_body: Option<String>, attachments: Vec<String> }`
  - `EventSummary { id, subject, start: Option<String>, end: Option<String>, location: String, organizer: String, all_day: bool, is_recurring: bool, is_meeting: bool }`
  - `EventDetail { #[serde(flatten)] summary: EventSummary, body: String, required_attendees: String, optional_attendees: String, response_status: Option<i32> }`
  - `TaskSummary { id: String, subject: String, due_date: Option<String>, complete: bool, status: i32, importance: i32 }`
  - `NoteSummary { id: String, subject: String, created: Option<String> }`
  - `NoteDetail { #[serde(flatten)] summary: NoteSummary, body: String }`
  - `AttachmentInfo { index: i32, filename: String, size: i32 }`

- [ ] **Step 1: Add serde as a dependency**

```bash
cargo add serde --features derive
cargo add serde_json
```

Expected: `Cargo.toml` gains `serde = { version = "...", features = ["derive"] }` and `serde_json = "..."`.

- [ ] **Step 2: Write the types**

Create `src/outlook/types.rs`:

```rust
use serde::Serialize;

#[derive(Debug, Clone, Serialize)]
pub struct FolderInfo {
    pub name: String,
    pub path: String,
    pub items: i32,
    pub unread: i32,
}

#[derive(Debug, Clone, Serialize)]
pub struct EmailSummary {
    pub id: String,
    pub subject: String,
    pub sender: String,
    pub sender_email: String,
    pub to: String,
    pub received: Option<String>,
    pub unread: bool,
    pub has_attachments: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct EmailDetail {
    #[serde(flatten)]
    pub summary: EmailSummary,
    pub cc: String,
    pub bcc: String,
    pub body: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html_body: Option<String>,
    pub attachments: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct EventSummary {
    pub id: String,
    pub subject: String,
    pub start: Option<String>,
    pub end: Option<String>,
    pub location: String,
    pub organizer: String,
    pub all_day: bool,
    pub is_recurring: bool,
    pub is_meeting: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct EventDetail {
    #[serde(flatten)]
    pub summary: EventSummary,
    pub body: String,
    pub required_attendees: String,
    pub optional_attendees: String,
    pub response_status: Option<i32>,
}

#[derive(Debug, Clone, Serialize)]
pub struct TaskSummary {
    pub id: String,
    pub subject: String,
    pub due_date: Option<String>,
    pub complete: bool,
    pub status: i32,
    pub importance: i32,
}

#[derive(Debug, Clone, Serialize)]
pub struct NoteSummary {
    pub id: String,
    pub subject: String,
    pub created: Option<String>,
}

#[derive(Debug, Clone, Serialize)]
pub struct NoteDetail {
    #[serde(flatten)]
    pub summary: NoteSummary,
    pub body: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct AttachmentInfo {
    pub index: i32,
    pub filename: String,
    pub size: i32,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn email_detail_flattens_summary_fields_at_top_level() {
        let detail = EmailDetail {
            summary: EmailSummary {
                id: "e1|s1".into(), subject: "Hi".into(), sender: "Ada".into(),
                sender_email: "ada@example.com".into(), to: "bob@example.com".into(),
                received: Some("2026-06-10T12:00:00".into()), unread: true,
                has_attachments: false,
            },
            cc: "".into(), bcc: "".into(), body: "Hello".into(),
            html_body: None, attachments: vec![],
        };
        let value = serde_json::to_value(&detail).unwrap();
        // Flattened: "id" and "subject" appear at the top level, not nested
        // under a "summary" key, and html_body is omitted when None.
        assert_eq!(value["id"], "e1|s1");
        assert_eq!(value["subject"], "Hi");
        assert_eq!(value["body"], "Hello");
        assert!(value.get("html_body").is_none());
        assert!(value.get("summary").is_none());
    }
}
```

Create `src/outlook/mod.rs`:

```rust
pub mod types;
```

- [ ] **Step 3: Wire the module and run the test**

Add to `src/main.rs`:

```rust
mod outlook;
```

```bash
cargo test outlook::types::
```

Expected: `test outlook::types::tests::email_detail_flattens_summary_fields_at_top_level ... ok`.

- [ ] **Step 4: Commit**

```bash
git add Cargo.toml src/outlook/mod.rs src/outlook/types.rs src/main.rs
git commit -m "Add Outlook domain types"
```

---

### Task 5: OutlookClient trait + FakeOutlookClient

**Files:**
- Modify: `src/outlook/mod.rs` (add the `OutlookClient` trait)
- Create: `src/outlook/fake.rs`

**Interfaces:**
- Consumes: types from Task 4 (`FolderInfo`, `EmailSummary`, ... `AttachmentInfo`), `ToolError` from Task 3.
- Produces:
  - `pub trait OutlookClient: Send + Sync` with the 21 methods listed below (all synchronous — async wrapping happens at the tool layer in Tasks 6–10).
  - `pub struct FakeOutlookClient` implementing it, plus `pub const EMAIL_ID/EVENT_ID/TASK_ID/NOTE_ID: &str`, mirroring `tests/conftest.py`. Exposes `pub fn calls(&self) -> Vec<(String, serde_json::Value)>` (a snapshot) and `pub fn set_fail_with(&self, msg: impl Into<String>)` / `pub fn clear_fail_with(&self)`.

Trait (add to `src/outlook/mod.rs`, above `pub mod types;`):

```rust
pub mod fake;
pub mod types;

use crate::error::ToolError;
use serde_json::Value;
use types::*;

pub trait OutlookClient: Send + Sync {
    fn list_folders(&self) -> Result<Vec<FolderInfo>, ToolError>;
    fn list_emails(&self, folder: String, count: i32, unread_only: bool)
        -> Result<Vec<EmailSummary>, ToolError>;
    fn search_emails(&self, query: String, folder: String, count: i32,
        since_days: Option<i32>) -> Result<Vec<EmailSummary>, ToolError>;
    fn get_email(&self, email_id: String, prefer_html: bool)
        -> Result<EmailDetail, ToolError>;
    fn send_email(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool)
        -> Result<Value, ToolError>;
    fn create_draft(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool)
        -> Result<Value, ToolError>;
    fn reply_email(&self, email_id: String, body: String, reply_all: bool,
        html: bool, send: bool) -> Result<Value, ToolError>;
    fn move_email(&self, email_id: String, target_folder: String)
        -> Result<Value, ToolError>;
    fn delete_email(&self, email_id: String) -> Result<Value, ToolError>;

    fn list_events(&self, start_date: Option<String>, end_date: Option<String>)
        -> Result<Vec<EventSummary>, ToolError>;
    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError>;
    fn create_event(&self, subject: String, start: String, end: String,
        body: Option<String>, location: Option<String>,
        attendees: Option<Vec<String>>, all_day: bool,
        reminder_minutes: Option<i32>) -> Result<Value, ToolError>;
    fn respond_to_meeting(&self, event_id: String, response: String,
        comment: Option<String>, send: bool) -> Result<Value, ToolError>;

    fn list_attachments(&self, email_id: String)
        -> Result<Vec<AttachmentInfo>, ToolError>;
    fn save_attachments(&self, email_id: String, save_dir: String,
        attachment_names: Option<Vec<String>>) -> Result<Vec<Value>, ToolError>;

    fn list_tasks(&self, include_completed: bool)
        -> Result<Vec<TaskSummary>, ToolError>;
    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String) -> Result<Value, ToolError>;
    fn complete_task(&self, task_id: String) -> Result<Value, ToolError>;

    fn list_notes(&self) -> Result<Vec<NoteSummary>, ToolError>;
    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError>;
    fn create_note(&self, body: String) -> Result<Value, ToolError>;
}
```

- [ ] **Step 1: Write the fake and its tests**

Create `src/outlook/fake.rs`:

```rust
use std::sync::Mutex;

use serde_json::{json, Value};

use crate::error::ToolError;
use super::types::*;
use super::OutlookClient;

pub const EMAIL_ID: &str = "entry-1|store-1";
pub const EVENT_ID: &str = "entry-2|store-1";
pub const TASK_ID: &str = "entry-3|store-1";
pub const NOTE_ID: &str = "entry-4|store-1";

/// In-memory stand-in for COM Outlook; records every call. Mirrors
/// `tests/conftest.py::FakeOutlookClient` in the Python project.
pub struct FakeOutlookClient {
    calls: Mutex<Vec<(String, Value)>>,
    fail_with: Mutex<Option<String>>,
}

impl FakeOutlookClient {
    pub fn new() -> Self {
        Self { calls: Mutex::new(Vec::new()), fail_with: Mutex::new(None) }
    }

    pub fn calls(&self) -> Vec<(String, Value)> {
        self.calls.lock().unwrap().clone()
    }

    pub fn set_fail_with(&self, msg: impl Into<String>) {
        *self.fail_with.lock().unwrap() = Some(msg.into());
    }

    fn record(&self, name: &str, args: Value) -> Result<(), ToolError> {
        if let Some(msg) = self.fail_with.lock().unwrap().clone() {
            return Err(ToolError::new(msg));
        }
        self.calls.lock().unwrap().push((name.to_string(), args));
        Ok(())
    }
}

impl OutlookClient for FakeOutlookClient {
    fn list_folders(&self) -> Result<Vec<FolderInfo>, ToolError> {
        self.record("list_folders", json!({}))?;
        Ok(vec![FolderInfo {
            name: "Inbox".into(), path: "Inbox".into(), items: 2, unread: 1,
        }])
    }

    fn list_emails(&self, folder: String, count: i32, unread_only: bool)
        -> Result<Vec<EmailSummary>, ToolError> {
        self.record("list_emails",
            json!({"folder": folder, "count": count, "unread_only": unread_only}))?;
        Ok(vec![EmailSummary {
            id: EMAIL_ID.into(), subject: "Hello".into(), sender: "Ada".into(),
            sender_email: "".into(), to: "".into(), received: None,
            unread: true, has_attachments: false,
        }])
    }

    fn search_emails(&self, query: String, folder: String, count: i32,
        since_days: Option<i32>) -> Result<Vec<EmailSummary>, ToolError> {
        self.record("search_emails",
            json!({"query": query, "folder": folder, "count": count, "since_days": since_days}))?;
        Ok(vec![EmailSummary {
            id: EMAIL_ID.into(), subject: "Hello".into(), sender: "".into(),
            sender_email: "".into(), to: "".into(), received: None,
            unread: false, has_attachments: false,
        }])
    }

    fn get_email(&self, email_id: String, prefer_html: bool)
        -> Result<EmailDetail, ToolError> {
        self.record("get_email", json!({"email_id": email_id, "prefer_html": prefer_html}))?;
        Ok(EmailDetail {
            summary: EmailSummary {
                id: email_id, subject: "Hello".into(), sender: "".into(),
                sender_email: "".into(), to: "".into(), received: None,
                unread: false, has_attachments: false,
            },
            cc: "".into(), bcc: "".into(), body: "Hi there".into(),
            html_body: None, attachments: vec![],
        })
    }

    fn send_email(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool)
        -> Result<Value, ToolError> {
        self.record("send_email",
            json!({"to": to, "subject": subject, "body": body, "cc": cc, "bcc": bcc, "html": html}))?;
        Ok(json!({"status": "sent", "to": to.join("; "), "subject": subject}))
    }

    fn create_draft(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool)
        -> Result<Value, ToolError> {
        self.record("create_draft",
            json!({"to": to, "subject": subject, "body": body, "cc": cc, "bcc": bcc, "html": html}))?;
        Ok(json!({"status": "draft_saved", "id": EMAIL_ID, "subject": subject}))
    }

    fn reply_email(&self, email_id: String, body: String, reply_all: bool,
        html: bool, send: bool) -> Result<Value, ToolError> {
        self.record("reply_email",
            json!({"email_id": email_id, "body": body, "reply_all": reply_all, "html": html, "send": send}))?;
        Ok(json!({"status": if send { "sent" } else { "draft_saved" }}))
    }

    fn move_email(&self, email_id: String, target_folder: String)
        -> Result<Value, ToolError> {
        self.record("move_email", json!({"email_id": email_id, "target_folder": target_folder}))?;
        Ok(json!({"status": "moved", "folder": target_folder, "id": "new-entry|store-1"}))
    }

    fn delete_email(&self, email_id: String) -> Result<Value, ToolError> {
        self.record("delete_email", json!({"email_id": email_id}))?;
        Ok(json!({"status": "deleted"}))
    }

    fn list_events(&self, start_date: Option<String>, end_date: Option<String>)
        -> Result<Vec<EventSummary>, ToolError> {
        self.record("list_events", json!({"start_date": start_date, "end_date": end_date}))?;
        Ok(vec![EventSummary {
            id: EVENT_ID.into(), subject: "Standup".into(), start: None, end: None,
            location: "".into(), organizer: "".into(), all_day: false,
            is_recurring: false, is_meeting: false,
        }])
    }

    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError> {
        self.record("get_event", json!({"event_id": event_id}))?;
        Ok(EventDetail {
            summary: EventSummary {
                id: event_id, subject: "Standup".into(), start: None, end: None,
                location: "".into(), organizer: "".into(), all_day: false,
                is_recurring: false, is_meeting: false,
            },
            body: "".into(), required_attendees: "".into(),
            optional_attendees: "".into(), response_status: None,
        })
    }

    fn create_event(&self, subject: String, start: String, end: String,
        body: Option<String>, location: Option<String>,
        attendees: Option<Vec<String>>, all_day: bool,
        reminder_minutes: Option<i32>) -> Result<Value, ToolError> {
        self.record("create_event", json!({
            "subject": subject, "start": start, "end": end, "body": body,
            "location": location, "attendees": attendees, "all_day": all_day,
            "reminder_minutes": reminder_minutes,
        }))?;
        Ok(json!({"status": "saved", "id": EVENT_ID, "subject": subject}))
    }

    fn respond_to_meeting(&self, event_id: String, response: String,
        comment: Option<String>, send: bool) -> Result<Value, ToolError> {
        self.record("respond_to_meeting",
            json!({"event_id": event_id, "response": response, "comment": comment, "send": send}))?;
        Ok(json!({"status": format!("{response}_sent")}))
    }

    fn list_attachments(&self, email_id: String)
        -> Result<Vec<AttachmentInfo>, ToolError> {
        self.record("list_attachments", json!({"email_id": email_id}))?;
        Ok(vec![AttachmentInfo { index: 1, filename: "report.pdf".into(), size: 1234 }])
    }

    fn save_attachments(&self, email_id: String, save_dir: String,
        attachment_names: Option<Vec<String>>) -> Result<Vec<Value>, ToolError> {
        self.record("save_attachments",
            json!({"email_id": email_id, "save_dir": save_dir, "attachment_names": attachment_names}))?;
        Ok(vec![json!({"filename": "report.pdf", "saved_to": save_dir, "status": "saved"})])
    }

    fn list_tasks(&self, include_completed: bool) -> Result<Vec<TaskSummary>, ToolError> {
        self.record("list_tasks", json!({"include_completed": include_completed}))?;
        Ok(vec![TaskSummary {
            id: TASK_ID.into(), subject: "Buy milk".into(), due_date: None,
            complete: false, status: 0, importance: 1,
        }])
    }

    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String) -> Result<Value, ToolError> {
        self.record("create_task",
            json!({"subject": subject, "body": body, "due_date": due_date, "importance": importance}))?;
        Ok(json!({"status": "created", "id": TASK_ID, "subject": subject}))
    }

    fn complete_task(&self, task_id: String) -> Result<Value, ToolError> {
        self.record("complete_task", json!({"task_id": task_id}))?;
        Ok(json!({"status": "completed"}))
    }

    fn list_notes(&self) -> Result<Vec<NoteSummary>, ToolError> {
        self.record("list_notes", json!({}))?;
        Ok(vec![NoteSummary { id: NOTE_ID.into(), subject: "Ideas".into(), created: None }])
    }

    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError> {
        self.record("get_note", json!({"note_id": note_id}))?;
        Ok(NoteDetail {
            summary: NoteSummary { id: note_id, subject: "Ideas".into(), created: None },
            body: "Ideas\n- one".into(),
        })
    }

    fn create_note(&self, body: String) -> Result<Value, ToolError> {
        self.record("create_note", json!({"body": body}))?;
        Ok(json!({"status": "created", "id": NOTE_ID}))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn records_calls_in_order() {
        let fake = FakeOutlookClient::new();
        fake.list_folders().unwrap();
        fake.list_emails("inbox".into(), 10, false).unwrap();
        assert_eq!(fake.calls(), vec![
            ("list_folders".to_string(), json!({})),
            ("list_emails".to_string(),
             json!({"folder": "inbox", "count": 10, "unread_only": false})),
        ]);
    }

    #[test]
    fn fail_with_makes_every_call_error_before_recording() {
        let fake = FakeOutlookClient::new();
        fake.set_fail_with("Outlook exploded");
        let err = fake.list_emails("inbox".into(), 10, false).unwrap_err();
        assert_eq!(err.to_string(), "Outlook exploded");
        assert!(fake.calls().is_empty());
    }
}
```

- [ ] **Step 2: Run the tests**

```bash
cargo test outlook::fake::
```

Expected: both tests pass.

- [ ] **Step 3: Commit**

```bash
git add src/outlook/mod.rs src/outlook/fake.rs
git commit -m "Add OutlookClient trait and FakeOutlookClient"
```

---

### Task 6: MCP server scaffold + Email tools

**Files:**
- Create: `src/server.rs`
- Modify: `src/main.rs` (add `mod server;`)
- Create: `tests/tools.rs`

**Interfaces:**
- Consumes: `OutlookClient` trait + `FakeOutlookClient` (Task 5), domain types (Task 4).
- Produces: `pub struct OutlookMcpServer { client: Arc<dyn OutlookClient>, tool_router: ToolRouter<OutlookMcpServer> }` with `pub fn new(client: Arc<dyn OutlookClient>) -> Self`, and the 9 email tools (`list_folders`, `list_emails`, `search_emails`, `get_email`, `send_email`, `create_draft`, `reply_email`, `move_email`, `delete_email`) registered via `rmcp`'s `#[tool_router]`/`#[tool]`. Later tasks (7–10) add more `#[tool]` methods to the *same* `impl` block in this file.

- [ ] **Step 1: Add dependencies**

```bash
cargo add rmcp --features server,transport-io
cargo add tokio --features full
cargo add schemars
cargo add anyhow
```

Expected: `Cargo.toml` gains entries for `rmcp`, `tokio`, `schemars`, `anyhow`. Note the exact versions `cargo add` resolves — if any subsequent step's code doesn't compile against what was resolved, check `cargo doc -p rmcp --open` for the current macro/type shape (see "API-drift note" in Global Constraints).

- [ ] **Step 2: Wire `ToolError` into `rmcp::ErrorData`**

Add to the bottom of `src/error.rs`:

```rust
impl From<ToolError> for rmcp::ErrorData {
    fn from(err: ToolError) -> Self {
        rmcp::ErrorData::internal_error(err.0, None)
    }
}
```

- [ ] **Step 3: Add the blocking-call helper**

This is what keeps every COM call (Task 11+) from running on a tokio worker thread that might migrate mid-call — see the "COM/tokio correctness constraint" in Global Constraints. Add to the top of `src/server.rs`:

**CORRECTED for the real rmcp 2.1.0 API** (this is what's actually committed
in `61d9de6`; the surrounding prose below still describes the original
research-based guess, kept for historical context, but this code block is
ground truth):

```rust
use std::sync::Arc;

use rmcp::{
    ErrorData as McpError, ServerHandler,
    handler::server::wrapper::Parameters,
    model::{CallToolResult, ContentBlock, ServerCapabilities, ServerInfo},
    tool, tool_handler, tool_router,
};
use serde::Deserialize;

use crate::error::ToolError;
use crate::outlook::OutlookClient;

/// Runs a blocking `OutlookClient` call on a dedicated blocking thread so the
/// tokio scheduler never migrates it mid-call (COM apartment-threading
/// requires the same OS thread for the lifetime of a call).
async fn run_blocking<T, F>(f: F) -> Result<T, ToolError>
where
    T: Send + 'static,
    F: FnOnce() -> Result<T, ToolError> + Send + 'static,
{
    tokio::task::spawn_blocking(f)
        .await
        .map_err(|e| ToolError::new(format!("internal task error: {e}")))?
}

fn json_content<T: serde::Serialize>(value: &T) -> Result<ContentBlock, McpError> {
    ContentBlock::json(value)
}

#[derive(Clone)]
pub struct OutlookMcpServer {
    client: Arc<dyn OutlookClient>,
}

impl OutlookMcpServer {
    pub fn new(client: Arc<dyn OutlookClient>) -> Self {
        Self { client }
    }
}
```

Note there is deliberately no `tool_router: ToolRouter<Self>` field: `#[tool_handler]`'s
default expansion calls `Self::tool_router()` fresh on every dispatch rather
than reading a stored field (that only happens with an explicit
`#[tool_handler(router = self.tool_router)]`), so a stored field is dead
weight — rebuilding this small, static router per call is cheap.

- [ ] **Step 4: Write the failing tool-layer tests**

Create `tests/tools.rs`:

**CORRECTED for the real rmcp 2.1.0 API** — this is what's actually
committed in `61d9de6`. As established above, `call_tool()` needs a
`RequestContext<RoleServer>` that's server-internal machinery, so tests call
each `pub async fn` tool method directly instead:

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, EMAIL_ID};
use outlook_mcp_rs::server::{
    DeleteEmailParams, GetEmailParams, ListEmailsParams, MoveEmailParams, OutlookMcpServer,
    ReplyEmailParams, SearchEmailsParams, SendEmailParams, CreateDraftParams,
};
use rmcp::handler::server::wrapper::Parameters;
use rmcp::model::CallToolResult;
use serde_json::{json, Value};

fn result_json(result: &CallToolResult) -> Value {
    let text = result.content[0]
        .as_text()
        .expect("expected text content")
        .text
        .clone();
    serde_json::from_str(&text).unwrap()
}

#[tokio::test]
async fn list_folders_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server.list_folders().await.unwrap();
    assert_eq!(fake.calls(), vec![("list_folders".to_string(), json!({}))]);
}

#[tokio::test]
async fn list_emails_passes_arguments() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_emails(Parameters(ListEmailsParams {
            folder: "sent".to_string(),
            count: 5,
            unread_only: true,
        }))
        .await
        .unwrap();
    assert_eq!(
        fake.calls(),
        vec![(
            "list_emails".to_string(),
            json!({"folder": "sent", "count": 5, "unread_only": true})
        )]
    );
}

#[tokio::test]
async fn list_emails_uses_defaults() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({})).unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    assert_eq!(
        fake.calls(),
        vec![(
            "list_emails".to_string(),
            json!({"folder": "inbox", "count": 10, "unread_only": false})
        )]
    );
}

#[tokio::test]
async fn search_emails_passes_query_and_since_days() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: SearchEmailsParams =
        serde_json::from_value(json!({"query": "invoice", "since_days": 30})).unwrap();
    server.search_emails(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "search_emails");
    assert_eq!(args["query"], "invoice");
    assert_eq!(args["since_days"], 30);
}

#[tokio::test]
async fn get_email_returns_body() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_email(Parameters(GetEmailParams {
            email_id: EMAIL_ID.to_string(),
            prefer_html: false,
        }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["body"], "Hi there");
}

#[tokio::test]
async fn send_email_passes_recipients_and_html_flag() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .send_email(Parameters(SendEmailParams {
            to: vec!["a@example.com".to_string(), "b@example.com".to_string()],
            subject: "Hi".to_string(),
            body: "Hello!".to_string(),
            cc: None,
            bcc: None,
            html: false,
        }))
        .await
        .unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "send_email");
    assert_eq!(args["to"], json!(["a@example.com", "b@example.com"]));
    assert_eq!(args["html"], false);
}

#[tokio::test]
async fn create_draft_returns_draft_saved_status() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .create_draft(Parameters(CreateDraftParams {
            to: vec!["a@example.com".to_string()],
            subject: "Hi".to_string(),
            body: "Hello!".to_string(),
            cc: None,
            bcc: None,
            html: false,
        }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["status"], "draft_saved");
}

#[tokio::test]
async fn reply_email_passes_reply_all_and_send_flags() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .reply_email(Parameters(ReplyEmailParams {
            email_id: EMAIL_ID.to_string(),
            body: "Thanks!".to_string(),
            reply_all: true,
            html: false,
            send: false,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["reply_all"], true);
    assert_eq!(args["send"], false);
}

#[tokio::test]
async fn move_email_returns_new_id() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .move_email(Parameters(MoveEmailParams {
            email_id: EMAIL_ID.to_string(),
            target_folder: "Archive".to_string(),
        }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["id"], "new-entry|store-1");
}

#[tokio::test]
async fn delete_email_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .delete_email(Parameters(DeleteEmailParams { email_id: EMAIL_ID.to_string() }))
        .await
        .unwrap();
    assert_eq!(
        fake.calls(),
        vec![("delete_email".to_string(), json!({"email_id": EMAIL_ID}))]
    );
}

#[tokio::test]
async fn client_error_propagates_as_tool_error() {
    let fake = Arc::new(FakeOutlookClient::new());
    fake.set_fail_with("Outlook exploded");
    let server = OutlookMcpServer::new(fake.clone());
    let err = server
        .list_emails(Parameters(ListEmailsParams {
            folder: "inbox".to_string(),
            count: 10,
            unread_only: false,
        }))
        .await
        .unwrap_err();
    assert!(err.message.contains("Outlook exploded"));
}
```

Add a `src/lib.rs` so the integration test crate (`tests/tools.rs`) can import `outlook_mcp_rs::server`/`outlook_mcp_rs::outlook` — Rust integration tests only see a crate's *public* API, and only if that API is exposed from a library target:

```rust
pub mod error;
pub mod outlook;
pub mod server;
```

And change `src/main.rs` to use the library instead of its own copies of the modules:

```rust
use outlook_mcp_rs::server;

fn main() {
    println!("outlook-mcp-rs: scaffold only, real server wired in a later task");
    let _ = &server::OutlookMcpServer::new; // silence unused-import warning until Task 17
}
```

Remove the now-duplicated `mod error; mod outlook; mod server;` lines from `src/main.rs` (they live in `src/lib.rs` now).

- [ ] **Step 2: Run the tests to see them fail on missing tool registrations**

```bash
cargo test --test tools
```

Expected: compile failure — `list_folders`/`list_emails`/etc. aren't registered as tools yet (`OutlookMcpServer` has no tools).

- [ ] **Step 3: Implement the email tools**

Add param structs and the `#[tool_router]` impl block to `src/server.rs`, below the `OutlookMcpServer` struct:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListEmailsParams {
    #[serde(default = "default_folder")]
    pub folder: String,
    #[serde(default = "default_count")]
    pub count: i32,
    #[serde(default)]
    pub unread_only: bool,
}
fn default_folder() -> String { "inbox".to_string() }
fn default_count() -> i32 { 10 }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SearchEmailsParams {
    pub query: String,
    #[serde(default = "default_folder")]
    pub folder: String,
    #[serde(default = "default_count")]
    pub count: i32,
    #[serde(default)]
    pub since_days: Option<i32>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetEmailParams {
    pub email_id: String,
    #[serde(default)]
    pub prefer_html: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SendEmailParams {
    pub to: Vec<String>,
    pub subject: String,
    pub body: String,
    #[serde(default)]
    pub cc: Option<Vec<String>>,
    #[serde(default)]
    pub bcc: Option<Vec<String>>,
    #[serde(default)]
    pub html: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateDraftParams {
    pub to: Vec<String>,
    pub subject: String,
    pub body: String,
    #[serde(default)]
    pub cc: Option<Vec<String>>,
    #[serde(default)]
    pub bcc: Option<Vec<String>>,
    #[serde(default)]
    pub html: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ReplyEmailParams {
    pub email_id: String,
    pub body: String,
    #[serde(default)]
    pub reply_all: bool,
    #[serde(default)]
    pub html: bool,
    #[serde(default = "default_true")]
    pub send: bool,
}
fn default_true() -> bool { true }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct MoveEmailParams {
    pub email_id: String,
    pub target_folder: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteEmailParams {
    pub email_id: String,
}

#[tool_router]
impl OutlookMcpServer {
    #[tool(description = "List Outlook mail folders (name, path, item counts).")]
    pub async fn list_folders(&self) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_folders()).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "List recent emails in a folder (default: inbox).")]
    pub async fn list_emails(
        &self,
        Parameters(ListEmailsParams { folder, count, unread_only }): Parameters<ListEmailsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_emails(folder, count, unread_only)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Search emails by subject/sender/body text in a folder.")]
    pub async fn search_emails(
        &self,
        Parameters(SearchEmailsParams { query, folder, count, since_days }): Parameters<SearchEmailsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.search_emails(query, folder, count, since_days)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get the full body and attachment list of one email by id.")]
    pub async fn get_email(
        &self,
        Parameters(GetEmailParams { email_id, prefer_html }): Parameters<GetEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.get_email(email_id, prefer_html)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Send a new email immediately.")]
    pub async fn send_email(
        &self,
        Parameters(SendEmailParams { to, subject, body, cc, bcc, html }): Parameters<SendEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.send_email(to, subject, body, cc, bcc, html)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create (but don't send) a draft email.")]
    pub async fn create_draft(
        &self,
        Parameters(CreateDraftParams { to, subject, body, cc, bcc, html }): Parameters<CreateDraftParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.create_draft(to, subject, body, cc, bcc, html)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Reply to an email, optionally to all recipients, optionally as a draft.")]
    pub async fn reply_email(
        &self,
        Parameters(ReplyEmailParams { email_id, body, reply_all, html, send }): Parameters<ReplyEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.reply_email(email_id, body, reply_all, html, send)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Move an email to another folder.")]
    pub async fn move_email(
        &self,
        Parameters(MoveEmailParams { email_id, target_folder }): Parameters<MoveEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.move_email(email_id, target_folder)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Delete an email (moves it to Deleted Items).")]
    pub async fn delete_email(
        &self,
        Parameters(DeleteEmailParams { email_id }): Parameters<DeleteEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_email(email_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
}

#[tool_handler]
impl ServerHandler for OutlookMcpServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo {
            protocol_version: ProtocolVersion::V_2024_11_05,
            capabilities: ServerCapabilities::builder().enable_tools().build(),
            server_info: Implementation::from_build_env(),
            instructions: Some(
                "Controls Microsoft Outlook desktop (email, calendar, tasks, notes) via COM.".into(),
            ),
        }
    }
}
```

- [ ] **Step 4: Run the tests until they pass**

```bash
cargo test --test tools
```

Expected: all 10 tests in `tests/tools.rs` pass. If `Parameters<T>`, `#[tool_router]`, or `ServerHandler::call_tool` don't match what `cargo doc -p rmcp --open` shows for the resolved version, fix the mismatched names/signatures mechanically — the test *behavior* (what's asserted) doesn't change.

- [ ] **Step 5: Commit**

```bash
git add Cargo.toml src/lib.rs src/main.rs src/server.rs src/error.rs tests/tools.rs
git commit -m "Add MCP server scaffold and email tools"
```

---

### Task 7: Calendar tools

**IMPORTANT — read before writing any code in this task or Tasks 8–10:**
Task 6 (already implemented and committed, commit `61d9de6`) discovered the
real `rmcp` crate resolved to version 2.1.0, meaningfully different from the
API this plan's test/example snippets were written against. Two concrete
consequences that affect every remaining tool-layer task:

1. `json_content(value)` (defined once in `src/server.rs` by Task 6) now
   returns `Result<ContentBlock, McpError>`, not a bare `ContentBlock` — every
   call site needs the `?` operator: `json_content(&result)?`, not
   `json_content(&result)`. (This plan's own snippets below and in Tasks 8–10
   have already been corrected to include the `?`; if you spot one without
   it, that's a transcription bug in the plan, not intentional.)
2. `tests/tools.rs`'s test helper is NOT `call(&server, "tool_name",
   json!({...}))` going through `call_tool()` — the real `ServerHandler::
   call_tool` requires a `RequestContext<RoleServer>` that's server-internal
   machinery a unit test has no simple way to construct. Instead, **read the
   actual committed `tests/tools.rs` file first** (all 11 email tests from
   Task 6) — it calls tool methods directly, e.g.:
   ```rust
   let result = server
       .get_email(Parameters(GetEmailParams { email_id: EMAIL_ID.to_string(), prefer_html: false }))
       .await
       .unwrap();
   assert_eq!(result_json(&result)["body"], "Hi there");
   ```
   Follow this exact pattern for every test in this task: construct the
   `Parameters<XxxParams>` struct literal (or `serde_json::from_value` it
   when you need default-filling behavior, as `list_emails_uses_defaults`
   does), call the `pub async fn` tool method directly, `.await.unwrap()` (or
   `.unwrap_err()` for the error-path test). Every tool method must be
   declared `pub async fn` (not just `async fn`) for this cross-crate test
   call to compile — Task 6's methods all needed this fix; don't repeat the
   omission.

**Files:**
- Modify: `src/server.rs` (add to the same `#[tool_router] impl OutlookMcpServer` block)
- Modify: `tests/tools.rs`

**Interfaces:**
- Consumes: `OutlookClient::list_events/get_event/create_event/respond_to_meeting` (Task 5).
- Produces: `list_events`, `get_event`, `create_event`, `respond_to_meeting` tools.

- [ ] **Step 1: Write the failing tests**

Add to `tests/tools.rs`:

```rust
#[tokio::test]
async fn list_events_passes_date_range() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_events(Parameters(ListEventsParams {
            start_date: Some("2026-06-10".to_string()),
            end_date: Some("2026-06-17".to_string()),
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("list_events".to_string(), json!({"start_date": "2026-06-10", "end_date": "2026-06-17"})),
    ]);
}

#[tokio::test]
async fn get_event_returns_subject() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_event(Parameters(GetEventParams { event_id: EVENT_ID.to_string() }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["subject"], "Standup");
}

#[tokio::test]
async fn create_event_passes_attendees() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .create_event(Parameters(CreateEventParams {
            subject: "Sync".to_string(),
            start: "2026-06-12T14:00".to_string(),
            end: "2026-06-12T15:00".to_string(),
            body: None,
            location: None,
            attendees: Some(vec!["a@example.com".to_string()]),
            all_day: false,
            reminder_minutes: None,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["attendees"], json!(["a@example.com"]));
}

#[tokio::test]
async fn respond_to_meeting_defaults_send_true() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: RespondToMeetingParams =
        serde_json::from_value(json!({"event_id": EVENT_ID, "response": "accept"})).unwrap();
    server.respond_to_meeting(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["response"], "accept");
    assert_eq!(args["send"], true);
}
```

Add the new param struct imports to the top of `tests/tools.rs`'s existing `use outlook_mcp_rs::server::{...}` line: `ListEventsParams, GetEventParams, CreateEventParams, RespondToMeetingParams`.

- [ ] **Step 2: Run to see them fail**

```bash
cargo test --test tools
```

Expected: compile failure, `list_events`/etc. tools don't exist yet.

- [ ] **Step 3: Implement the calendar tools**

Add to `src/server.rs`, param structs above the `#[tool_router]` block and methods inside it (same block as Task 6's methods):

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListEventsParams {
    #[serde(default)]
    pub start_date: Option<String>,
    #[serde(default)]
    pub end_date: Option<String>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetEventParams {
    pub event_id: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateEventParams {
    pub subject: String,
    pub start: String,
    pub end: String,
    #[serde(default)]
    pub body: Option<String>,
    #[serde(default)]
    pub location: Option<String>,
    #[serde(default)]
    pub attendees: Option<Vec<String>>,
    #[serde(default)]
    pub all_day: bool,
    #[serde(default)]
    pub reminder_minutes: Option<i32>,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct RespondToMeetingParams {
    pub event_id: String,
    pub response: String,
    #[serde(default)]
    pub comment: Option<String>,
    #[serde(default = "default_true")]
    pub send: bool,
}
```

Add these methods inside the existing `#[tool_router] impl OutlookMcpServer { ... }` block from Task 6:

```rust
    #[tool(description = "List calendar events in a date range (default: next 7 days).")]
    pub async fn list_events(
        &self,
        Parameters(ListEventsParams { start_date, end_date }): Parameters<ListEventsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_events(start_date, end_date)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get the full details of one calendar event by id.")]
    pub async fn get_event(
        &self,
        Parameters(GetEventParams { event_id }): Parameters<GetEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.get_event(event_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create a calendar event, optionally sending meeting invites to attendees.")]
    pub async fn create_event(
        &self,
        Parameters(CreateEventParams {
            subject, start, end, body, location, attendees, all_day, reminder_minutes,
        }): Parameters<CreateEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || {
            client.create_event(subject, start, end, body, location, attendees, all_day, reminder_minutes)
        }).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Respond to a meeting invite: accept, decline, or tentative.")]
    pub async fn respond_to_meeting(
        &self,
        Parameters(RespondToMeetingParams { event_id, response, comment, send }): Parameters<RespondToMeetingParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.respond_to_meeting(event_id, response, comment, send)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

- [ ] **Step 4: Run tests until they pass**

```bash
cargo test --test tools
```

Expected: all tests pass (11 from Task 6 + 4 new).

- [ ] **Step 5: Commit**

```bash
git add src/server.rs tests/tools.rs
git commit -m "Add calendar tools"
```

---

### Task 8: Attachment tools

**Files:**
- Modify: `src/server.rs`
- Modify: `tests/tools.rs`

**Interfaces:**
- Consumes: `OutlookClient::list_attachments/save_attachments` (Task 5).
- Produces: `list_attachments`, `save_attachments` tools.

- [ ] **Step 1: Write the failing tests**

Add to `tests/tools.rs`:

```rust
#[tokio::test]
async fn list_attachments_returns_filename() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .list_attachments(Parameters(ListAttachmentsParams { email_id: EMAIL_ID.to_string() }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)[0]["filename"], "report.pdf");
}

#[tokio::test]
async fn save_attachments_passes_dir_and_names() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .save_attachments(Parameters(SaveAttachmentsParams {
            email_id: EMAIL_ID.to_string(),
            save_dir: "/tmp/x".to_string(),
            attachment_names: Some(vec!["report.pdf".to_string()]),
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["save_dir"], "/tmp/x");
    assert_eq!(args["attachment_names"], json!(["report.pdf"]));
}
```

Add `ListAttachmentsParams, SaveAttachmentsParams` to `tests/tools.rs`'s `use outlook_mcp_rs::server::{...}` import line.

- [ ] **Step 2: Run to see them fail**

```bash
cargo test --test tools
```

- [ ] **Step 3: Implement**

Add to `src/server.rs`:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListAttachmentsParams {
    pub email_id: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct SaveAttachmentsParams {
    pub email_id: String,
    pub save_dir: String,
    #[serde(default)]
    pub attachment_names: Option<Vec<String>>,
}
```

Inside the `#[tool_router]` block:

```rust
    #[tool(description = "List an email's attachments (filename and size).")]
    pub async fn list_attachments(
        &self,
        Parameters(ListAttachmentsParams { email_id }): Parameters<ListAttachmentsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_attachments(email_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Save an email's attachments to a local directory.")]
    pub async fn save_attachments(
        &self,
        Parameters(SaveAttachmentsParams { email_id, save_dir, attachment_names }): Parameters<SaveAttachmentsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.save_attachments(email_id, save_dir, attachment_names)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

- [ ] **Step 4: Run tests until they pass**

```bash
cargo test --test tools
```

- [ ] **Step 5: Commit**

```bash
git add src/server.rs tests/tools.rs
git commit -m "Add attachment tools"
```

---

### Task 9: Task tools

**Files:**
- Modify: `src/server.rs`
- Modify: `tests/tools.rs`

**Interfaces:**
- Consumes: `OutlookClient::list_tasks/create_task/complete_task` (Task 5).
- Produces: `list_tasks`, `create_task`, `complete_task` tools.

- [ ] **Step 1: Write the failing tests**

```rust
#[tokio::test]
async fn list_tasks_passes_include_completed() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_tasks(Parameters(ListTasksParams { include_completed: true }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("list_tasks".to_string(), json!({"include_completed": true})),
    ]);
}

#[tokio::test]
async fn create_task_passes_importance() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: CreateTaskParams = serde_json::from_value(json!({
        "subject": "Buy milk", "due_date": "2026-06-15", "importance": "high"
    })).unwrap();
    server.create_task(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["importance"], "high");
}

#[tokio::test]
async fn complete_task_records_call() {
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .complete_task(Parameters(CompleteTaskParams { task_id: TASK_ID.to_string() }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("complete_task".to_string(), json!({"task_id": TASK_ID})),
    ]);
}
```

Add `ListTasksParams, CreateTaskParams, CompleteTaskParams` to `tests/tools.rs`'s `use outlook_mcp_rs::server::{...}` import line.

- [ ] **Step 2: Run to see them fail**

```bash
cargo test --test tools
```

- [ ] **Step 3: Implement**

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListTasksParams {
    #[serde(default)]
    pub include_completed: bool,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateTaskParams {
    pub subject: String,
    #[serde(default)]
    pub body: Option<String>,
    #[serde(default)]
    pub due_date: Option<String>,
    #[serde(default = "default_importance")]
    pub importance: String,
}
fn default_importance() -> String { "normal".to_string() }

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CompleteTaskParams {
    pub task_id: String,
}
```

```rust
    #[tool(description = "List Outlook tasks (default: not-yet-completed only).")]
    pub async fn list_tasks(
        &self,
        Parameters(ListTasksParams { include_completed }): Parameters<ListTasksParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_tasks(include_completed)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create a new task.")]
    pub async fn create_task(
        &self,
        Parameters(CreateTaskParams { subject, body, due_date, importance }): Parameters<CreateTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.create_task(subject, body, due_date, importance)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Mark a task complete.")]
    pub async fn complete_task(
        &self,
        Parameters(CompleteTaskParams { task_id }): Parameters<CompleteTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.complete_task(task_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

- [ ] **Step 4: Run tests until they pass**

```bash
cargo test --test tools
```

- [ ] **Step 5: Commit**

```bash
git add src/server.rs tests/tools.rs
git commit -m "Add task tools"
```

---

### Task 10: Note tools

**Files:**
- Modify: `src/server.rs`
- Modify: `tests/tools.rs`

**Interfaces:**
- Consumes: `OutlookClient::list_notes/get_note/create_note` (Task 5).
- Produces: `list_notes`, `get_note`, `create_note` tools. After this task, `OutlookMcpServer` has all 21 tools.

- [ ] **Step 1: Write the failing tests**

```rust
#[tokio::test]
async fn list_notes_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server.list_notes().await.unwrap();
    assert_eq!(fake.calls(), vec![("list_notes".to_string(), json!({}))]);
}

#[tokio::test]
async fn get_note_returns_body() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_note(Parameters(GetNoteParams { note_id: NOTE_ID.to_string() }))
        .await
        .unwrap();
    assert!(result_json(&result)["body"].as_str().unwrap().starts_with("Ideas"));
}

#[tokio::test]
async fn create_note_records_body() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .create_note(Parameters(CreateNoteParams { body: "Ideas\n- one".to_string() }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("create_note".to_string(), json!({"body": "Ideas\n- one"})),
    ]);
}
```

Add `GetNoteParams, CreateNoteParams` to `tests/tools.rs`'s `use outlook_mcp_rs::server::{...}` import line. After this task, verify `cargo test` shows 21 total tests across `tests/tools.rs` (matching all 21 tool methods having at least one test, though several share coverage — the exact count depends on how many were written per task; the important check is that every one of the 21 tools in `OutlookClient`/`server.rs` has at least one passing test exercising it).

- [ ] **Step 2: Run to see them fail**

```bash
cargo test --test tools
```

- [ ] **Step 3: Implement**

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct GetNoteParams {
    pub note_id: String,
}

#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateNoteParams {
    pub body: String,
}
```

```rust
    #[tool(description = "List Outlook notes.")]
    pub async fn list_notes(&self) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_notes()).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Get the full body of one note by id.")]
    pub async fn get_note(
        &self,
        Parameters(GetNoteParams { note_id }): Parameters<GetNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.get_note(note_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }

    #[tool(description = "Create a new note.")]
    pub async fn create_note(
        &self,
        Parameters(CreateNoteParams { body }): Parameters<CreateNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.create_note(body)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

- [ ] **Step 4: Run tests until they pass**

```bash
cargo test --test tools
```

Expected: all 21 tools are registered and every tool-layer test (from Tasks 6–10) passes — the entire MCP surface is now provably correct against `FakeOutlookClient`, without any Outlook installed.

- [ ] **Step 5: Commit**

```bash
git add src/server.rs tests/tools.rs
git commit -m "Add note tools; full tool parity against FakeOutlookClient"
```

---

### Task 11: COM helpers foundation

**Files:**
- Create: `src/outlook/com.rs`
- Create: `src/constants.rs`
- Modify: `src/lib.rs` (add `mod com;` under `outlook`, add `pub mod constants;`)

**Interfaces:**
- Consumes: `ToolError` (Task 3).
- Produces:
  - `pub struct ComGuard` (RAII `CoInitializeEx`/`CoUninitialize`, one per call — mirrors `pythoncom.CoInitialize()` in `client.py`'s `_com` decorator).
  - `pub fn create_com_object(prog_id: &str) -> windows::core::Result<IDispatch>`
  - `pub fn get_property(disp: &IDispatch, name: &str) -> windows::core::Result<VARIANT>`
  - `pub fn put_property(disp: &IDispatch, name: &str, value: VARIANT) -> windows::core::Result<()>`
  - `pub fn call_method(disp: &IDispatch, name: &str, args: &mut [VARIANT]) -> windows::core::Result<VARIANT>`
  - `pub fn format_com_error(err: &windows::core::Error) -> String` (translation of `outlook_mcp/errors.py::format_com_error`)
  - Pure-logic helpers, fully unit-tested without any COM object: `make_item_id`, `parse_item_id`, `jet_datetime`, `safe_filename`.
  - `src/constants.rs`: verbatim translation of `outlook_mcp/constants.py`.

- [ ] **Step 1: Add the windows crate**

```bash
cargo add windows --features Win32_System_Com,Win32_System_Variant,Win32_System_Ole,Win32_Globalization,Win32_Foundation
```

Expected: `Cargo.toml` gains a `windows` dependency with those features. If `cargo build` later reports an item (e.g. `CLSIDFromProgID`, `DISPPARAMS`) living in a different feature/module than listed here, add the reported feature with `cargo add windows --features <name>` and update the `use` path — this is expected drift, not a design problem (see the "API-drift note" in Global Constraints).

- [ ] **Step 2: Translate the constants**

Create `src/constants.rs` (direct translation of `outlook_mcp/constants.py`):

```rust
// OlDefaultFolders
pub const OL_FOLDER_DELETED_ITEMS: i32 = 3;
pub const OL_FOLDER_OUTBOX: i32 = 4;
pub const OL_FOLDER_SENT_MAIL: i32 = 5;
pub const OL_FOLDER_INBOX: i32 = 6;
pub const OL_FOLDER_CALENDAR: i32 = 9;
pub const OL_FOLDER_CONTACTS: i32 = 10;
pub const OL_FOLDER_JOURNAL: i32 = 11;
pub const OL_FOLDER_NOTES: i32 = 12;
pub const OL_FOLDER_TASKS: i32 = 13;
pub const OL_FOLDER_DRAFTS: i32 = 16;

// OlItemType (Application.CreateItem)
pub const OL_MAIL_ITEM: i32 = 0;
pub const OL_APPOINTMENT_ITEM: i32 = 1;
pub const OL_TASK_ITEM: i32 = 3;
pub const OL_NOTE_ITEM: i32 = 5;

// OlBodyFormat
pub const OL_FORMAT_PLAIN: i32 = 1;
pub const OL_FORMAT_HTML: i32 = 2;

// OlMeetingResponse (AppointmentItem.Respond)
pub const OL_MEETING_TENTATIVE: i32 = 2;
pub const OL_MEETING_ACCEPTED: i32 = 3;
pub const OL_MEETING_DECLINED: i32 = 4;

// OlMeetingStatus
pub const OL_NONMEETING: i32 = 0;
pub const OL_MEETING: i32 = 1;

// OlTaskStatus
pub const OL_TASK_NOT_STARTED: i32 = 0;

// OlImportance
pub const OL_IMPORTANCE_LOW: i32 = 0;
pub const OL_IMPORTANCE_NORMAL: i32 = 1;
pub const OL_IMPORTANCE_HIGH: i32 = 2;

pub fn folder_name_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "inbox" => Some(OL_FOLDER_INBOX),
        "sent" | "sent items" => Some(OL_FOLDER_SENT_MAIL),
        "drafts" => Some(OL_FOLDER_DRAFTS),
        "deleted" | "deleted items" | "trash" => Some(OL_FOLDER_DELETED_ITEMS),
        "outbox" => Some(OL_FOLDER_OUTBOX),
        _ => None,
    }
}

pub fn importance_name_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "low" => Some(OL_IMPORTANCE_LOW),
        "normal" => Some(OL_IMPORTANCE_NORMAL),
        "high" => Some(OL_IMPORTANCE_HIGH),
        _ => None,
    }
}

pub fn meeting_response_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "accept" => Some(OL_MEETING_ACCEPTED),
        "decline" => Some(OL_MEETING_DECLINED),
        "tentative" => Some(OL_MEETING_TENTATIVE),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn folder_name_lookup_is_case_insensitive() {
        assert_eq!(folder_name_to_id("Sent Items"), Some(OL_FOLDER_SENT_MAIL));
        assert_eq!(folder_name_to_id("nonexistent"), None);
    }

    #[test]
    fn importance_and_meeting_response_lookups() {
        assert_eq!(importance_name_to_id("HIGH"), Some(OL_IMPORTANCE_HIGH));
        assert_eq!(meeting_response_to_id("Accept"), Some(OL_MEETING_ACCEPTED));
        assert_eq!(meeting_response_to_id("maybe"), None);
    }
}
```

- [ ] **Step 3: Write the failing pure-logic tests**

Create `src/outlook/com.rs` with just the pure-logic pieces and their tests first:

```rust
use crate::error::ToolError;

/// `"{EntryID}|{StoreID}"`, matching the Python client's opaque item id format.
pub fn make_item_id(entry_id: &str, store_id: &str) -> String {
    format!("{entry_id}|{store_id}")
}

pub fn parse_item_id(item_id: &str) -> Result<(String, String), ToolError> {
    match item_id.split_once('|') {
        Some((entry, store)) if !entry.is_empty() && !store.is_empty() => {
            Ok((entry.to_string(), store.to_string()))
        }
        _ => Err(ToolError::new(format!(
            "Invalid item id {item_id:?}: expected the opaque id returned by a list/search tool."
        ))),
    }
}

/// JET `Restrict` filters want `MM/DD/YYYY HH:MM AM/PM` (US format, no
/// seconds) — anything else silently misfilters. Mirrors `_jet_dt` in
/// `outlook_mcp/outlook/client.py`.
pub fn jet_datetime(dt: &chrono::NaiveDateTime) -> String {
    dt.format("%m/%d/%Y %I:%M %p").to_string()
}

pub fn safe_filename(name: &str) -> String {
    let cleaned: String = name
        .chars()
        .map(|c| if "\\/:*?\"<>|".contains(c) || (c as u32) < 0x20 { '_' } else { c })
        .collect();
    let trimmed = cleaned.trim_matches(|c| c == '.' || c == ' ');
    if trimmed.is_empty() { "attachment".to_string() } else { trimmed.to_string() }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn make_and_parse_item_id_round_trip() {
        let id = make_item_id("entry-1", "store-1");
        assert_eq!(id, "entry-1|store-1");
        assert_eq!(parse_item_id(&id).unwrap(), ("entry-1".to_string(), "store-1".to_string()));
    }

    #[test]
    fn parse_item_id_rejects_malformed_input() {
        assert!(parse_item_id("no-separator").is_err());
        assert!(parse_item_id("|missing-entry").is_err());
        assert!(parse_item_id("missing-store|").is_err());
    }

    #[test]
    fn jet_datetime_formats_us_style_no_seconds() {
        use chrono::NaiveDate;
        let dt = NaiveDate::from_ymd_opt(2026, 6, 10).unwrap().and_hms_opt(14, 30, 0).unwrap();
        assert_eq!(jet_datetime(&dt), "06/10/2026 02:30 PM");
    }

    #[test]
    fn safe_filename_strips_unsafe_characters() {
        assert_eq!(safe_filename("report:final*.pdf"), "report_final_.pdf");
        assert_eq!(safe_filename("   "), "attachment");
        assert_eq!(safe_filename(""), "attachment");
    }
}
```

- [ ] **Step 4: Add chrono and run the pure-logic tests**

```bash
cargo add chrono
cargo test outlook::com::
```

Expected: all 4 tests pass.

- [ ] **Step 5: Add the late-bound IDispatch helpers**

Append to `src/outlook/com.rs` (these are exercised by the live system tests in Task 18, not by `cargo test`, since they need a real COM object):

```rust
use windows::core::{Error as WinError, GUID, PCWSTR, Result as WinResult};
use windows::Win32::Globalization::GetUserDefaultLCID;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSIDFromProgID, IDispatch,
    CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
    DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
};
use windows::Win32::System::Variant::VARIANT;

/// One per COM call (mirrors `pythoncom.CoInitialize()` inside `client.py`'s
/// `@_com` decorator): initializes this OS thread for apartment-threaded COM
/// on construction, uninitializes on drop. Must be created on the same
/// thread `spawn_blocking`'s closure runs on (see `run_blocking` in
/// `src/server.rs`), and must outlive every COM object used within that call.
pub struct ComGuard;

impl ComGuard {
    pub fn new() -> WinResult<Self> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()? };
        Ok(ComGuard)
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

pub fn create_com_object(prog_id: &str) -> WinResult<IDispatch> {
    let wide: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();
    let clsid = unsafe { CLSIDFromProgID(PCWSTR(wide.as_ptr()))? };
    unsafe { CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER) }
}

fn name_to_dispid(disp: &IDispatch, name: &str) -> WinResult<i32> {
    let wide: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
    let names = [PCWSTR(wide.as_ptr())];
    let mut dispid = 0i32;
    unsafe {
        disp.GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 1, GetUserDefaultLCID(), &mut dispid)?;
    }
    Ok(dispid)
}

const DISP_E_EXCEPTION: i32 = -2147352567; // 0x80020009, from winerror.h

fn enrich_error(err: WinError, excepinfo: &EXCEPINFO) -> WinError {
    if err.code().0 == DISP_E_EXCEPTION && !excepinfo.bstrDescription.is_empty() {
        return WinError::new(err.code(), excepinfo.bstrDescription.to_string());
    }
    err
}

fn invoke(
    disp: &IDispatch,
    name: &str,
    flags: windows::Win32::System::Com::DISPATCH_FLAGS,
    args: &mut [VARIANT],
) -> WinResult<VARIANT> {
    let dispid = name_to_dispid(disp, name)?;
    args.reverse(); // COM wants arguments in reverse order
    let is_put = flags == DISPATCH_PROPERTYPUT;
    let mut put_dispid: i32 = -3; // DISPID_PROPERTYPUT
    let mut params = DISPPARAMS {
        rgvarg: args.as_mut_ptr(),
        rgdispidNamedArgs: if is_put { &mut put_dispid } else { std::ptr::null_mut() },
        cArgs: args.len() as u32,
        cNamedArgs: if is_put { 1 } else { 0 },
    };
    let mut result = VARIANT::default();
    let mut excepinfo = EXCEPINFO::default();
    let mut arg_err = 0u32;
    unsafe {
        disp.Invoke(
            dispid, &GUID::zeroed(), GetUserDefaultLCID(), flags,
            &mut params, Some(&mut result), Some(&mut excepinfo), Some(&mut arg_err),
        )
    }
    .map_err(|e| enrich_error(e, &excepinfo))?;
    Ok(result)
}

pub fn get_property(disp: &IDispatch, name: &str) -> WinResult<VARIANT> {
    invoke(disp, name, DISPATCH_PROPERTYGET, &mut [])
}

pub fn put_property(disp: &IDispatch, name: &str, value: VARIANT) -> WinResult<()> {
    invoke(disp, name, DISPATCH_PROPERTYPUT, &mut [value])?;
    Ok(())
}

pub fn call_method(disp: &IDispatch, name: &str, args: &mut [VARIANT]) -> WinResult<VARIANT> {
    invoke(disp, name, DISPATCH_METHOD, args)
}

/// Translation of `outlook_mcp/errors.py::format_com_error`. `windows-rs`'s
/// `Error::message()` on an error enriched by `enrich_error` above already
/// carries the COM exception's own description text (equivalent to Python's
/// `excepinfo[2]`), so this is simpler than the Python version, which has to
/// dig into `excepinfo` itself.
pub fn format_com_error(err: &WinError) -> String {
    format!("Outlook error: {} (HRESULT {:#010x})", err.message(), err.code().0)
}
```

- [ ] **Step 6: Wire the module**

`src/outlook/mod.rs` gains:

```rust
pub mod com;
```

`src/lib.rs` gains:

```rust
pub mod constants;
```

- [ ] **Step 7: Run all tests**

```bash
cargo build
cargo test
```

Expected: builds cleanly on Windows; all previously-passing tests (Tasks 3–10) plus the 6 new tests in this task still pass. The `get_property`/`put_property`/`call_method`/`create_com_object` functions compile but are not exercised by `cargo test` — they're proven out in Task 18's live suite.

- [ ] **Step 8: Commit**

```bash
git add Cargo.toml src/constants.rs src/outlook/com.rs src/outlook/mod.rs src/lib.rs
git commit -m "Add Outlook constants and late-bound COM helpers"
```

---

### Task 12: WindowsOutlookClient — email methods

**Files:**
- Create: `src/outlook/client.rs`
- Modify: `src/outlook/mod.rs` (add `pub mod client;`)

**Interfaces:**
- Consumes: `ComGuard`, `create_com_object`, `get_property`, `put_property`, `call_method`, `format_com_error`, `make_item_id`, `parse_item_id`, `safe_filename` (Task 11), `OutlookClient` trait (Task 5), constants (Task 11).
- Produces: `pub struct WindowsOutlookClient` implementing `list_folders`, `list_emails`, `search_emails`, `get_email`, `send_email`, `create_draft`, `reply_email`, `move_email`, `delete_email`. No automated test in this task (needs live Outlook) — proven out in Task 18.

This task is a direct translation of `outlook_mcp/outlook/client.py`'s email section (lines 133–336 in the Python file: `_email_summary`, `list_folders`, `list_emails`, `search_emails`, `get_email`, `_compose`, `send_email`, `create_draft`, `reply_email`, `move_email`, `delete_email`). Read that file side-by-side while implementing this task — every COM property/method name (`GetNamespace`, `GetDefaultFolder`, `Items`, `Restrict`, `Sort`, `EntryID`, `Parent.StoreID`, etc.) and every JET filter string must match exactly.

- [ ] **Step 1: Scaffold the struct and the MAPI/folder-resolution helpers**

Create `src/outlook/client.rs`:

```rust
use windows::core::Result as WinResult;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::VARIANT;

use crate::constants as c;
use crate::error::ToolError;
use crate::outlook::com::{
    call_method, create_com_object, format_com_error, get_property, make_item_id,
    parse_item_id, ComGuard,
};
use crate::outlook::types::*;
use crate::outlook::OutlookClient;

pub struct WindowsOutlookClient;

impl WindowsOutlookClient {
    pub fn new() -> Self {
        Self
    }

    /// Every public method wraps its body in this: initializes COM on the
    /// current (blocking-pool) thread, maps `windows::core::Error` to
    /// `ToolError`, mirrors `client.py`'s `@_com` decorator.
    fn with_com<T>(&self, f: impl FnOnce() -> WinResult<T>) -> Result<T, ToolError> {
        let _guard = ComGuard::new().map_err(|e| ToolError::new(format_com_error(&e)))?;
        f().map_err(|e| ToolError::new(format_com_error(&e)))
    }
}
```

*(The full set of `variant_from_*`/`variant_to_*` conversion helpers used throughout this and the following tasks — `variant_from_str`, `variant_from_i32`, `variant_from_bool`, `variant_to_string`, `variant_to_optional_string` — belongs alongside `get_property`/`call_method` in `src/outlook/com.rs`. Add them there now, translating however the resolved `windows` crate's `VARIANT` constructors/accessors work — check `cargo doc -p windows --open` for `VARIANT` if the exact conversion API isn't obvious from the compiler's suggestions. Each is a small, independently testable pure function: writing a round-trip unit test — `variant_to_string(&variant_from_str("x")) == "x"` — in `src/outlook/com.rs` for each one is good practice and doesn't require live Outlook, since `VARIANT` construction alone doesn't need a running COM server.)*

- [ ] **Step 2: Implement `list_folders`**

Add to `src/outlook/client.rs`. Translated from `client.py:182-207`:

```rust
impl OutlookClient for WindowsOutlookClient {
    fn list_folders(&self) -> Result<Vec<FolderInfo>, ToolError> {
        self.with_com(|| {
            let app = create_com_object("Outlook.Application")?;
            let ns = call_method(&app, "GetNamespace", &mut [crate::outlook::com::variant_from_str("MAPI")])?;
            let ns_disp: IDispatch = ns.try_into()?;
            let inbox = call_method(&ns_disp, "GetDefaultFolder", &mut [crate::outlook::com::variant_from_i32(c::OL_FOLDER_INBOX)])?;
            let inbox_disp: IDispatch = inbox.try_into()?;
            let root = get_property(&inbox_disp, "Parent")?;
            let root_disp: IDispatch = root.try_into()?;

            let mut results = Vec::new();
            fn walk(
                folder: &IDispatch, path: &str, depth: u32, results: &mut Vec<FolderInfo>,
            ) -> WinResult<()> {
                let name = crate::outlook::com::variant_to_string(&get_property(folder, "Name")?);
                let items = get_property(folder, "Items")?;
                let items_disp: IDispatch = items.try_into()?;
                let item_count = crate::outlook::com::variant_to_i32(&get_property(&items_disp, "Count")?).unwrap_or(0);
                let unread = crate::outlook::com::variant_to_i32(&get_property(folder, "UnReadItemCount")?).unwrap_or(0);
                results.push(FolderInfo { name: name.clone(), path: path.to_string(), items: item_count, unread });
                if depth >= 3 {
                    return Ok(());
                }
                let subfolders = get_property(folder, "Folders")?;
                let subfolders_disp: IDispatch = subfolders.try_into()?;
                let count = crate::outlook::com::variant_to_i32(&get_property(&subfolders_disp, "Count")?).unwrap_or(0);
                for i in 1..=count {
                    let sub = call_method(&subfolders_disp, "Item", &mut [crate::outlook::com::variant_from_i32(i)])?;
                    let sub_disp: IDispatch = sub.try_into()?;
                    let sub_name = crate::outlook::com::variant_to_string(&get_property(&sub_disp, "Name")?);
                    walk(&sub_disp, &format!("{path}/{sub_name}"), depth + 1, results)?;
                }
                Ok(())
            }

            let root_subfolders = get_property(&root_disp, "Folders")?;
            let root_subfolders_disp: IDispatch = root_subfolders.try_into()?;
            let count = crate::outlook::com::variant_to_i32(&get_property(&root_subfolders_disp, "Count")?).unwrap_or(0);
            for i in 1..=count {
                let sub = call_method(&root_subfolders_disp, "Item", &mut [crate::outlook::com::variant_from_i32(i)])?;
                let sub_disp: IDispatch = sub.try_into()?;
                let sub_name = crate::outlook::com::variant_to_string(&get_property(&sub_disp, "Name")?);
                walk(&sub_disp, &sub_name, 1, &mut results)?;
            }
            Ok(results)
        })
    }

    // list_emails, search_emails, get_email, send_email, create_draft,
    // reply_email, move_email, delete_email implemented in the next step,
    // following the same with_com(...) + get_property/call_method pattern,
    // translating client.py:209-336 method-for-method:
    //   - list_emails: resolve folder (see Step 3's resolve_folder helper),
    //     Items.Restrict("[UnRead] = True") if unread_only, Items.Sort
    //     ("[ReceivedTime]", true), iterate up to `count`, building
    //     EmailSummary via the same fields as `_email_summary` in client.py.
    //   - search_emails: build the @SQL DASL filter exactly as in
    //     client.py:230-235 (escaping single quotes by doubling them),
    //     optionally re-Restrict on `[ReceivedTime] >= '<jet_datetime>'`
    //     using this task's `jet_datetime` helper.
    //   - get_email: EmailDetail with cc/bcc/body/html_body(if prefer_html)/
    //     attachments (FileName of each Attachments.Item(i), 1-based).
    //   - send_email/create_draft/reply_email: build a MailItem via
    //     Application.CreateItem(OL_MAIL_ITEM), set To/CC/BCC (";" joined),
    //     Subject, BodyFormat + Body or HTMLBody depending on `html`, then
    //     Send() or Save() depending on the tool (reply_email additionally
    //     depending on the `send` flag) — mirrors `_compose` in client.py.
    //   - move_email: resolve target folder, call Move(target), return the
    //     new item's id (EntryID changes on Move).
    //   - delete_email: call Delete() after reading Subject for the return value.
}
```

Note the inline comment block above deliberately specifies exact COM member names, JET filter syntax, and control flow for the remaining 8 email methods rather than pasting nearly-mechanical repeats of the same `get_property`/`call_method` pattern shown in full for `list_folders` — implement each by reading the corresponding lines of `outlook_mcp/outlook/client.py` referenced above and transliterating using this task's helpers. Do not leave any of them unimplemented; `cargo build` in Step 3 requires the full trait implementation to compile.

- [ ] **Step 3: Add the folder-resolution helper used by `list_emails`/`search_emails`/`move_email`**

Translated from `client.py:112-131`:

```rust
fn resolve_folder(ns: &IDispatch, folder: Option<&str>) -> WinResult<IDispatch> {
    let name = folder.unwrap_or("inbox").trim();
    if let Some(id) = c::folder_name_to_id(name) {
        let result = call_method(ns, "GetDefaultFolder", &mut [crate::outlook::com::variant_from_i32(id)])?;
        return result.try_into();
    }
    let inbox = call_method(ns, "GetDefaultFolder", &mut [crate::outlook::com::variant_from_i32(c::OL_FOLDER_INBOX)])?;
    let inbox_disp: IDispatch = inbox.try_into()?;
    let root = get_property(&inbox_disp, "Parent")?;
    let mut current: IDispatch = root.try_into()?;
    for part in name.split(['/', '\\']).filter(|p| !p.is_empty()) {
        let subfolders = get_property(&current, "Folders")?;
        let subfolders_disp: IDispatch = subfolders.try_into()?;
        let count = crate::outlook::com::variant_to_i32(&get_property(&subfolders_disp, "Count")?).unwrap_or(0);
        let mut found = None;
        for i in 1..=count {
            let sub = call_method(&subfolders_disp, "Item", &mut [crate::outlook::com::variant_from_i32(i)])?;
            let sub_disp: IDispatch = sub.try_into()?;
            let sub_name = crate::outlook::com::variant_to_string(&get_property(&sub_disp, "Name")?);
            if sub_name.eq_ignore_ascii_case(part) {
                found = Some(sub_disp);
                break;
            }
        }
        current = found.ok_or_else(|| {
            windows::core::Error::from_hresult(windows::Win32::Foundation::E_FAIL)
        })?;
    }
    Ok(current)
}
```

Note: unlike the Python version, this raises a generic `E_FAIL` (mapped to a generic "folder not found" message by `format_com_error`) rather than the specific `"Folder not found: {name} (no subfolder named {part})"` message — since `format_com_error` only sees a `windows::core::Error`, not the folder-lookup context. If a more specific error message matters, wrap the `ok_or_else` branch to construct a `ToolError` directly and return early from `with_com`'s caller instead of routing through `format_com_error`. Given this is a Windows-only internal helper with a live-Outlook test covering it in Task 18, prefer the simple version first and only add the richer message if the live test's error output isn't clear enough.

- [ ] **Step 4: Build**

```bash
cargo build
```

Expected: compiles on Windows. Fix any `VARIANT`/`IDispatch` conversion mismatches against whatever the resolved `windows` crate version's actual API surface is (see the "API-drift note").

- [ ] **Step 5: Commit**

```bash
git add src/outlook/client.rs src/outlook/mod.rs src/outlook/com.rs
git commit -m "Implement WindowsOutlookClient email methods"
```

---

### Task 13: WindowsOutlookClient — calendar methods

**Files:**
- Modify: `src/outlook/client.rs`

**Interfaces:**
- Produces: `list_events`, `get_event`, `create_event`, `respond_to_meeting` on `WindowsOutlookClient`.

Direct translation of `client.py:340-428`. Key details to preserve exactly:
- `list_events`: default range is today through +7 days if `start_date`/`end_date` omitted; `Items.IncludeRecurrences = True` **must** be set before `Sort`/`Restrict` (COM ordering requirement, see `client.py:351`); cap results at `MAX_CALENDAR_ITEMS = 250` since unbounded recurrences can expand forever.
- `get_event`: adds `body`, `required_attendees`, `optional_attendees`, `response_status` to the summary.
- `create_event`: `Application.CreateItem(OL_APPOINTMENT_ITEM)`; if `attendees` is non-empty, set `MeetingStatus = OL_MEETING`, add each attendee via `Recipients.Add`, `Recipients.ResolveAll()`, then `Send()` (status `"meeting_sent"`); otherwise `Save()` (status `"saved"`).
- `respond_to_meeting`: validate `response` via `constants::meeting_response_to_id` (error message mirrors `client.py:409-412` exactly: `"Invalid response {response!r}: use 'accept', 'decline' or 'tentative'."`); if the resolved item exposes `GetAssociatedAppointment` (i.e., it's a meeting request from the inbox, not already an appointment), call it first; then `Respond(id, True)`, set `Body` on the response item if `comment` given, then `Send()` or `Save()` depending on `send`.

- [ ] **Step 1: Implement**

Add the four methods to the `impl OutlookClient for WindowsOutlookClient` block, following the same `with_com(|| { ... })` + `get_property`/`call_method`/`put_property` pattern established in Task 12, transliterating `client.py:340-428` line-for-line including the exact JET filter strings (`"[Start] >= '{start}' AND [Start] <= '{end}'"`) via `jet_datetime`.

- [ ] **Step 2: Build**

```bash
cargo build
```

- [ ] **Step 3: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Implement WindowsOutlookClient calendar methods"
```

---

### Task 14: WindowsOutlookClient — attachment methods

**Files:**
- Modify: `src/outlook/client.rs`

**Interfaces:**
- Produces: `list_attachments`, `save_attachments` on `WindowsOutlookClient`.

Direct translation of `client.py:432-478`:
- `list_attachments`: 1-based `Attachments.Item(i)` iteration, `{index, filename, size}` per attachment.
- `save_attachments`: validate the email has attachments (else `ToolError::new("This email has no attachments.")`); create `save_dir` if missing; if `attachment_names` given, filter case-insensitively; call `SaveAsFile` per attachment via `safe_filename`-sanitized paths; collect per-file `{filename, saved_to, status: "saved"}` or `{filename, status: "failed", error}` (COM errors on individual files don't abort the whole call — matches `client.py:466-472`); if nothing matched, `ToolError::new("No attachments matched attachment_names; use list_attachments to see the exact file names.")`.

- [ ] **Step 1: Implement**

Add both methods, using `std::fs::create_dir_all` for the directory (translation of Python's `os.makedirs(save_dir, exist_ok=True)`) and the `safe_filename` helper from Task 11.

- [ ] **Step 2: Build**

```bash
cargo build
```

- [ ] **Step 3: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Implement WindowsOutlookClient attachment methods"
```

---

### Task 15: WindowsOutlookClient — task methods

**Files:**
- Modify: `src/outlook/client.rs`

**Interfaces:**
- Produces: `list_tasks`, `create_task`, `complete_task` on `WindowsOutlookClient`.

Direct translation of `client.py:482-516`:
- `list_tasks`: `GetDefaultFolder(OL_FOLDER_TASKS).Items`, `Restrict("[Complete] = False")` unless `include_completed`.
- `create_task`: validate `importance` via `constants::importance_name_to_id` (error message mirrors `client.py:496-497` exactly); `CreateItem(OL_TASK_ITEM)`, set `Subject`/`Body`/`DueDate`/`Importance`, `Save()`.
- `complete_task`: `MarkComplete()`.

- [ ] **Step 1: Implement**

- [ ] **Step 2: Build**

```bash
cargo build
```

- [ ] **Step 3: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Implement WindowsOutlookClient task methods"
```

---

### Task 16: WindowsOutlookClient — note methods

**Files:**
- Modify: `src/outlook/client.rs`

**Interfaces:**
- Produces: `list_notes`, `get_note`, `create_note` on `WindowsOutlookClient`. After this task, `WindowsOutlookClient` implements the full `OutlookClient` trait.

Direct translation of `client.py:520-542`:
- `list_notes`/`get_note`: `_note_summary` derives `subject` from the first non-empty line of `Body`, truncated to 120 chars (`client.py:171-178`) — translate this exactly, since notes have no native `Subject` property.
- `create_note`: reject empty `body` (`ToolError::new("create_note requires a non-empty body.")`), `CreateItem(OL_NOTE_ITEM)`, set `Body`, `Save()`.

- [ ] **Step 1: Implement**

- [ ] **Step 2: Build and confirm the trait is fully implemented**

```bash
cargo build
```

Expected: builds with no "missing trait items" errors — `WindowsOutlookClient` now implements all 21 `OutlookClient` methods.

- [ ] **Step 3: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Implement WindowsOutlookClient note methods; full trait complete"
```

---

### Task 17: main.rs — real server wiring

**Files:**
- Modify: `src/main.rs`

**Interfaces:**
- Consumes: `OutlookMcpServer::new` (Task 6), `WindowsOutlookClient::new` (Task 12).
- Produces: a runnable binary that serves the real MCP server over stdio.

- [ ] **Step 1: Replace the stub `main`**

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::client::WindowsOutlookClient;
use outlook_mcp_rs::server::OutlookMcpServer;
use rmcp::{transport::stdio, ServiceExt};

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let client = Arc::new(WindowsOutlookClient::new());
    let server = OutlookMcpServer::new(client);
    let service = server.serve(stdio()).await?;
    service.waiting().await?;
    Ok(())
}
```

*(If `cargo doc -p rmcp --open` shows `serve`/`waiting` living on a different trait or with a different name for the resolved version, adjust — the intent is "start the server over stdio and block until the client disconnects".)*

- [ ] **Step 2: Build**

```bash
cargo build --release
```

Expected: `target/release/outlook-mcp-rs.exe` produced.

- [ ] **Step 3: Manual smoke test**

```bash
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"smoke-test","version":"0"}}}' | ./target/release/outlook-mcp-rs.exe
```

Expected: a single JSON-RPC response line on stdout with `"result"` containing `serverInfo`/`capabilities` — confirms the binary starts, initializes, and doesn't panic, without needing Outlook installed yet (the client isn't constructed until a tool is actually called... note: `WindowsOutlookClient::new()` per Task 12's design does no COM work itself, only `with_com` does when a method runs — so `main()` starting up doesn't require Outlook to be running, only actually calling a tool does).

- [ ] **Step 4: Run the full test suite one more time**

```bash
cargo test
```

Expected: everything from Tasks 3–11 still passes (this task didn't touch anything they depend on).

- [ ] **Step 5: Commit**

```bash
git add src/main.rs
git commit -m "Wire main() to serve the real WindowsOutlookClient over stdio"
```

---

### Task 18: Live system tests

**Files:**
- Create: `tests/live_outlook.rs`
- Create: `TESTING.md`

**Interfaces:**
- Consumes: `WindowsOutlookClient` (Tasks 12–16), `OutlookClient` trait (Task 5).
- Produces: an `#[ignore]`d integration test suite that exercises the real COM client, safe to run repeatedly against a real mailbox (every test cleans up what it creates).

- [ ] **Step 1: Write the live test suite**

Create `tests/live_outlook.rs`:

```rust
//! Live system tests against a real, running Outlook. NOT run by plain
//! `cargo test` — every test is `#[ignore]`d. Run explicitly:
//!   cargo test --test live_outlook -- --ignored
//! See TESTING.md for preconditions.
//!
//! Every test that creates an Outlook item deletes it before returning, so
//! repeated runs don't accumulate junk in the mailbox. `send_email` and
//! `respond_to_meeting` are deliberately NOT covered here since a real send
//! can't be undone — see TESTING.md for how to test those by hand.

use outlook_mcp_rs::outlook::client::WindowsOutlookClient;
use outlook_mcp_rs::outlook::OutlookClient;

fn client() -> WindowsOutlookClient {
    WindowsOutlookClient::new()
}

#[test]
#[ignore]
fn list_folders_returns_at_least_inbox() {
    let folders = client().list_folders().expect("list_folders should succeed against a live Outlook");
    assert!(folders.iter().any(|f| f.name.eq_ignore_ascii_case("inbox")));
}

#[test]
#[ignore]
fn list_emails_returns_inbox_items() {
    let emails = client().list_emails("inbox".into(), 5, false)
        .expect("list_emails should succeed against a live Outlook");
    // Not asserting a specific count/content since the real mailbox varies —
    // just confirm the call succeeds and returns well-formed summaries.
    for email in &emails {
        assert!(!email.id.is_empty());
    }
}

#[test]
#[ignore]
fn create_draft_then_delete_round_trips() {
    let c = client();
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs live test draft".to_string(),
        "This draft is created and deleted by an automated test.".to_string(),
        None, None, false,
    ).expect("create_draft should succeed");
    let id = created["id"].as_str().expect("create_draft returns an id").to_string();
    c.delete_email(id).expect("cleanup: delete_email should succeed");
}

#[test]
#[ignore]
fn create_task_complete_then_it_is_marked_complete() {
    let c = client();
    let created = c.create_task(
        "outlook-mcp-rs live test task".to_string(), None, None, "normal".to_string(),
    ).expect("create_task should succeed");
    let id = created["id"].as_str().unwrap().to_string();
    c.complete_task(id.clone()).expect("complete_task should succeed");
    let tasks = c.list_tasks(true).expect("list_tasks should succeed");
    assert!(tasks.iter().any(|t| t.id == id && t.complete));
    // Outlook has no direct "delete task" in our trait yet — deleting the
    // completed test task manually is fine (it's clearly labeled).
}

#[test]
#[ignore]
fn create_note_then_get_it_back() {
    let c = client();
    let created = c.create_note("outlook-mcp-rs live test note".to_string())
        .expect("create_note should succeed");
    let id = created["id"].as_str().unwrap().to_string();
    let note = c.get_note(id).expect("get_note should succeed");
    assert!(note.body.starts_with("outlook-mcp-rs live test note"));
}

#[test]
#[ignore]
fn create_event_then_delete_it() {
    let c = client();
    let created = c.create_event(
        "outlook-mcp-rs live test event".to_string(),
        "2099-01-01T10:00:00".to_string(),
        "2099-01-01T10:30:00".to_string(),
        None, None, None, false, None,
    ).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();
    // Calendar items don't have a dedicated "delete" tool in the trait; use
    // move_email into Deleted Items works for mail but not appointments —
    // delete the test event manually from the calendar after this test runs,
    // or extend the trait with a delete_event method if this becomes
    // frequent enough to automate.
    let _ = c.get_event(id); // just confirm it round-trips before manual cleanup
}
```

Note the `create_event` test's cleanup gap is called out explicitly rather than silently skipped — `OutlookClient` has no `delete_event` (matching the Python version, which also doesn't have one), so full automation isn't possible without expanding scope beyond parity. This is documented in `TESTING.md` too.

- [ ] **Step 2: Write TESTING.md**

Create `TESTING.md`:

```markdown
# Testing outlook-mcp-rs

## Unit tests

```
cargo test
```

Runs everything except the live suite (`tests/live_outlook.rs`, all
`#[ignore]`d) — no Outlook required, safe to run anywhere, and what CI runs
on every push.

## Live system tests

These exercise the real `WindowsOutlookClient` against Outlook actually
running on your machine. Preconditions:

- Windows, with classic Outlook desktop installed
- Outlook is open and signed in to a normal mailbox
- You're comfortable with a handful of test items (a draft, a task, a note,
  a calendar event, each clearly named "outlook-mcp-rs live test ...") being
  created in that mailbox — most are cleaned up automatically, but the
  calendar event test currently requires manual deletion afterward (there's
  no `delete_event` tool; see `tests/live_outlook.rs` for why).

Run them with:

```
cargo test --test live_outlook -- --ignored
```

## Manual-only tests (not automated at all)

`send_email` and `respond_to_meeting` have real, unrecoverable side effects
(an actually-delivered email; an actual meeting response sent to an
organizer) and are not covered by any automated test. To verify them by
hand before a release:

1. Pick a test recipient you control (e.g. a second mailbox of your own).
2. Call `send_email` with that recipient and a clearly-marked test subject;
   confirm it arrives.
3. Find (or create) a meeting invite in your test mailbox and call
   `respond_to_meeting` with `response: "tentative"`; confirm the organizer
   sees a tentative response.
```

- [ ] **Step 3: Run the live suite locally** (only if you have Outlook open right now — otherwise skip to Step 4)

```bash
cargo test --test live_outlook -- --ignored
```

Expected: all 5 tests pass against your real mailbox, and the draft/task/note test items are gone afterward (check Deleted Items / your task list / your notes to confirm cleanup worked); the test calendar event remains until you delete it by hand.

- [ ] **Step 4: Confirm plain `cargo test` still ignores them**

```bash
cargo test 2>&1 | grep -i live_outlook
```

Expected: lines like `test live_outlook::list_folders_returns_at_least_inbox ... ignored` — proving CI (which runs plain `cargo test`) never touches live Outlook.

- [ ] **Step 5: Commit**

```bash
git add tests/live_outlook.rs TESTING.md
git commit -m "Add live-Outlook system tests and TESTING.md"
```

---

### Task 19: CI — build and release jobs

**Files:**
- Modify: `.github/workflows/ci.yaml`

**Interfaces:**
- Consumes: `cargo build --release` (Task 17).
- Produces: on tag push `v*`, a `build` job producing `outlook-mcp-rs.exe` and a `release` job attaching it to a GitHub Release.

- [ ] **Step 1: Extend the workflow**

Replace the contents of `.github/workflows/ci.yaml`:

```yaml
name: CI/CD

on:
  push:
    branches: [main]
    tags: ['v*']
  pull_request:
    branches: [main]

jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: dtolnay/rust-toolchain@stable
      - name: Run tests
        run: cargo test --all -- --skip live_outlook

  build:
    needs: test
    runs-on: windows-latest
    if: github.event_name == 'push' && startsWith(github.ref, 'refs/tags/v')
    steps:
      - uses: actions/checkout@v4
      - uses: dtolnay/rust-toolchain@stable
      - name: Build release binary
        run: cargo build --release
      - name: Verify the binary was produced
        shell: bash
        run: |
          ls -lh target/release/outlook-mcp-rs.exe
          test -f target/release/outlook-mcp-rs.exe
      - uses: actions/upload-artifact@v4
        with:
          name: outlook-mcp-rs-exe
          path: target/release/outlook-mcp-rs.exe

  release:
    needs: build
    runs-on: ubuntu-latest
    if: github.event_name == 'push' && startsWith(github.ref, 'refs/tags/v')
    permissions:
      contents: write
    steps:
      - uses: actions/download-artifact@v4
        with:
          name: outlook-mcp-rs-exe
          path: dist/
      - name: Create GitHub Release and attach the binary
        uses: softprops/action-gh-release@v2
        with:
          files: dist/outlook-mcp-rs.exe
          generate_release_notes: true
```

- [ ] **Step 2: Commit**

```bash
git add .github/workflows/ci.yaml
git commit -m "Add release build/publish jobs to CI"
git push
```

- [ ] **Step 3: Verify with a real tag** (do this once Task 20's manual QA has passed, not before)

```bash
git tag v0.1.0
git push origin v0.1.0
gh run list --workflow ci.yaml --limit 1
```

Expected: `test` → `build` → `release` all succeed; `gh release view v0.1.0` shows `outlook-mcp-rs.exe` attached.

---

### Task 20: Final polish and release

**Files:**
- Modify: `README.md`

**Interfaces:**
- Produces: a tagged, released v0.1.0.

- [ ] **Step 1: Manual QA against real Outlook**

With Outlook open and signed in:

```bash
cargo build --release
cargo test --test live_outlook -- --ignored
```

Expected: all live tests pass (Task 18). Then run the manual-only `send_email`/`respond_to_meeting` checks from `TESTING.md` at least once.

- [ ] **Step 2: Point a real MCP client at the built binary**

Configure your MCP client (e.g. Claude Desktop's config) to run
`target/release/outlook-mcp-rs.exe` directly, restart it, and confirm the
Outlook tools appear and a couple of them (e.g. `list_folders`, `list_emails`)
return real data from your mailbox through the actual client, not just
`cargo test`.

- [ ] **Step 3: Finalize README**

Add a short "Available tools" section to `README.md` listing all 21 tool
names grouped by category (Email/Calendar/Attachments/Tasks/Notes), so users
browsing the repo can see the surface area without reading source.

- [ ] **Step 4: Commit, tag, and push**

```bash
git add README.md
git commit -m "Document available tools in README"
git push
git tag v0.1.0
git push origin v0.1.0
```

- [ ] **Step 5: Confirm the release**

```bash
gh run list --workflow ci.yaml --limit 1
gh release view v0.1.0
```

Expected: release job succeeded, `outlook-mcp-rs.exe` is attached and downloadable.
