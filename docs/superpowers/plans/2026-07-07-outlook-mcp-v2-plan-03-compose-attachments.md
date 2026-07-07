# v2 Plan 3 — Compose attachments Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development or superpowers:executing-plans. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Let `send_email`, `create_draft`, and `reply_email` attach local files, given file paths, validated to exist before anything is sent.

**Architecture:** Add an `attachments: Option<Vec<String>>` param (list of local file paths) to all three trait methods. A shared `attach_files` helper validates every path exists (fail fast — nothing sends if a path is bad) then calls `MailItem.Attachments.Add(path)`. `send_email`/`create_draft` attach after `compose`; `reply_email` attaches to the reply item. Both trait implementors + server + tests updated together.

**Tech Stack:** Rust 2024, `windows` 0.62.2 COM, `std::path`, `serde`/`schemars`.

**Depends on:** nothing (independent of Plan 1); Plans 1–2 are already shipped.

## Global Constraints

- Target crate: `C:\Users\adamk\projects\outlook-mcp-rs`.
- Trait has TWO implementors (`WindowsOutlookClient` in `client.rs`, `FakeOutlookClient` in `fake.rs`); both change, plus `src/server.rs` and `tests/tools.rs`.
- Validate ALL paths exist BEFORE attaching/sending — a missing path returns a `ToolError` and nothing is sent (send is irreversible).
- `send_email` keeps its existing empty-`to` guard.
- Commit after each task; `cargo test` green before commit. No push (controller pushes at plan end).

---

### Task 1: Thread `attachments` through trait, client, fake, server, tests

**Files:**
- Modify: `src/outlook/mod.rs` (add `attachments: Option<Vec<String>>` to the three trait methods)
- Modify: `src/outlook/client.rs` (add `attach_files` helper; wire into `send_email`/`create_draft`/`reply_email`)
- Modify: `src/outlook/fake.rs` (add the param + record it in the three methods)
- Modify: `src/server.rs` (add `attachments` to `SendEmailParams`/`CreateDraftParams`/`ReplyEmailParams` + pass through)
- Modify: `tests/tools.rs` (update the three tests' param literals; add a forwarding assertion)

**Interfaces:**
- Produces: `send_email(&self, to, subject, body, cc, bcc, html, attachments: Option<Vec<String>>)`, `create_draft(&self, …, attachments: Option<Vec<String>>)`, `reply_email(&self, email_id, body, reply_all, html, send, attachments: Option<Vec<String>>)`.
- Produces (client.rs module-level): `fn attach_files(mail: &IDispatch, paths: &[String]) -> Result<(), ToolError>`.

- [ ] **Step 1: Change the trait in `src/outlook/mod.rs`**

Add `attachments: Option<Vec<String>>` as the LAST parameter of each of `send_email`, `create_draft`, `reply_email`. Example:
```rust
    fn send_email(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool,
        attachments: Option<Vec<String>>) -> Result<Value, ToolError>;
    fn create_draft(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool,
        attachments: Option<Vec<String>>) -> Result<Value, ToolError>;
    fn reply_email(&self, email_id: String, body: String, reply_all: bool,
        html: bool, send: bool, attachments: Option<Vec<String>>)
        -> Result<Value, ToolError>;
```
(Match the exact existing formatting of these signatures in the file; only add the new param.)

- [ ] **Step 2: Add the `attach_files` helper in `src/outlook/client.rs`**

Place it near the other module-level helpers (e.g. right after `compose`):
```rust
/// Attach local files to a mail/reply item. Validates every path exists
/// FIRST (so a bad path fails before anything is sent), then adds each via
/// `MailItem.Attachments.Add(path)`.
fn attach_files(mail: &IDispatch, paths: &[String]) -> Result<(), ToolError> {
    for p in paths {
        if !std::path::Path::new(p).is_file() {
            return Err(ToolError::new(format!("attachment not found: {p}")));
        }
    }
    let atts = to_disp(get_property(mail, "Attachments")?)?;
    for p in paths {
        call_method(&atts, "Add", &mut [variant_from_str(p)])?;
    }
    Ok(())
}
```

- [ ] **Step 3: Wire into the three client methods**

`send_email`: add `attachments: Option<Vec<String>>` to the signature; inside `with_com`, after `let mail = compose(...)?;` and BEFORE `call_method(&mail, "Send", ...)`:
```rust
            if let Some(atts) = attachments.as_deref() {
                attach_files(&mail, atts)?;
            }
```
`create_draft`: same — add the param, attach after `compose` and before `Save`.
`reply_email`: add the param; inside `with_com`, after setting the reply body (the `if html { … } else { … }` block) and BEFORE the `if send { … }` block:
```rust
            if let Some(atts) = attachments.as_deref() {
                attach_files(&reply, atts)?;
            }
```

- [ ] **Step 4: Update the fake in `src/outlook/fake.rs`**

Add `attachments: Option<Vec<String>>` to the three fake methods' signatures and include it in the recorded JSON. E.g. for `send_email`:
```rust
    fn send_email(&self, to: Vec<String>, subject: String, body: String,
        cc: Option<Vec<String>>, bcc: Option<Vec<String>>, html: bool,
        attachments: Option<Vec<String>>) -> Result<Value, ToolError> {
        self.record("send_email", json!({
            "to": to, "subject": subject, "body": body, "cc": cc, "bcc": bcc,
            "html": html, "attachments": attachments,
        }))?;
        Ok(json!({"status": "sent", "to": to.join("; "), "subject": subject}))
    }
```
Do the same for `create_draft` and `reply_email` (add the param + `"attachments": attachments` to their recorded json).

- [ ] **Step 5: Update the server in `src/server.rs`**

Add to each params struct:
```rust
    #[serde(default)]
    pub attachments: Option<Vec<String>>,
```
(to `SendEmailParams`, `CreateDraftParams`, `ReplyEmailParams`). Then pass `p.attachments`/`params.attachments` through in each tool method's client call (add it as the final argument). Follow the existing destructuring/passing style in those three tool methods.

- [ ] **Step 6: Update `tests/tools.rs`**

The existing `send_email_*`, `create_draft_*`, `reply_email_*` tests construct the params structs — add `attachments: None` to those literals (or switch to `serde_json::from_value(json!({...})).unwrap()`). Add one forwarding test:
```rust
#[tokio::test]
async fn send_email_forwards_attachments() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: SendEmailParams = serde_json::from_value(json!({
        "to": ["a@x.com"], "subject": "Hi", "body": "yo",
        "attachments": ["C:/tmp/a.pdf", "C:/tmp/b.png"]
    })).unwrap();
    server.send_email(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["attachments"], serde_json::json!(["C:/tmp/a.pdf", "C:/tmp/b.png"]));
}
```

- [ ] **Step 7: Build + test**

Run: `cargo build` (clean, no warnings). Run: `cargo test` (all green incl. `send_email_forwards_attachments`).

- [ ] **Step 8: Commit**

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/server.rs tests/tools.rs
git commit -m "Add attachments param to send_email/create_draft/reply_email"
```

---

### Task 2: Live attachment round-trip test

**Files:**
- Modify: `tests/live_outlook.rs`

**Interfaces:**
- Consumes: real `WindowsOutlookClient::create_draft` + `delete_email`.

- [ ] **Step 1: Add an `#[ignore]`d live test that drafts with an attachment then deletes it**

Uses a real temp file so the path validation + `Attachments.Add` actually run; cleans up the draft afterward:
```rust
#[test]
#[ignore]
fn create_draft_with_attachment_round_trips() {
    let dir = std::env::temp_dir();
    let path = dir.join("outlook-mcp-rs-live-attach.txt");
    std::fs::write(&path, b"live attachment test").expect("write temp file");
    let path_str = path.to_string_lossy().to_string();

    let c = WindowsOutlookClient::new();
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs attachment test".to_string(),
        "see attached".to_string(),
        None, None, false,
        Some(vec![path_str]),
    ).expect("create_draft with attachment should succeed");
    let id = created["id"].as_str().expect("draft id").to_string();
    c.delete_email(id).expect("cleanup: delete the draft");
    let _ = std::fs::remove_file(&path);
}

#[test]
#[ignore]
fn send_with_missing_attachment_errors_before_sending() {
    let c = WindowsOutlookClient::new();
    let err = c.send_email(
        vec!["nobody@example.invalid".to_string()],
        "should not send".to_string(), "body".to_string(),
        None, None, false,
        Some(vec!["C:/definitely/does/not/exist/nope.pdf".to_string()]),
    ).unwrap_err();
    assert!(err.to_string().contains("attachment not found"));
}
```
(Note: the second test proves the fail-fast path errors before Send — it never actually sends because the path check fails first. Safe to run.)

- [ ] **Step 2: Confirm compile + ignored**

Run: `cargo build --tests` (clean). Run: `cargo test 2>&1 | grep -E "create_draft_with_attachment|send_with_missing_attachment"` → both show `ignored`.

- [ ] **Step 3: (If Outlook available) run live**

Run: `cargo test --test live_outlook -- --ignored create_draft_with_attachment_round_trips send_with_missing_attachment_errors_before_sending`
Expected: PASS; the temp file and draft are cleaned up. Skip if no Outlook.

- [ ] **Step 4: Commit**

```bash
git add tests/live_outlook.rs
git commit -m "Add live attachment round-trip and fail-fast tests"
```

---

## Self-Review

- **Spec coverage:** `attachments` on send/draft/reply ✅ (T1); paths validated before send ✅ (`attach_files` validates all first); shared logic (single `attach_files` helper) ✅.
- **Placeholder scan:** none — full code throughout.
- **Type consistency:** `attach_files(&IDispatch, &[String])` signature matches its three call sites; the new trait param `Option<Vec<String>>` is consistent across trait/client/fake/server.
- **Safety:** validation happens before `Send`/`Save`; a bad path errors out with nothing sent.

## Execution Handoff

Plan 3 of 12. After green, controller pushes to main and proceeds to Plan 4 (Meeting-aware get_email). Models: T1 sonnet, T2 haiku.
