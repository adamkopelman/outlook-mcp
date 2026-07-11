# v2 Plan 5 — update_email (absorbs move_email) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the single-purpose `move_email` tool with one `update_email` tool that performs every non-destructive change to an existing email — move, mark read/unread, flag, add/remove categories, set importance — in a single call.

**Architecture:** Introduce an `EmailUpdate` params struct (mirroring the existing `EmailQuery` pattern). The trait method `move_email` is removed and `update_email(&self, u: EmailUpdate)` takes its place, implemented in both `FakeOutlookClient` (records the call, returns the `{status,id,changed}` shape) and `WindowsOutlookClient` (real COM: apply state changes first, then `move_to` **last** because `Move` changes the EntryID). The MCP tool `move_email` + `MoveEmailParams` are retired in the same plan that adds `update_email` + `UpdateEmailParams`, so there is never a broken intermediate tool set.

**Tech Stack:** Rust, `windows` 0.62.2 COM, `rmcp` 2.1.0 tool macros, `serde_json`.

## Global Constraints

- **Target crate:** `C:\Users\adamk\projects\outlook-mcp-rs` (the Rust impl, NOT the Python `outlook-mcp`). Edition 2024, rustc 1.95.0.
- **Two implementors per trait change.** `OutlookClient` lives in `src/outlook/mod.rs`; every signature change touches BOTH `WindowsOutlookClient` (`src/outlook/client.rs`) and `FakeOutlookClient` (`src/outlook/fake.rs`), plus the tool layer (`src/server.rs`) and tests (`tests/tools.rs`). Also scan `tests/live_outlook.rs` for call sites.
- **Tolerance:** new COM property reads in summary/detail builders use `.unwrap_or_default()`, never `?`. (State *writes* in `update_email` may use `?` — a write failure is a real error worth surfacing.)
- **Return shape for mutators:** `serde_json::Value` — a `{"status": ...}` object.
- **Zero warnings** on `cargo build` / `cargo test` before the plan is pushed.
- **Model policy:** Task 1 = **sonnet** (interface + fake + tool ripple), Task 2 = **opus** (real COM state mutation), Task 3 = **haiku** (live test).

---

### Task 1: Interface + fake + tool layer + unit tests

Swap the trait method, implement the fake, retire the `move_email` tool and add the `update_email` tool, and cover it with fake-client tool tests. The real COM impl is stubbed here and filled in Task 2.

**Files:**
- Modify: `src/outlook/mod.rs` (add `EmailUpdate` struct; swap trait method)
- Modify: `src/outlook/fake.rs` (replace `move_email` with `update_email`)
- Modify: `src/outlook/client.rs` (replace `move_email` with a `todo!()` stub — real impl in Task 2)
- Modify: `src/server.rs` (retire `MoveEmailParams` + `move_email` tool; add `UpdateEmailParams` + `update_email` tool)
- Modify: `tests/tools.rs` (replace `move_email_returns_new_id` test with `update_email` tests)
- Modify: `README.md` (replace the `move_email` bullet)

**Interfaces:**
- Produces: `pub struct EmailUpdate { pub email_id: String, pub move_to: Option<String>, pub mark_read: Option<bool>, pub flag: Option<String>, pub add_categories: Option<Vec<String>>, pub remove_categories: Option<Vec<String>>, pub importance: Option<String> }`
- Produces: trait method `fn update_email(&self, u: EmailUpdate) -> Result<Value, ToolError>;`
- Produces (fake return / real return contract): `{"status": "updated", "id": "<current-or-new>", "changed": [ ... ]}` where `changed` lists, in application order, the fields that were touched: any of `"mark_read"`, `"flag"`, `"add_categories"`, `"remove_categories"`, `"importance"`, `"move_to"`.

- [ ] **Step 1: Add the `EmailUpdate` struct to `src/outlook/mod.rs`**

Add directly below the existing `EmailQuery` struct (before `pub trait OutlookClient`):

```rust
/// All changes `update_email` can apply to one existing email. Every field
/// except `email_id` is optional; supplying several applies all of them.
/// State changes are applied first and `move_to` last (Move changes the
/// EntryID, so it must come after everything that addresses the item by id).
#[derive(Debug, Clone, Default)]
pub struct EmailUpdate {
    pub email_id: String,
    pub move_to: Option<String>,
    pub mark_read: Option<bool>,
    pub flag: Option<String>,               // "follow_up" | "complete" | "clear"
    pub add_categories: Option<Vec<String>>,
    pub remove_categories: Option<Vec<String>>,
    pub importance: Option<String>,         // "low" | "normal" | "high"
}
```

- [ ] **Step 2: Swap the trait method in `src/outlook/mod.rs`**

Replace:

```rust
    fn move_email(&self, email_id: String, target_folder: String)
        -> Result<Value, ToolError>;
```

with:

```rust
    fn update_email(&self, u: EmailUpdate) -> Result<Value, ToolError>;
```

- [ ] **Step 3: Run the build to confirm it now fails (both implementors + tool missing)**

Run: `cargo build 2>&1 | Select-String "not.*implemented|move_email|update_email" | Select-Object -First 5`
Expected: FAIL — `FakeOutlookClient`/`WindowsOutlookClient` no longer satisfy the trait, and `server.rs` references the removed method. (Red.)

- [ ] **Step 4: Implement `update_email` in `src/outlook/fake.rs`**

Replace the whole `move_email` fn (lines ~111–115) with:

```rust
    fn update_email(&self, u: EmailUpdate) -> Result<Value, ToolError> {
        self.record("update_email", json!({
            "email_id": u.email_id, "move_to": u.move_to, "mark_read": u.mark_read,
            "flag": u.flag, "add_categories": u.add_categories,
            "remove_categories": u.remove_categories, "importance": u.importance,
        }))?;
        // Mirror the real client's `changed` ordering: state changes first, move last.
        let mut changed: Vec<&str> = Vec::new();
        if u.mark_read.is_some() { changed.push("mark_read"); }
        if u.flag.is_some() { changed.push("flag"); }
        if u.add_categories.is_some() { changed.push("add_categories"); }
        if u.remove_categories.is_some() { changed.push("remove_categories"); }
        if u.importance.is_some() { changed.push("importance"); }
        // Move changes the EntryID; simulate a new id only when we moved.
        let id = if u.move_to.is_some() {
            changed.push("move_to");
            "new-entry|store-1".to_string()
        } else {
            u.email_id.clone()
        };
        Ok(json!({"status": "updated", "id": id, "changed": changed}))
    }
```

- [ ] **Step 5: Stub `update_email` in `src/outlook/client.rs` (real impl lands in Task 2)**

Replace the whole `move_email` fn (lines ~668–684) with:

```rust
    fn update_email(&self, _u: EmailUpdate) -> Result<Value, ToolError> {
        // Real COM implementation added in Plan 5 Task 2.
        todo!("update_email real COM impl — Plan 5 Task 2")
    }
```

Confirm `EmailUpdate` is in scope. `client.rs` already brings the outlook module types into scope the same way it uses `EmailQuery`; if `EmailQuery` is imported by an explicit `use`, add `EmailUpdate` alongside it. (Check the existing `use` lines near the top of `client.rs` — match whatever pattern `EmailQuery` uses.)

- [ ] **Step 6: Retire `MoveEmailParams` and add `UpdateEmailParams` in `src/server.rs`**

Find the `MoveEmailParams` struct and replace it with:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct UpdateEmailParams {
    pub email_id: String,
    /// Destination folder name (e.g. "Archive"). Applied last; changes the id.
    #[serde(default)]
    pub move_to: Option<String>,
    /// true = mark read, false = mark unread.
    #[serde(default)]
    pub mark_read: Option<bool>,
    /// "follow_up" | "complete" | "clear".
    #[serde(default)]
    pub flag: Option<String>,
    /// Category names to add (existing categories are preserved).
    #[serde(default)]
    pub add_categories: Option<Vec<String>>,
    /// Category names to remove.
    #[serde(default)]
    pub remove_categories: Option<Vec<String>>,
    /// "low" | "normal" | "high".
    #[serde(default)]
    pub importance: Option<String>,
}
```

- [ ] **Step 7: Swap the tool method in `src/server.rs`**

Replace the `move_email` `#[tool]` method (lines ~295–303) with:

```rust
    #[tool(description = "Update an existing email: move to a folder, mark read/unread, flag (follow_up/complete/clear), add/remove categories, or set importance. Combine any of these in one call.")]
    pub async fn update_email(
        &self,
        Parameters(p): Parameters<UpdateEmailParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let u = EmailUpdate {
            email_id: p.email_id, move_to: p.move_to, mark_read: p.mark_read,
            flag: p.flag, add_categories: p.add_categories,
            remove_categories: p.remove_categories, importance: p.importance,
        };
        let result = run_blocking(move || client.update_email(u)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Ensure `EmailUpdate` is imported in `server.rs` wherever `EmailQuery` is imported (same `use crate::outlook::{...}` line).

- [ ] **Step 8: Replace the tool test in `tests/tools.rs`**

Delete `move_email_returns_new_id` (and remove `MoveEmailParams` from the imports at the top of the file; add `UpdateEmailParams`). Add:

```rust
#[tokio::test]
async fn update_email_move_returns_new_id() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_email(Parameters(UpdateEmailParams {
            email_id: EMAIL_ID.to_string(),
            move_to: Some("Archive".to_string()),
            mark_read: None, flag: None, add_categories: None,
            remove_categories: None, importance: None,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["id"], "new-entry|store-1");
    assert_eq!(v["status"], "updated");
    assert_eq!(v["changed"], serde_json::json!(["move_to"]));
}

#[tokio::test]
async fn update_email_state_only_keeps_same_id_and_lists_changes() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_email(Parameters(UpdateEmailParams {
            email_id: EMAIL_ID.to_string(),
            move_to: None,
            mark_read: Some(true),
            flag: Some("follow_up".to_string()),
            add_categories: Some(vec!["Work".to_string()]),
            remove_categories: None,
            importance: Some("high".to_string()),
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    // No move → id unchanged.
    assert_eq!(v["id"], EMAIL_ID);
    assert_eq!(v["changed"], serde_json::json!(["mark_read", "flag", "add_categories", "importance"]));
    // The client saw the full update.
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "update_email");
    assert_eq!(args["flag"], "follow_up");
    assert_eq!(args["importance"], "high");
}
```

- [ ] **Step 9: Update the `README.md` tool list**

Replace the line:

```
- `move_email` — move an email to a different folder
```

with:

```
- `update_email` — change an existing email: move to a folder, mark read/unread, flag (follow_up/complete/clear), add/remove categories, set importance
```

- [ ] **Step 10: Build and test (green)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished` with zero warnings.
Run: `cargo test 2>&1 | Select-String "test result|update_email"`
Expected: all unit + tool tests pass; the two new `update_email_*` tests show `ok`. (The `todo!()` in `client.rs` is never reached because unit tests use the fake and live tests are `#[ignore]`d.)

- [ ] **Step 11: Commit**

```bash
git add src/outlook/mod.rs src/outlook/fake.rs src/outlook/client.rs src/server.rs tests/tools.rs README.md
git commit -m "Add update_email tool and retire move_email (interface + fake + tool layer)"
```

---

### Task 2: Real COM implementation in `WindowsOutlookClient`

Fill in the `todo!()` stub with real COM: apply state changes first, `move_to` last, building the `changed` list.

**Files:**
- Modify: `src/constants.rs` (add flag constants)
- Modify: `src/outlook/client.rs` (implement `update_email`)

**Interfaces:**
- Consumes: `EmailUpdate` (Task 1), `importance_name_to_id` (`src/constants.rs`), `get_item_categories`/`set_item_categories` (`src/outlook/com.rs`), and the `client.rs` plumbing `with_com`, `mapi()`, `get_item`, `resolve_folder`, `to_disp`, `make_id`, plus COM helpers `get_property`, `put_property`, `call_method`, `variant_from_i32`, `variant_from_bool`, `variant_to_string`.
- Produces: real `update_email` returning `{"status":"updated","id":<current-or-new>,"changed":[...]}`.

- [ ] **Step 1: Add flag constants to `src/constants.rs`**

Add after the `OlImportance` block (after line 42):

```rust
// OlFlagStatus (MailItem.FlagStatus)
pub const OL_NO_FLAG: i32 = 0;
pub const OL_FLAG_COMPLETE: i32 = 1;
pub const OL_FLAG_MARKED: i32 = 2;

// OlMarkInterval (MailItem.MarkAsTask)
pub const OL_MARK_NO_DATE: i32 = 0;
```

- [ ] **Step 2: Implement `update_email` in `src/outlook/client.rs`**

Replace the `todo!()` stub from Task 1 with:

```rust
    fn update_email(&self, u: EmailUpdate) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &u.email_id)?;
            let mut changed: Vec<&str> = Vec::new();

            // ---- state changes first (they address the item by its current id) ----

            if let Some(read) = u.mark_read {
                // UnRead is the inverse of "read".
                put_property(&item, "UnRead", variant_from_bool(!read))?;
                changed.push("mark_read");
            }

            if let Some(flag) = &u.flag {
                match flag.to_lowercase().as_str() {
                    "follow_up" => {
                        // MarkAsTask flags for follow-up with no due date.
                        call_method(&item, "MarkAsTask", &mut [variant_from_i32(c::OL_MARK_NO_DATE)])?;
                    }
                    "complete" => {
                        put_property(&item, "FlagStatus", variant_from_i32(c::OL_FLAG_COMPLETE))?;
                    }
                    "clear" => {
                        // ClearTaskFlag removes the follow-up flag entirely.
                        call_method(&item, "ClearTaskFlag", &mut [])?;
                    }
                    other => {
                        return Err(ToolError::new(format!(
                            "invalid flag {other:?}: expected \"follow_up\", \"complete\", or \"clear\""
                        )));
                    }
                }
                call_method(&item, "Save", &mut [])?;
                changed.push("flag");
            }

            // Categories: read the current set once, then add/remove against it,
            // so tagging never wipes existing categories.
            if u.add_categories.is_some() || u.remove_categories.is_some() {
                let mut cats = get_item_categories(&item);
                if let Some(add) = &u.add_categories {
                    for a in add {
                        if !cats.iter().any(|c| c.eq_ignore_ascii_case(a)) {
                            cats.push(a.clone());
                        }
                    }
                    changed.push("add_categories");
                }
                if let Some(remove) = &u.remove_categories {
                    cats.retain(|c| !remove.iter().any(|r| r.eq_ignore_ascii_case(c)));
                    changed.push("remove_categories");
                }
                set_item_categories(&item, &cats)?;
                call_method(&item, "Save", &mut [])?;
            }

            if let Some(imp) = &u.importance {
                let id = c::importance_name_to_id(imp).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid importance {imp:?}: expected \"low\", \"normal\", or \"high\""
                    ))
                })?;
                put_property(&item, "Importance", variant_from_i32(id))?;
                call_method(&item, "Save", &mut [])?;
                changed.push("importance");
            }

            // ---- move last (Move changes the EntryID) ----

            let id = if let Some(dest) = &u.move_to {
                let target = resolve_folder(&ns, Some(dest))?;
                let moved = to_disp(call_method(
                    &item, "Move", &mut [VARIANT::from(target.clone())],
                )?)?;
                changed.push("move_to");
                make_id(&moved)? // EntryID changed — return the new id.
            } else {
                u.email_id.clone()
            };

            Ok(json!({"status": "updated", "id": id, "changed": changed}))
        })
    }
```

Note on the `c::` alias and `VARIANT` import: `move_email` already used `constants` as `c` and `VARIANT` in this file, so both are in scope. If `variant_from_bool` is not yet imported in `client.rs`, add it to the `com::{...}` use list (it is defined in `com.rs`).

- [ ] **Step 3: Build (green, zero warnings)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings. The `todo!()` is gone.

- [ ] **Step 4: Run the full test suite (unit + tool tests unaffected, still green)**

Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass. (No new unit test here — the real COM path is covered by the live test in Task 3.)

- [ ] **Step 5: Commit**

```bash
git add src/constants.rs src/outlook/client.rs
git commit -m "Implement update_email real COM (mark_read, flag, categories, importance, move-last)"
```

---

### Task 3: Live test

An `#[ignore]`d end-to-end test against real Outlook: create a draft, update several fields, verify, move it, then delete for cleanup.

**Files:**
- Modify: `tests/live_outlook.rs`

- [ ] **Step 1: Add the live test**

Add before the final `send_with_missing_attachment_errors_before_sending` test (imports `EmailUpdate` — the file already imports the outlook module types; add `EmailUpdate` to the existing `use outlook_mcp_rs::outlook::{...}` line):

```rust
#[test]
#[ignore]
fn update_email_applies_state_then_moves() {
    let c = WindowsOutlookClient::new();
    // A draft is a safe, disposable target (never sent).
    let created = c.create_draft(
        vec!["nobody@example.invalid".to_string()],
        "outlook-mcp-rs update_email live test".to_string(),
        "body".to_string(),
        None, None, false, None,
    ).expect("create_draft");
    let id = created["id"].as_str().expect("draft id").to_string();

    // Apply state changes only (no move yet) so we can read them back by the same id.
    let res = c.update_email(EmailUpdate {
        email_id: id.clone(),
        move_to: None,
        mark_read: Some(true),
        flag: Some("follow_up".to_string()),
        add_categories: Some(vec!["Work".to_string()]),
        remove_categories: None,
        importance: Some("high".to_string()),
    }).expect("update_email state");
    assert_eq!(res["status"], "updated");
    assert_eq!(res["id"], id); // no move → id unchanged
    let changed = res["changed"].as_array().unwrap();
    assert!(changed.iter().any(|v| v == "importance"));
    assert!(changed.iter().any(|v| v == "flag"));

    // Verify importance + category landed.
    let detail = c.get_email(id.clone(), false).expect("get_email");
    let dv = serde_json::to_value(&detail).unwrap();
    assert_eq!(dv["summary"]["importance"], "high");
    assert!(dv["summary"]["categories"].as_array().unwrap().iter().any(|v| v == "Work"));
    // mark_read(true) → the item must now read as read (unread == false).
    assert_eq!(dv["summary"]["unread"], false);

    // A standalone mark_read (no other field, so nothing else Saves afterward)
    // must still persist — set it back to unread and confirm it stuck.
    let unread = c.update_email(EmailUpdate {
        email_id: id.clone(),
        mark_read: Some(false),
        ..Default::default()
    }).expect("update_email mark unread");
    assert_eq!(unread["changed"], serde_json::json!(["mark_read"]));
    let redetail = c.get_email(id.clone(), false).expect("get_email after unread");
    let rv = serde_json::to_value(&redetail).unwrap();
    assert_eq!(rv["summary"]["unread"], true);

    // Now move it; the id must change, then delete via the new id for cleanup.
    let moved = c.update_email(EmailUpdate {
        email_id: id.clone(),
        move_to: Some("Deleted Items".to_string()),
        ..Default::default()
    }).expect("update_email move");
    assert_eq!(moved["changed"], serde_json::json!(["move_to"]));
    let new_id = moved["id"].as_str().expect("moved id").to_string();
    c.delete_email(new_id).expect("cleanup delete");
}
```

- [ ] **Step 2: Confirm compile + ignored**

Run: `cargo build --tests 2>&1 | Select-Object -Last 2` → `Finished`.
Run: `cargo test --test live_outlook 2>&1 | Select-String "update_email_applies_state_then_moves"` → `ignored`.

- [ ] **Step 3: (If Outlook available) run live**

Run: `cargo test --test live_outlook -- --ignored update_email_applies_state_then_moves 2>&1 | Select-Object -Last 8`
Expected: `test result: ok. 1 passed`. If the `importance`/`categories` read-back assertions fail, the property write didn't stick — investigate against real Outlook (this is the point of the live test). Skip this step only if no Outlook is available.

- [ ] **Step 4: Commit**

```bash
git add tests/live_outlook.rs
git commit -m "Add live update_email round-trip test"
```

---

## Self-Review

**1. Spec coverage** (spec §`update_email`):
- `email_id` (required) ✅ struct field
- `move_to` ✅ applied last, returns new id
- `mark_read` true/false ✅ via `UnRead` inverse
- `flag` follow_up/complete/clear ✅ via MarkAsTask / FlagStatus=complete / ClearTaskFlag
- `add_categories`/`remove_categories` with add/remove (non-destructive) semantics ✅ read-modify-write against current set, case-insensitive dedup
- `importance` low/normal/high ✅ via `importance_name_to_id`
- Behavior "state first, move last; return `{status:'updated', id:<current-or-new>, changed:[...]}`" ✅ Task 2 ordering + return
- Retire `move_email` ✅ trait, fake, tool, params, test, README all swapped in Task 1

**2. Placeholder scan:** The only `todo!()` is the deliberate Task 1 → Task 2 handoff, replaced in Task 2 Step 2. No TBD/TODO copy left after Task 2.

**3. Type consistency:** `EmailUpdate` field names/types identical across mod.rs, fake.rs, client.rs, server.rs (`UpdateEmailParams` → `EmailUpdate` mapping is 1:1), and `tests/`. `changed` ordering is identical in fake (Task 1 Step 4) and real (Task 2 Step 2): mark_read, flag, add_categories, remove_categories, importance, move_to. `importance_name_to_id` matches the existing signature in `constants.rs`. `EMAIL_ID` fake id and `"new-entry|store-1"` moved-id match the existing fake conventions.

## Execution Handoff

Plan 5 of 12. Models: T1 sonnet, T2 opus (COM state mutation), T3 haiku. After all three are green with zero warnings, controller pushes `main` → Plan 6 (Calendar finder). Remember the trait-ripple checklist: mod.rs + client.rs + fake.rs + server.rs + tests/tools.rs (+ scan live_outlook.rs).
