# v2 Plan 2 — Email finder (merge search + filters) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Merge `search_emails` into `list_emails` and give it rich composable filters (`query`, `from`, `category`, `received_after/before`, `since_days`, `has_attachments`, `flagged`, `high_importance`), so there is one powerful email-finder tool.

**Architecture:** Replace the trait's two methods (`list_emails(folder,count,unread_only)` + `search_emails(...)`) with a single `list_emails(&self, q: EmailQuery)` taking a filter struct. The Windows client applies cheap filters as sequential COM `Restrict` calls (they AND together) and the fuzzy ones (`category`, `has_attachments`) client-side while collecting summaries. `search_emails` is removed from the trait, both impls, the server, and its params struct.

**Tech Stack:** Rust 2024, `windows` 0.62.2 COM, `chrono`, `serde`/`schemars`.

**Depends on:** Plan 1 (Foundations) — complete.

## Global Constraints

- Target crate: `C:\Users\adamk\projects\outlook-mcp-rs`.
- Trait has TWO implementors: `WindowsOutlookClient` (`src/outlook/client.rs`) and `FakeOutlookClient` (`src/outlook/fake.rs`). Every trait change touches both + `tests/tools.rs`.
- Count is clamped `1..=MAX_EMAIL_COUNT` (50). Default folder `"inbox"`, default count 10 — defaults set in the server params struct, not the trait.
- DASL single quotes are escaped by doubling (`query.replace('\'', "''")`) — existing pattern, preserve it.
- Fuzzy filters done client-side: `category` (case-insensitive match against the item's categories, which `email_summary` already reads), `has_attachments` (compare `has_attachments` bool on the built summary).
- Commit after every task; `cargo test` green before each commit. No push steps here (controller pushes at plan end).

---

### Task 1: Introduce `EmailQuery`, change the trait, keep everything compiling

**Files:**
- Modify: `src/outlook/mod.rs` (add `pub struct EmailQuery`; change trait: replace `list_emails` + remove `search_emails`)
- Modify: `src/outlook/fake.rs` (new `list_emails(EmailQuery)`, remove `search_emails`)
- Modify: `src/server.rs` (update `ListEmailsParams`, build an `EmailQuery`, remove `SearchEmailsParams` + the `search_emails` tool)
- Modify: `tests/tools.rs` (update `list_emails*` tests to the new struct; remove/convert the `search_emails` test)
- Modify: `src/outlook/client.rs` (temporary: make `list_emails` take `EmailQuery` but keep current behavior; delete `search_emails` impl — real filtering lands in Task 2)

**Interfaces:**
- Produces: `pub struct EmailQuery { query: Option<String>, folder: String, count: i32, unread_only: bool, from: Option<String>, category: Option<String>, received_after: Option<String>, received_before: Option<String>, since_days: Option<i32>, has_attachments: Option<bool>, flagged: bool, high_importance: bool }` and trait method `fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError>`.

- [ ] **Step 1: Add `EmailQuery` and change the trait in `src/outlook/mod.rs`**

Add near the top of the trait file (after the `use` lines, before `pub trait OutlookClient`):

```rust
/// All filters for `list_emails`. All optional except `folder`/`count`
/// (which the server fills with defaults). Supplying several ANDs them.
#[derive(Debug, Clone)]
pub struct EmailQuery {
    pub query: Option<String>,
    pub folder: String,
    pub count: i32,
    pub unread_only: bool,
    pub from: Option<String>,
    pub category: Option<String>,
    pub received_after: Option<String>,
    pub received_before: Option<String>,
    pub since_days: Option<i32>,
    pub has_attachments: Option<bool>,
    pub flagged: bool,
    pub high_importance: bool,
}
```

In the trait, replace the two lines:
```rust
    fn list_emails(&self, folder: String, count: i32, unread_only: bool)
        -> Result<Vec<EmailSummary>, ToolError>;
    fn search_emails(&self, query: String, folder: String, count: i32,
        since_days: Option<i32>) -> Result<Vec<EmailSummary>, ToolError>;
```
with:
```rust
    fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError>;
```

- [ ] **Step 2: Update the fake in `src/outlook/fake.rs`**

Replace the fake's `list_emails` and delete its `search_emails`. The fake records the call and returns canned data. New `list_emails` records the full query as JSON so tests can assert forwarding:

```rust
    fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError> {
        self.record("list_emails", json!({
            "query": q.query, "folder": q.folder, "count": q.count,
            "unread_only": q.unread_only, "from": q.from, "category": q.category,
            "received_after": q.received_after, "received_before": q.received_before,
            "since_days": q.since_days, "has_attachments": q.has_attachments,
            "flagged": q.flagged, "high_importance": q.high_importance,
        }))?;
        Ok(vec![EmailSummary {
            id: EMAIL_ID.into(), subject: "Hello".into(), sender: "Ada".into(),
            sender_email: "".into(), to: "".into(), received: None,
            unread: true, has_attachments: false,
            categories: vec!["Work".to_string()],
        }])
    }
```
Delete the entire `fn search_emails(...)` from the fake. Add `use crate::outlook::EmailQuery;` (or `super::EmailQuery`) to the fake's imports.

- [ ] **Step 3: Update the server in `src/server.rs`**

Replace `ListEmailsParams` with the full filter set and delete `SearchEmailsParams`:
```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListEmailsParams {
    #[serde(default)]
    pub query: Option<String>,
    #[serde(default = "default_folder")]
    pub folder: String,
    #[serde(default = "default_count")]
    pub count: i32,
    #[serde(default)]
    pub unread_only: bool,
    #[serde(default)]
    pub from: Option<String>,
    #[serde(default)]
    pub category: Option<String>,
    #[serde(default)]
    pub received_after: Option<String>,
    #[serde(default)]
    pub received_before: Option<String>,
    #[serde(default)]
    pub since_days: Option<i32>,
    #[serde(default)]
    pub has_attachments: Option<bool>,
    #[serde(default)]
    pub flagged: bool,
    #[serde(default)]
    pub high_importance: bool,
}
```
In the `list_emails` tool method, build an `EmailQuery` from the params and pass it. Import `EmailQuery`:
```rust
use crate::outlook::{EmailQuery, OutlookClient};
```
```rust
    #[tool(description = "Find emails in a folder with optional text query and filters (sender, category, date range, attachments, flagged, importance).")]
    pub async fn list_emails(
        &self,
        Parameters(p): Parameters<ListEmailsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = EmailQuery {
            query: p.query, folder: p.folder, count: p.count, unread_only: p.unread_only,
            from: p.from, category: p.category, received_after: p.received_after,
            received_before: p.received_before, since_days: p.since_days,
            has_attachments: p.has_attachments, flagged: p.flagged,
            high_importance: p.high_importance,
        };
        let result = run_blocking(move || client.list_emails(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```
Delete the entire `search_emails` `#[tool]` method and the `SearchEmailsParams` struct.

- [ ] **Step 4: Temporarily adapt `src/outlook/client.rs` (real filtering comes in Task 2)**

Replace the client's `list_emails` signature to take `EmailQuery`, preserving today's behavior (folder/count/unread only) so the build passes; delete the client's `search_emails`:
```rust
    fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let count = q.count.clamp(1, MAX_EMAIL_COUNT);
            let folder_obj = resolve_folder(&ns, Some(&q.folder))?;
            let mut items = to_disp(get_property(&folder_obj, "Items")?)?;
            if q.unread_only {
                items = to_disp(call_method(&items, "Restrict",
                    &mut [variant_from_str("[UnRead] = True")])?)?;
            }
            call_method(&items, "Sort",
                &mut [variant_from_str("[ReceivedTime]"), variant_from_bool(true)])?;
            collect_summaries(&items, count)
        })
    }
```
Add `use crate::outlook::EmailQuery;` to client.rs imports. Delete the entire `fn search_emails(...)` impl.

- [ ] **Step 5: Update `tests/tools.rs`**

- `list_emails_passes_arguments` / `list_emails_uses_defaults`: build the new `ListEmailsParams` (all the new fields default to `None`/`false`). For defaults, use `serde_json::from_value(json!({})).unwrap()` to get a `ListEmailsParams` with defaults, then call. Assert the recorded call's `folder`/`count`/`unread_only` still match; the extra keys are present as null/false.
- `client_error_propagates_as_tool_error` uses `list_emails` — update to the new params (defaults).
- `list_emails_returns_categories` (from Plan 1) — update its `ListEmailsParams` literal to include the new fields (or switch to `serde_json::from_value(json!({})).unwrap()`).
- Delete `search_emails_passes_query_and_since_days` (the tool no longer exists) OR convert it to a `list_emails` call passing `query`/`since_days` and asserting they're recorded.

Example updated default test:
```rust
#[tokio::test]
async fn list_emails_uses_defaults() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({})).unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_emails");
    assert_eq!(args["folder"], "inbox");
    assert_eq!(args["count"], 10);
    assert_eq!(args["unread_only"], false);
}
```
And a new forwarding test for a filter:
```rust
#[tokio::test]
async fn list_emails_forwards_query_and_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListEmailsParams = serde_json::from_value(json!({
        "query": "invoice", "from": "ada@x.com", "category": "Work",
        "since_days": 30, "has_attachments": true, "flagged": true, "high_importance": true
    })).unwrap();
    server.list_emails(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["query"], "invoice");
    assert_eq!(args["from"], "ada@x.com");
    assert_eq!(args["category"], "Work");
    assert_eq!(args["since_days"], 30);
    assert_eq!(args["has_attachments"], true);
    assert_eq!(args["flagged"], true);
    assert_eq!(args["high_importance"], true);
}
```

Remove the `SearchEmailsParams` import if present.

- [ ] **Step 6: Build then run the full suite**

Run: `cargo build` → clean. Run: `cargo test` → all pass (the new forwarding test + updated defaults test).

- [ ] **Step 7: Commit**

```bash
git add src/outlook/mod.rs src/outlook/fake.rs src/server.rs src/outlook/client.rs tests/tools.rs
git commit -m "Merge search_emails into list_emails with EmailQuery filter struct"
```

---

### Task 2: Implement the real filtering in the Windows client

**Files:**
- Modify: `src/outlook/client.rs` (the `list_emails` impl — apply all filters)

**Interfaces:**
- Consumes: `EmailQuery` (Task 1), `collect_summaries`, `resolve_folder`, `jet_datetime`, existing COM helpers.
- Produces: `list_emails` that honors every filter.

- [ ] **Step 1: Replace the client `list_emails` body with full filtering**

Cheap filters become sequential `Restrict` calls (they AND); `category` and `has_attachments` are filtered client-side. Replace the Task-1 placeholder body:

```rust
    fn list_emails(&self, q: EmailQuery) -> Result<Vec<EmailSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let count = q.count.clamp(1, MAX_EMAIL_COUNT);
            let folder_obj = resolve_folder(&ns, Some(&q.folder))?;
            let mut items = to_disp(get_property(&folder_obj, "Items")?)?;

            // Text query: DASL @SQL across subject/sender/body (escaped).
            if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
                let e = query.replace('\'', "''");
                let dasl = format!(
                    "@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{e}%' \
                     OR \"urn:schemas:httpmail:fromname\" LIKE '%{e}%' \
                     OR \"urn:schemas:httpmail:textdescription\" LIKE '%{e}%')"
                );
                items = to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&dasl)])?)?;
            }
            // Sender: DASL @SQL against fromname + fromemail.
            if let Some(from) = q.from.as_deref().filter(|s| !s.is_empty()) {
                let e = from.replace('\'', "''");
                let dasl = format!(
                    "@SQL=(\"urn:schemas:httpmail:fromname\" LIKE '%{e}%' \
                     OR \"urn:schemas:httpmail:fromemail\" LIKE '%{e}%')"
                );
                items = to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&dasl)])?)?;
            }
            if q.unread_only {
                items = to_disp(call_method(&items, "Restrict",
                    &mut [variant_from_str("[UnRead] = True")])?)?;
            }
            if q.flagged {
                // FlagStatus 2 = flagged/marked.
                items = to_disp(call_method(&items, "Restrict",
                    &mut [variant_from_str("[FlagStatus] = 2")])?)?;
            }
            if q.high_importance {
                items = to_disp(call_method(&items, "Restrict",
                    &mut [variant_from_str("[Importance] = 2")])?)?;
            }
            // Date filters: since_days (relative), received_after/before (absolute).
            if q.since_days.is_some_and(|d| d != 0) {
                let cutoff = chrono::Local::now().naive_local()
                    - chrono::Duration::days(q.since_days.unwrap() as i64);
                let f = format!("[ReceivedTime] >= '{}'", jet_datetime(&cutoff));
                items = to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&f)])?)?;
            }
            if let Some(after) = q.received_after.as_deref().filter(|s| !s.is_empty()) {
                let dt = parse_dt(after, "received_after")?;
                let f = format!("[ReceivedTime] >= '{}'", jet_datetime(&dt));
                items = to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&f)])?)?;
            }
            if let Some(before) = q.received_before.as_deref().filter(|s| !s.is_empty()) {
                let dt = parse_dt(before, "received_before")?;
                let f = format!("[ReceivedTime] <= '{}'", jet_datetime(&dt));
                items = to_disp(call_method(&items, "Restrict", &mut [variant_from_str(&f)])?)?;
            }

            call_method(&items, "Sort",
                &mut [variant_from_str("[ReceivedTime]"), variant_from_bool(true)])?;

            // Client-side fuzzy filters: category + has_attachments. Iterate,
            // build each summary, keep it only if it passes, stop at count.
            let cat_want = q.category.as_deref().map(|c| c.to_lowercase());
            let total = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let mut results = Vec::new();
            for i in 1..=total {
                let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                let summary = email_summary(&item)?;
                if let Some(want) = &cat_want {
                    if !summary.categories.iter().any(|c| c.to_lowercase() == *want) {
                        continue;
                    }
                }
                if let Some(want_att) = q.has_attachments {
                    if summary.has_attachments != want_att {
                        continue;
                    }
                }
                results.push(summary);
                if results.len() as i32 >= count {
                    break;
                }
            }
            Ok(results)
        })
    }
```
Note: `parse_dt` already exists in `client.rs` (added for calendar in v1). Confirm it's in scope (module-level `fn parse_dt`); if not, use the same one the calendar methods use.

- [ ] **Step 2: Build**

Run: `cargo build` → clean, no warnings. (No new fake-client test — the fake returns canned data; real filtering is verified live in Step 3/Task 3.)

- [ ] **Step 3: Run the suite**

Run: `cargo test` → all still green (fake-backed tests unaffected; this task only changes the real COM path).

- [ ] **Step 4: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Implement list_emails filtering (query, from, dates, flagged, importance, category, attachments)"
```

---

### Task 3: Live verification + docs

**Files:**
- Modify: `tests/live_outlook.rs` (add an `#[ignore]`d filter smoke test)

**Interfaces:**
- Consumes: the real `WindowsOutlookClient::list_emails`.

- [ ] **Step 1: Add a live filter test**

```rust
#[test]
#[ignore]
fn list_emails_query_filter_narrows_results() {
    use outlook_mcp_rs::outlook::EmailQuery;
    let c = WindowsOutlookClient::new();
    let all = c.list_emails(EmailQuery {
        query: None, folder: "inbox".into(), count: 25, unread_only: false,
        from: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("plain list should work");
    // A query that almost certainly matches nothing should return <= all.
    let filtered = c.list_emails(EmailQuery {
        query: Some("zzqx-improbable-token-9137".into()),
        folder: "inbox".into(), count: 25, unread_only: false,
        from: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("query list should work");
    assert!(filtered.len() <= all.len());
}
```

- [ ] **Step 2: Confirm plain `cargo test` still ignores it**

Run: `cargo test 2>&1 | grep list_emails_query_filter`
Expected: shows `... ignored`.

- [ ] **Step 3: (If Outlook available) run it live**

Run: `cargo test --test live_outlook -- --ignored list_emails_query_filter_narrows_results`
Expected: PASS. If Outlook isn't available, skip — it's `#[ignore]`d and CI won't run it.

- [ ] **Step 4: Commit**

```bash
git add tests/live_outlook.rs
git commit -m "Add live filter smoke test for list_emails"
```

---

## Self-Review

- **Spec coverage:** Email finder — `search_emails` merged ✅ (T1 removes it, T2 folds `query` in); filters `from`/`category`/`received_after`/`received_before`/`since_days`/`has_attachments`/`flagged`/`high_importance` ✅ (T2); `categories` already in output from Plan 1. 
- **Placeholder scan:** none — full code in every step.
- **Type consistency:** `EmailQuery` fields defined in T1 match the client reads in T2 and the server build in T1 Step 3; `list_emails(EmailQuery)` signature consistent across trait/fake/client/server.
- **Retirement completeness:** `search_emails` removed from trait (T1.1), fake (T1.2), server tool + params (T1.3), client (T1.4), tests (T1.5) — no dangling references.

## Execution Handoff

Plan 2 of 12. After it ships green, controller pushes to main and proceeds to Plan 3 (Compose attachments). Execute with subagent-driven-development (model per task: T1 sonnet, T2 opus [COM filtering], T3 sonnet).
