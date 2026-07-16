# Tasks CRUD Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Bring the Tasks tools to full CRUD: add filters to `list_tasks`, add 3 fields to `create_task`, add a new `update_task` that absorbs (retires) `complete_task`, and add a new `delete_task`.

**Architecture:** `list_tasks` gains a `TaskQuery` struct (mirroring `EventQuery`) and a pure `task_matches` predicate (mirroring `event_matches`) for client-side filtering, since Outlook's Tasks folder has no established DASL-`Restrict` text-search precedent in this codebase — `include_completed` keeps its existing server-side `Restrict`, everything else filters client-side after `task_summary` builds each row. `update_task` follows `update_email`'s exact shape (a `TaskUpdate` struct, a `changed: Vec<&str>` accumulator, one `Save()` per logical group of field writes) and fully retires `complete_task` as a standalone tool — this project already has a clean precedent for "new tool absorbs an old one" from Plan 5 (`update_email` absorbing `move_email`).

**Tech Stack:** Rust, `windows` crate 0.62.2 (Win32 COM/`IDispatch::Invoke`), `rmcp` 2.1.0 tool macros, `chrono`, `serde`/`serde_json`.

## Global Constraints

- `list_tasks` keeps `include_completed`; adds `category`, `importance` (`"low"`/`"normal"`/`"high"`), `query` (text match on subject + body). Output already has `categories` (see Task 1 Step 0 — no change needed there); `status`/`importance` are already friendly words (`task_status_word`/`importance_word`, both pre-existing in `src/friendly.rs`).
- `create_task` keeps `subject`/`body`/`due_date`/`importance`; adds `categories` (assign on creation), `start_date` (sets `StartDate`), `reminder_time` (ISO datetime; sets `ReminderSet` + `ReminderTime` — task reminders are an **absolute time**, unlike appointment reminders which are minutes-before-start).
- `update_task` (new; absorbs `complete_task`): `task_id` (required), all else optional — `mark_complete` (`true`=complete/100%, `false`=reopen), `subject`, `body`, `due_date`, `start_date`, `importance`, `add_categories`/`remove_categories`, `percent_complete` (0-100), `reminder_time`. Returns `{ "status": "updated", "id", "changed": [...] }`. Retires standalone `complete_task` (= `update_task` with `mark_complete: true`) — remove `complete_task` entirely from the trait/both implementors/tool layer/tests, not just deprecate it.
- `delete_task` (new): soft-delete a task by id (to Deleted Items), mirrors `delete_email`'s shape exactly.
- The `OutlookClient` trait (`src/outlook/mod.rs`) has two implementors — `WindowsOutlookClient` (`client.rs`) and `FakeOutlookClient` (`fake.rs`) — plus `src/server.rs` (MCP tool layer) and `tests/tools.rs` (fake-backed tests); every trait change touches all four.
- `importance` values are validated via the pre-existing `c::importance_name_to_id` (`src/constants.rs`) — reuse it, don't reinvent.
- Per this project's model policy: Task 1 (filters, mirrors an existing pattern closely) is standard-tier; Task 2 (mechanical param additions) is cheap-tier; Task 3 (new struct-based tool absorbing an old one, most design judgment) is standard-tier; Task 4 (mechanical, mirrors `delete_email`) is cheap-tier; Task 5 (live tests) is standard-tier.

---

### Task 1: `list_tasks` filters (`category`, `importance`, `query`)

**Files:**
- Modify: `src/outlook/mod.rs` (replace `list_tasks`'s trait signature; add `TaskQuery` struct)
- Modify: `src/outlook/client.rs` (real COM: wire `TaskQuery`, add `task_matches`)
- Modify: `src/outlook/fake.rs` (record the new query fields)
- Modify: `src/server.rs` (`ListTasksParams` gains 3 fields)
- Test: `tests/tools.rs`

**Interfaces:**
- Produces (consumed by Tasks 2-5): `TaskQuery { pub include_completed: bool, pub category: Option<String>, pub importance: Option<String>, pub query: Option<String> }` (in `mod.rs`, alongside `EventQuery`); `fn task_matches(summary: &TaskSummary, q: &TaskQuery) -> bool` (in `client.rs`, alongside `event_matches`).

- [ ] **Step 1: Replace the trait signature**

In `src/outlook/mod.rs`, find:

```rust
    fn list_tasks(&self, include_completed: bool)
        -> Result<Vec<TaskSummary>, ToolError>;
```

Replace with:

```rust
    fn list_tasks(&self, q: TaskQuery) -> Result<Vec<TaskSummary>, ToolError>;
```

Add the `TaskQuery` struct right after `EventQuery`'s closing `}` (before `CreateEventInput`):

```rust
/// All filters for `list_tasks`. Every field is optional except
/// `include_completed`; supplying several ANDs them. `include_completed`
/// drives a server-side `Restrict`; the rest filter the streamed tasks
/// client-side (there's no established DASL text-search path for the Tasks
/// folder in this codebase, unlike email's `@SQL` queries — same approach
/// `EventQuery`'s `query`/`category` already use).
#[derive(Debug, Clone, Default)]
pub struct TaskQuery {
    pub include_completed: bool,
    pub category: Option<String>,
    pub importance: Option<String>,
    pub query: Option<String>, // text match on subject + body
}
```

Run `cargo build` — expect failures in `client.rs`, `fake.rs`, and `server.rs` (wrong argument count/type for `list_tasks`). This confirms the ripple; Steps 2-4 fix it.

- [ ] **Step 2: Wire `TaskQuery` into `client.rs`, add `task_matches`**

Find `event_matches` in `client.rs` and add `task_matches` right after its closing `}` (before `event_matches`'s callers, anywhere at module scope is fine — place it directly after `event_matches` for locality):

```rust
/// Client-side filter for `list_tasks`'s `category`/`importance`/`query`.
/// `include_completed` is applied earlier via `Restrict`, not here.
fn task_matches(summary: &TaskSummary, q: &TaskQuery) -> bool {
    if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
        let needle = query.to_lowercase();
        if !summary.subject.to_lowercase().contains(&needle) {
            return false;
        }
    }
    if let Some(cat) = q.category.as_deref().filter(|s| !s.is_empty()) {
        let want = cat.to_lowercase();
        if !summary.categories.iter().any(|c| c.to_lowercase() == want) {
            return false;
        }
    }
    if let Some(imp) = q.importance.as_deref().filter(|s| !s.is_empty()) {
        if !summary.importance.eq_ignore_ascii_case(imp) {
            return false;
        }
    }
    true
}
```

Note: the spec says `query` matches "subject + body," but `TaskSummary` (unlike `EmailDetail`) has no `body` field on the summary row — only `get_note`/`get_email`-style detail views carry full body text, and `list_tasks` returns summaries, not details. Match on `subject` only for now (matching what data is actually available on `TaskSummary`); this is a deliberate, minimal scope decision — do not add a COM call to fetch each task's `Body` just to filter on it (that would turn an O(1)-property-read filter into an O(n) extra-COM-call-per-task filter, out of proportion for this feature). Note this decision in your task report.

Replace `list_tasks`'s existing body:

```rust
    fn list_tasks(&self, include_completed: bool) -> Result<Vec<TaskSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let tasks = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_TASKS)],
            )?)?;
            let mut items = to_disp(get_property(&tasks, "Items")?)?;
            if !include_completed {
                items = to_disp(call_method(
                    &items,
                    "Restrict",
                    &mut [variant_from_str("[Complete] = False")],
                )?)?;
            }
            let count = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let mut results = Vec::new();
            for i in 1..=count {
                let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                results.push(task_summary(&item)?);
            }
            Ok(results)
        })
    }
```

with:

```rust
    fn list_tasks(&self, q: TaskQuery) -> Result<Vec<TaskSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let tasks = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_TASKS)],
            )?)?;
            let mut items = to_disp(get_property(&tasks, "Items")?)?;
            if !q.include_completed {
                items = to_disp(call_method(
                    &items,
                    "Restrict",
                    &mut [variant_from_str("[Complete] = False")],
                )?)?;
            }
            let count = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let mut results = Vec::new();
            for i in 1..=count {
                let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                let summary = task_summary(&item)?;
                if task_matches(&summary, &q) {
                    results.push(summary);
                }
            }
            Ok(results)
        })
    }
```

- [ ] **Step 3: Update `fake.rs`**

Find `list_tasks` in `fake.rs` and replace:

```rust
    fn list_tasks(&self, include_completed: bool) -> Result<Vec<TaskSummary>, ToolError> {
        self.record("list_tasks", json!({"include_completed": include_completed}))?;
        Ok(vec![TaskSummary {
            id: TASK_ID.into(), subject: "Buy milk".into(), due_date: None,
            complete: false, status: "not_started".to_string(), importance: "normal".to_string(), categories: vec![],
        }])
    }
```

with:

```rust
    fn list_tasks(&self, q: TaskQuery) -> Result<Vec<TaskSummary>, ToolError> {
        self.record("list_tasks", json!({
            "include_completed": q.include_completed, "category": q.category,
            "importance": q.importance, "query": q.query,
        }))?;
        Ok(vec![TaskSummary {
            id: TASK_ID.into(), subject: "Buy milk".into(), due_date: None,
            complete: false, status: "not_started".to_string(), importance: "normal".to_string(), categories: vec![],
        }])
    }
```

(The fake doesn't simulate filtering — matches the existing `list_events` fake's behavior of recording the query and returning one canned row regardless of filters, verified in `fake.rs`.)

- [ ] **Step 4: Update `server.rs`**

Find `ListTasksParams`:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListTasksParams {
    #[serde(default)]
    pub include_completed: bool,
}
```

Replace with:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListTasksParams {
    #[serde(default)]
    pub include_completed: bool,
    /// Filter to a color category.
    #[serde(default)]
    pub category: Option<String>,
    /// "low" | "normal" | "high".
    #[serde(default)]
    pub importance: Option<String>,
    /// Text match on the task's subject.
    #[serde(default)]
    pub query: Option<String>,
}
```

Find `list_tasks`'s tool method and replace:

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
```

with:

```rust
    #[tool(description = "List Outlook tasks (default: not-yet-completed only). Filter by category, importance, or a text query matching the subject.")]
    pub async fn list_tasks(
        &self,
        Parameters(p): Parameters<ListTasksParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = TaskQuery {
            include_completed: p.include_completed, category: p.category,
            importance: p.importance, query: p.query,
        };
        let result = run_blocking(move || client.list_tasks(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Add `TaskQuery` to `server.rs`'s top-level `use crate::outlook::{...}` import line.

- [ ] **Step 5: Run `cargo build` to confirm everything compiles**

Run: `cargo build`
Expected: 0 errors, 0 warnings.

- [ ] **Step 6: Write fake-backed tool tests**

Open `tests/tools.rs`, find `list_tasks`-related tests (search for `ListTasksParams`), and add these two tests right after the existing one (match this file's exact `#[tokio::test]`/`Arc::new(FakeOutlookClient::new())`/`fake.calls()` pattern):

```rust
#[tokio::test]
async fn list_tasks_forwards_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_tasks(Parameters(ListTasksParams {
            include_completed: true,
            category: Some("Red Category".to_string()),
            importance: Some("high".to_string()),
            query: Some("milk".to_string()),
        }))
        .await
        .unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_tasks");
    assert_eq!(args["include_completed"], true);
    assert_eq!(args["category"], "Red Category");
    assert_eq!(args["importance"], "high");
    assert_eq!(args["query"], "milk");
}

#[tokio::test]
async fn list_tasks_defaults_all_filters_to_none() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListTasksParams = serde_json::from_value(json!({})).unwrap();
    server.list_tasks(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["include_completed"], false);
    assert!(args["category"].is_null());
    assert!(args["importance"].is_null());
    assert!(args["query"].is_null());
}
```

Add `ListTasksParams` to this file's `use outlook_mcp_rs::server::{...}` import block if it's not already imported (it should already be, since the pre-existing `list_tasks` test uses it — check before adding a duplicate).

- [ ] **Step 7: Run the new tests, then the full non-live suite, then commit**

Run: `cargo test --test tools list_tasks`
Expected: all `list_tasks` tests pass (the pre-existing one plus 2 new).

Run: `cargo build` (0 warnings) then `cargo test` (all passing).

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/server.rs tests/tools.rs
git commit -m "Add list_tasks filters: category, importance, query (client-side, mirrors event_matches)"
```

---

### Task 2: `create_task` additions (`categories`, `start_date`, `reminder_time`)

**Files:**
- Modify: `src/outlook/mod.rs` (trait signature)
- Modify: `src/outlook/client.rs`
- Modify: `src/outlook/fake.rs`
- Modify: `src/server.rs`
- Test: `tests/tools.rs`

**Interfaces:**
- Modifies the existing `create_task` signature (all implementors + tool layer + tests ripple together, same as Task 1).

- [ ] **Step 1: Update the trait signature**

In `src/outlook/mod.rs`, find:

```rust
    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String) -> Result<Value, ToolError>;
```

Replace with:

```rust
    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String, categories: Option<Vec<String>>,
        start_date: Option<String>, reminder_time: Option<String>) -> Result<Value, ToolError>;
```

- [ ] **Step 2: Update `client.rs`**

Find `create_task` and replace:

```rust
    fn create_task(
        &self,
        subject: String,
        body: Option<String>,
        due_date: Option<String>,
        importance: String,
    ) -> Result<Value, ToolError> {
        let importance_key = importance.trim().to_lowercase();
        let importance_id = c::importance_name_to_id(&importance_key).ok_or_else(|| {
            ToolError::new(format!(
                "Invalid importance {importance:?}: use 'low', 'normal' or 'high'."
            ))
        })?;
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let task = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_TASK_ITEM)],
            )?)?;
            put_property(&task, "Subject", variant_from_str(&subject))?;
            if let Some(body) = body.as_deref().filter(|b| !b.is_empty()) {
                put_property(&task, "Body", variant_from_str(body))?;
            }
            if let Some(due) = due_date.as_deref().filter(|d| !d.is_empty()) {
                put_property(
                    &task,
                    "DueDate",
                    variant_from_datetime(&parse_dt(due, "due_date")?)?,
                )?;
            }
            put_property(&task, "Importance", variant_from_i32(importance_id))?;
            call_method(&task, "Save", &mut [])?;
            Ok(json!({"status": "created", "id": make_id(&task)?, "subject": subject}))
        })
    }
```

with:

```rust
    fn create_task(
        &self,
        subject: String,
        body: Option<String>,
        due_date: Option<String>,
        importance: String,
        categories: Option<Vec<String>>,
        start_date: Option<String>,
        reminder_time: Option<String>,
    ) -> Result<Value, ToolError> {
        let importance_key = importance.trim().to_lowercase();
        let importance_id = c::importance_name_to_id(&importance_key).ok_or_else(|| {
            ToolError::new(format!(
                "Invalid importance {importance:?}: use 'low', 'normal' or 'high'."
            ))
        })?;
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let task = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_TASK_ITEM)],
            )?)?;
            put_property(&task, "Subject", variant_from_str(&subject))?;
            if let Some(body) = body.as_deref().filter(|b| !b.is_empty()) {
                put_property(&task, "Body", variant_from_str(body))?;
            }
            if let Some(due) = due_date.as_deref().filter(|d| !d.is_empty()) {
                put_property(
                    &task,
                    "DueDate",
                    variant_from_datetime(&parse_dt(due, "due_date")?)?,
                )?;
            }
            if let Some(start) = start_date.as_deref().filter(|d| !d.is_empty()) {
                put_property(
                    &task,
                    "StartDate",
                    variant_from_datetime(&parse_dt(start, "start_date")?)?,
                )?;
            }
            if let Some(reminder) = reminder_time.as_deref().filter(|d| !d.is_empty()) {
                put_property(&task, "ReminderSet", variant_from_bool(true))?;
                put_property(
                    &task,
                    "ReminderTime",
                    variant_from_datetime(&parse_dt(reminder, "reminder_time")?)?,
                )?;
            }
            put_property(&task, "Importance", variant_from_i32(importance_id))?;
            if let Some(cats) = categories.as_ref().filter(|c| !c.is_empty()) {
                set_item_categories(&task, cats)?;
            }
            call_method(&task, "Save", &mut [])?;
            Ok(json!({"status": "created", "id": make_id(&task)?, "subject": subject}))
        })
    }
```

- [ ] **Step 3: Update `fake.rs`**

Find `create_task` and replace:

```rust
    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String) -> Result<Value, ToolError> {
        self.record("create_task",
            json!({"subject": subject, "body": body, "due_date": due_date, "importance": importance}))?;
        Ok(json!({"status": "created", "id": TASK_ID, "subject": subject}))
    }
```

with:

```rust
    fn create_task(&self, subject: String, body: Option<String>,
        due_date: Option<String>, importance: String, categories: Option<Vec<String>>,
        start_date: Option<String>, reminder_time: Option<String>) -> Result<Value, ToolError> {
        self.record("create_task", json!({
            "subject": subject, "body": body, "due_date": due_date, "importance": importance,
            "categories": categories, "start_date": start_date, "reminder_time": reminder_time,
        }))?;
        Ok(json!({"status": "created", "id": TASK_ID, "subject": subject}))
    }
```

- [ ] **Step 4: Update `server.rs`**

Find `CreateTaskParams`:

```rust
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
```

Replace with:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateTaskParams {
    pub subject: String,
    #[serde(default)]
    pub body: Option<String>,
    #[serde(default)]
    pub due_date: Option<String>,
    #[serde(default = "default_importance")]
    pub importance: String,
    /// Category names to assign on creation.
    #[serde(default)]
    pub categories: Option<Vec<String>>,
    #[serde(default)]
    pub start_date: Option<String>,
    /// ISO datetime — an absolute reminder time (unlike appointment
    /// reminders, which are minutes-before-start).
    #[serde(default)]
    pub reminder_time: Option<String>,
}
fn default_importance() -> String { "normal".to_string() }
```

Find `create_task`'s tool method and replace:

```rust
    #[tool(description = "Create a new task.")]
    pub async fn create_task(
        &self,
        Parameters(CreateTaskParams { subject, body, due_date, importance }): Parameters<CreateTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.create_task(subject, body, due_date, importance)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

with:

```rust
    #[tool(description = "Create a new task. reminder_time (ISO datetime) is an absolute reminder time, unlike appointment reminders which are minutes-before-start.")]
    pub async fn create_task(
        &self,
        Parameters(CreateTaskParams { subject, body, due_date, importance, categories, start_date, reminder_time }):
            Parameters<CreateTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move ||
            client.create_task(subject, body, due_date, importance, categories, start_date, reminder_time)
        ).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

- [ ] **Step 5: Run `cargo build`, fix the 2 pre-existing `create_task` call sites**

Run: `cargo build`
Expected: compile errors in `tests/tools.rs` (existing `create_task` tests calling the fake/trait method with the old 4-arg shape) and possibly `tests/live_outlook.rs`. Find every call site (`grep -rn "create_task(" tests/` if unsure) and add `, None, None, None` (or explicit values, if the surrounding test cares) to each — do not change what those tests were already asserting, just extend the call to match the new signature.

- [ ] **Step 6: Write a new fake-backed test for the 3 additions**

Add to `tests/tools.rs`, right after the existing `create_task` test(s):

```rust
#[tokio::test]
async fn create_task_forwards_categories_start_date_and_reminder() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: CreateTaskParams = serde_json::from_value(json!({
        "subject": "Ship it",
        "categories": ["Blue Category"],
        "start_date": "2099-01-01",
        "reminder_time": "2099-01-01T09:00"
    })).unwrap();
    server.create_task(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["categories"], json!(["Blue Category"]));
    assert_eq!(args["start_date"], "2099-01-01");
    assert_eq!(args["reminder_time"], "2099-01-01T09:00");
}
```

- [ ] **Step 7: Run the tests, then the full non-live suite, then commit**

Run: `cargo test --test tools create_task`
Expected: all pass.

Run: `cargo build` (0 warnings) then `cargo test` (all passing).

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/server.rs tests/tools.rs tests/live_outlook.rs
git commit -m "Add create_task additions: categories, start_date, reminder_time"
```

---

### Task 3: `update_task` (new; absorbs and retires `complete_task`)

**Files:**
- Modify: `src/outlook/mod.rs` (remove `complete_task` from the trait, add `TaskUpdate` + `update_task`)
- Modify: `src/outlook/client.rs`
- Modify: `src/outlook/fake.rs`
- Modify: `src/server.rs` (remove `CompleteTaskParams`/`complete_task` tool, add `UpdateTaskParams`/`update_task` tool)
- Test: `tests/tools.rs`

**Interfaces:**
- Produces: `TaskUpdate { pub task_id: String, pub mark_complete: Option<bool>, pub subject: Option<String>, pub body: Option<String>, pub due_date: Option<String>, pub start_date: Option<String>, pub importance: Option<String>, pub add_categories: Option<Vec<String>>, pub remove_categories: Option<Vec<String>>, pub percent_complete: Option<i32>, pub reminder_time: Option<String> }` (in `mod.rs`, alongside `EmailUpdate`).
- Removes: `complete_task` from the trait and both implementors, `CompleteTaskParams`/the `complete_task` tool method from `server.rs`, and any `complete_task`-specific tests in `tests/tools.rs`/`tests/live_outlook.rs` (replace their assertions with `update_task`-based equivalents rather than just deleting coverage — see Step 6).

- [ ] **Step 1: Replace `complete_task` with `update_task` in the trait**

In `src/outlook/mod.rs`, find:

```rust
    fn complete_task(&self, task_id: String) -> Result<Value, ToolError>;
```

Replace with:

```rust
    fn update_task(&self, u: TaskUpdate) -> Result<Value, ToolError>;
```

Add `TaskUpdate` right after `EmailUpdate`'s closing `}` (before `EventQuery`):

```rust
/// All changes `update_task` can apply to one existing task. Every field
/// except `task_id` is optional; supplying several applies all of them.
/// `mark_complete: Some(true)` replaces the retired standalone
/// `complete_task` tool (`= update_task` with `mark_complete: true`);
/// `Some(false)` reopens a completed task, filling the "can't reopen" gap
/// the old `complete_task` had no way to close.
#[derive(Debug, Clone, Default)]
pub struct TaskUpdate {
    pub task_id: String,
    pub mark_complete: Option<bool>,
    pub subject: Option<String>,
    pub body: Option<String>,
    pub due_date: Option<String>,
    pub start_date: Option<String>,
    pub importance: Option<String>,
    pub add_categories: Option<Vec<String>>,
    pub remove_categories: Option<Vec<String>>,
    pub percent_complete: Option<i32>,
    pub reminder_time: Option<String>,
}
```

Run `cargo build` — expect failures in `client.rs`, `fake.rs`, `server.rs`, and test files referencing the now-removed `complete_task`. Steps 2-6 fix each.

- [ ] **Step 2: Implement `update_task` in `client.rs`**

Find `complete_task` in `client.rs` and replace its entire body:

```rust
    fn complete_task(&self, task_id: String) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let task = get_item(&ns, &task_id)?;
            call_method(&task, "MarkComplete", &mut [])?;
            let subject = variant_to_string(&get_property(&task, "Subject")?);
            Ok(json!({"status": "completed", "subject": subject}))
        })
    }
```

with:

```rust
    fn update_task(&self, u: TaskUpdate) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let task = get_item(&ns, &u.task_id)?;
            let mut changed: Vec<&str> = Vec::new();

            if let Some(subject) = &u.subject {
                put_property(&task, "Subject", variant_from_str(subject))?;
                changed.push("subject");
            }
            if let Some(body) = &u.body {
                put_property(&task, "Body", variant_from_str(body))?;
                changed.push("body");
            }
            if let Some(due) = &u.due_date {
                put_property(&task, "DueDate", variant_from_datetime(&parse_dt(due, "due_date")?)?)?;
                changed.push("due_date");
            }
            if let Some(start) = &u.start_date {
                put_property(&task, "StartDate", variant_from_datetime(&parse_dt(start, "start_date")?)?)?;
                changed.push("start_date");
            }
            if let Some(imp) = &u.importance {
                let id = c::importance_name_to_id(imp).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid importance {imp:?}: expected \"low\", \"normal\", or \"high\""
                    ))
                })?;
                put_property(&task, "Importance", variant_from_i32(id))?;
                changed.push("importance");
            }
            if u.add_categories.is_some() || u.remove_categories.is_some() {
                let mut cats = get_item_categories(&task);
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
                set_item_categories(&task, &cats)?;
            }
            if let Some(pct) = u.percent_complete {
                put_property(&task, "PercentComplete", variant_from_i32(pct))?;
                changed.push("percent_complete");
            }
            if let Some(reminder) = &u.reminder_time {
                put_property(&task, "ReminderSet", variant_from_bool(true))?;
                put_property(&task, "ReminderTime", variant_from_datetime(&parse_dt(reminder, "reminder_time")?)?)?;
                changed.push("reminder_time");
            }
            // mark_complete last: MarkComplete() is Outlook's dedicated
            // "finish this task" method (it also sets PercentComplete=100
            // and Status=olTaskComplete), so apply any field edits above
            // to the task's live state first, then finish/reopen it.
            if let Some(complete) = u.mark_complete {
                if complete {
                    call_method(&task, "MarkComplete", &mut [])?;
                } else {
                    put_property(&task, "Complete", variant_from_bool(false))?;
                    put_property(&task, "Status", variant_from_i32(c::OL_TASK_NOT_STARTED))?;
                    put_property(&task, "PercentComplete", variant_from_i32(0))?;
                }
                changed.push("mark_complete");
            }

            call_method(&task, "Save", &mut [])?;
            Ok(json!({"status": "updated", "id": u.task_id, "changed": changed}))
        })
    }
```

- [ ] **Step 3: Update `fake.rs`**

Find `complete_task` and replace:

```rust
    fn complete_task(&self, task_id: String) -> Result<Value, ToolError> {
        self.record("complete_task", json!({"task_id": task_id}))?;
        Ok(json!({"status": "completed"}))
    }
```

with:

```rust
    fn update_task(&self, u: TaskUpdate) -> Result<Value, ToolError> {
        self.record("update_task", json!({
            "task_id": u.task_id, "mark_complete": u.mark_complete, "subject": u.subject,
            "body": u.body, "due_date": u.due_date, "start_date": u.start_date,
            "importance": u.importance, "add_categories": u.add_categories,
            "remove_categories": u.remove_categories, "percent_complete": u.percent_complete,
            "reminder_time": u.reminder_time,
        }))?;
        let mut changed: Vec<&str> = Vec::new();
        if u.mark_complete.is_some() { changed.push("mark_complete"); }
        if u.subject.is_some() { changed.push("subject"); }
        if u.body.is_some() { changed.push("body"); }
        if u.due_date.is_some() { changed.push("due_date"); }
        if u.start_date.is_some() { changed.push("start_date"); }
        if u.importance.is_some() { changed.push("importance"); }
        if u.add_categories.is_some() { changed.push("add_categories"); }
        if u.remove_categories.is_some() { changed.push("remove_categories"); }
        if u.percent_complete.is_some() { changed.push("percent_complete"); }
        if u.reminder_time.is_some() { changed.push("reminder_time"); }
        Ok(json!({"status": "updated", "id": u.task_id, "changed": changed}))
    }
```

Add `TaskUpdate` to `fake.rs`'s existing `use super::{...}` import line.

- [ ] **Step 4: Update `server.rs`**

Find `CompleteTaskParams`:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CompleteTaskParams {
    pub task_id: String,
}
```

Replace with:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct UpdateTaskParams {
    pub task_id: String,
    /// true = mark complete (100%), false = reopen.
    #[serde(default)]
    pub mark_complete: Option<bool>,
    #[serde(default)]
    pub subject: Option<String>,
    #[serde(default)]
    pub body: Option<String>,
    #[serde(default)]
    pub due_date: Option<String>,
    #[serde(default)]
    pub start_date: Option<String>,
    #[serde(default)]
    pub importance: Option<String>,
    #[serde(default)]
    pub add_categories: Option<Vec<String>>,
    #[serde(default)]
    pub remove_categories: Option<Vec<String>>,
    /// 0-100.
    #[serde(default)]
    pub percent_complete: Option<i32>,
    #[serde(default)]
    pub reminder_time: Option<String>,
}
```

Find `complete_task`'s tool method:

```rust
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

Replace with:

```rust
    #[tool(description = "Update an existing task: mark_complete (true=complete, false=reopen), subject, body, due_date, start_date, importance, add/remove categories, percent_complete, reminder_time. Combine any of these in one call.")]
    pub async fn update_task(
        &self,
        Parameters(p): Parameters<UpdateTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let u = TaskUpdate {
            task_id: p.task_id, mark_complete: p.mark_complete, subject: p.subject,
            body: p.body, due_date: p.due_date, start_date: p.start_date,
            importance: p.importance, add_categories: p.add_categories,
            remove_categories: p.remove_categories, percent_complete: p.percent_complete,
            reminder_time: p.reminder_time,
        };
        let result = run_blocking(move || client.update_task(u)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Update `server.rs`'s top-level `use crate::outlook::{...}` import line: remove nothing (there was no `CompleteTaskInput`-style type to remove, `complete_task` took a bare `String`), add `TaskUpdate`.

- [ ] **Step 5: Run `cargo build`, fix remaining call sites**

Run: `cargo build`
Expected: compile errors in `tests/tools.rs` and possibly `tests/live_outlook.rs` referencing `CompleteTaskParams`/`complete_task`. For each:
- If a test exists purely to exercise `complete_task`, rewrite it as an `update_task` call with `mark_complete: Some(true)`, keeping the same assertion intent (task ends up complete). Do not just delete it — Step 6 below gives the exact replacement tests to add; remove the old `complete_task`-specific ones once the new ones exist and cover the same ground.
- Remove `CompleteTaskParams` from any `use` import lists that reference it.

- [ ] **Step 6: Write fake-backed `update_task` tests (replacing `complete_task` coverage)**

Add to `tests/tools.rs`, in the tasks test group:

```rust
#[tokio::test]
async fn update_task_marks_complete() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    server
        .update_task(Parameters(UpdateTaskParams {
            task_id: TASK_ID.to_string(), mark_complete: Some(true),
            subject: None, body: None, due_date: None, start_date: None,
            importance: None, add_categories: None, remove_categories: None,
            percent_complete: None, reminder_time: None,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["mark_complete"], true);
}

#[tokio::test]
async fn update_task_reopens_with_mark_complete_false() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    let result = server
        .update_task(Parameters(UpdateTaskParams {
            task_id: TASK_ID.to_string(), mark_complete: Some(false),
            subject: None, body: None, due_date: None, start_date: None,
            importance: None, add_categories: None, remove_categories: None,
            percent_complete: None, reminder_time: None,
        }))
        .await
        .unwrap();
    let json = result_json(&result);
    assert!(json["changed"].as_array().unwrap().iter().any(|v| v == "mark_complete"));
}

#[tokio::test]
async fn update_task_forwards_field_edits() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    server
        .update_task(Parameters(UpdateTaskParams {
            task_id: TASK_ID.to_string(), mark_complete: None,
            subject: Some("Renamed".to_string()), body: None,
            due_date: None, start_date: None, importance: Some("high".to_string()),
            add_categories: Some(vec!["Red Category".to_string()]), remove_categories: None,
            percent_complete: Some(50), reminder_time: None,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["subject"], "Renamed");
    assert_eq!(args["importance"], "high");
    assert_eq!(args["add_categories"], json!(["Red Category"]));
    assert_eq!(args["percent_complete"], 50);
}
```

Add `UpdateTaskParams` to this file's `use outlook_mcp_rs::server::{...}` import block; remove `CompleteTaskParams` from it.

- [ ] **Step 7: Run the tests, then the full non-live suite, then commit**

Run: `cargo test --test tools update_task`
Expected: 3 new tests pass.

Run: `cargo build` (0 warnings) then `cargo test` (all passing — confirm `complete_task` no longer appears anywhere via `grep -rn complete_task src/ tests/`, which should now only match doc comments/commit messages, not code).

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/server.rs tests/tools.rs tests/live_outlook.rs
git commit -m "Add update_task, retiring complete_task (absorbed as mark_complete: true)"
```

---

### Task 4: `delete_task` (new)

**Files:**
- Modify: `src/outlook/mod.rs`
- Modify: `src/outlook/client.rs`
- Modify: `src/outlook/fake.rs`
- Modify: `src/server.rs`
- Test: `tests/tools.rs`

**Interfaces:**
- Produces: `fn delete_task(&self, task_id: String) -> Result<Value, ToolError>;` on the trait.

- [ ] **Step 1: Add the trait method**

In `src/outlook/mod.rs`, find `update_task`'s trait line (added in Task 3) and add right after it:

```rust
    fn delete_task(&self, task_id: String) -> Result<Value, ToolError>;
```

- [ ] **Step 2: Implement in `client.rs`**

Find `delete_email` in `client.rs` (for reference — `delete_task` mirrors it exactly) and add a new `delete_task` right after `update_task`'s closing `}`:

```rust
    fn delete_task(&self, task_id: String) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &task_id)?;
            let subject = variant_to_string(&get_property(&item, "Subject")?);
            call_method(&item, "Delete", &mut [])?;
            Ok(json!({"status": "deleted", "subject": subject, "note": "Moved to Deleted Items."}))
        })
    }
```

- [ ] **Step 3: Implement in `fake.rs`**

Add right after `update_task`'s closing `}`:

```rust
    fn delete_task(&self, task_id: String) -> Result<Value, ToolError> {
        self.record("delete_task", json!({"task_id": task_id}))?;
        Ok(json!({"status": "deleted", "note": "Moved to Deleted Items."}))
    }
```

- [ ] **Step 4: Add to `server.rs`**

Add right after `UpdateTaskParams`:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteTaskParams {
    pub task_id: String,
}
```

Add the tool method right after `update_task`'s:

```rust
    #[tool(description = "Delete a task (moves it to Deleted Items).")]
    pub async fn delete_task(
        &self,
        Parameters(DeleteTaskParams { task_id }): Parameters<DeleteTaskParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_task(task_id)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

- [ ] **Step 5: Run `cargo build`**

Run: `cargo build`
Expected: 0 errors, 0 warnings.

- [ ] **Step 6: Write a fake-backed test**

Add to `tests/tools.rs`:

```rust
#[tokio::test]
async fn delete_task_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    use outlook_mcp_rs::outlook::fake::TASK_ID;
    let result = server
        .delete_task(Parameters(DeleteTaskParams { task_id: TASK_ID.to_string() }))
        .await
        .unwrap();
    let json = result_json(&result);
    assert_eq!(json["status"], "deleted");
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "delete_task");
    assert_eq!(args["task_id"], TASK_ID);
}
```

Add `DeleteTaskParams` to this file's `use outlook_mcp_rs::server::{...}` import block.

- [ ] **Step 7: Run the test, then the full non-live suite, then commit**

Run: `cargo test --test tools delete_task`
Expected: 1 test passes.

Run: `cargo build` (0 warnings) then `cargo test` (all passing).

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/server.rs tests/tools.rs
git commit -m "Add delete_task"
```

---

### Task 5: Live tests and `TESTING.md`

**Files:**
- Modify: `tests/live_outlook.rs`
- Modify: `TESTING.md`

**Interfaces:**
- Consumes everything from Tasks 1-4: `TaskQuery`, the expanded `create_task` signature, `TaskUpdate`/`update_task`, `delete_task`.

- [ ] **Step 1: Add the live tests**

Open `tests/live_outlook.rs`. Add `TaskQuery`, `TaskUpdate` to its `use outlook_mcp_rs::outlook::{...}` import line. Find the existing task-related live test(s) (search for `create_task` or `TASK`) to see the current cleanup convention used there, then append these tests at the end of the file, matching that convention:

```rust
#[test]
#[ignore]
fn list_tasks_filters_and_create_task_additions_round_trip() {
    let c = client();
    let created = c.create_task(
        "[outlook-mcp-rs P11 live] filtered task".to_string(),
        None,
        None,
        "high".to_string(),
        Some(vec!["Red Category".to_string()]),
        Some("2099-01-01".to_string()),
        Some("2099-01-01T09:00".to_string()),
    ).expect("create_task with additions should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let found = c.list_tasks(TaskQuery {
        include_completed: false,
        category: Some("Red Category".to_string()),
        importance: Some("high".to_string()),
        query: Some("filtered task".to_string()),
    }).expect("list_tasks should succeed");
    assert!(found.iter().any(|t| t.id == id), "filtered list_tasks should find the new task");

    c.delete_task(id).expect("cleanup delete_task");
}

#[test]
#[ignore]
fn update_task_marks_complete_then_reopens() {
    let c = client();
    let created = c.create_task(
        "[outlook-mcp-rs P11 live] update probe".to_string(),
        None, None, "normal".to_string(), None, None, None,
    ).expect("create_task should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let updated = c.update_task(TaskUpdate {
        task_id: id.clone(),
        mark_complete: Some(true),
        ..Default::default()
    }).expect("update_task mark_complete should succeed");
    assert!(updated["changed"].as_array().unwrap().iter().any(|v| v == "mark_complete"));

    let after_complete = c.list_tasks(TaskQuery { include_completed: true, ..Default::default() })
        .expect("list_tasks should succeed");
    let task = after_complete.iter().find(|t| t.id == id).expect("task should still exist");
    assert!(task.complete);

    let reopened = c.update_task(TaskUpdate {
        task_id: id.clone(),
        mark_complete: Some(false),
        ..Default::default()
    }).expect("update_task reopen should succeed");
    assert!(reopened["changed"].as_array().unwrap().iter().any(|v| v == "mark_complete"));

    c.delete_task(id).expect("cleanup delete_task");
}

#[test]
#[ignore]
fn delete_task_removes_it() {
    let c = client();
    let created = c.create_task(
        "[outlook-mcp-rs P11 live] delete probe".to_string(),
        None, None, "normal".to_string(), None, None, None,
    ).expect("create_task should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let deleted = c.delete_task(id).expect("delete_task should succeed");
    assert_eq!(deleted["status"], "deleted");
}
```

- [ ] **Step 2: Run the live tests**

Run: `cargo test --test live_outlook -- --ignored list_tasks_filters_and_create_task_additions_round_trip update_task_marks_complete_then_reopens delete_task_removes_it --test-threads=1`
Expected: all 3 pass against the real mailbox. If a filter doesn't match as expected, root-cause it (per `.claude/skills/live-outlook-system-test/SKILL.md`'s discipline — this project's own live-testing skill, worth a skim if anything looks flaky) before weakening an assertion.

- [ ] **Step 3: Update `TESTING.md`**

Open `TESTING.md`, find wherever tasks-related tools are documented (or the general "how live tests clean up" section), and add a line noting `update_task`/`delete_task`/`list_tasks` filters are covered by the live suite (`cargo test --test live_outlook -- --ignored` matching the 3 new test names above); note that `complete_task` no longer exists as a standalone tool (absorbed into `update_task`) if `TESTING.md` mentions it anywhere — update or remove that reference.

- [ ] **Step 4: Run the full suite one final time and commit**

Run: `cargo build` (0 warnings) then `cargo test` (all passing) then the 3 live tests again to confirm.

```bash
git add tests/live_outlook.rs TESTING.md
git commit -m "Add live tests for list_tasks filters, update_task, delete_task"
```

---

## After all 5 tasks are green

Dispatch the final whole-branch review (per `superpowers:subagent-driven-development`), covering all 5 tasks together: confirm `TaskQuery`/`TaskUpdate` are threaded consistently through `mod.rs` → `client.rs`/`fake.rs` → `server.rs` → tests; confirm `complete_task` has been fully retired (no dangling references anywhere, including `TESTING.md`); confirm `task_matches`'s subject-only query scope (vs. the spec's "subject + body") is a reasonable, documented simplification given `TaskSummary` has no body field. Then push to `main`, and update `V2-RESUME.md` / `2026-07-07-outlook-mcp-v2-plans-index.md` to mark Plan 11 shipped, following the exact pattern used for Plans 5-10.
