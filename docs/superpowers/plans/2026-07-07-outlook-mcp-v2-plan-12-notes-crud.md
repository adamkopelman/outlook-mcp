# Notes CRUD Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Bring the Notes tools to full CRUD: add filters to `list_notes`, add a `modified` field to `get_note`, add `categories`/`color` to `create_note`, add a new `update_note`, and add a new `delete_note`. This is the last plan in the v2 feature build.

**Architecture:** `list_notes` gains a `NoteQuery` struct and a pure `note_matches` predicate, following the exact precedent Plan 11 established for `list_tasks` (`TaskQuery`/`task_matches`, itself modeled on `EventQuery`/`event_matches`) — client-side filtering, since there's no DASL text-search precedent for a non-email folder in this codebase. Unlike tasks, notes' `query` filter genuinely needs body content (a note's only content), so `note_matches` takes the raw body string as a parameter rather than being scoped to subject-only. `update_note`/`delete_note` mirror `update_task`/`delete_task`'s shapes exactly (both themselves modeled on `update_email`/`delete_email`) — this plan is almost entirely pattern transcription from Plan 11, not new design.

**Tech Stack:** Rust, `windows` crate 0.62.2 (Win32 COM/`IDispatch::Invoke`), `rmcp` 2.1.0 tool macros, `chrono`, `serde`/`serde_json`.

## Global Constraints

- `list_notes` adds `category` and `query` (text match on the note's body — a note's *only* content, so unlike Plan 11's `list_tasks` this one genuinely searches body, not subject). Output already has `categories` (`NoteSummary` already carries it — verify, don't re-add).
- `get_note` adds `modified` (`LastModificationTime`, ISO string) to its output, alongside the existing `created`. `categories` is already present via `NoteSummary` (flattened into `NoteDetail`) — no change needed there.
- `create_note` keeps `body` (required); adds `categories` (assign on creation) and `color` — one of `"blue"`/`"green"`/`"pink"`/`"yellow"`/`"white"`, mapping to the real `OlNoteColor` enum values (confirmed against Microsoft's official enum reference, not guessed): `olBlue=0, olGreen=1, olPink=2, olYellow=3, olWhite=4`. Note the enum's *declared* order (blue, green, pink, yellow, white) differs from a naive alphabetical listing — use the exact values above, not a re-derived guess.
- `update_note` (new): `note_id` (required), `body`, `add_categories`/`remove_categories`, `color` all optional. Returns `{ "status": "updated", "id", "changed": [...] }`.
- `delete_note` (new): soft-delete a note by id (to Deleted Items), mirrors `delete_task`/`delete_email`'s shape exactly.
- The `OutlookClient` trait (`src/outlook/mod.rs`) has two implementors — `WindowsOutlookClient` (`client.rs`) and `FakeOutlookClient` (`fake.rs`) — plus `src/server.rs` (MCP tool layer) and `tests/tools.rs` (fake-backed tests); every trait change touches all four. Also check `tests/live_outlook.rs` for existing `create_note`/`get_note` call sites that need updating for signature changes — Plan 11 found an unplanned call site in every single task that changed a trait signature; do not assume the brief's named file list is exhaustive, grep the whole repo.
- Note colors are validated the same way importance is: a new `note_color_to_id` in `src/constants.rs`, mirroring the existing `importance_name_to_id` exactly (same file, same style, same `#[cfg(test)] mod tests` group).
- Per this project's model policy: Task 1 (filters, mirrors an existing pattern closely) is standard-tier; Task 2 (small field additions to 2 existing tools) is cheap-tier; Task 3 (new struct-based tool, most design judgment though still highly patterned) is standard-tier; Task 4 (mechanical, mirrors `delete_task`) is cheap-tier; Task 5 (live tests) is standard-tier.

---

### Task 1: `list_notes` filters (`category`, `query`)

**Files:**
- Modify: `src/outlook/mod.rs` (replace `list_notes`'s trait signature; add `NoteQuery` struct)
- Modify: `src/outlook/client.rs` (real COM: wire `NoteQuery`, add `note_matches`)
- Modify: `src/outlook/fake.rs` (record the new query fields)
- Modify: `src/server.rs` (`ListNotesParams` — new struct, this tool currently takes no parameters at all)
- Test: `tests/tools.rs`

**Interfaces:**
- Produces (consumed by Tasks 2-5): `NoteQuery { pub category: Option<String>, pub query: Option<String> }` (in `mod.rs`, alongside `TaskQuery`); `fn note_matches(body: &str, summary: &NoteSummary, q: &NoteQuery) -> bool` (in `client.rs`, alongside `task_matches`).

- [ ] **Step 1: Replace the trait signature**

In `src/outlook/mod.rs`, find:

```rust
    fn list_notes(&self) -> Result<Vec<NoteSummary>, ToolError>;
```

Replace with:

```rust
    fn list_notes(&self, q: NoteQuery) -> Result<Vec<NoteSummary>, ToolError>;
```

Add the `NoteQuery` struct right after `TaskQuery`'s closing `}` (before `CreateEventInput`):

```rust
/// All filters for `list_notes`. Both fields optional; supplying both ANDs
/// them. Unlike `TaskQuery`'s `query` (subject-only, since tasks have a
/// separate subject), a note's *only* content is its body — `note_matches`
/// reads the real body text to match `query`, not just the derived subject.
#[derive(Debug, Clone, Default)]
pub struct NoteQuery {
    pub category: Option<String>,
    pub query: Option<String>,
}
```

Run `cargo build` — expect failures in `client.rs`, `fake.rs`, and `server.rs`. Steps 2-4 fix them.

- [ ] **Step 2: Wire `NoteQuery` into `client.rs`, add `note_matches`**

Find `task_matches` in `client.rs` and add `note_matches` right after its closing `}`:

```rust
/// Client-side filter for `list_notes`'s `category`/`query`. `body` is the
/// note's real, untruncated body text (read once per item by the caller —
/// see `list_notes` below) — unlike `task_matches`, this genuinely searches
/// content, since a note's body IS its content.
fn note_matches(body: &str, summary: &NoteSummary, q: &NoteQuery) -> bool {
    if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
        if !body.to_lowercase().contains(&query.to_lowercase()) {
            return false;
        }
    }
    if let Some(cat) = q.category.as_deref().filter(|s| !s.is_empty()) {
        let want = cat.to_lowercase();
        if !summary.categories.iter().any(|c| c.to_lowercase() == want) {
            return false;
        }
    }
    true
}
```

Replace `list_notes`'s existing body:

```rust
    fn list_notes(&self) -> Result<Vec<NoteSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let notes = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_NOTES)],
            )?)?;
            let items = to_disp(get_property(&notes, "Items")?)?;
            let count = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let mut results = Vec::new();
            for i in 1..=count {
                let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                results.push(note_summary(&item)?);
            }
            Ok(results)
        })
    }
```

with:

```rust
    fn list_notes(&self, q: NoteQuery) -> Result<Vec<NoteSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let notes = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_NOTES)],
            )?)?;
            let items = to_disp(get_property(&notes, "Items")?)?;
            let count = variant_to_i32(&get_property(&items, "Count")?).unwrap_or(0);
            let mut results = Vec::new();
            for i in 1..=count {
                let item = to_disp(call_method(&items, "Item", &mut [variant_from_i32(i)])?)?;
                let summary = note_summary(&item)?;
                // Read the real body directly for query matching — `note_summary`
                // only exposes the derived (120-char-truncated) subject, not the
                // full body, so this is a second, deliberate property read (same
                // pattern `get_note` already uses: it re-reads `Body` outside
                // `note_summary` too, for its own untruncated-body purpose).
                let body = variant_to_string(&get_property(&item, "Body").unwrap_or_default());
                if note_matches(&body, &summary, &q) {
                    results.push(summary);
                }
            }
            Ok(results)
        })
    }
```

- [ ] **Step 3: Update `fake.rs`**

Find `list_notes` in `fake.rs`:

```rust
    fn list_notes(&self) -> Result<Vec<NoteSummary>, ToolError> {
        self.record("list_notes", json!({}))?;
        Ok(vec![NoteSummary { id: NOTE_ID.into(), subject: "Ideas".into(), created: None, categories: vec![] }])
    }
```

Replace with:

```rust
    fn list_notes(&self, q: NoteQuery) -> Result<Vec<NoteSummary>, ToolError> {
        self.record("list_notes", json!({"category": q.category, "query": q.query}))?;
        Ok(vec![NoteSummary { id: NOTE_ID.into(), subject: "Ideas".into(), created: None, categories: vec![] }])
    }
```

(Match this file's actual existing canned `NoteSummary` field values verbatim — do not invent new placeholder text; open `fake.rs`, find the current `list_notes` body, and keep every field's existing value, changing only the function signature and adding the `record(...)` call's query fields.)

- [ ] **Step 4: Update `server.rs`**

`list_notes` currently takes no parameters at all (`pub async fn list_notes(&self) -> Result<CallToolResult, McpError>`). Add a new `ListNotesParams` struct right before it:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListNotesParams {
    /// Filter to a color category.
    #[serde(default)]
    pub category: Option<String>,
    /// Text match on the note's body.
    #[serde(default)]
    pub query: Option<String>,
}
```

Replace `list_notes`'s tool method:

```rust
    #[tool(description = "List Outlook notes.")]
    pub async fn list_notes(&self) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.list_notes()).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

with:

```rust
    #[tool(description = "List Outlook notes. Filter by category or a text query matching the note's body.")]
    pub async fn list_notes(
        &self,
        Parameters(ListNotesParams { category, query }): Parameters<ListNotesParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = NoteQuery { category, query };
        let result = run_blocking(move || client.list_notes(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Add `NoteQuery` to `server.rs`'s top-level `use crate::outlook::{...}` import line.

- [ ] **Step 5: Run `cargo build --tests`, find and fix every `list_notes` call site**

Run: `cargo build --tests`
Expected failures include (at minimum) `tests/tools.rs`'s pre-existing test:

```rust
async fn list_notes_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server.list_notes().await.unwrap();
    assert_eq!(fake.calls(), vec![("list_notes".to_string(), json!({}))]);
}
```

Fix it to:

```rust
async fn list_notes_records_call() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server.list_notes(Parameters(ListNotesParams { category: None, query: None })).await.unwrap();
    assert_eq!(fake.calls(), vec![("list_notes".to_string(), json!({"category": null, "query": null}))]);
}
```

Do not rely only on compiler errors from `cargo build` alone (without `--tests`) — a call inside a `#[ignore]`d test still needs fixing even though it won't surface without building the test binaries. Run `grep -rn "list_notes(" src/ tests/` after the build to confirm you found every call site, not just the one above.

- [ ] **Step 6: Write fake-backed tool tests**

Open `tests/tools.rs`, find the existing `list_notes`-related test(s), and add these two:

```rust
#[tokio::test]
async fn list_notes_forwards_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_notes(Parameters(ListNotesParams {
            category: Some("Green Category".to_string()),
            query: Some("renew".to_string()),
        }))
        .await
        .unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "list_notes");
    assert_eq!(args["category"], "Green Category");
    assert_eq!(args["query"], "renew");
}

#[tokio::test]
async fn list_notes_defaults_filters_to_none() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: ListNotesParams = serde_json::from_value(json!({})).unwrap();
    server.list_notes(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert!(args["category"].is_null());
    assert!(args["query"].is_null());
}
```

Add `ListNotesParams` to this file's `use outlook_mcp_rs::server::{...}` import block (check whether the pre-existing `list_notes`-related test already imports something that needs adjusting alongside this — read the current test before editing).

- [ ] **Step 7: Run the tests, then the full non-live suite, then commit**

Run: `cargo test --test tools list_notes`
Expected: all pass (pre-existing test(s), fixed for the new signature, plus 2 new).

Run: `cargo build` (0 warnings) then `cargo test` (all passing).

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/server.rs tests/tools.rs tests/live_outlook.rs
git commit -m "Add list_notes filters: category, query (client-side, mirrors task_matches but searches real body)"
```

---

### Task 2: `get_note` gains `modified`; `create_note` gains `categories`/`color`

**Files:**
- Modify: `src/outlook/types.rs` (`modified` field on `NoteDetail`)
- Modify: `src/outlook/mod.rs` (trait signature for `create_note`)
- Modify: `src/outlook/client.rs`
- Modify: `src/outlook/fake.rs`
- Modify: `src/constants.rs` (`OL_NOTE_COLOR_*` constants + `note_color_to_id`)
- Modify: `src/server.rs`
- Test: `tests/tools.rs`

**Interfaces:**
- Modifies `NoteDetail` (adds `modified: Option<String>` — NOT on `NoteSummary`, so `list_notes`'s output shape is unaffected).
- Modifies the existing `create_note` signature: `fn create_note(&self, body: String, categories: Option<Vec<String>>, color: Option<String>) -> Result<Value, ToolError>;`.
- Produces: `pub fn note_color_to_id(name: &str) -> Option<i32>` in `src/constants.rs`.

- [ ] **Step 1: Add `OL_NOTE_COLOR_*` constants and `note_color_to_id`**

In `src/constants.rs`, add near the other `OL_*` color-ish/enum constants (e.g. near `OL_IMPORTANCE_*`):

```rust
// OlNoteColor (NoteItem.Color) — confirmed against Microsoft's official
// enum reference; note the declared order (blue, green, pink, yellow,
// white) is NOT alphabetical.
pub const OL_NOTE_COLOR_BLUE: i32 = 0;
pub const OL_NOTE_COLOR_GREEN: i32 = 1;
pub const OL_NOTE_COLOR_PINK: i32 = 2;
pub const OL_NOTE_COLOR_YELLOW: i32 = 3;
pub const OL_NOTE_COLOR_WHITE: i32 = 4;
```

Add `note_color_to_id`, right after `importance_name_to_id`:

```rust
pub fn note_color_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "blue" => Some(OL_NOTE_COLOR_BLUE),
        "green" => Some(OL_NOTE_COLOR_GREEN),
        "pink" => Some(OL_NOTE_COLOR_PINK),
        "yellow" => Some(OL_NOTE_COLOR_YELLOW),
        "white" => Some(OL_NOTE_COLOR_WHITE),
        _ => None,
    }
}
```

Add a unit test in `constants.rs`'s existing `#[cfg(test)] mod tests`, right after `importance_and_meeting_response_lookups`:

```rust
    #[test]
    fn note_color_lookup_is_case_insensitive() {
        assert_eq!(note_color_to_id("BLUE"), Some(OL_NOTE_COLOR_BLUE));
        assert_eq!(note_color_to_id("Yellow"), Some(OL_NOTE_COLOR_YELLOW));
        assert_eq!(note_color_to_id("purple"), None);
    }
```

- [ ] **Step 2: Run the new test**

Run: `cargo test --lib note_color`
Expected: 1 test passes.

- [ ] **Step 3: Add `modified` to `NoteDetail`**

In `src/outlook/types.rs`, find:

```rust
#[derive(Debug, Clone, Serialize)]
pub struct NoteDetail {
    #[serde(flatten)]
    pub summary: NoteSummary,
    pub body: String,
}
```

Replace with:

```rust
#[derive(Debug, Clone, Serialize)]
pub struct NoteDetail {
    #[serde(flatten)]
    pub summary: NoteSummary,
    pub body: String,
    pub modified: Option<String>,
}
```

- [ ] **Step 4: Wire `modified` into `get_note`, extend `create_note`'s signature**

In `src/outlook/mod.rs`, find:

```rust
    fn create_note(&self, body: String) -> Result<Value, ToolError>;
```

Replace with:

```rust
    fn create_note(&self, body: String, categories: Option<Vec<String>>, color: Option<String>) -> Result<Value, ToolError>;
```

In `src/outlook/client.rs`, find `get_note`:

```rust
    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let note = get_item(&ns, &note_id)?;
            let summary = note_summary(&note)?;
            Ok(NoteDetail {
                summary,
                body: truncate(&variant_to_string(&get_property(&note, "Body")?)),
            })
        })
    }
```

Replace with:

```rust
    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let note = get_item(&ns, &note_id)?;
            let summary = note_summary(&note)?;
            Ok(NoteDetail {
                summary,
                body: truncate(&variant_to_string(&get_property(&note, "Body")?)),
                modified: variant_to_iso_string(&get_property(&note, "LastModificationTime").unwrap_or_default()),
            })
        })
    }
```

Find `create_note`:

```rust
    fn create_note(&self, body: String) -> Result<Value, ToolError> {
        // Validate before touching COM (fail-fast, like `create_task`).
        if body.is_empty() {
            return Err(ToolError::new("create_note requires a non-empty body."));
        }
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let note = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_NOTE_ITEM)],
            )?)?;
            put_property(&note, "Body", variant_from_str(&body))?;
            call_method(&note, "Save", &mut [])?;
            Ok(json!({"status": "created", "id": make_id(&note)?}))
        })
    }
```

Replace with:

```rust
    fn create_note(&self, body: String, categories: Option<Vec<String>>, color: Option<String>) -> Result<Value, ToolError> {
        // Validate before touching COM (fail-fast, like `create_task`).
        if body.is_empty() {
            return Err(ToolError::new("create_note requires a non-empty body."));
        }
        let color_id = color.as_deref().map(|c| {
            c::note_color_to_id(c).ok_or_else(|| {
                ToolError::new(format!(
                    "invalid color {c:?}: expected \"blue\", \"green\", \"pink\", \"yellow\", or \"white\""
                ))
            })
        }).transpose()?;
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let note = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_NOTE_ITEM)],
            )?)?;
            put_property(&note, "Body", variant_from_str(&body))?;
            if let Some(id) = color_id {
                put_property(&note, "Color", variant_from_i32(id))?;
            }
            call_method(&note, "Save", &mut [])?;
            if let Some(cats) = categories.as_ref().filter(|c| !c.is_empty()) {
                set_item_categories(&note, cats)?;
                call_method(&note, "Save", &mut [])?;
            }
            Ok(json!({"status": "created", "id": make_id(&note)?}))
        })
    }
```

Note: `color` is set and saved **before** `categories`, with a second `Save()` after categories — `NoteItem.Color` is a display property that some Outlook versions only apply cleanly on the initial `Save()` of a freshly created item, so it's set pre-save; categories are applied the same way every other `set_item_categories` call site in this codebase already does (set property, then `Save()`).

- [ ] **Step 5: Update `fake.rs`**

Find `get_note`:

```rust
    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError> {
        self.record("get_note", json!({"note_id": note_id}))?;
        Ok(NoteDetail {
            summary: NoteSummary { id: note_id, subject: "Ideas".into(), created: None, categories: vec![] },
            body: "Ideas\n- one".into(),
        })
    }
```

Replace with (adding `modified: None` — the fake has no meaningful COM-backed modification time to simulate):

```rust
    fn get_note(&self, note_id: String) -> Result<NoteDetail, ToolError> {
        self.record("get_note", json!({"note_id": note_id}))?;
        Ok(NoteDetail {
            summary: NoteSummary { id: note_id, subject: "Ideas".into(), created: None, categories: vec![] },
            body: "Ideas\n- one".into(),
            modified: None,
        })
    }
```

Find `create_note`:

```rust
    fn create_note(&self, body: String) -> Result<Value, ToolError> {
        self.record("create_note", json!({"body": body}))?;
        Ok(json!({"status": "created", "id": NOTE_ID}))
    }
```

Replace with:

```rust
    fn create_note(&self, body: String, categories: Option<Vec<String>>, color: Option<String>) -> Result<Value, ToolError> {
        self.record("create_note", json!({"body": body, "categories": categories, "color": color}))?;
        Ok(json!({"status": "created", "id": NOTE_ID}))
    }
```

- [ ] **Step 6: Update `server.rs`**

Find `CreateNoteParams`:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateNoteParams {
    pub body: String,
}
```

Replace with:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateNoteParams {
    pub body: String,
    /// Category names to assign on creation.
    #[serde(default)]
    pub categories: Option<Vec<String>>,
    /// "blue" | "green" | "pink" | "yellow" | "white".
    #[serde(default)]
    pub color: Option<String>,
}
```

Find `create_note`'s tool method:

```rust
    pub async fn create_note(
        &self,
        Parameters(CreateNoteParams { body }): Parameters<CreateNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.create_note(body)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Replace with:

```rust
    pub async fn create_note(
        &self,
        Parameters(CreateNoteParams { body, categories, color }): Parameters<CreateNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.create_note(body, categories, color)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

- [ ] **Step 7: Run `cargo build --tests`, find and fix every `create_note` call site**

Run: `cargo build --tests`
Expected failures include (at minimum) `tests/tools.rs`'s pre-existing test:

```rust
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

Fix it to:

```rust
async fn create_note_records_body() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .create_note(Parameters(CreateNoteParams {
            body: "Ideas\n- one".to_string(), categories: None, color: None,
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls(), vec![
        ("create_note".to_string(), json!({"body": "Ideas\n- one", "categories": null, "color": null})),
    ]);
}
```

and every `create_note(...)` call in `tests/live_outlook.rs` called with the old 1-argument shape needs `, None, None` appended. Run `grep -rn "create_note(" tests/` after fixing to confirm you found every call site, not just the ones above.

- [ ] **Step 8: Write fake-backed tests**

Add to `tests/tools.rs`, near the existing `create_note`/`get_note` tests:

```rust
#[tokio::test]
async fn create_note_forwards_categories_and_color() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let params: CreateNoteParams = serde_json::from_value(json!({
        "body": "Remember to renew the domain",
        "categories": ["Yellow Category"],
        "color": "yellow"
    })).unwrap();
    server.create_note(Parameters(params)).await.unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["categories"], json!(["Yellow Category"]));
    assert_eq!(args["color"], "yellow");
}

#[tokio::test]
async fn get_note_includes_modified() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_note(Parameters(GetNoteParams { note_id: NOTE_ID.to_string() }))
        .await
        .unwrap();
    let v = result_json(&result);
    // The fake may return null for a note that was never "modified" —
    // assert the key exists in the JSON shape, not a specific non-null value.
    assert!(v.as_object().unwrap().contains_key("modified"));
}
```

Add `CreateNoteParams` to this file's import list if not already present (it should already be there from the pre-existing `create_note` test).

- [ ] **Step 9: Run the tests, then the full non-live suite, then commit**

Run: `cargo test --test tools create_note get_note`
Expected: all pass.

Run: `cargo build` (0 warnings) then `cargo test` (all passing).

```bash
git add src/outlook/types.rs src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/constants.rs src/server.rs tests/tools.rs tests/live_outlook.rs
git commit -m "Add modified to get_note; add categories/color to create_note"
```

---

### Task 3: `update_note` (new)

**Files:**
- Modify: `src/outlook/mod.rs` (add `NoteUpdate` + trait method)
- Modify: `src/outlook/client.rs`
- Modify: `src/outlook/fake.rs`
- Modify: `src/server.rs`
- Test: `tests/tools.rs`

**Interfaces:**
- Produces: `NoteUpdate { pub note_id: String, pub body: Option<String>, pub add_categories: Option<Vec<String>>, pub remove_categories: Option<Vec<String>>, pub color: Option<String> }` (in `mod.rs`); `fn update_note(&self, u: NoteUpdate) -> Result<Value, ToolError>;` on the trait.

- [ ] **Step 1: Add `NoteUpdate` and the trait method**

In `src/outlook/mod.rs`, add right after `TaskUpdate`'s closing `}`:

```rust
/// All changes `update_note` can apply to one existing note. Every field
/// except `note_id` is optional; supplying several applies all of them.
#[derive(Debug, Clone, Default)]
pub struct NoteUpdate {
    pub note_id: String,
    pub body: Option<String>,
    pub add_categories: Option<Vec<String>>,
    pub remove_categories: Option<Vec<String>>,
    pub color: Option<String>,
}
```

Add the trait method right after `create_note`'s line:

```rust
    fn update_note(&self, u: NoteUpdate) -> Result<Value, ToolError>;
```

- [ ] **Step 2: Implement in `client.rs`**

Add right after `create_note`'s closing `}` (read the file to find the exact spot — `create_note` is the last method before whatever comes next in the notes section):

```rust
    fn update_note(&self, u: NoteUpdate) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let note = get_item(&ns, &u.note_id)?;
            let mut changed: Vec<&str> = Vec::new();

            if let Some(body) = &u.body {
                put_property(&note, "Body", variant_from_str(body))?;
                changed.push("body");
            }
            if u.add_categories.is_some() || u.remove_categories.is_some() {
                let mut cats = get_item_categories(&note);
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
                set_item_categories(&note, &cats)?;
            }
            if let Some(color) = &u.color {
                let id = c::note_color_to_id(color).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid color {color:?}: expected \"blue\", \"green\", \"pink\", \"yellow\", or \"white\""
                    ))
                })?;
                put_property(&note, "Color", variant_from_i32(id))?;
                changed.push("color");
            }

            call_method(&note, "Save", &mut [])?;
            Ok(json!({"status": "updated", "id": u.note_id, "changed": changed}))
        })
    }
```

- [ ] **Step 3: Update `fake.rs`**

Add right after `create_note`'s closing `}`:

```rust
    fn update_note(&self, u: NoteUpdate) -> Result<Value, ToolError> {
        self.record("update_note", json!({
            "note_id": u.note_id, "body": u.body, "add_categories": u.add_categories,
            "remove_categories": u.remove_categories, "color": u.color,
        }))?;
        let mut changed: Vec<&str> = Vec::new();
        if u.body.is_some() { changed.push("body"); }
        if u.add_categories.is_some() { changed.push("add_categories"); }
        if u.remove_categories.is_some() { changed.push("remove_categories"); }
        if u.color.is_some() { changed.push("color"); }
        Ok(json!({"status": "updated", "id": u.note_id, "changed": changed}))
    }
```

Add `NoteUpdate` to `fake.rs`'s existing `use super::{...}` import line.

- [ ] **Step 4: Add to `server.rs`**

Add right after `CreateNoteParams`:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct UpdateNoteParams {
    pub note_id: String,
    #[serde(default)]
    pub body: Option<String>,
    #[serde(default)]
    pub add_categories: Option<Vec<String>>,
    #[serde(default)]
    pub remove_categories: Option<Vec<String>>,
    /// "blue" | "green" | "pink" | "yellow" | "white".
    #[serde(default)]
    pub color: Option<String>,
}
```

Add the tool method right after `create_note`'s:

```rust
    #[tool(description = "Update an existing note: body, add/remove categories, color. Combine any of these in one call.")]
    pub async fn update_note(
        &self,
        Parameters(p): Parameters<UpdateNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let u = NoteUpdate {
            note_id: p.note_id, body: p.body, add_categories: p.add_categories,
            remove_categories: p.remove_categories, color: p.color,
        };
        let result = run_blocking(move || client.update_note(u)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Add `NoteUpdate` to `server.rs`'s top-level `use crate::outlook::{...}` import line.

- [ ] **Step 5: Run `cargo build`**

Run: `cargo build`
Expected: 0 errors, 0 warnings.

- [ ] **Step 6: Write fake-backed tests**

Add to `tests/tools.rs`:

```rust
#[tokio::test]
async fn update_note_forwards_body_and_color() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .update_note(Parameters(UpdateNoteParams {
            note_id: NOTE_ID.to_string(),
            body: Some("Updated body".to_string()),
            add_categories: None, remove_categories: None,
            color: Some("pink".to_string()),
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["body"], "Updated body");
    assert_eq!(args["color"], "pink");
}

#[tokio::test]
async fn update_note_manages_categories() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_note(Parameters(UpdateNoteParams {
            note_id: NOTE_ID.to_string(),
            body: None,
            add_categories: Some(vec!["Blue Category".to_string()]),
            remove_categories: None,
            color: None,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert!(v["changed"].as_array().unwrap().iter().any(|c| c == "add_categories"));
}
```

Add `UpdateNoteParams` to this file's `use outlook_mcp_rs::server::{...}` import block.

- [ ] **Step 7: Run the tests, then the full non-live suite, then commit**

Run: `cargo test --test tools update_note`
Expected: 2 tests pass.

Run: `cargo build` (0 warnings) then `cargo test` (all passing).

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/server.rs tests/tools.rs
git commit -m "Add update_note: body, add/remove categories, color"
```

---

### Task 4: `delete_note` (new)

**Files:**
- Modify: `src/outlook/mod.rs`
- Modify: `src/outlook/client.rs`
- Modify: `src/outlook/fake.rs`
- Modify: `src/server.rs`
- Test: `tests/tools.rs`

**Interfaces:**
- Produces: `fn delete_note(&self, note_id: String) -> Result<Value, ToolError>;` on the trait.

- [ ] **Step 1: Add the trait method**

In `src/outlook/mod.rs`, add right after `update_note`'s trait line:

```rust
    fn delete_note(&self, note_id: String) -> Result<Value, ToolError>;
```

- [ ] **Step 2: Implement in `client.rs`**

Add right after `update_note`'s closing `}` (mirrors `delete_task`/`delete_email` exactly):

```rust
    fn delete_note(&self, note_id: String) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &note_id)?;
            call_method(&item, "Delete", &mut [])?;
            Ok(json!({"status": "deleted", "note": "Moved to Deleted Items."}))
        })
    }
```

Note: unlike `delete_email`/`delete_task`, this doesn't read `Subject` before deleting — notes have no native `Subject` property (per `note_summary`'s own doc comment: "Notes have no native `Subject` property"), so there's nothing meaningful to report back. Don't call `get_property(&item, "Subject")` here — it would either error or return an empty string, neither useful.

- [ ] **Step 3: Implement in `fake.rs`**

Add right after `update_note`'s closing `}`:

```rust
    fn delete_note(&self, note_id: String) -> Result<Value, ToolError> {
        self.record("delete_note", json!({"note_id": note_id}))?;
        Ok(json!({"status": "deleted", "note": "Moved to Deleted Items."}))
    }
```

- [ ] **Step 4: Add to `server.rs`**

Add right after `UpdateNoteParams`:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteNoteParams {
    pub note_id: String,
}
```

Add the tool method right after `update_note`'s:

```rust
    #[tool(description = "Delete a note (moves it to Deleted Items).")]
    pub async fn delete_note(
        &self,
        Parameters(DeleteNoteParams { note_id }): Parameters<DeleteNoteParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_note(note_id)).await?;
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
async fn delete_note_records_call() {
    use outlook_mcp_rs::outlook::fake::NOTE_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .delete_note(Parameters(DeleteNoteParams { note_id: NOTE_ID.to_string() }))
        .await
        .unwrap();
    let json = result_json(&result);
    assert_eq!(json["status"], "deleted");
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "delete_note");
    assert_eq!(args["note_id"], NOTE_ID);
}
```

Add `DeleteNoteParams` to this file's `use outlook_mcp_rs::server::{...}` import block.

- [ ] **Step 7: Run the test, then the full non-live suite, then commit**

Run: `cargo test --test tools delete_note`
Expected: 1 test passes.

Run: `cargo build` (0 warnings) then `cargo test` (all passing).

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/server.rs tests/tools.rs
git commit -m "Add delete_note"
```

---

### Task 5: Live tests and `TESTING.md`

**Files:**
- Modify: `tests/live_outlook.rs`
- Modify: `TESTING.md`
- Modify: `README.md` (the Notes tool listing is now stale — it still shows the pre-Plan-12 3-tool set)

**Interfaces:**
- Consumes everything from Tasks 1-4: `NoteQuery`, the expanded `create_note`/`get_note` shapes, `NoteUpdate`/`update_note`, `delete_note`.

- [ ] **Step 1: Add the live tests**

Open `tests/live_outlook.rs`. Add `NoteQuery`, `NoteUpdate` to its `use outlook_mcp_rs::outlook::{...}` import line. Find the existing `create_note_then_get_it_back` test (already fixed for the new `create_note` signature in Task 2) to see its current shape, then append these tests at the end of the file:

```rust
#[test]
#[ignore]
fn list_notes_filters_and_create_note_additions_round_trip() {
    let c = client();
    let created = c.create_note(
        "outlook-mcp-rs P12 live filtered note - remember to renew".to_string(),
        Some(vec!["Green Category".to_string()]),
        Some("green".to_string()),
    ).expect("create_note with additions should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let found = c.list_notes(NoteQuery {
        category: Some("Green Category".to_string()),
        query: Some("renew".to_string()),
    }).expect("list_notes should succeed");
    assert!(found.iter().any(|n| n.id == id), "filtered list_notes should find the new note");

    c.delete_note(id).expect("cleanup delete_note");
}

#[test]
#[ignore]
fn get_note_includes_modified_after_update() {
    let c = client();
    let created = c.create_note(
        "outlook-mcp-rs P12 live modified probe".to_string(), None, None,
    ).expect("create_note should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    c.update_note(NoteUpdate {
        note_id: id.clone(),
        body: Some("outlook-mcp-rs P12 live modified probe (edited)".to_string()),
        ..Default::default()
    }).expect("update_note should succeed");

    let note = c.get_note(id.clone()).expect("get_note should succeed");
    assert!(note.modified.is_some(), "modified should be populated after an edit");
    assert!(note.body.starts_with("outlook-mcp-rs P12 live modified probe (edited)"));

    c.delete_note(id).expect("cleanup delete_note");
}

#[test]
#[ignore]
fn update_note_manages_categories_and_color() {
    let c = client();
    let created = c.create_note(
        "outlook-mcp-rs P12 live category probe".to_string(), None, None,
    ).expect("create_note should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let updated = c.update_note(NoteUpdate {
        note_id: id.clone(),
        add_categories: Some(vec!["Blue Category".to_string()]),
        color: Some("blue".to_string()),
        ..Default::default()
    }).expect("update_note should succeed");
    assert!(updated["changed"].as_array().unwrap().iter().any(|v| v == "add_categories"));
    assert!(updated["changed"].as_array().unwrap().iter().any(|v| v == "color"));

    let note = c.get_note(id.clone()).expect("get_note should succeed");
    assert!(note.summary.categories.iter().any(|cat| cat == "Blue Category"));

    c.delete_note(id).expect("cleanup delete_note");
}

#[test]
#[ignore]
fn delete_note_removes_it() {
    let c = client();
    let created = c.create_note(
        "outlook-mcp-rs P12 live delete probe".to_string(), None, None,
    ).expect("create_note should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let deleted = c.delete_note(id).expect("delete_note should succeed");
    assert_eq!(deleted["status"], "deleted");
}
```

- [ ] **Step 2: Run the live tests**

Run: `cargo test --test live_outlook -- --ignored list_notes_filters_and_create_note_additions_round_trip get_note_includes_modified_after_update update_note_manages_categories_and_color delete_note_removes_it --test-threads=1`
Expected: all 4 pass against the real mailbox. If anything doesn't match as expected, root-cause it per `.claude/skills/live-outlook-system-test/SKILL.md`'s discipline (this project's own live-testing skill) before weakening an assertion — do NOT touch that skill file itself as part of this task; skill updates for the whole v2 build are handled separately, after this plan ships.

- [ ] **Step 3: Update `TESTING.md`**

Open `TESTING.md`, find the paragraph Plan 11 added documenting `list_tasks`/`update_task`/`delete_task`'s live coverage, and add an equivalent paragraph for notes right after it: the 4 new test names and the exact `cargo test` command to run them.

- [ ] **Step 4: Update `README.md`**

Find the Notes section in `README.md`'s tool listing:

```markdown
**Notes**
- `list_notes` — list Outlook notes
- `get_note` — get a note's content
- `create_note` — create a new note
```

Replace with:

```markdown
**Notes**
- `list_notes` — list Outlook notes (filter by category or a text query on the body)
- `get_note` — get a note's content
- `create_note` — create a new note (with optional categories/color)
- `update_note` — change an existing note: body, add/remove categories, color
- `delete_note` — delete a note
```

(Read the file first to confirm the exact current text before replacing — match its existing formatting style precisely, don't guess.)

- [ ] **Step 5: Run the full suite one final time and commit**

Run: `cargo build` (0 warnings) then `cargo test` (all passing) then the 4 live tests again to confirm.

```bash
git add tests/live_outlook.rs TESTING.md README.md
git commit -m "Add live tests for list_notes filters, get_note modified, update_note, delete_note"
```

---

## After all 5 tasks are green

Dispatch the final whole-branch review (per `superpowers:subagent-driven-development`), covering all 5 tasks together: confirm `NoteQuery`/`NoteUpdate` are threaded consistently through `mod.rs` → `client.rs`/`fake.rs` → `server.rs` → tests; confirm the `OL_NOTE_COLOR_*` values match the plan's Microsoft-confirmed values exactly (this is the one place in this plan where getting a raw enum value wrong would silently write the wrong color, not fail loudly — worth an explicit spot-check); confirm `note_matches`'s real-body search (vs. Plan 11's task-subject-only precedent) is correctly implemented and doesn't accidentally read a truncated/derived value instead of the real body. Then push to `main`, and update `V2-RESUME.md` / `2026-07-07-outlook-mcp-v2-plans-index.md` to mark Plan 12 shipped — and mark the **whole v2 feature build complete**, since this is the last plan (all 26 tools from the spec will be done).

**Do not touch `.claude/skills/live-outlook-system-test/SKILL.md` as part of this plan's execution.** Per explicit user instruction, that skill gets one consolidated review-and-update pass after Plan 12 ships — not incremental touches during Plans 10-12. That pass happens as a separate, final step after this plan's whole-branch review and push, not as part of any task above.
