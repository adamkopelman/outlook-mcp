# v2 Plan 1 — Foundations (friendly words + categories) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add the two cross-cutting foundations every later v2 plan depends on: a friendly-word enum module (so outputs say `"accepted"`/`"busy"`, not `3`/`2`), and category read/write support (so color categories are visible in output, and settable/filterable later).

**Architecture:** A new pure-logic module `src/friendly.rs` converts Outlook enum integers ↔ lowercase strings. Category string parse/join helpers go in `src/outlook/com.rs` (pure) with thin COM read/write wrappers. The `categories: Vec<String>` field is added to every summary type in `src/outlook/types.rs`, and the raw-integer fields (`response_status`, task `status`/`importance`) become friendly `String`s. Both trait implementors (`WindowsOutlookClient`, `FakeOutlookClient`) are updated together.

**Tech Stack:** Rust 2024, `serde`, `windows` 0.62.2 COM, `schemars`; test with `cargo test`.

**Spec:** `docs/superpowers/specs/2026-07-07-outlook-mcp-v2-features-design.md` (cross-cutting principles §1–§2).

## Global Constraints

- Target crate: `C:\Users\adamk\projects\outlook-mcp-rs` only.
- The `OutlookClient` trait has TWO implementors — `WindowsOutlookClient` (`src/outlook/client.rs`) and `FakeOutlookClient` (`src/outlook/fake.rs`); any type change touches both plus `tests/tools.rs`.
- Friendly words are lowercase snake: response → `"organizer"`/`"accepted"`/`"declined"`/`"tentative"`/`"not_responded"`/`"none"`; busy → `"free"`/`"tentative"`/`"busy"`/`"out_of_office"`/`"working_elsewhere"`; importance → `"low"`/`"normal"`/`"high"`; task status → `"not_started"`/`"in_progress"`/`"complete"`/`"waiting"`/`"deferred"`.
- Missing-property tolerance: summary/detail builders read new properties with `.unwrap_or_default()`, never `?`.
- The Outlook `Categories` property is a single string of category names joined by `", "` (comma-space); empty string = no categories.
- Commit after every task. Run `cargo test` (all 34 existing tests must stay green) before each commit.

---

### Task 1: Friendly-word conversion module

**Files:**
- Create: `src/friendly.rs`
- Modify: `src/lib.rs` (add `pub mod friendly;`)
- Modify: `src/constants.rs` (add the OlBusyStatus + OlTaskStatus + OlResponseStatus enum values used by the mappings)

**Interfaces:**
- Produces:
  - `friendly::importance_word(i32) -> &'static str`
  - `friendly::response_word(i32) -> &'static str`
  - `friendly::busy_status_word(i32) -> &'static str`
  - `friendly::task_status_word(i32) -> &'static str`
  - `friendly::busy_status_to_id(&str) -> Option<i32>`
  - `friendly::task_status_to_id(&str) -> Option<i32>`
  - (importance/response name→id already exist in `constants.rs` as `importance_name_to_id` / `meeting_response_to_id`.)

- [ ] **Step 1: Add the enum constants to `src/constants.rs`**

Append after the existing `OL_TASK_NOT_STARTED` / importance block (around line 38):

```rust
// OlBusyStatus (AppointmentItem.BusyStatus)
pub const OL_FREE: i32 = 0;
pub const OL_TENTATIVE: i32 = 1;
pub const OL_BUSY: i32 = 2;
pub const OL_OUT_OF_OFFICE: i32 = 3;
pub const OL_WORKING_ELSEWHERE: i32 = 4;

// OlTaskStatus (full set)
pub const OL_TASK_IN_PROGRESS: i32 = 1;
pub const OL_TASK_COMPLETE: i32 = 2;
pub const OL_TASK_WAITING: i32 = 3;
pub const OL_TASK_DEFERRED: i32 = 4;

// OlResponseStatus (AppointmentItem.ResponseStatus)
pub const OL_RESPONSE_NONE: i32 = 0;
pub const OL_RESPONSE_ORGANIZED: i32 = 1;
pub const OL_RESPONSE_TENTATIVE: i32 = 2;
pub const OL_RESPONSE_ACCEPTED: i32 = 3;
pub const OL_RESPONSE_DECLINED: i32 = 4;
pub const OL_RESPONSE_NOT_RESPONDED: i32 = 5;
```

- [ ] **Step 2: Write the failing test module** — create `src/friendly.rs`:

```rust
//! Convert Outlook enum integers to/from the lowercase friendly words the
//! MCP API exposes, so callers see "accepted" / "busy" rather than 3 / 2.

use crate::constants as c;

pub fn importance_word(v: i32) -> &'static str {
    match v {
        c::OL_IMPORTANCE_LOW => "low",
        c::OL_IMPORTANCE_HIGH => "high",
        _ => "normal",
    }
}

pub fn response_word(v: i32) -> &'static str {
    match v {
        c::OL_RESPONSE_ORGANIZED => "organizer",
        c::OL_RESPONSE_TENTATIVE => "tentative",
        c::OL_RESPONSE_ACCEPTED => "accepted",
        c::OL_RESPONSE_DECLINED => "declined",
        c::OL_RESPONSE_NOT_RESPONDED => "not_responded",
        _ => "none",
    }
}

pub fn busy_status_word(v: i32) -> &'static str {
    match v {
        c::OL_FREE => "free",
        c::OL_TENTATIVE => "tentative",
        c::OL_OUT_OF_OFFICE => "out_of_office",
        c::OL_WORKING_ELSEWHERE => "working_elsewhere",
        _ => "busy",
    }
}

pub fn task_status_word(v: i32) -> &'static str {
    match v {
        c::OL_TASK_IN_PROGRESS => "in_progress",
        c::OL_TASK_COMPLETE => "complete",
        c::OL_TASK_WAITING => "waiting",
        c::OL_TASK_DEFERRED => "deferred",
        _ => "not_started",
    }
}

pub fn busy_status_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "free" => Some(c::OL_FREE),
        "tentative" => Some(c::OL_TENTATIVE),
        "busy" => Some(c::OL_BUSY),
        "out_of_office" => Some(c::OL_OUT_OF_OFFICE),
        "working_elsewhere" => Some(c::OL_WORKING_ELSEWHERE),
        _ => None,
    }
}

pub fn task_status_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "not_started" => Some(c::OL_TASK_NOT_STARTED),
        "in_progress" => Some(c::OL_TASK_IN_PROGRESS),
        "complete" => Some(c::OL_TASK_COMPLETE),
        "waiting" => Some(c::OL_TASK_WAITING),
        "deferred" => Some(c::OL_TASK_DEFERRED),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn words_map_known_and_unknown_values() {
        assert_eq!(importance_word(2), "high");
        assert_eq!(importance_word(99), "normal"); // unknown → default
        assert_eq!(response_word(3), "accepted");
        assert_eq!(response_word(5), "not_responded");
        assert_eq!(busy_status_word(0), "free");
        assert_eq!(busy_status_word(3), "out_of_office");
        assert_eq!(busy_status_word(99), "busy"); // unknown → default
        assert_eq!(task_status_word(1), "in_progress");
        assert_eq!(task_status_word(99), "not_started");
    }

    #[test]
    fn reverse_lookups_are_case_insensitive_and_reject_garbage() {
        assert_eq!(busy_status_to_id("Out_Of_Office"), Some(3));
        assert_eq!(busy_status_to_id("nope"), None);
        assert_eq!(task_status_to_id("COMPLETE"), Some(2));
        assert_eq!(task_status_to_id("nope"), None);
    }
}
```

- [ ] **Step 3: Wire the module** — add to `src/lib.rs` (alongside the existing `pub mod constants;`):

```rust
pub mod friendly;
```

- [ ] **Step 4: Run the tests**

Run: `cargo test friendly::`
Expected: `words_map_known_and_unknown_values ... ok`, `reverse_lookups_are_case_insensitive_and_reject_garbage ... ok`.

- [ ] **Step 5: Commit**

```bash
git add src/friendly.rs src/lib.rs src/constants.rs
git commit -m "Add friendly-word enum conversion module"
```

---

### Task 2: Category parse/join helpers + COM read/write wrappers

**Files:**
- Modify: `src/outlook/com.rs` (add pure `parse_categories`/`join_categories` with tests, and COM `get_item_categories`/`set_item_categories` wrappers)

**Interfaces:**
- Consumes: existing `com.rs` helpers `get_property`, `put_property`, `variant_from_str`, `variant_to_string`.
- Produces:
  - `com::parse_categories(&str) -> Vec<String>` (pure)
  - `com::join_categories(&[String]) -> String` (pure)
  - `com::get_item_categories(disp: &IDispatch) -> Vec<String>` (reads `Categories` property; empty vec on missing/empty)
  - `com::set_item_categories(disp: &IDispatch, cats: &[String]) -> windows::core::Result<()>` (writes joined string)

- [ ] **Step 1: Write the failing pure-logic test** — add near the existing pure tests in `src/outlook/com.rs` (the module already has a `#[cfg(test)] mod tests`; add these two functions above it and two asserts inside it):

Add the pure functions in the non-test area of `com.rs` (near `safe_filename`):

```rust
/// Outlook stores categories as one `", "`-joined string. Split it into names,
/// trimming whitespace and dropping empties. Mirrors how the Python client
/// would `.split(", ")`.
pub fn parse_categories(raw: &str) -> Vec<String> {
    raw.split(',')
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string())
        .collect()
}

/// Join category names back into the `", "`-separated string Outlook expects.
pub fn join_categories(cats: &[String]) -> String {
    cats.join(", ")
}
```

Add to the existing `#[cfg(test)] mod tests` in `com.rs`:

```rust
#[test]
fn categories_round_trip_and_trim() {
    assert_eq!(parse_categories("Work, Receipts"), vec!["Work", "Receipts"]);
    assert_eq!(parse_categories("  Work ,  Personal "), vec!["Work", "Personal"]);
    assert_eq!(parse_categories(""), Vec::<String>::new());
    assert_eq!(join_categories(&["Work".into(), "Personal".into()]), "Work, Personal");
    assert_eq!(join_categories(&[]), "");
}
```

- [ ] **Step 2: Run it fail-then-pass**

Run: `cargo test outlook::com::tests::categories_round_trip_and_trim`
Expected: PASS (pure functions, no COM).

- [ ] **Step 3: Add the COM wrappers** — append to the COM section of `src/outlook/com.rs` (after `call_method`):

```rust
/// Read an item's color categories (empty vec if the property is missing or blank).
pub fn get_item_categories(disp: &IDispatch) -> Vec<String> {
    let raw = get_property(disp, "Categories")
        .map(|v| variant_to_string(&v))
        .unwrap_or_default();
    parse_categories(&raw)
}

/// Overwrite an item's categories with the given list.
pub fn set_item_categories(disp: &IDispatch, cats: &[String]) -> windows::core::Result<()> {
    put_property(disp, "Categories", variant_from_str(&join_categories(cats)))
}
```

- [ ] **Step 4: Build (COM wrappers are compile-checked here; exercised live in later plans)**

Run: `cargo build`
Expected: compiles clean, no warnings.

- [ ] **Step 5: Commit**

```bash
git add src/outlook/com.rs
git commit -m "Add category parse/join and COM read/write helpers"
```

---

### Task 3: Add `categories` to summary types and populate them

**Files:**
- Modify: `src/outlook/types.rs` (add `categories: Vec<String>` to `EmailSummary`, `EventSummary`, `TaskSummary`, `NoteSummary`)
- Modify: `src/outlook/client.rs` (populate `categories` in `email_summary`, `event_summary`, `task_summary`, `note_summary` via `get_item_categories`)
- Modify: `src/outlook/fake.rs` (add `categories` to every canned summary the fake returns)
- Modify: `src/outlook/types.rs` test (the flatten test constructs `EmailSummary` — add the field)

**Interfaces:**
- Consumes: `com::get_item_categories` (Task 2).
- Produces: every summary now serializes a `"categories": [...]` array.

- [ ] **Step 1: Add the field to the four structs in `src/outlook/types.rs`**

Add `pub categories: Vec<String>,` to `EmailSummary` (after `has_attachments`), `EventSummary` (after `is_meeting`), `TaskSummary` (after `importance`), and `NoteSummary` (after `created`). Example for `EmailSummary`:

```rust
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
    pub categories: Vec<String>,
}
```

- [ ] **Step 2: Run the build to see what breaks**

Run: `cargo build`
Expected: FAIL — `client.rs` and `fake.rs` construct these structs without the new field (`missing field 'categories'`), and the `types.rs` flatten test too. This lists every construction site to fix.

- [ ] **Step 3: Populate in `src/outlook/client.rs`**

In each summary builder, add the field, reading via the Task 2 helper. Import it in the `use crate::outlook::com::{...}` block (add `get_item_categories`). For `email_summary` add as the last field:

```rust
        categories: crate::outlook::com::get_item_categories(item),
```

Do the same in `event_summary`, `task_summary`, and `note_summary` (each takes an `item: &IDispatch`, so the call is identical).

- [ ] **Step 4: Populate the fake in `src/outlook/fake.rs`**

Every method that returns a summary constructs the struct literally. Add `categories: vec![]` (empty) to each — EXCEPT give one representative non-empty value so a test can assert it flows through. In the `list_emails` fake return, use:

```rust
        categories: vec!["Work".to_string()],
```

and `categories: vec![]` in the others. (The exact fake methods to edit: `list_emails`, `search_emails`, `get_email`'s summary, `list_events`, `get_event`, `list_tasks`, `list_notes`, `get_note` — every place a `*Summary` is built.)

- [ ] **Step 5: Fix the `types.rs` flatten test** — add `categories: vec![],` to the `EmailSummary { .. }` literal in `email_detail_flattens_summary_fields_at_top_level`.

- [ ] **Step 6: Add a fake-client assertion** — in `tests/tools.rs`, extend an existing email test (e.g. `list_emails_uses_defaults` or add a small new test) to assert categories flow through:

```rust
#[tokio::test]
async fn list_emails_returns_categories() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .list_emails(Parameters(ListEmailsParams {
            folder: "inbox".to_string(),
            count: 10,
            unread_only: false,
        }))
        .await
        .unwrap();
    let json = result_json(&result);
    assert_eq!(json[0]["categories"], serde_json::json!(["Work"]));
}
```

- [ ] **Step 7: Run the full suite**

Run: `cargo test`
Expected: all existing tests + the new `list_emails_returns_categories` pass; build clean.

- [ ] **Step 8: Commit**

```bash
git add src/outlook/types.rs src/outlook/client.rs src/outlook/fake.rs tests/tools.rs
git commit -m "Add categories field to all summary types"
```

---

### Task 4: Convert raw-number output fields to friendly words

**Files:**
- Modify: `src/outlook/types.rs` (`EventDetail.response_status: Option<i32>` → `response: String`; `TaskSummary.status: i32` → `String` and `importance: i32` → `String`)
- Modify: `src/outlook/client.rs` (`event_summary`/`get_event`, `task_summary` produce friendly words via `friendly::*`)
- Modify: `src/outlook/fake.rs` (fake returns friendly-word strings)
- Modify: `tests/tools.rs` (any test asserting the old numeric fields)

**Interfaces:**
- Consumes: `friendly::response_word`, `friendly::task_status_word`, `friendly::importance_word` (Task 1).
- Produces: `EventDetail.response` (String, friendly), `TaskSummary.status`/`importance` (String, friendly).

- [ ] **Step 1: Change the types in `src/outlook/types.rs`**

`EventDetail`: replace `pub response_status: Option<i32>,` with `pub response: String,`.
`TaskSummary`: change `pub status: i32,` → `pub status: String,` and `pub importance: i32,` → `pub importance: String,`.

- [ ] **Step 2: Build to enumerate breakage**

Run: `cargo build`
Expected: FAIL at the client + fake construction sites (type mismatch). This lists exactly what to fix.

- [ ] **Step 3: Fix `src/outlook/client.rs`**

In `get_event` (builds `EventDetail`), replace the `response_status` read with:

```rust
        response: crate::friendly::response_word(
            crate::outlook::com::variant_to_i32(&get_property(&item, "ResponseStatus").unwrap_or_default())
                .unwrap_or(crate::constants::OL_RESPONSE_NONE),
        )
        .to_string(),
```

In `task_summary`, replace the raw `status`/`importance` reads with:

```rust
        status: crate::friendly::task_status_word(
            variant_to_i32(&get_property(item, "Status").unwrap_or_default())
                .unwrap_or(c::OL_TASK_NOT_STARTED),
        )
        .to_string(),
        importance: crate::friendly::importance_word(
            variant_to_i32(&get_property(item, "Importance").unwrap_or_default())
                .unwrap_or(c::OL_IMPORTANCE_NORMAL),
        )
        .to_string(),
```

- [ ] **Step 4: Fix `src/outlook/fake.rs`**

`get_event` fake: replace `response_status: Some(3)` (or similar) with `response: "accepted".to_string()`.
`list_tasks` fake: replace `status: 0, importance: 1` with `status: "not_started".to_string(), importance: "normal".to_string()`.

- [ ] **Step 5: Fix any `tests/tools.rs` assertions**

Search for tests asserting `response_status`, task `status`, or `importance` as numbers and update to the friendly strings (e.g. `create_task_passes_importance` asserts the *input* arg, not output — leave input assertions alone; only fix output-shape assertions). Run the next step to find them.

- [ ] **Step 6: Run the full suite**

Run: `cargo test`
Expected: all pass. If a test fails asserting an old number, update it to the friendly word.

- [ ] **Step 7: Commit**

```bash
git add src/outlook/types.rs src/outlook/client.rs src/outlook/fake.rs tests/tools.rs
git commit -m "Return friendly words for response/status/importance instead of raw enum numbers"
```

---

## Self-Review

- **Spec coverage:** Cross-cutting §1 (categories first-class — *visible* part) ✅ Tasks 2–3; the *filter/settable* parts land in later plans. §2 (friendly words) ✅ Tasks 1 & 4. Tolerance principle ✅ (new reads use `.unwrap_or_default()`).
- **Placeholder scan:** none — every step has concrete code or an exact command.
- **Type consistency:** `friendly::*` signatures in Task 1 match their call sites in Task 4; `get_item_categories`/`set_item_categories` in Task 2 match the call in Task 3; the four struct field additions in Task 3 match the populate sites.
- **Known follow-on:** Tasks 11 (tasks) and 12 (notes) will accept friendly-word *inputs* for status/importance via `friendly::*_to_id`; those reverse helpers are built here in Task 1 so they're ready.

## Execution Handoff

This is Plan 1 of 12 (see `2026-07-07-outlook-mcp-v2-plans-index.md`). Once complete and green, proceed to Plan 2 (Email finder). Execute this plan with superpowers:subagent-driven-development (fresh subagent per task + review) or superpowers:executing-plans (inline with checkpoints).
