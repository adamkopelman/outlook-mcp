# v2 Plan 7 — create_event enhancements (attendee tiers, categories, show_as, send flag) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give `create_event` optional-attendee tiers (`required_attendees`/`optional_attendees`, with `attendees` kept as a legacy alias for required), `categories`, `show_as` (busy status), and a `send` flag that lets a meeting be saved without notifying attendees — matching the spec's "create_event (+ 5 additions)" section. Recurrence is explicitly OUT of scope (Plan 9).

**Architecture:** `create_event`'s positional-argument signature is already at 8 params; adding 5 more would make it unreadable, so this plan replaces it with a `CreateEventInput` struct (mirroring the existing `EmailUpdate`/`EventQuery` pattern). The alias merge (`attendees` → `required_attendees`) happens once, at the MCP tool boundary in `server.rs`, so the trait/COM layer only ever deals with two clean tiers. The three-way return status (`"meeting_sent"`/`"meeting_saved"`/`"saved"`) is pure string logic independent of COM, so it's factored into a standalone function in `src/outlook/mod.rs` that both `FakeOutlookClient` and `WindowsOutlookClient` call — this is the one piece of this plan's business logic that gets a real fake-backed tool test instead of being live-test-only.

**Tech Stack:** Rust, `windows` 0.62.2 COM, `rmcp` 2.1.0 tool macros, `serde_json`, `chrono`.

## Global Constraints

- **Target crate:** `C:\Users\adamk\projects\outlook-mcp-rs` (the Rust impl, NOT the Python `outlook-mcp`). Edition 2024, rustc 1.95.0.
- **Two implementors per trait change.** `OutlookClient` lives in `src/outlook/mod.rs`; every signature change touches BOTH `WindowsOutlookClient` (`src/outlook/client.rs`) and `FakeOutlookClient` (`src/outlook/fake.rs`), plus the tool layer (`src/server.rs`) and tests (`tests/tools.rs`). Also scan `tests/live_outlook.rs` for call sites — this plan's Task 1 breaks TWO existing `create_event(...)` positional call sites there (`create_event_then_delete_it`, `list_events_filters_by_query_and_category`).
- **Tolerance:** new COM property reads in summary/detail builders use `.unwrap_or_default()`, never `?` — unchanged by this plan (it only touches `create_event`, which writes properties, not reads).
- **Reuse existing helpers.** `friendly::busy_status_to_id(&str) -> Option<i32>` already exists (`src/friendly.rs`) for the `show_as` validation — do not hand-roll a mapping. `com::set_item_categories(&IDispatch, &[String])` already exists for categories — do not duplicate it.
- **No recurrence in this plan.** `recurrence` is Plan 9 and out of scope here entirely.
- **Zero warnings** on `cargo build` / `cargo test` before the plan is pushed.
- **Model policy:** Task 1 = **sonnet** (struct + interface ripple), Task 2 = **opus** (per-recipient COM: capturing `Recipients.Add()`'s return value and setting `.Type`), Task 3 = **sonnet** (categories, reuses existing helper), Task 4 = **sonnet** (show_as, reuses existing helper), Task 5 = **sonnet** (pure status helper + wiring, real fake-backed tests), Task 6 = **haiku** (live tests).

---

### Task 1: `CreateEventInput` struct — interface swap (interim: behavior unchanged)

Replace `create_event`'s 8 positional params with a `CreateEventInput` struct carrying all 12 fields (7 old + `required_attendees`, `optional_attendees`, `categories`, `show_as`, `send`). The real client's logic is otherwise **unchanged** in this task: it still only honors a single required-attendee tier (now read from `input.required_attendees`, which `server.rs` populates by merging the legacy `attendees` param into it) and unconditionally sends when any attendee is present. `optional_attendees`, `categories`, `show_as`, and `send` are accepted into the struct and forwarded to the fake's recorded call, but have no COM effect yet — Tasks 2–5 add that, one field group at a time.

**Files:**
- Modify: `src/outlook/mod.rs:81-84` (replace the `create_event` trait method; add the `CreateEventInput` struct above it)
- Modify: `src/outlook/fake.rs:170-180` (swap `create_event` to take `CreateEventInput`, record every field)
- Modify: `src/outlook/client.rs:938-990` (swap `create_event` to take `CreateEventInput`; interim logic unchanged)
- Modify: `src/server.rs:194-208` (expand `CreateEventParams`; keep `attendees` as a legacy alias field)
- Modify: `src/server.rs:392-405` (the `create_event` tool method: merge `attendees` into `required_attendees`, build `CreateEventInput`)
- Modify: `tests/tools.rs` (fix `create_event_passes_attendees` to compile against the new params shape)
- Modify: `tests/live_outlook.rs:82-96` (`create_event_then_delete_it` — update the call site)
- Modify: `tests/live_outlook.rs:101-108` (`list_events_filters_by_query_and_category` — update the call site)

**Interfaces:**
- Produces: `pub struct CreateEventInput { pub subject: String, pub start: String, pub end: String, pub body: Option<String>, pub location: Option<String>, pub required_attendees: Option<Vec<String>>, pub optional_attendees: Option<Vec<String>>, pub all_day: bool, pub reminder_minutes: Option<i32>, pub categories: Option<Vec<String>>, pub show_as: Option<String>, pub send: bool }` — no `#[derive(Default)]` (forces every call site to state `send` explicitly rather than silently getting `false`).
- Produces: trait method `fn create_event(&self, input: CreateEventInput) -> Result<Value, ToolError>;`

- [ ] **Step 1: Add the `CreateEventInput` struct to `src/outlook/mod.rs`**

Add directly above the existing `create_event` trait method (currently line 81), right after the `EventQuery` struct:

```rust
/// All inputs for `create_event`. `required_attendees`/`optional_attendees`
/// are the two invite tiers Outlook shows a meeting organizer; any attendee
/// in either tier makes the item a meeting. `send` (default true in the
/// tool layer) controls whether a meeting is actually sent to attendees or
/// merely saved for later review — see `create_event_status` below for the
/// resulting status string.
#[derive(Debug, Clone)]
pub struct CreateEventInput {
    pub subject: String,
    pub start: String,
    pub end: String,
    pub body: Option<String>,
    pub location: Option<String>,
    pub required_attendees: Option<Vec<String>>,
    pub optional_attendees: Option<Vec<String>>,
    pub all_day: bool,
    pub reminder_minutes: Option<i32>,
    pub categories: Option<Vec<String>>,
    pub show_as: Option<String>,
    pub send: bool,
}
```

- [ ] **Step 2: Swap the trait method in `src/outlook/mod.rs`**

Replace:

```rust
    fn create_event(&self, subject: String, start: String, end: String,
        body: Option<String>, location: Option<String>,
        attendees: Option<Vec<String>>, all_day: bool,
        reminder_minutes: Option<i32>) -> Result<Value, ToolError>;
```

with:

```rust
    fn create_event(&self, input: CreateEventInput) -> Result<Value, ToolError>;
```

- [ ] **Step 3: Run the build to confirm it now fails**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo build 2>&1 | Select-String "create_event|CreateEventInput" | Select-Object -First 8`
Expected: FAIL — `FakeOutlookClient`/`WindowsOutlookClient` no longer satisfy the trait, and `server.rs` + both test files still call the old positional signature. (Red.)

- [ ] **Step 4: Swap `create_event` in `src/outlook/fake.rs`**

Replace the whole `create_event` fn (currently lines 170–180) with:

```rust
    fn create_event(&self, input: CreateEventInput) -> Result<Value, ToolError> {
        self.record("create_event", json!({
            "subject": input.subject, "start": input.start, "end": input.end,
            "body": input.body, "location": input.location,
            "required_attendees": input.required_attendees,
            "optional_attendees": input.optional_attendees,
            "all_day": input.all_day, "reminder_minutes": input.reminder_minutes,
            "categories": input.categories, "show_as": input.show_as,
            "send": input.send,
        }))?;
        Ok(json!({"status": "saved", "id": EVENT_ID, "subject": input.subject}))
    }
```

(The hardcoded `"status": "saved"` is fixed to reflect attendees/send in Task 5 — leave it as-is here.)

Add `CreateEventInput` to the fake's imports: change `use super::{EmailQuery, EmailUpdate, EventQuery, OutlookClient};` to `use super::{CreateEventInput, EmailQuery, EmailUpdate, EventQuery, OutlookClient};`.

- [ ] **Step 5: Swap `create_event` in `src/outlook/client.rs` (interim — same logic, new field names)**

Replace the whole `create_event` fn (currently lines 938–990) with:

```rust
    fn create_event(&self, input: CreateEventInput) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (app, _ns) = mapi()?;
            let appt = to_disp(call_method(
                &app,
                "CreateItem",
                &mut [variant_from_i32(c::OL_APPOINTMENT_ITEM)],
            )?)?;
            put_property(&appt, "Subject", variant_from_str(&input.subject))?;
            put_property(&appt, "Start", variant_from_datetime(&parse_dt(&input.start, "start")?)?)?;
            put_property(&appt, "End", variant_from_datetime(&parse_dt(&input.end, "end")?)?)?;
            if input.all_day {
                put_property(&appt, "AllDayEvent", variant_from_bool(true))?;
            }
            if let Some(body) = input.body.as_deref().filter(|b| !b.is_empty()) {
                put_property(&appt, "Body", variant_from_str(body))?;
            }
            if let Some(location) = input.location.as_deref().filter(|l| !l.is_empty()) {
                put_property(&appt, "Location", variant_from_str(location))?;
            }
            if let Some(minutes) = input.reminder_minutes {
                put_property(&appt, "ReminderSet", variant_from_bool(true))?;
                put_property(&appt, "ReminderMinutesBeforeStart", variant_from_i32(minutes))?;
            }
            // Interim: only the required tier has an effect (optional_attendees,
            // categories, show_as, send are accepted but not yet applied — see
            // Tasks 2-5). Matches the pre-Plan-7 behavior of always sending when
            // any attendee is present.
            let status = match input.required_attendees.filter(|a| !a.is_empty()) {
                Some(addresses) => {
                    put_property(&appt, "MeetingStatus", variant_from_i32(c::OL_MEETING))?;
                    let recipients = to_disp(get_property(&appt, "Recipients")?)?;
                    for address in &addresses {
                        call_method(&recipients, "Add", &mut [variant_from_str(address)])?;
                    }
                    call_method(&recipients, "ResolveAll", &mut [])?;
                    call_method(&appt, "Send", &mut [])?;
                    "meeting_sent"
                }
                None => {
                    call_method(&appt, "Save", &mut [])?;
                    "saved"
                }
            };
            Ok(json!({"status": status, "id": make_id(&appt)?, "subject": input.subject}))
        })
    }
```

- [ ] **Step 6: Expand `CreateEventParams` in `src/server.rs`**

Replace the whole `CreateEventParams` struct (currently lines 194–208) with:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CreateEventParams {
    pub subject: String,
    pub start: String,
    pub end: String,
    #[serde(default)]
    pub body: Option<String>,
    #[serde(default)]
    pub location: Option<String>,
    /// Legacy alias for `required_attendees`; merged in if both are given.
    #[serde(default)]
    pub attendees: Option<Vec<String>>,
    #[serde(default)]
    pub required_attendees: Option<Vec<String>>,
    #[serde(default)]
    pub optional_attendees: Option<Vec<String>>,
    #[serde(default)]
    pub all_day: bool,
    #[serde(default)]
    pub reminder_minutes: Option<i32>,
    #[serde(default)]
    pub categories: Option<Vec<String>>,
    /// "free" | "tentative" | "busy" | "out_of_office" | "working_elsewhere".
    #[serde(default)]
    pub show_as: Option<String>,
    /// If false, a meeting with attendees is saved (not sent) for later review.
    #[serde(default = "default_true")]
    pub send: bool,
}
```

- [ ] **Step 7: Update the `create_event` tool method in `src/server.rs`**

Replace the whole method (currently lines 392–405) with:

```rust
    #[tool(description = "Create a calendar event. required_attendees/optional_attendees invite two tiers (attendees is a legacy alias merged into required_attendees); any attendee makes it a meeting. categories and show_as (busy status) can be set on creation. send (default true) controls whether a meeting is actually sent to attendees or just saved for review.")]
    pub async fn create_event(
        &self,
        Parameters(CreateEventParams {
            subject, start, end, body, location, attendees, required_attendees,
            optional_attendees, all_day, reminder_minutes, categories, show_as, send,
        }): Parameters<CreateEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        // `attendees` is a legacy alias for the required tier; merge it in.
        let mut required = required_attendees.unwrap_or_default();
        required.extend(attendees.unwrap_or_default());
        let required_attendees = (!required.is_empty()).then_some(required);
        let input = CreateEventInput {
            subject, start, end, body, location, required_attendees, optional_attendees,
            all_day, reminder_minutes, categories, show_as, send,
        };
        let result = run_blocking(move || client.create_event(input)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Add `CreateEventInput` to `server.rs`'s outlook imports — the same `use crate::outlook::{...}` line that already brings in `EventQuery`/`EmailUpdate`.

- [ ] **Step 8: Fix `create_event_passes_attendees` in `tests/tools.rs`**

Replace the whole test with a version against the new params shape, asserting the alias merge:

```rust
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
            required_attendees: None,
            optional_attendees: None,
            all_day: false,
            reminder_minutes: None,
            categories: None,
            show_as: None,
            send: true,
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    // The legacy `attendees` alias merges into `required_attendees`.
    assert_eq!(args["required_attendees"], json!(["a@example.com"]));
}
```

- [ ] **Step 9: Fix the two `create_event(...)` call sites in `tests/live_outlook.rs`**

In `create_event_then_delete_it` (currently lines 82–96), replace the `c.create_event(...)` call:

```rust
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs live test event".to_string(),
        start: "2099-01-01T10:00:00".to_string(),
        end: "2099-01-01T10:30:00".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
    }).expect("create_event should succeed");
```

In `list_events_filters_by_query_and_category` (currently lines 101–108), replace the `c.create_event(...)` call:

```rust
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P6 filter probe".to_string(),
        start: "2099-01-05T09:00".to_string(),
        end: "2099-01-05T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
    }).expect("create_event");
```

Add `CreateEventInput` to the file's `use outlook_mcp_rs::outlook::{EmailQuery, EventQuery, OutlookClient, EmailUpdate};` line.

- [ ] **Step 10: Build and run the full suite (green)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.
Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass, including `create_event_passes_attendees`.

- [ ] **Step 11: Commit**

```bash
git add src/outlook/mod.rs src/outlook/fake.rs src/outlook/client.rs src/server.rs tests/tools.rs tests/live_outlook.rs
git commit -m "Replace create_event's positional params with CreateEventInput"
```

---

### Task 2: Attendee tiers (`required_attendees` / `optional_attendees`)

Wire both attendee tiers into the real COM call: each address added via `Recipients.Add()` gets its `.Type` set to `olRequired` (1) or `olOptional` (2) depending on which list it came from. Any attendee in either tier still makes the item a meeting.

**Files:**
- Modify: `src/constants.rs` (add `OL_RECIPIENT_REQUIRED`/`OL_RECIPIENT_OPTIONAL`)
- Modify: `src/outlook/client.rs` (rework the attendee block in `create_event` to use both tiers)

**Interfaces:**
- Consumes: `CreateEventInput.required_attendees`/`.optional_attendees` (Task 1).
- Produces: `create_event` sets attendee `.Type` per tier; `is_meeting` (surfaced via `event_summary`, already built in Plan 6) is true whenever either tier is non-empty.

- [ ] **Step 1: Add the two `OlMeetingRecipientType` constants to `src/constants.rs`**

Add directly below the `OlMeetingStatus` block:

```rust
// OlMeetingRecipientType (Recipient.Type on an AppointmentItem)
pub const OL_RECIPIENT_REQUIRED: i32 = 1;
pub const OL_RECIPIENT_OPTIONAL: i32 = 2;
```

- [ ] **Step 2: Add a private `add_meeting_recipient` helper in `src/outlook/client.rs`**

Place this free fn next to `event_summary` (or any other module-level helper):

```rust
/// Adds `address` to `recipients` and marks it required or optional. The
/// `Recipient` object `Recipients.Add()` returns must have its `.Type` set
/// explicitly — Outlook does not infer tier from call order.
fn add_meeting_recipient(recipients: &IDispatch, address: &str, role: i32) -> Result<(), ToolError> {
    let recipient = to_disp(call_method(recipients, "Add", &mut [variant_from_str(address)])?)?;
    put_property(&recipient, "Type", variant_from_i32(role))?;
    Ok(())
}
```

- [ ] **Step 3: Rework the attendee block in `create_event` (`src/outlook/client.rs`)**

Replace the `match input.required_attendees...` block added in Task 1 with:

```rust
            let required = input.required_attendees.unwrap_or_default();
            let optional = input.optional_attendees.unwrap_or_default();
            let status = if !required.is_empty() || !optional.is_empty() {
                put_property(&appt, "MeetingStatus", variant_from_i32(c::OL_MEETING))?;
                let recipients = to_disp(get_property(&appt, "Recipients")?)?;
                for address in &required {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_REQUIRED)?;
                }
                for address in &optional {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_OPTIONAL)?;
                }
                call_method(&recipients, "ResolveAll", &mut [])?;
                // Interim: still always sends when any attendee is present —
                // Task 5 makes this honor `input.send`.
                call_method(&appt, "Send", &mut [])?;
                "meeting_sent"
            } else {
                call_method(&appt, "Save", &mut [])?;
                "saved"
            };
```

- [ ] **Step 4: Build (green, zero warnings)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.

- [ ] **Step 5: Run the full suite**

Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass. (The fake path is unaffected — the two-tier COM behavior is verified live in Task 6.)

- [ ] **Step 6: Commit**

```bash
git add src/constants.rs src/outlook/client.rs
git commit -m "Add required/optional attendee tiers to create_event"
```

---

### Task 3: `categories` on `create_event`

Assign color categories at creation time by reusing the existing `com::set_item_categories` helper — the same one `update_email`/`update_event`-family code already uses.

**Files:**
- Modify: `src/outlook/client.rs` (set categories in `create_event`, before the attendee/save-or-send block)

**Interfaces:**
- Consumes: `CreateEventInput.categories` (Task 1), `com::set_item_categories(&IDispatch, &[String]) -> WinResult<()>` (existing, from Plan 1).

- [ ] **Step 1: Set categories in `create_event` (`src/outlook/client.rs`)**

Directly after the `reminder_minutes` block (before the attendee/`required`/`optional` block from Task 2), add:

```rust
            if let Some(categories) = input.categories.as_ref().filter(|c| !c.is_empty()) {
                set_item_categories(&appt, categories)?;
            }
```

`set_item_categories` is already imported in `client.rs` (used by `update_email`); no new `use` needed.

- [ ] **Step 2: Build (green, zero warnings)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.

- [ ] **Step 3: Run the full suite**

Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass. (Verified live in Task 6 via `get_event`'s existing `categories` field from Plan 1.)

- [ ] **Step 4: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Set categories on create_event"
```

---

### Task 4: `show_as` on `create_event`

Set the appointment's `BusyStatus` from the friendly `show_as` word, reusing `friendly::busy_status_to_id`. Outlook's own default when `BusyStatus` is never touched is `olBusy` (2), which already matches the spec's stated default of `"busy"` — so an absent `show_as` needs no explicit write.

**Files:**
- Modify: `src/outlook/client.rs` (set `BusyStatus` in `create_event` from `show_as`, with the same invalid-value error style used by `update_email`'s `importance`)

**Interfaces:**
- Consumes: `CreateEventInput.show_as` (Task 1), `friendly::busy_status_to_id(&str) -> Option<i32>` (existing, from Plan 6).

- [ ] **Step 1: Set `BusyStatus` in `create_event` (`src/outlook/client.rs`)**

Directly after the categories block added in Task 3, add:

```rust
            if let Some(show_as) = input.show_as.as_deref().filter(|s| !s.is_empty()) {
                let busy_status = crate::friendly::busy_status_to_id(show_as).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid show_as {show_as:?}: expected \"free\", \"tentative\", \"busy\", \"out_of_office\", or \"working_elsewhere\""
                    ))
                })?;
                put_property(&appt, "BusyStatus", variant_from_i32(busy_status))?;
            }
```

- [ ] **Step 2: Build (green, zero warnings)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.

- [ ] **Step 3: Run the full suite**

Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass. (Verified live in Task 6 via `get_event`'s existing `show_as` field from Plan 6; the invalid-value error path follows the same untested-by-design precedent as `update_email`'s `importance` validation — real-client-only, not fake-testable.)

- [ ] **Step 4: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Set show_as (BusyStatus) on create_event"
```

---

### Task 5: `send` flag + shared `create_event_status` helper (fake-backed tests)

Make `send: false` actually save-without-sending when the event is a meeting, and replace the two hardcoded status strings with a single pure helper shared by both implementors — this is the one piece of Plan 7 logic simple enough to unit-test and fake-test directly, so do that instead of deferring to Task 6's live tests.

**Files:**
- Modify: `src/outlook/mod.rs` (add `create_event_status` + its unit tests)
- Modify: `src/outlook/client.rs` (honor `input.send`; use the shared helper)
- Modify: `src/outlook/fake.rs` (use the shared helper so its recorded status is realistic)
- Modify: `tests/tools.rs` (add a tool-level test covering all three status outcomes through the fake)

**Interfaces:**
- Produces: `pub fn create_event_status(has_attendees: bool, send: bool) -> &'static str` in `src/outlook/mod.rs`.
- Consumes: `CreateEventInput.send` (Task 1).

- [ ] **Step 1: Add the failing unit test for `create_event_status` in `src/outlook/mod.rs`**

Add a `#[cfg(test)] mod tests` block at the bottom of `src/outlook/mod.rs` (create one if none exists yet):

```rust
#[cfg(test)]
mod tests {
    use super::create_event_status;

    #[test]
    fn create_event_status_covers_all_three_outcomes() {
        assert_eq!(create_event_status(true, true), "meeting_sent");
        assert_eq!(create_event_status(true, false), "meeting_saved");
        assert_eq!(create_event_status(false, true), "saved");
        assert_eq!(create_event_status(false, false), "saved");
    }
}
```

- [ ] **Step 2: Run the test to confirm it fails**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --lib create_event_status_covers_all_three_outcomes 2>&1 | Select-Object -Last 8`
Expected: FAIL to compile (`create_event_status` doesn't exist yet). (Red.)

- [ ] **Step 3: Add `create_event_status` to `src/outlook/mod.rs`**

Add directly above the `#[cfg(test)]` block from Step 1 (or anywhere at module scope, e.g. right after the `CreateEventInput` struct):

```rust
/// The status string `create_event` returns: `"meeting_sent"` (attendees +
/// send), `"meeting_saved"` (attendees + no send), or `"saved"` (no
/// attendees, regardless of `send` — there's nothing to send or withhold).
pub fn create_event_status(has_attendees: bool, send: bool) -> &'static str {
    match (has_attendees, send) {
        (true, true) => "meeting_sent",
        (true, false) => "meeting_saved",
        (false, _) => "saved",
    }
}
```

- [ ] **Step 4: Run the test (green)**

Run: `cargo test --lib create_event_status_covers_all_three_outcomes 2>&1 | Select-String "test result"`
Expected: `test result: ok. 1 passed`.

- [ ] **Step 5: Use the helper and honor `send` in `src/outlook/client.rs`**

Replace the `if !required.is_empty() || !optional.is_empty() { ... } else { ... }` block from Task 2 with:

```rust
            let has_attendees = !required.is_empty() || !optional.is_empty();
            if has_attendees {
                put_property(&appt, "MeetingStatus", variant_from_i32(c::OL_MEETING))?;
                let recipients = to_disp(get_property(&appt, "Recipients")?)?;
                for address in &required {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_REQUIRED)?;
                }
                for address in &optional {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_OPTIONAL)?;
                }
                call_method(&recipients, "ResolveAll", &mut [])?;
                if input.send {
                    call_method(&appt, "Send", &mut [])?;
                } else {
                    call_method(&appt, "Save", &mut [])?;
                }
            } else {
                call_method(&appt, "Save", &mut [])?;
            }
            let status = create_event_status(has_attendees, input.send);
```

Add `create_event_status` to `client.rs`'s outlook imports — the same `use crate::outlook::{...}` line that brings in `EventQuery`/`CreateEventInput`.

- [ ] **Step 6: Use the helper in `src/outlook/fake.rs`**

Replace the fake's hardcoded `Ok(json!({"status": "saved", ...}))` line with:

```rust
        let has_attendees = input.required_attendees.as_ref().is_some_and(|v| !v.is_empty())
            || input.optional_attendees.as_ref().is_some_and(|v| !v.is_empty());
        let status = super::create_event_status(has_attendees, input.send);
        Ok(json!({"status": status, "id": EVENT_ID, "subject": input.subject}))
```

(This replaces only the final `Ok(...)` line of the Task-1 `create_event` fn — the `self.record(...)` call above it is unchanged.)

- [ ] **Step 7: Add the tool-level status test in `tests/tools.rs`**

```rust
#[tokio::test]
async fn create_event_status_reflects_attendees_and_send() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());

    let base = |required: Option<Vec<String>>, send: bool| CreateEventParams {
        subject: "Sync".to_string(),
        start: "2026-06-12T14:00".to_string(),
        end: "2026-06-12T15:00".to_string(),
        body: None, location: None, attendees: None,
        required_attendees: required, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send,
    };

    let r = server.create_event(Parameters(base(Some(vec!["a@example.com".to_string()]), true)))
        .await.unwrap();
    assert_eq!(result_json(&r)["status"], "meeting_sent");

    let r = server.create_event(Parameters(base(Some(vec!["a@example.com".to_string()]), false)))
        .await.unwrap();
    assert_eq!(result_json(&r)["status"], "meeting_saved");

    let r = server.create_event(Parameters(base(None, true))).await.unwrap();
    assert_eq!(result_json(&r)["status"], "saved");
}
```

- [ ] **Step 8: Build and run the full suite (green)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.
Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass, including `create_event_status_covers_all_three_outcomes` and `create_event_status_reflects_attendees_and_send`.

- [ ] **Step 9: Commit**

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs tests/tools.rs
git commit -m "Honor create_event's send flag via a shared status helper"
```

---

### Task 6: Live tests

`#[ignore]`d end-to-end tests against real Outlook: create a meeting with both attendee tiers, categories, and a non-default `show_as`, with `send: false` (so nothing is actually delivered), then verify every new field round-trips through `get_event`. Also confirm the no-attendee path is still `"saved"`.

**Files:**
- Modify: `tests/live_outlook.rs`
- Modify: `TESTING.md` (document that `send: true` with real attendees is manual-only, same precedent as `send_email`/`respond_to_meeting`)

**Interfaces:**
- Consumes: `CreateEventInput` (Task 1), `WindowsOutlookClient::create_event`/`get_event`.

- [ ] **Step 1: Add the live test**

Add after `create_event_then_delete_it`:

```rust
#[test]
#[ignore]
fn create_event_with_tiers_categories_and_show_as() {
    let c = client();
    // send:false means nothing is ever delivered, so a placeholder address
    // for the invite tiers is safe — Outlook stores it without resolving
    // for delivery.
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P7 tiers probe".to_string(),
        start: "2099-01-06T09:00".to_string(),
        end: "2099-01-06T09:30".to_string(),
        body: None, location: None,
        required_attendees: Some(vec!["required-probe@example.com".to_string()]),
        optional_attendees: Some(vec!["optional-probe@example.com".to_string()]),
        all_day: false, reminder_minutes: None,
        categories: Some(vec!["Work".to_string()]),
        show_as: Some("tentative".to_string()),
        send: false,
    }).expect("create_event should succeed");
    assert_eq!(created["status"], "meeting_saved");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_event(id).expect("get_event should succeed");
    assert!(detail.summary.required_attendees.contains("required-probe@example.com"));
    assert!(detail.summary.optional_attendees.contains("optional-probe@example.com"));
    assert!(detail.summary.categories.iter().any(|cat| cat == "Work"));
    assert_eq!(detail.summary.show_as, "tentative");
    assert!(detail.summary.is_meeting);
    // Calendar items have no dedicated delete tool yet (Plan 8's delete_event);
    // delete the probe manually from the calendar after this test runs.
}
```

- [ ] **Step 2: Confirm compile + ignored**

Run: `cargo build --tests 2>&1 | Select-Object -Last 2` → `Finished`.
Run: `cargo test --test live_outlook 2>&1 | Select-String "create_event_with_tiers_categories_and_show_as"` → shows `ignored`.

- [ ] **Step 3: (If Outlook available) run live**

Run: `cargo test --test live_outlook -- --ignored create_event_with_tiers_categories_and_show_as 2>&1 | Select-Object -Last 12`
Expected: `test result: ok. 1 passed`. If any assertion fails, the tier/category/show_as write or the `get_event` round-trip is wrong — investigate against real Outlook. Skip this step only if no Outlook is available; delete the probe manually from the calendar afterward regardless.

- [ ] **Step 4: Document the manual-only `send: true` check in `TESTING.md`**

Add a bullet under the manual-only section: verifying `create_event` with `send: true` and real attendees actually delivers an invite requires a real recipient and can't be automated (same reasoning as `send_email`/`respond_to_meeting`) — test by hand with your own address in `required_attendees`.

- [ ] **Step 5: Commit**

```bash
git add tests/live_outlook.rs TESTING.md
git commit -m "Add live create_event tiers/categories/show_as test; document manual send:true check"
```

---

## Self-Review

**1. Spec coverage** (spec §`create_event (+ 5 additions)`, recurrence excluded):
- Optional attendees (`required_attendees` + `optional_attendees`, `attendees` alias for required, any attendee → meeting) ✅ Tasks 1 (struct + alias merge), 2 (COM tiers)
- `categories` ✅ Task 3
- `show_as` (default `"busy"`, i.e. untouched `BusyStatus`) ✅ Task 4
- `send` (default true; false = save without sending) ✅ Task 5
- Return status `meeting_sent`/`meeting_saved`/`saved` ✅ Task 5 (`create_event_status`, fake-tested)
- `recurrence` — explicitly deferred to Plan 9, not touched here.

**2. Placeholder scan:** No `todo!()`/TBD. Task 1's "interim, not yet applied" state for `optional_attendees`/`categories`/`show_as`/`send` is deliberate and fully resolved by Task 5 — the plan isn't pushed until all 6 tasks are green, so no placeholder ships.

**3. Type consistency:** `CreateEventInput` field names/types are identical across `mod.rs` (struct), `fake.rs` (destructure + record), `client.rs` (destructure + COM writes), `server.rs` (`CreateEventParams` → `CreateEventInput` mapping incl. the `attendees`-alias merge), and both test files. `create_event_status(bool, bool) -> &'static str` signature matches at all three call sites (`mod.rs` unit test, `client.rs`, `fake.rs`). `OL_RECIPIENT_REQUIRED`/`OL_RECIPIENT_OPTIONAL` are defined once in `constants.rs` and used only in `client.rs`. `friendly::busy_status_to_id` and `com::set_item_categories` signatures match their existing definitions (Plans 1 and 6) — not redefined here.

## Execution Handoff

Plan 7 of 12. Models: T1 sonnet (struct + interface ripple), T2 opus (per-recipient COM), T3 sonnet (categories, reuses helper), T4 sonnet (show_as, reuses helper), T5 sonnet (pure helper + real fake-backed tests), T6 haiku (live tests). After all six are green with zero warnings, controller pushes `main` → Plan 8 (`update_event` + `delete_event`). Trait-ripple checklist for the Task 1 signature change: mod.rs + client.rs + fake.rs + server.rs + tests/tools.rs + tests/live_outlook.rs (TWO call sites there, both listed in Task 1).
