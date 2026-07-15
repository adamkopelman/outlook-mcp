# v2 Plan 8 — update_event + delete_event Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give calendar events full CRUD by adding two new tools: `update_event` (edit any field of an existing event — subject/time/location/body/show_as/categories/attendees/reminder/all_day — with a `send_update` flag controlling whether a meeting notifies attendees) and `delete_event` (soft-delete an event, with a `send_cancellation` flag controlling whether an organized meeting notifies attendees of the cancellation). Matches the spec's "Calendar" row: `update_event (new)`, `delete_event (new)`.

**Architecture:** `update_event` follows the `update_email` pattern (Plan 5): an `EventUpdate` params struct (mirroring `EmailUpdate`) carrying every optional field, applied field-by-field with a `changed` list, no `#[derive(Default)]` (mirrors `CreateEventInput`'s rationale — `send_update` must be stated explicitly at every call site rather than silently defaulting to `false`). Unlike `update_email`, there is no `move_to`/id-changing step — calendar items don't change folders in this API — so field order is simply cosmetic, not a correctness constraint. The one genuinely new COM technique is **removing** a recipient (`remove_attendees`): `Recipients.Remove(index)` by index, iterated in reverse so removing doesn't shift not-yet-visited indices — split into its own task (3) since it's the trickiest part, reusing Task 2's simpler field-write task as a stepping stone (mirrors Plan 7's task-per-capability split). `delete_event` follows the `delete_email` pattern (a single small task): detect whether you organize the meeting (`MeetingStatus == olMeeting`), and if so mark it `olMeetingCanceled` and optionally `Send()` (which delivers the cancellation) before `Delete()`; a personal appointment or a meeting you don't organize just calls `Delete()` — soft-delete to Deleted Items either way, matching `delete_email`'s note field.

**Tech Stack:** Rust, `windows` 0.62.2 COM, `rmcp` 2.1.0 tool macros, `serde_json`, `chrono`.

## Global Constraints

- **Target crate:** `C:\Users\adamk\projects\outlook-mcp-rs` (the Rust impl, NOT the Python `outlook-mcp`). Edition 2024, rustc 1.95.0.
- **Two implementors per trait change.** `OutlookClient` lives in `src/outlook/mod.rs`; every signature change touches BOTH `WindowsOutlookClient` (`src/outlook/client.rs`) and `FakeOutlookClient` (`src/outlook/fake.rs`), plus the tool layer (`src/server.rs`) and tests (`tests/tools.rs`). This plan does NOT touch any existing call sites in `tests/live_outlook.rs` (no existing trait method's signature changes) — it only adds new ones, plus Task 5 opportunistically improves two existing tests once `delete_event` exists (see Task 5).
- **Reuse existing helpers.** `friendly::busy_status_to_id(&str) -> Option<i32>` (show_as), `com::get_item_categories`/`com::set_item_categories` (categories), and `add_meeting_recipient` (adding a tiered attendee, defined in Plan 7's Task 2 next to `event_summary` in `client.rs`) all already exist — do not duplicate them.
- **No recurrence.** Recurring events are edited/deleted as a whole series automatically (there is no per-occurrence targeting in this API) — no special handling needed; Plan 9 (Recurrence) is a separate, later concern and doesn't touch this plan's code.
- **Zero warnings** on `cargo build` / `cargo test` before the plan is pushed.
- **Model policy:** Task 1 = **sonnet** (struct + interface ripple, mirrors Plan 5 Task 1), Task 2 = **sonnet** (simple property writes, reuses existing helpers), Task 3 = **opus** (attendee removal by index + the Save-vs-Send decision), Task 4 = **sonnet** (delete_event, mirrors delete_email plus one conditional branch), Task 5 = **haiku** (live tests + docs).

---

### Task 1: `EventUpdate` struct — interface + fake + tool layer (interim: real client stubbed)

Add the `EventUpdate` struct, the `update_event` trait method, a fully-working fake implementation (so tool-level tests can run immediately), and the `update_event` MCP tool. The real COM body is a `todo!()` stub, filled in by Tasks 2–3.

**Files:**
- Modify: `src/outlook/mod.rs` (add `EventUpdate` struct after `CreateEventInput`; add `update_event` trait method after `respond_to_meeting`, line ~105)
- Modify: `src/outlook/fake.rs` (add `update_event` after `respond_to_meeting`, line ~191)
- Modify: `src/outlook/client.rs` (add `update_event` as a `todo!()` stub after `respond_to_meeting`, line ~1052)
- Modify: `src/server.rs` (add `UpdateEventParams` after `RespondToMeetingParams`, line ~232; add the `update_event` tool method after `respond_to_meeting`, line ~435)
- Modify: `tests/tools.rs` (add `update_event` tool tests)

**Interfaces:**
- Produces: `pub struct EventUpdate { pub event_id: String, pub subject: Option<String>, pub start: Option<String>, pub end: Option<String>, pub location: Option<String>, pub body: Option<String>, pub all_day: Option<bool>, pub reminder_minutes: Option<i32>, pub show_as: Option<String>, pub add_categories: Option<Vec<String>>, pub remove_categories: Option<Vec<String>>, pub add_required_attendees: Option<Vec<String>>, pub add_optional_attendees: Option<Vec<String>>, pub remove_attendees: Option<Vec<String>>, pub send_update: bool }` — no `#[derive(Default)]` (same rationale as `CreateEventInput`: forces every call site to state `send_update` explicitly).
- Produces: trait method `fn update_event(&self, u: EventUpdate) -> Result<Value, ToolError>;`
- Produces (fake/real return contract): `{"status": "updated", "id": event_id, "changed": [...]}` where `changed` lists, in field-declaration order, the fields that were touched: any of `"subject"`, `"start"`, `"end"`, `"location"`, `"body"`, `"all_day"`, `"reminder_minutes"`, `"show_as"`, `"add_categories"`, `"remove_categories"`, `"add_required_attendees"`, `"add_optional_attendees"`, `"remove_attendees"`. (No `"move_to"`/id-change — events don't change folders here.)

- [ ] **Step 1: Add the `EventUpdate` struct to `src/outlook/mod.rs`**

Add directly below the `CreateEventInput` struct (after its closing `}`, before `pub trait OutlookClient`):

```rust
/// All changes `update_event` can apply to one existing calendar event. Every
/// field except `event_id` is optional; supplying several applies all of
/// them. There is no `move_to` — events don't change folders in this API —
/// so, unlike `EmailUpdate`, field application order is cosmetic, not a
/// correctness constraint. Adding either attendee tier converts a personal
/// appointment into a meeting. `send_update` (no default here — the tool
/// layer defaults it to `true`) controls whether a meeting's edits are
/// delivered to attendees or applied quietly to your own copy only; a
/// personal (non-meeting) appointment always just saves, regardless.
#[derive(Debug, Clone)]
pub struct EventUpdate {
    pub event_id: String,
    pub subject: Option<String>,
    pub start: Option<String>,
    pub end: Option<String>,
    pub location: Option<String>,
    pub body: Option<String>,
    pub all_day: Option<bool>,
    pub reminder_minutes: Option<i32>,
    pub show_as: Option<String>,
    pub add_categories: Option<Vec<String>>,
    pub remove_categories: Option<Vec<String>>,
    pub add_required_attendees: Option<Vec<String>>,
    pub add_optional_attendees: Option<Vec<String>>,
    pub remove_attendees: Option<Vec<String>>,
    pub send_update: bool,
}
```

- [ ] **Step 2: Add the trait method in `src/outlook/mod.rs`**

Directly below `fn respond_to_meeting(...)` (currently line 105), add:

```rust
    fn update_event(&self, u: EventUpdate) -> Result<Value, ToolError>;
```

- [ ] **Step 3: Run the build to confirm it now fails**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo build 2>&1 | Select-String "update_event|EventUpdate" | Select-Object -First 8`
Expected: FAIL — `FakeOutlookClient`/`WindowsOutlookClient` don't satisfy the trait yet. (Red.)

- [ ] **Step 4: Implement `update_event` in `src/outlook/fake.rs`**

Add directly below `fn respond_to_meeting(...)` (currently line 191):

```rust
    fn update_event(&self, u: EventUpdate) -> Result<Value, ToolError> {
        self.record("update_event", json!({
            "event_id": u.event_id, "subject": u.subject, "start": u.start, "end": u.end,
            "location": u.location, "body": u.body, "all_day": u.all_day,
            "reminder_minutes": u.reminder_minutes, "show_as": u.show_as,
            "add_categories": u.add_categories, "remove_categories": u.remove_categories,
            "add_required_attendees": u.add_required_attendees,
            "add_optional_attendees": u.add_optional_attendees,
            "remove_attendees": u.remove_attendees, "send_update": u.send_update,
        }))?;
        let mut changed: Vec<&str> = Vec::new();
        if u.subject.is_some() { changed.push("subject"); }
        if u.start.is_some() { changed.push("start"); }
        if u.end.is_some() { changed.push("end"); }
        if u.location.is_some() { changed.push("location"); }
        if u.body.is_some() { changed.push("body"); }
        if u.all_day.is_some() { changed.push("all_day"); }
        if u.reminder_minutes.is_some() { changed.push("reminder_minutes"); }
        if u.show_as.is_some() { changed.push("show_as"); }
        if u.add_categories.is_some() { changed.push("add_categories"); }
        if u.remove_categories.is_some() { changed.push("remove_categories"); }
        if u.add_required_attendees.is_some() { changed.push("add_required_attendees"); }
        if u.add_optional_attendees.is_some() { changed.push("add_optional_attendees"); }
        if u.remove_attendees.is_some() { changed.push("remove_attendees"); }
        Ok(json!({"status": "updated", "id": u.event_id, "changed": changed}))
    }
```

Add `EventUpdate` to the fake's imports: change `use super::{CreateEventInput, EmailQuery, EmailUpdate, EventQuery, OutlookClient};` to also include `EventUpdate` (keep alphabetical: `...EmailUpdate, EventQuery, EventUpdate, OutlookClient`).

- [ ] **Step 5: Stub `update_event` in `src/outlook/client.rs` (real impl lands in Tasks 2-3)**

Add directly below `fn respond_to_meeting(...)` (currently ending line 1052, before the `// ---- Attachments (Task 14) ----` comment):

```rust
    fn update_event(&self, _u: EventUpdate) -> Result<Value, ToolError> {
        // Real COM implementation added in Plan 8 Tasks 2-3.
        todo!("update_event real COM impl — Plan 8 Tasks 2-3")
    }
```

Add `EventUpdate` to `client.rs`'s outlook imports (the `use crate::outlook::{create_event_status, CreateEventInput, EmailQuery, EmailUpdate, EventQuery, OutlookClient};` line) — same alphabetical slot as in `fake.rs`.

- [ ] **Step 6: Add `UpdateEventParams` in `src/server.rs`**

Add directly below `RespondToMeetingParams` (currently ending line 232):

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct UpdateEventParams {
    pub event_id: String,
    #[serde(default)]
    pub subject: Option<String>,
    #[serde(default)]
    pub start: Option<String>,
    #[serde(default)]
    pub end: Option<String>,
    #[serde(default)]
    pub location: Option<String>,
    #[serde(default)]
    pub body: Option<String>,
    #[serde(default)]
    pub all_day: Option<bool>,
    #[serde(default)]
    pub reminder_minutes: Option<i32>,
    /// "free" | "tentative" | "busy" | "out_of_office" | "working_elsewhere".
    #[serde(default)]
    pub show_as: Option<String>,
    /// Category names to add (existing categories are preserved).
    #[serde(default)]
    pub add_categories: Option<Vec<String>>,
    /// Category names to remove.
    #[serde(default)]
    pub remove_categories: Option<Vec<String>>,
    /// Adding either attendee list converts a personal appointment into a meeting.
    #[serde(default)]
    pub add_required_attendees: Option<Vec<String>>,
    #[serde(default)]
    pub add_optional_attendees: Option<Vec<String>>,
    /// Names/emails to remove from either attendee tier.
    #[serde(default)]
    pub remove_attendees: Option<Vec<String>>,
    /// If the event is a meeting, notify attendees of these changes (default true).
    /// false = apply quietly to your own copy only. Ignored for non-meetings.
    #[serde(default = "default_true")]
    pub send_update: bool,
}
```

- [ ] **Step 7: Add the `update_event` tool method in `src/server.rs`**

Add directly below the `respond_to_meeting` tool method (currently ending line 435, before `// ---- Attachments ----`):

```rust
    #[tool(description = "Update an existing calendar event: subject, start/end, location, body, show_as, add/remove categories, add/remove attendees, reminder, all_day. Adding an attendee converts a personal appointment into a meeting. send_update (default true) notifies attendees if the event is a meeting.")]
    pub async fn update_event(
        &self,
        Parameters(p): Parameters<UpdateEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let u = EventUpdate {
            event_id: p.event_id, subject: p.subject, start: p.start, end: p.end,
            location: p.location, body: p.body, all_day: p.all_day,
            reminder_minutes: p.reminder_minutes, show_as: p.show_as,
            add_categories: p.add_categories, remove_categories: p.remove_categories,
            add_required_attendees: p.add_required_attendees,
            add_optional_attendees: p.add_optional_attendees,
            remove_attendees: p.remove_attendees, send_update: p.send_update,
        };
        let result = run_blocking(move || client.update_event(u)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Add `EventUpdate` to `server.rs`'s outlook import line (`use crate::outlook::{CreateEventInput, EmailQuery, EmailUpdate, EventQuery, EventUpdate, OutlookClient};`).

- [ ] **Step 8: Add tool tests in `tests/tools.rs`**

Add `UpdateEventParams` to the `use outlook_mcp_rs::server::{...}` import list (keep alphabetical). Add:

```rust
#[tokio::test]
async fn update_event_lists_changed_fields() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_event(Parameters(UpdateEventParams {
            event_id: EVENT_ID.to_string(),
            subject: Some("Renamed sync".to_string()),
            start: None, end: None, location: None, body: None, all_day: None,
            reminder_minutes: None, show_as: Some("tentative".to_string()),
            add_categories: Some(vec!["Work".to_string()]),
            remove_categories: None,
            add_required_attendees: Some(vec!["a@example.com".to_string()]),
            add_optional_attendees: None, remove_attendees: None,
            send_update: true,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["status"], "updated");
    assert_eq!(v["id"], EVENT_ID);
    assert_eq!(
        v["changed"],
        json!(["subject", "show_as", "add_categories", "add_required_attendees"])
    );
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "update_event");
    assert_eq!(args["send_update"], true);
}

#[tokio::test]
async fn update_event_remove_attendees_is_tracked() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_event(Parameters(UpdateEventParams {
            event_id: EVENT_ID.to_string(),
            subject: None, start: None, end: None, location: None, body: None,
            all_day: None, reminder_minutes: None, show_as: None,
            add_categories: None, remove_categories: None,
            add_required_attendees: None, add_optional_attendees: None,
            remove_attendees: Some(vec!["a@example.com".to_string()]),
            send_update: false,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["changed"], json!(["remove_attendees"]));
}
```

- [ ] **Step 9: Build and run the full suite (green; `todo!()` unreached)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.
Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass, including the two new `update_event_*` tests. (The fake path never reaches `client.rs`'s `todo!()`.)

- [ ] **Step 10: Commit**

```bash
git add src/outlook/mod.rs src/outlook/fake.rs src/outlook/client.rs src/server.rs tests/tools.rs
git commit -m "Add update_event tool (interface + fake + tool layer; real COM stubbed)"
```

---

### Task 2: Real COM implementation — simple fields (subject/time/location/body/all_day/reminder/show_as/categories)

Fill in everything except attendee add/remove and the send_update-aware Save-vs-Send decision (Task 3). This task always ends with a plain `Save()`, matching the pre-attendee-support behavior — Task 3 replaces that final line.

**Files:**
- Modify: `src/outlook/client.rs` (implement the non-attendee part of `update_event`)

**Interfaces:**
- Consumes: `EventUpdate` (Task 1, minus the three attendee fields and `send_update` — used in Task 3), `crate::friendly::busy_status_to_id`, `com::get_item_categories`/`set_item_categories`, `client.rs` plumbing (`with_com`, `mapi`, `get_item`, `parse_dt`, COM helpers).

- [ ] **Step 1: Replace the `todo!()` stub in `src/outlook/client.rs`**

```rust
    fn update_event(&self, u: EventUpdate) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &u.event_id)?;
            let mut changed: Vec<&str> = Vec::new();

            if let Some(subject) = &u.subject {
                put_property(&item, "Subject", variant_from_str(subject))?;
                changed.push("subject");
            }
            if let Some(start) = &u.start {
                put_property(&item, "Start", variant_from_datetime(&parse_dt(start, "start")?)?)?;
                changed.push("start");
            }
            if let Some(end) = &u.end {
                put_property(&item, "End", variant_from_datetime(&parse_dt(end, "end")?)?)?;
                changed.push("end");
            }
            if let Some(location) = &u.location {
                put_property(&item, "Location", variant_from_str(location))?;
                changed.push("location");
            }
            if let Some(body) = &u.body {
                put_property(&item, "Body", variant_from_str(body))?;
                changed.push("body");
            }
            if let Some(all_day) = u.all_day {
                put_property(&item, "AllDayEvent", variant_from_bool(all_day))?;
                changed.push("all_day");
            }
            if let Some(minutes) = u.reminder_minutes {
                put_property(&item, "ReminderSet", variant_from_bool(true))?;
                put_property(&item, "ReminderMinutesBeforeStart", variant_from_i32(minutes))?;
                changed.push("reminder_minutes");
            }
            if let Some(show_as) = u.show_as.as_deref().filter(|s| !s.is_empty()) {
                let busy_status = crate::friendly::busy_status_to_id(show_as).ok_or_else(|| {
                    ToolError::new(format!(
                        "invalid show_as {show_as:?}: expected \"free\", \"tentative\", \"busy\", \"out_of_office\", or \"working_elsewhere\""
                    ))
                })?;
                put_property(&item, "BusyStatus", variant_from_i32(busy_status))?;
                changed.push("show_as");
            }
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
            }

            // Attendee add/remove and the Save-vs-Send decision land in Task 3.
            // Interim: always just Save (matches pre-attendee-support behavior).
            call_method(&item, "Save", &mut [])?;

            Ok(json!({"status": "updated", "id": u.event_id, "changed": changed}))
        })
    }
```

- [ ] **Step 2: Build (green, zero warnings)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings. The `todo!()` is gone.

- [ ] **Step 3: Run the full suite**

Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass. (Fake path unaffected; this real path is verified live in Task 5.)

- [ ] **Step 4: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Implement update_event real COM for subject/time/location/body/reminder/show_as/categories"
```

---

### Task 3: Attendee add/remove + honor `send_update`

The trickiest COM technique in this plan: removing a recipient by index, iterated in reverse so an in-progress removal never shifts an index not yet visited. Also replaces Task 2's unconditional `Save()` with a Save-vs-Send decision based on whether the event is (now) a meeting and whether `send_update` is true.

**Files:**
- Modify: `src/outlook/client.rs` (add `remove_meeting_recipients` helper; rework the attendee/save-or-send tail of `update_event`)

**Interfaces:**
- Consumes: `EventUpdate.add_required_attendees`/`.add_optional_attendees`/`.remove_attendees`/`.send_update` (Task 1), `add_meeting_recipient` (existing, Plan 7 Task 2), `c::OL_RECIPIENT_REQUIRED`/`OL_RECIPIENT_OPTIONAL`/`OL_MEETING`/`OL_NONMEETING` (existing).
- Produces: `fn remove_meeting_recipients(recipients: &IDispatch, addresses: &[String]) -> Result<(), ToolError>` (private, `client.rs`).

- [ ] **Step 1: Add the `remove_meeting_recipients` helper in `src/outlook/client.rs`**

Place directly below `add_meeting_recipient`:

```rust
/// Removes every recipient whose `Name` or `Address` case-insensitively
/// matches any entry in `addresses`. Iterates from `Count` down to `1` —
/// `Recipients.Remove(index)` is 1-based and shifts every later index down
/// by one, so removing in reverse means an index we haven't visited yet is
/// never invalidated by an earlier removal.
fn remove_meeting_recipients(recipients: &IDispatch, addresses: &[String]) -> Result<(), ToolError> {
    let count = variant_to_i32(&get_property(recipients, "Count")?).unwrap_or(0);
    for i in (1..=count).rev() {
        let recipient = to_disp(call_method(recipients, "Item", &mut [variant_from_i32(i)])?)?;
        let name = variant_to_string(&get_property(&recipient, "Name").unwrap_or_default());
        let address = variant_to_string(&get_property(&recipient, "Address").unwrap_or_default());
        if addresses.iter().any(|a| a.eq_ignore_ascii_case(&name) || a.eq_ignore_ascii_case(&address)) {
            call_method(recipients, "Remove", &mut [variant_from_i32(i)])?;
        }
    }
    Ok(())
}
```

- [ ] **Step 2: Rework the attendee/save-or-send tail of `update_event`**

Replace the two lines added at the end of Task 2 —

```rust
            // Attendee add/remove and the Save-vs-Send decision land in Task 3.
            // Interim: always just Save (matches pre-attendee-support behavior).
            call_method(&item, "Save", &mut [])?;
```

— with:

```rust
            // Adding either tier converts a personal appointment into a
            // meeting; MeetingStatus must be set before Recipients.Add for a
            // previously-non-meeting item.
            let adding_attendees = u.add_required_attendees.as_ref().is_some_and(|v| !v.is_empty())
                || u.add_optional_attendees.as_ref().is_some_and(|v| !v.is_empty());
            if adding_attendees {
                let current_status =
                    variant_to_i32(&get_property(&item, "MeetingStatus")?).unwrap_or(c::OL_NONMEETING);
                if current_status == c::OL_NONMEETING {
                    put_property(&item, "MeetingStatus", variant_from_i32(c::OL_MEETING))?;
                }
                let recipients = to_disp(get_property(&item, "Recipients")?)?;
                for address in u.add_required_attendees.as_deref().unwrap_or(&[]) {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_REQUIRED)?;
                }
                for address in u.add_optional_attendees.as_deref().unwrap_or(&[]) {
                    add_meeting_recipient(&recipients, address, c::OL_RECIPIENT_OPTIONAL)?;
                }
                call_method(&recipients, "ResolveAll", &mut [])?;
                if u.add_required_attendees.is_some() { changed.push("add_required_attendees"); }
                if u.add_optional_attendees.is_some() { changed.push("add_optional_attendees"); }
            }
            if let Some(remove) = u.remove_attendees.as_ref().filter(|v| !v.is_empty()) {
                let recipients = to_disp(get_property(&item, "Recipients")?)?;
                remove_meeting_recipients(&recipients, remove)?;
                changed.push("remove_attendees");
            }

            // Save vs Send: only a meeting can notify attendees; a personal
            // appointment always just saves, regardless of send_update.
            let is_meeting =
                variant_to_i32(&get_property(&item, "MeetingStatus")?).unwrap_or(c::OL_NONMEETING)
                    != c::OL_NONMEETING;
            if is_meeting && u.send_update {
                call_method(&item, "Send", &mut [])?;
            } else {
                call_method(&item, "Save", &mut [])?;
            }
```

- [ ] **Step 3: Build (green, zero warnings)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.

- [ ] **Step 4: Run the full suite**

Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass. (The two-tier add + reverse-index remove + Save-vs-Send branching is COM-only logic verified live in Task 5, matching the precedent set by `create_event`'s attendee tiers in Plan 7.)

- [ ] **Step 5: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Add attendee add/remove and honor send_update in update_event"
```

---

### Task 4: `delete_event`

New tool: soft-delete a calendar event. If you organize the meeting (`MeetingStatus == olMeeting`), mark it canceled and — unless `send_cancellation` is `false` — `Send()` the cancellation before removing your own copy; anything else (a personal appointment, or a meeting you attend but don't organize) is simply `Delete()`d.

**Files:**
- Modify: `src/constants.rs` (add `OL_MEETING_CANCELED`)
- Modify: `src/outlook/mod.rs` (add `delete_event` trait method after `update_event`)
- Modify: `src/outlook/fake.rs` (add `delete_event` after `update_event`)
- Modify: `src/outlook/client.rs` (add real `delete_event` after `update_event`)
- Modify: `src/server.rs` (add `DeleteEventParams` + `delete_event` tool method after `update_event`)
- Modify: `tests/tools.rs` (add `delete_event` tool tests)

**Interfaces:**
- Produces: trait method `fn delete_event(&self, event_id: String, send_cancellation: bool) -> Result<Value, ToolError>;`
- Produces (fake/real return contract): `{"status": "deleted", "subject": ..., "note": "..."}` (mirrors `delete_email`'s shape). `note` is one of: `"Meeting canceled; attendees notified. Moved to Deleted Items."` (organizer + `send_cancellation: true`), `"Meeting canceled without notifying attendees. Moved to Deleted Items."` (organizer + `send_cancellation: false`), or `"Moved to Deleted Items."` (not the organizer, or a personal appointment).

- [ ] **Step 1: Add `OL_MEETING_CANCELED` to `src/constants.rs`**

Add directly below the `OlMeetingStatus` block:

```rust
pub const OL_MEETING_CANCELED: i32 = 5;
```

- [ ] **Step 2: Add the trait method in `src/outlook/mod.rs`**

Directly below `fn update_event(...)`, add:

```rust
    fn delete_event(&self, event_id: String, send_cancellation: bool) -> Result<Value, ToolError>;
```

- [ ] **Step 3: Run the build to confirm it now fails**

Run: `cargo build 2>&1 | Select-String "delete_event" | Select-Object -First 8`
Expected: FAIL. (Red.)

- [ ] **Step 4: Implement `delete_event` in `src/outlook/fake.rs`**

Directly below the fake's `update_event`, add:

```rust
    fn delete_event(&self, event_id: String, send_cancellation: bool) -> Result<Value, ToolError> {
        self.record("delete_event", json!({"event_id": event_id, "send_cancellation": send_cancellation}))?;
        Ok(json!({"status": "deleted", "note": "Moved to Deleted Items."}))
    }
```

- [ ] **Step 5: Implement `delete_event` in `src/outlook/client.rs`**

Directly below the real `update_event`, add:

```rust
    fn delete_event(&self, event_id: String, send_cancellation: bool) -> Result<Value, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &event_id)?;
            let subject = variant_to_string(&get_property(&item, "Subject")?);
            let meeting_status =
                variant_to_i32(&get_property(&item, "MeetingStatus")?).unwrap_or(c::OL_NONMEETING);
            let note = if meeting_status == c::OL_MEETING {
                // You organize this meeting: mark it canceled, optionally
                // notify attendees, then remove your own copy.
                put_property(&item, "MeetingStatus", variant_from_i32(c::OL_MEETING_CANCELED))?;
                if send_cancellation {
                    call_method(&item, "Send", &mut [])?;
                    "Meeting canceled; attendees notified. Moved to Deleted Items."
                } else {
                    "Meeting canceled without notifying attendees. Moved to Deleted Items."
                }
            } else {
                "Moved to Deleted Items."
            };
            call_method(&item, "Delete", &mut [])?;
            Ok(json!({"status": "deleted", "subject": subject, "note": note}))
        })
    }
```

- [ ] **Step 6: Add `DeleteEventParams` and the tool method in `src/server.rs`**

Directly below `UpdateEventParams`, add:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct DeleteEventParams {
    pub event_id: String,
    /// If you organize the meeting, notify attendees of the cancellation (default true).
    #[serde(default = "default_true")]
    pub send_cancellation: bool,
}
```

Directly below the `update_event` tool method, add:

```rust
    #[tool(description = "Delete/cancel a calendar event (moves it to Deleted Items). If you organize the meeting, send_cancellation (default true) notifies attendees; if false, it's canceled quietly.")]
    pub async fn delete_event(
        &self,
        Parameters(DeleteEventParams { event_id, send_cancellation }): Parameters<DeleteEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let result = run_blocking(move || client.delete_event(event_id, send_cancellation)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

- [ ] **Step 7: Add tool tests in `tests/tools.rs`**

Add `DeleteEventParams` to the `use outlook_mcp_rs::server::{...}` import list. Add:

```rust
#[tokio::test]
async fn delete_event_returns_deleted_status() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .delete_event(Parameters(DeleteEventParams {
            event_id: EVENT_ID.to_string(),
            send_cancellation: true,
        }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["status"], "deleted");
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "delete_event");
    assert_eq!(args["send_cancellation"], true);
}
```

- [ ] **Step 8: Build and run the full suite (green)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.
Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass, including `delete_event_returns_deleted_status`.

- [ ] **Step 9: Commit**

```bash
git add src/constants.rs src/outlook/mod.rs src/outlook/fake.rs src/outlook/client.rs src/server.rs tests/tools.rs
git commit -m "Add delete_event tool (soft-delete, cancels + notifies if you organize the meeting)"
```

---

### Task 5: Live tests + docs

`#[ignore]`d end-to-end tests against real Outlook, plus two opportunistic improvements now that `delete_event` exists: the two existing Plan-7 live tests that currently document "delete the probe manually from the calendar" can finally clean up after themselves.

**Files:**
- Modify: `tests/live_outlook.rs`
- Modify: `TESTING.md`

**Interfaces:**
- Consumes: `EventUpdate` (Task 1), `WindowsOutlookClient::update_event`/`delete_event`/`create_event`/`get_event`.

- [ ] **Step 1: Add the `update_event` live test**

Add after `create_event_with_tiers_categories_and_show_as`:

```rust
#[test]
#[ignore]
fn update_event_edits_fields_and_manages_attendees() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P8 update probe".to_string(),
        start: "2099-01-07T09:00".to_string(),
        end: "2099-01-07T09:30".to_string(),
        body: None, location: None,
        required_attendees: Some(vec!["required-probe@example.com".to_string()]),
        optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: false,
    }).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    // Edit fields, add an optional attendee, remove the required one, quietly
    // (send_update: false — nothing is ever delivered).
    let updated = c.update_event(EventUpdate {
        event_id: id.clone(),
        subject: Some("outlook-mcp-rs P8 update probe (renamed)".to_string()),
        start: None, end: None,
        location: Some("Room 42".to_string()),
        body: None, all_day: None, reminder_minutes: Some(15),
        show_as: Some("tentative".to_string()),
        add_categories: Some(vec!["Work".to_string()]),
        remove_categories: None,
        add_required_attendees: None,
        add_optional_attendees: Some(vec!["optional-probe@example.com".to_string()]),
        remove_attendees: Some(vec!["required-probe@example.com".to_string()]),
        send_update: false,
    }).expect("update_event should succeed");
    assert_eq!(updated["status"], "updated");
    let changed = updated["changed"].as_array().unwrap();
    for field in ["subject", "location", "reminder_minutes", "show_as", "add_categories",
                  "add_optional_attendees", "remove_attendees"] {
        assert!(changed.iter().any(|v| v == field), "expected {field} in changed: {changed:?}");
    }

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    assert_eq!(detail.summary.subject, "outlook-mcp-rs P8 update probe (renamed)");
    assert_eq!(detail.summary.location, "Room 42");
    assert_eq!(detail.summary.show_as, "tentative");
    assert!(detail.summary.categories.iter().any(|cat| cat == "Work"));
    assert!(!detail.summary.required_attendees.contains("required-probe@example.com"));
    assert!(detail.summary.optional_attendees.contains("optional-probe@example.com"));

    c.delete_event(id, false).expect("cleanup delete_event");
}
```

- [ ] **Step 2: Add the `delete_event` live test**

Add directly after:

```rust
#[test]
#[ignore]
fn delete_event_removes_a_personal_appointment() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P8 delete probe".to_string(),
        start: "2099-01-08T09:00".to_string(),
        end: "2099-01-08T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true, // no attendees present, so this just Saves — nothing is sent
    }).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let deleted = c.delete_event(id.clone(), true).expect("delete_event should succeed");
    assert_eq!(deleted["status"], "deleted");
    assert_eq!(deleted["note"], "Moved to Deleted Items.");

    // Soft-deleted: get_event on the original id should now fail (moved to
    // Deleted Items changes its EntryID, same as delete_email's behavior).
    assert!(c.get_event(id).is_err());
}
```

- [ ] **Step 3: Clean up the two Plan-7 live tests now that `delete_event` exists**

In `create_event_then_delete_it`, replace the trailing comment block —

```rust
    // Calendar items don't have a dedicated "delete" tool in the trait; moving
    // into Deleted Items works for mail but not appointments — delete the test
    // event manually from the calendar after this test runs, or extend the
    // trait with a delete_event method if this becomes frequent enough to
    // automate.
    let _ = c.get_event(id); // just confirm it round-trips before manual cleanup
```

— with:

```rust
    let _ = c.get_event(&id).expect("get_event should round-trip before cleanup");
    c.delete_event(id, true).expect("cleanup delete_event");
```

(Adjust `get_event`'s call to match its actual signature — it takes `id` by value, so clone first if `id` is needed again: `c.get_event(id.clone()).expect(...); c.delete_event(id, true).expect(...);`.)

In `create_event_with_tiers_categories_and_show_as`, replace the trailing comment —

```rust
    // Calendar items have no dedicated delete tool yet (Plan 8's delete_event);
    // delete the probe manually from the calendar after this test runs.
```

— with:

```rust
    c.delete_event(id, false).expect("cleanup delete_event");
```

(`send_cancellation: false` — this event's attendee tiers are placeholder addresses that were never actually invited (`send: false` at creation), so there's nothing real to notify.)

- [ ] **Step 4: Confirm compile + ignored**

Run: `cargo build --tests 2>&1 | Select-Object -Last 2` → `Finished`.
Run: `cargo test --test live_outlook 2>&1 | Select-String "update_event_edits_fields_and_manages_attendees|delete_event_removes_a_personal_appointment"` → both show `ignored`.

- [ ] **Step 5: (If Outlook available) run live**

Run: `cargo test --test live_outlook -- --ignored update_event_edits_fields_and_manages_attendees delete_event_removes_a_personal_appointment create_event_then_delete_it create_event_with_tiers_categories_and_show_as 2>&1 | Select-Object -Last 15`
Expected: all pass. If any assertion fails, the property write, attendee add/remove, or cancel/delete path is wrong — investigate against real Outlook. Skip only if no Outlook is available.

- [ ] **Step 6: Update `TESTING.md`**

In the "Live system tests" preconditions bullet, remove the now-stale caveat: change

```
- Windows, with classic Outlook desktop installed
- Outlook is open and signed in to a normal mailbox
- You're comfortable with a handful of test items (a draft, a task, a note,
  a calendar event, each clearly named "outlook-mcp-rs live test ...") being
  created in that mailbox — most are cleaned up automatically, but the
  calendar event test currently requires manual deletion afterward (there's
  no `delete_event` tool; see `tests/live_outlook.rs` for why).
```

to:

```
- Windows, with classic Outlook desktop installed
- Outlook is open and signed in to a normal mailbox
- You're comfortable with a handful of test items (a draft, a task, a note,
  a calendar event, each clearly named "outlook-mcp-rs live test ...") being
  created in that mailbox — every live test cleans up after itself
  (calendar events via `delete_event`, since Plan 8).
```

Add a new bullet to the "Manual-only tests" section (below the existing `create_event` bullet), documenting that a *real* meeting update/cancellation notification still can't be automated:

```
5. Call `update_event` on a meeting you organize with `send_update: true` and
   real attendees; confirm they receive the update email. Call `delete_event`
   on a meeting you organize with `send_cancellation: true`; confirm they
   receive the cancellation. (The automated live test
   `update_event_edits_fields_and_manages_attendees` uses placeholder
   attendee addresses with `send_update: false`, so nothing is ever
   delivered — this is why real-recipient delivery still needs a manual check.)
```

- [ ] **Step 7: Commit**

```bash
git add tests/live_outlook.rs TESTING.md
git commit -m "Add live update_event/delete_event tests; clean up Plan 7 tests now that delete_event exists"
```

---

## Self-Review

**1. Spec coverage** (spec §`update_event (new)` and §`delete_event (new)`):
- `update_event`: `event_id` (required) ✅; `subject`/`start`/`end`/`location`/`body` ✅ Task 2; `show_as` ✅ Task 2; `add_categories`/`remove_categories` ✅ Task 2; `add_required_attendees`/`add_optional_attendees`/`remove_attendees` ✅ Task 3; `reminder_minutes`/`all_day` ✅ Task 2; `send_update` default `true`, gates notify-vs-quiet for meetings, no-op for personal appointments ✅ Task 3; return shape `{"status":"updated","id","changed":[...]}` ✅ Task 1 (fake) / Task 2-3 (real); recurring events edit the whole series — true by construction, since this plan never touches `RecurrencePattern` or expands occurrences, it always edits the single item identified by `event_id`.
- `delete_event`: `event_id` (required) ✅; `send_cancellation` default `true`, organizer-only cancel+notify, else quiet removal ✅ Task 4; soft-delete (Deleted Items) ✅ Task 4 (`Delete()`, same mechanism as `delete_email`); kept separate from `update_event` ✅ (distinct trait method, distinct tool).

**2. Placeholder scan:** The only `todo!()` is the deliberate Task 1 → Task 2/3 handoff for `update_event`'s real COM body, fully resolved by the end of Task 3. No TBD/FIXME left in any task's final state — the plan isn't pushed until all five tasks are green.

**3. Type consistency:** `EventUpdate` field names/types are identical across `mod.rs` (struct), `fake.rs` (destructure + record + `changed` logic), `client.rs` (destructure + COM writes + `changed` logic — same field-declaration order as `fake.rs`), `server.rs` (`UpdateEventParams` → `EventUpdate` mapping is 1:1), and `tests/tools.rs`/`tests/live_outlook.rs`. `delete_event(&self, event_id: String, send_cancellation: bool)` signature matches at all four call sites (trait, fake, real, tool). `OL_MEETING_CANCELED` is defined once in `constants.rs`, used only in `client.rs`. `remove_meeting_recipients`/`add_meeting_recipient` both take `&IDispatch` + address list/single address — consistent with existing COM helper conventions (borrow, not own). `busy_status_to_id`/`get_item_categories`/`set_item_categories` signatures match their existing Plan 1/6/7 definitions — not redefined here.

## Execution Handoff

Plan 8 of 12. Models: T1 sonnet (struct + interface ripple), T2 sonnet (simple property writes, reuses existing helpers), T3 opus (attendee removal by reverse index + Save-vs-Send branching), T4 sonnet (delete_event, mirrors delete_email), T5 haiku (live tests + docs). After all five are green with zero warnings, controller pushes `main` → Plan 9 (Recurrence — heaviest remaining plan, touches `create_event`/`update_event` via `GetRecurrencePattern()`). Trait-ripple checklist for Task 1's `update_event` addition and Task 4's `delete_event` addition: mod.rs + client.rs + fake.rs + server.rs + tests/tools.rs (no `tests/live_outlook.rs` call-site fixes needed — these are new trait methods, not signature changes to existing ones).
