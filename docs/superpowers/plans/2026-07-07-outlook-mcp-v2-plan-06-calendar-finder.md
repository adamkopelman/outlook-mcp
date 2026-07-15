# v2 Plan 6 — Calendar finder (list_events filters + calendar_of; enriched get_event) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn `list_events` into a rich calendar finder — text/category/show_as/response/attendee/meeting/all-day filters plus the ability to view another person's shared calendar (`calendar_of`) — and enrich `EventSummary`/`EventDetail`/`get_event` output with the friendly fields (`show_as`, `my_response`, `required_attendees`, `optional_attendees`) they currently lack.

**Architecture:** Introduce an `EventQuery` params struct (mirroring the existing `EmailQuery` pattern) that replaces the two-arg `list_events(start_date, end_date)`. Enrich `EventSummary` with four new fields populated by the shared `event_summary()` builder, and slim `EventDetail` to `{summary, body}` (its former `required_attendees`/`optional_attendees`/`response` fields move into the summary, with `response` renamed to the friendly `my_response`). Because every new filter is expressible against an already-built `EventSummary`, the filters run **client-side** over the existing `GetFirst`/`GetNext` recurrence stream — exactly how `list_emails` already filters category/attachments — so no new COM reads are needed for filtering. `calendar_of` swaps the enumerated folder from the default calendar to a resolved recipient's shared calendar; all filters then apply identically.

**Tech Stack:** Rust, `windows` 0.62.2 COM, `rmcp` 2.1.0 tool macros, `serde_json`, `chrono`.

## Global Constraints

- **Target crate:** `C:\Users\adamk\projects\outlook-mcp-rs` (the Rust impl, NOT the Python `outlook-mcp`). Edition 2024, rustc 1.95.0.
- **Two implementors per trait change.** `OutlookClient` lives in `src/outlook/mod.rs`; every signature change touches BOTH `WindowsOutlookClient` (`src/outlook/client.rs`) and `FakeOutlookClient` (`src/outlook/fake.rs`), plus the tool layer (`src/server.rs`) and tests (`tests/tools.rs`). Also scan `tests/live_outlook.rs` for call sites (this plan: `get_event`/`create_event` are called there; `list_events` is not).
- **Tolerance:** new COM property reads in summary/detail builders use `.unwrap_or_default()`, never `?`, so a property missing on an odd item type yields an empty/default value rather than an error.
- **Reuse the friendly module.** `friendly::busy_status_word(i32)` and `friendly::response_word(i32)` already exist and are the ONLY way raw `BusyStatus`/`ResponseStatus` integers become words. Do not hand-roll mappings.
- **Return types.** `list_events` returns `Vec<EventSummary>`; `get_event` returns `EventDetail`.
- **Zero warnings** on `cargo build` / `cargo test` before the plan is pushed.
- **Model policy:** Task 1 = **sonnet** (types ripple + COM reads), Task 2 = **sonnet** (interface swap + fake + tool + tests), Task 3 = **sonnet** (client-side filter logic), Task 4 = **opus** (shared-calendar recipient resolution COM), Task 5 = **haiku** (live tests).

---

### Task 1: Enrich `EventSummary` / `EventDetail` output

Add `show_as`, `my_response`, `required_attendees`, `optional_attendees` to `EventSummary` (populated by the shared `event_summary()` builder), and slim `EventDetail` to `{summary, body}`. This ships richer `list_events` AND `get_event` output with no new filters yet.

**Files:**
- Modify: `src/outlook/types.rs` (add 4 fields to `EventSummary`; remove 3 fields from `EventDetail`)
- Modify: `src/outlook/client.rs` (`event_summary()` populates new fields; `get_event()` returns `{summary, body}`)
- Modify: `src/outlook/fake.rs` (both event literals gain the new summary fields; `get_event` literal slims)
- Modify: `tests/tools.rs` (assert `get_event` now surfaces `show_as`/`my_response`)

**Interfaces:**
- Produces: `EventSummary { id, subject, start, end, location, organizer, all_day, is_recurring, is_meeting, categories, show_as, my_response, required_attendees, optional_attendees }` (last four are new; `show_as`/`my_response` are `String` friendly words, `required_attendees`/`optional_attendees` are `String`).
- Produces: `EventDetail { #[serde(flatten)] summary, body }` — the former `required_attendees`, `optional_attendees`, and `response` fields are GONE (they now live in `summary`, with `response` renamed `my_response`).

- [ ] **Step 1: Add the failing tool test in `tests/tools.rs`**

Replace the existing `get_event_returns_subject` test (around lines 280–290) with a version that also asserts the new fields:

```rust
#[tokio::test]
async fn get_event_returns_subject_and_friendly_fields() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_event(Parameters(GetEventParams { event_id: EVENT_ID.to_string() }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["subject"], "Standup");
    // New enriched fields surface at the top level (EventDetail flattens the summary).
    assert_eq!(v["show_as"], "busy");
    assert_eq!(v["my_response"], "accepted");
    assert_eq!(v["required_attendees"], "");
    assert_eq!(v["optional_attendees"], "");
    // The old nested "response" key is gone (renamed to my_response in the summary).
    assert!(v.get("response").is_none());
}
```

- [ ] **Step 2: Run the test to confirm it fails (fields/methods don't exist yet)**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --test tools get_event_returns_subject_and_friendly_fields 2>&1 | Select-Object -Last 8`
Expected: FAIL to compile (`EventSummary` has no `show_as`/`my_response` field; `EventDetail` still has `response`). (Red.)

- [ ] **Step 3: Add the four fields to `EventSummary` in `src/outlook/types.rs`**

Replace the whole `EventSummary` struct (currently lines ~52–64) with:

```rust
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
    pub categories: Vec<String>,
    /// Busy status as a friendly word: "free"/"tentative"/"busy"/"out_of_office"/"working_elsewhere".
    pub show_as: String,
    /// This mailbox's response as a friendly word: "organizer"/"accepted"/"declined"/"tentative"/"not_responded"/"none".
    pub my_response: String,
    pub required_attendees: String,
    pub optional_attendees: String,
}
```

- [ ] **Step 4: Slim `EventDetail` in `src/outlook/types.rs`**

Replace the whole `EventDetail` struct (currently lines ~66–74) with:

```rust
#[derive(Debug, Clone, Serialize)]
pub struct EventDetail {
    #[serde(flatten)]
    pub summary: EventSummary,
    pub body: String,
}
```

- [ ] **Step 5: Populate the new fields in `event_summary()` (`src/outlook/client.rs`)**

Replace the whole `event_summary` fn (currently lines ~209–226) with:

```rust
/// `client.py::_event_summary`, enriched for v2 with show_as/my_response and the
/// attendee strings so every calendar filter can operate on the built summary.
fn event_summary(item: &IDispatch) -> Result<EventSummary, ToolError> {
    let meeting_status = variant_to_i32(&get_property(item, "MeetingStatus").unwrap_or_default())
        .unwrap_or(c::OL_NONMEETING);
    Ok(EventSummary {
        id: make_id(item)?,
        subject: variant_to_string(&get_property(item, "Subject").unwrap_or_default()),
        start: variant_to_iso_string(&get_property(item, "Start").unwrap_or_default()),
        end: variant_to_iso_string(&get_property(item, "End").unwrap_or_default()),
        location: variant_to_string(&get_property(item, "Location").unwrap_or_default()),
        organizer: variant_to_string(&get_property(item, "Organizer").unwrap_or_default()),
        all_day: variant_to_bool(&get_property(item, "AllDayEvent").unwrap_or_default())
            .unwrap_or(false),
        is_recurring: variant_to_bool(&get_property(item, "IsRecurring").unwrap_or_default())
            .unwrap_or(false),
        is_meeting: meeting_status != c::OL_NONMEETING,
        categories: get_item_categories(item),
        show_as: crate::friendly::busy_status_word(
            variant_to_i32(&get_property(item, "BusyStatus").unwrap_or_default())
                .unwrap_or(c::OL_BUSY),
        )
        .to_string(),
        my_response: crate::friendly::response_word(
            variant_to_i32(&get_property(item, "ResponseStatus").unwrap_or_default())
                .unwrap_or(c::OL_RESPONSE_NONE),
        )
        .to_string(),
        required_attendees: variant_to_string(
            &get_property(item, "RequiredAttendees").unwrap_or_default(),
        ),
        optional_attendees: variant_to_string(
            &get_property(item, "OptionalAttendees").unwrap_or_default(),
        ),
    })
}
```

- [ ] **Step 6: Simplify `get_event()` in `src/outlook/client.rs`**

The attendee/response reads now live in `event_summary`, so `get_event` only adds `body`. Replace the whole `get_event` fn (currently lines ~824–849) with:

```rust
    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &event_id)?;
            let summary = event_summary(&item)?;
            Ok(EventDetail {
                summary,
                body: truncate(&variant_to_string(&get_property(&item, "Body").unwrap_or_default())),
            })
        })
    }
```

- [ ] **Step 7: Update the fake event literals in `src/outlook/fake.rs`**

The fake's `list_events` and `get_event` build `EventSummary`/`EventDetail` by hand and must match the new shape. Replace the whole `list_events` fn (currently lines ~139–147) with:

```rust
    fn list_events(&self, start_date: Option<String>, end_date: Option<String>)
        -> Result<Vec<EventSummary>, ToolError> {
        self.record("list_events", json!({"start_date": start_date, "end_date": end_date}))?;
        Ok(vec![EventSummary {
            id: EVENT_ID.into(), subject: "Standup".into(), start: None, end: None,
            location: "".into(), organizer: "".into(), all_day: false,
            is_recurring: false, is_meeting: false, categories: vec![],
            show_as: "busy".into(), my_response: "accepted".into(),
            required_attendees: "".into(), optional_attendees: "".into(),
        }])
    }
```

(Note: this fn's signature is still the two-arg form here — it is swapped to `EventQuery` in Task 2. Task 1 only touches the returned literal.)

Replace the whole `get_event` fn (currently lines ~149–160) with:

```rust
    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError> {
        self.record("get_event", json!({"event_id": event_id}))?;
        Ok(EventDetail {
            summary: EventSummary {
                id: event_id, subject: "Standup".into(), start: None, end: None,
                location: "".into(), organizer: "".into(), all_day: false,
                is_recurring: false, is_meeting: false, categories: vec![],
                show_as: "busy".into(), my_response: "accepted".into(),
                required_attendees: "".into(), optional_attendees: "".into(),
            },
            body: "".into(),
        })
    }
```

- [ ] **Step 8: Build and run the test (green)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.
Run: `cargo test --test tools get_event_returns_subject_and_friendly_fields 2>&1 | Select-String "test result"`
Expected: `test result: ok. 1 passed`.

- [ ] **Step 9: Run the full suite (nothing else regressed)**

Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass. (The `list_events_passes_date_range` test still passes — its recorded args are unchanged in Task 1.)

- [ ] **Step 10: Commit**

```bash
git add src/outlook/types.rs src/outlook/client.rs src/outlook/fake.rs tests/tools.rs
git commit -m "Enrich EventSummary with show_as/my_response/attendees; slim EventDetail to summary+body"
```

---

### Task 2: `EventQuery` struct + `list_events` signature swap (interface + fake + tool + tests)

Replace `list_events(start_date, end_date)` with `list_events(q: EventQuery)` across the trait, fake, real client, tool layer, and tests. The real client destructures `EventQuery` but for now applies **only** the date range (the new filter fields are plumbed through and ignored until Task 3). Fake records every field.

**Files:**
- Modify: `src/outlook/mod.rs` (add `EventQuery` struct; swap trait method)
- Modify: `src/outlook/fake.rs` (swap `list_events` signature; record all fields)
- Modify: `src/outlook/client.rs` (swap `list_events` signature; apply dates only, interim)
- Modify: `src/server.rs` (expand `ListEventsParams`; map to `EventQuery` in the tool)
- Modify: `tests/tools.rs` (update `list_events_passes_date_range` to the new params/record shape; add a forwarding test for the new filters)

**Interfaces:**
- Consumes: `EventSummary` (Task 1).
- Produces: `pub struct EventQuery { pub start_date: Option<String>, pub end_date: Option<String>, pub query: Option<String>, pub category: Option<String>, pub show_as: Option<String>, pub my_response: Option<String>, pub attendees: Option<Vec<String>>, pub attendee_role: Option<String>, pub meetings_only: bool, pub all_day: Option<bool>, pub calendar_of: Option<String> }`
- Produces: trait method `fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError>;`

- [ ] **Step 1: Add the `EventQuery` struct to `src/outlook/mod.rs`**

Add directly below the `EmailUpdate` struct (before `pub trait OutlookClient`):

```rust
/// All filters for `list_events`. Every field is optional; supplying several
/// ANDs them. `start_date`/`end_date` bound the (recurrence-expanded) scan;
/// the rest filter the streamed events client-side. `calendar_of` (an
/// email/name) opens another person's shared calendar instead of your own.
#[derive(Debug, Clone, Default)]
pub struct EventQuery {
    pub start_date: Option<String>,
    pub end_date: Option<String>,
    pub query: Option<String>,                 // text match on subject + location
    pub category: Option<String>,
    pub show_as: Option<String>,               // "free"|"tentative"|"busy"|"out_of_office"|"working_elsewhere"
    pub my_response: Option<String>,           // "organizer"|"accepted"|"declined"|"tentative"|"not_responded"
    pub attendees: Option<Vec<String>>,        // match events where ANY listed person participates
    pub attendee_role: Option<String>,         // "required"|"optional"|"any" (default "any")
    pub meetings_only: bool,
    pub all_day: Option<bool>,
    pub calendar_of: Option<String>,
}
```

- [ ] **Step 2: Swap the trait method in `src/outlook/mod.rs`**

Replace:

```rust
    fn list_events(&self, start_date: Option<String>, end_date: Option<String>)
        -> Result<Vec<EventSummary>, ToolError>;
```

with:

```rust
    fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError>;
```

- [ ] **Step 3: Run the build to confirm it now fails (both implementors + tool)**

Run: `cargo build 2>&1 | Select-String "list_events|EventQuery|not.*implemented" | Select-Object -First 5`
Expected: FAIL — the fake and real client no longer satisfy the trait and `server.rs` still calls the old signature. (Red.)

- [ ] **Step 4: Swap `list_events` in `src/outlook/fake.rs`**

Replace the whole `list_events` fn with a version that takes `EventQuery` and records every field:

```rust
    fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError> {
        self.record("list_events", json!({
            "start_date": q.start_date, "end_date": q.end_date, "query": q.query,
            "category": q.category, "show_as": q.show_as, "my_response": q.my_response,
            "attendees": q.attendees, "attendee_role": q.attendee_role,
            "meetings_only": q.meetings_only, "all_day": q.all_day,
            "calendar_of": q.calendar_of,
        }))?;
        Ok(vec![EventSummary {
            id: EVENT_ID.into(), subject: "Standup".into(), start: None, end: None,
            location: "".into(), organizer: "".into(), all_day: false,
            is_recurring: false, is_meeting: false, categories: vec![],
            show_as: "busy".into(), my_response: "accepted".into(),
            required_attendees: "".into(), optional_attendees: "".into(),
        }])
    }
```

Add `EventQuery` to the fake's imports: find the `use crate::outlook::{...}` line at the top of `fake.rs` (it already imports `EmailQuery`, `EmailUpdate`, `OutlookClient`) and add `EventQuery` to it.

- [ ] **Step 5: Swap `list_events` in `src/outlook/client.rs` (dates only — interim)**

Replace the `list_events` signature and its `start`/`end` derivation so it destructures `EventQuery`. Replace the fn header + date block (currently lines ~767–791) so it reads:

```rust
    fn list_events(&self, q: EventQuery) -> Result<Vec<EventSummary>, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let start = match &q.start_date {
                Some(s) => parse_dt(s, "start_date")?,
                None => chrono::Local::now()
                    .date_naive()
                    .and_hms_opt(0, 0, 0)
                    .unwrap(),
            };
            let mut end = match &q.end_date {
                Some(s) => parse_dt(s, "end_date")?,
                None => start + chrono::Duration::days(7),
            };
            // If only a bare date was given for the end, treat it as the whole
            // end day (Python: `end.time() == time.min and "T" not in end_date`).
            if let Some(ed) = &q.end_date {
                if end.time() == chrono::NaiveTime::MIN && !ed.contains('T') {
                    end = end.date().and_hms_micro_opt(23, 59, 59, 999_999).unwrap();
                }
            }
```

Everything from `let calendar = ...` through the `GetFirst`/`GetNext` loop and `Ok(results)` stays exactly as-is for now. (The new filter fields on `q` are accepted but not yet applied — Task 3 adds the filtering; Task 4 adds `calendar_of`.)

Add `EventQuery` to the client's imports: extend the existing `use crate::outlook::{EmailQuery, EmailUpdate, OutlookClient};` line to `use crate::outlook::{EmailQuery, EmailUpdate, EventQuery, OutlookClient};`.

- [ ] **Step 6: Expand `ListEventsParams` and map it in `src/server.rs`**

Replace the whole `ListEventsParams` struct (currently lines ~154–160) with:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct ListEventsParams {
    #[serde(default)]
    pub start_date: Option<String>,
    #[serde(default)]
    pub end_date: Option<String>,
    /// Text match on subject + location.
    #[serde(default)]
    pub query: Option<String>,
    /// Filter to a color category.
    #[serde(default)]
    pub category: Option<String>,
    /// "free" | "tentative" | "busy" | "out_of_office" | "working_elsewhere".
    #[serde(default)]
    pub show_as: Option<String>,
    /// This mailbox's response: "organizer" | "accepted" | "declined" | "tentative" | "not_responded".
    #[serde(default)]
    pub my_response: Option<String>,
    /// Names/emails; match events where ANY listed person participates.
    #[serde(default)]
    pub attendees: Option<Vec<String>>,
    /// "required" | "optional" | "any" (default "any").
    #[serde(default)]
    pub attendee_role: Option<String>,
    /// Only events that have other attendees (meetings).
    #[serde(default)]
    pub meetings_only: bool,
    /// Only all-day (true) or only non-all-day (false) events.
    #[serde(default)]
    pub all_day: Option<bool>,
    /// Email/name of another person whose shared calendar to view (default: your own).
    #[serde(default)]
    pub calendar_of: Option<String>,
}
```

Then replace the `list_events` tool method (currently lines ~339–347) with:

```rust
    #[tool(description = "List/search calendar events. Filter by date range, text (subject/location), category, show_as, your response, attendees (+role), meetings-only, all-day; or view another person's shared calendar via calendar_of.")]
    pub async fn list_events(
        &self,
        Parameters(p): Parameters<ListEventsParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let q = EventQuery {
            start_date: p.start_date, end_date: p.end_date, query: p.query,
            category: p.category, show_as: p.show_as, my_response: p.my_response,
            attendees: p.attendees, attendee_role: p.attendee_role,
            meetings_only: p.meetings_only, all_day: p.all_day,
            calendar_of: p.calendar_of,
        };
        let result = run_blocking(move || client.list_events(q)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Add `EventQuery` to `server.rs`'s outlook imports — the same `use crate::outlook::{...}` line that brings in `EmailQuery`/`EmailUpdate`.

- [ ] **Step 7: Update and add the tool tests in `tests/tools.rs`**

Replace the whole `list_events_passes_date_range` test (currently lines ~264–278) with a date test on the new params shape plus a filter-forwarding test:

```rust
#[tokio::test]
async fn list_events_passes_date_range() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_events(Parameters(ListEventsParams {
            start_date: Some("2026-06-10".to_string()),
            end_date: Some("2026-06-17".to_string()),
            query: None, category: None, show_as: None, my_response: None,
            attendees: None, attendee_role: None, meetings_only: false,
            all_day: None, calendar_of: None,
        }))
        .await
        .unwrap();
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "list_events");
    assert_eq!(args["start_date"], "2026-06-10");
    assert_eq!(args["end_date"], "2026-06-17");
}

#[tokio::test]
async fn list_events_forwards_all_filters() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .list_events(Parameters(ListEventsParams {
            start_date: None, end_date: None,
            query: Some("review".to_string()),
            category: Some("Work".to_string()),
            show_as: Some("busy".to_string()),
            my_response: Some("accepted".to_string()),
            attendees: Some(vec!["alice@example.com".to_string()]),
            attendee_role: Some("required".to_string()),
            meetings_only: true,
            all_day: Some(false),
            calendar_of: Some("bob@example.com".to_string()),
        }))
        .await
        .unwrap();
    let (name, args) = fake.calls().pop().unwrap();
    assert_eq!(name, "list_events");
    assert_eq!(args["query"], "review");
    assert_eq!(args["category"], "Work");
    assert_eq!(args["show_as"], "busy");
    assert_eq!(args["my_response"], "accepted");
    assert_eq!(args["attendees"], serde_json::json!(["alice@example.com"]));
    assert_eq!(args["attendee_role"], "required");
    assert_eq!(args["meetings_only"], true);
    assert_eq!(args["all_day"], false);
    assert_eq!(args["calendar_of"], "bob@example.com");
}
```

- [ ] **Step 8: Build and test (green)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.
Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass, including the two `list_events_*` tests.

- [ ] **Step 9: Commit**

```bash
git add src/outlook/mod.rs src/outlook/fake.rs src/outlook/client.rs src/server.rs tests/tools.rs
git commit -m "Replace list_events(start,end) with EventQuery filter struct (interface + fake + tool)"
```

---

### Task 3: Client-side filters in `list_events`

Apply the new filters against each built `EventSummary` while streaming, mirroring how `list_emails` filters category/attachments client-side. Because Task 1 put `show_as`, `my_response`, `required_attendees`, `optional_attendees`, `categories`, `is_meeting`, `all_day`, `subject`, and `location` on the summary, every filter is a plain comparison — no extra COM reads.

**Files:**
- Modify: `src/outlook/client.rs` (add filtering inside the `list_events` stream loop)

**Interfaces:**
- Consumes: `EventQuery` (Task 2), `EventSummary` (Task 1).
- Produces: `list_events` honoring `query`/`category`/`show_as`/`my_response`/`attendees`+`attendee_role`/`meetings_only`/`all_day`.

- [ ] **Step 1: Add a private filter helper above the `impl OutlookClient for WindowsOutlookClient` block in `src/outlook/client.rs`**

Place this free fn next to the other module-level helpers (e.g. just after `event_summary`):

```rust
/// True if `summary` passes every filter set on `q`. All comparisons are
/// case-insensitive. Attendee matching is a substring test against the
/// semicolon-separated `RequiredAttendees`/`OptionalAttendees` strings.
fn event_matches(summary: &EventSummary, q: &EventQuery) -> bool {
    if let Some(query) = q.query.as_deref().filter(|s| !s.is_empty()) {
        let needle = query.to_lowercase();
        if !summary.subject.to_lowercase().contains(&needle)
            && !summary.location.to_lowercase().contains(&needle)
        {
            return false;
        }
    }
    if let Some(cat) = q.category.as_deref().filter(|s| !s.is_empty()) {
        let want = cat.to_lowercase();
        if !summary.categories.iter().any(|c| c.to_lowercase() == want) {
            return false;
        }
    }
    if let Some(show_as) = q.show_as.as_deref().filter(|s| !s.is_empty()) {
        if !summary.show_as.eq_ignore_ascii_case(show_as) {
            return false;
        }
    }
    if let Some(resp) = q.my_response.as_deref().filter(|s| !s.is_empty()) {
        if !summary.my_response.eq_ignore_ascii_case(resp) {
            return false;
        }
    }
    if q.meetings_only && !summary.is_meeting {
        return false;
    }
    if let Some(want_all_day) = q.all_day {
        if summary.all_day != want_all_day {
            return false;
        }
    }
    if let Some(people) = q.attendees.as_ref().filter(|v| !v.is_empty()) {
        // Which attendee tier(s) to search, per attendee_role (default "any").
        let role = q.attendee_role.as_deref().unwrap_or("any").to_lowercase();
        let required = summary.required_attendees.to_lowercase();
        let optional = summary.optional_attendees.to_lowercase();
        let haystack = match role.as_str() {
            "required" => required,
            "optional" => optional,
            _ => format!("{required}; {optional}"), // "any"
        };
        if !people
            .iter()
            .any(|p| !p.is_empty() && haystack.contains(&p.to_lowercase()))
        {
            return false;
        }
    }
    true
}
```

- [ ] **Step 2: Apply the filter inside the stream loop in `list_events`**

In `list_events`, the loop currently pushes every summary:

```rust
            let mut current = call_method(&restricted, "GetFirst", &mut [])?;
            while let Ok(item) = IDispatch::try_from(&current) {
                results.push(event_summary(&item)?);
                if results.len() >= MAX_CALENDAR_ITEMS {
                    break;
                }
                current = call_method(&restricted, "GetNext", &mut [])?;
            }
```

Replace that loop with one that keeps a summary only if it passes `event_matches`:

```rust
            let mut current = call_method(&restricted, "GetFirst", &mut [])?;
            while let Ok(item) = IDispatch::try_from(&current) {
                let summary = event_summary(&item)?;
                if event_matches(&summary, &q) {
                    results.push(summary);
                    if results.len() >= MAX_CALENDAR_ITEMS {
                        break;
                    }
                }
                current = call_method(&restricted, "GetNext", &mut [])?;
            }
```

(The cap is now checked only for kept items, matching `list_emails`'s "count after filter" behavior.)

- [ ] **Step 3: Build (green, zero warnings)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.

- [ ] **Step 4: Run the full suite (unit + tool tests still green)**

Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass. (Filtering is exercised live in Task 5; the fake path is unaffected.)

- [ ] **Step 5: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Filter list_events client-side (query, category, show_as, response, attendees, meetings_only, all_day)"
```

---

### Task 4: `calendar_of` — view another person's shared calendar

When `calendar_of` is supplied, resolve the recipient and open their shared calendar via `Namespace.GetSharedDefaultFolder`, instead of your own `GetDefaultFolder(9)`. All filters then apply identically. An unresolvable name or missing permission yields a clear error, never a crash.

**Files:**
- Modify: `src/outlook/client.rs` (branch the calendar-folder acquisition in `list_events`)

**Interfaces:**
- Consumes: `EventQuery.calendar_of` (Task 2), the `mapi()` namespace, `resolve`/`GetSharedDefaultFolder` COM calls.
- Produces: `list_events` honoring `calendar_of`.

- [ ] **Step 1: Branch the calendar-folder acquisition in `list_events`**

In `list_events`, the folder is currently fetched unconditionally:

```rust
            let calendar = to_disp(call_method(
                &ns,
                "GetDefaultFolder",
                &mut [variant_from_i32(c::OL_FOLDER_CALENDAR)],
            )?)?;
```

Replace that with a branch on `q.calendar_of`:

```rust
            // `calendar_of`: open another person's shared calendar; otherwise
            // our own default calendar (current behavior).
            let calendar = match q.calendar_of.as_deref().filter(|s| !s.is_empty()) {
                Some(person) => {
                    let recipient = to_disp(call_method(
                        &ns, "CreateRecipient", &mut [variant_from_str(person)],
                    )?)?;
                    let resolved = variant_to_bool(&call_method(&recipient, "Resolve", &mut [])?)
                        .unwrap_or(false);
                    if !resolved {
                        return Err(ToolError::new(format!(
                            "Could not resolve {person:?} to a person — check the name/email."
                        )));
                    }
                    // olFolderCalendar = 9. Requires that person to have shared
                    // their calendar with you; otherwise COM errors with a
                    // permission message, surfaced as-is.
                    to_disp(call_method(
                        &ns,
                        "GetSharedDefaultFolder",
                        &mut [
                            VARIANT::from(recipient),
                            variant_from_i32(c::OL_FOLDER_CALENDAR),
                        ],
                    ).map_err(|e| ToolError::new(format!(
                        "Could not open {person:?}'s calendar — they may not have shared it with you. {}",
                        format_com_error(&e)
                    )))?)?
                }
                None => to_disp(call_method(
                    &ns,
                    "GetDefaultFolder",
                    &mut [variant_from_i32(c::OL_FOLDER_CALENDAR)],
                )?)?,
            };
```

`VARIANT`, `variant_from_str`, `variant_to_bool`, `variant_from_i32`, and `format_com_error` are already imported in `client.rs` (used elsewhere in this file). `VARIANT::from(recipient)` boxes the `IDispatch` recipient exactly as the `update_email` `Move` call boxes its target folder.

- [ ] **Step 2: Build (green, zero warnings)**

Run: `cargo build 2>&1 | Select-Object -Last 3`
Expected: `Finished`, zero warnings.

- [ ] **Step 3: Run the full suite**

Run: `cargo test 2>&1 | Select-String "test result"`
Expected: all pass. (`calendar_of` is verified live in Task 5; the fake ignores it.)

- [ ] **Step 4: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Support calendar_of: view another person's shared calendar in list_events"
```

---

### Task 5: Live tests

`#[ignore]`d end-to-end tests against real Outlook: create a categorized meeting, list it back with filters (matching and non-matching), and open your own calendar via `calendar_of` (self-share works without setup). Each test cleans up after itself. The cross-user sharing path is documented as a manual check.

**Files:**
- Modify: `tests/live_outlook.rs`

**Interfaces:**
- Consumes: `EventQuery` (Task 2), `WindowsOutlookClient::list_events`/`create_event`/`get_event`.

- [ ] **Step 1: Add the live tests**

Add after the existing `create_event_then_delete_it` test. The file already imports `WindowsOutlookClient` and the outlook module types; add `EventQuery` to the existing `use outlook_mcp_rs::outlook::{...}` line.

```rust
#[test]
#[ignore]
fn list_events_filters_by_query_and_category() {
    let c = WindowsOutlookClient::new();
    // A far-future, uniquely-named appointment we can pinpoint and clean up.
    let created = c.create_event(
        "outlook-mcp-rs P6 filter probe".to_string(),
        "2099-01-05T09:00".to_string(),
        "2099-01-05T09:30".to_string(),
        None, None, None, false, None,
    ).expect("create_event");
    let id = created["id"].as_str().expect("event id").to_string();

    // A matching query in the window finds it.
    let hits = c.list_events(EventQuery {
        start_date: Some("2099-01-05".to_string()),
        end_date: Some("2099-01-05".to_string()),
        query: Some("filter probe".to_string()),
        ..Default::default()
    }).expect("list_events query");
    assert!(hits.iter().any(|e| e.id == id), "query should match the probe");
    // Enriched fields are populated.
    let probe = hits.iter().find(|e| e.id == id).unwrap();
    assert_eq!(probe.show_as, "busy");

    // A non-matching query in the same window excludes it.
    let misses = c.list_events(EventQuery {
        start_date: Some("2099-01-05".to_string()),
        end_date: Some("2099-01-05".to_string()),
        query: Some("no-such-subject-xyz".to_string()),
        ..Default::default()
    }).expect("list_events non-matching query");
    assert!(!misses.iter().any(|e| e.id == id), "non-matching query must exclude the probe");

    // Cleanup: delete the probe.
    c.delete_email(id).expect("cleanup delete");
}

#[test]
#[ignore]
fn list_events_calendar_of_self_opens_own_calendar() {
    // Opening your OWN calendar via calendar_of exercises the recipient-resolve
    // + GetSharedDefaultFolder path without needing a second user's sharing grant.
    // Set OUTLOOK_MCP_TEST_EMAIL to your SMTP address to run this.
    let me = match std::env::var("OUTLOOK_MCP_TEST_EMAIL") {
        Ok(v) if !v.is_empty() => v,
        _ => {
            eprintln!("skipping: set OUTLOOK_MCP_TEST_EMAIL to your address");
            return;
        }
    };
    let c = WindowsOutlookClient::new();
    // Should resolve and return without error (contents may be empty — that's fine).
    let _events = c.list_events(EventQuery {
        calendar_of: Some(me),
        ..Default::default()
    }).expect("list_events calendar_of self should resolve and not error");
}
```

- [ ] **Step 2: Confirm compile + ignored**

Run: `cargo build --tests 2>&1 | Select-Object -Last 2` → `Finished`.
Run: `cargo test --test live_outlook 2>&1 | Select-String "list_events_filters_by_query_and_category|calendar_of_self"` → both show `ignored`.

- [ ] **Step 3: (If Outlook available) run live**

Run: `cargo test --test live_outlook -- --ignored list_events_filters_by_query_and_category 2>&1 | Select-Object -Last 8`
Expected: `test result: ok. 1 passed`. If the probe isn't found, the client-side filter or the date-window Restrict is wrong — investigate against real Outlook.
Run (optional, needs your address): `$env:OUTLOOK_MCP_TEST_EMAIL="you@company.com"; cargo test --test live_outlook -- --ignored list_events_calendar_of_self_opens_own_calendar 2>&1 | Select-Object -Last 8`
Expected: `ok. 1 passed`. Skip these steps only if no Outlook is available.

- [ ] **Step 4: Document the cross-user manual check in `TESTING.md`**

Add a bullet under the manual-only section noting that `calendar_of` against **another** user's calendar requires that user to have granted you calendar-sharing permission, which can't be set up from an automated test; verify manually by calling `list_events` with a colleague's address who has shared their calendar.

- [ ] **Step 5: Commit**

```bash
git add tests/live_outlook.rs TESTING.md
git commit -m "Add live list_events filter + calendar_of tests; document cross-user manual check"
```

---

## Self-Review

**1. Spec coverage** (spec §`list_events (+ filters + calendar_of)` and §`get_event`):
- `query` (subject + location) ✅ Task 3 `event_matches`
- `category` ✅ Task 3
- `show_as` (friendly) ✅ output Task 1, filter Task 3
- `my_response` (friendly) ✅ output Task 1, filter Task 3
- `attendees` (ANY listed person) ✅ Task 3 substring match
- `attendee_role` required/optional/any (default any) ✅ Task 3 tier selection
- `meetings_only` ✅ Task 3 via `is_meeting`
- `all_day` ✅ Task 3
- `calendar_of` (CreateRecipient + Resolve + GetSharedDefaultFolder=9; clear error on unresolvable/no-permission) ✅ Task 4
- Output gains `categories`, `show_as`, `my_response`, `required_attendees`, `optional_attendees` ✅ Task 1 (`categories` already present from Plan 1)
- `get_event` gains `categories`, `show_as`, `my_response` (raw ResponseStatus number replaced by friendly word) ✅ Task 1 (EventDetail flattens the enriched summary; old `response` field renamed `my_response`)

**2. Placeholder scan:** No `todo!()`/TBD in this plan. The Task 2 "dates only, filters plumbed but not applied" state is a deliberate, documented interim fully resolved in Task 3 — not a placeholder left in shipped code (the plan isn't pushed until all 5 tasks are green).

**3. Type consistency:** `EventQuery` field names/types are identical across `mod.rs` (struct), `fake.rs` (record), `client.rs` (destructure), `server.rs` (`ListEventsParams`→`EventQuery` 1:1), and `tests/tools.rs`. `EventSummary`'s four new fields (`show_as`, `my_response`, `required_attendees`, `optional_attendees`) are constructed identically in `event_summary()` (real) and both fake literals (Task 1). `EventDetail { summary, body }` shape matches in `types.rs`, `client.rs::get_event`, and `fake.rs::get_event`. `friendly::busy_status_word`/`friendly::response_word` signatures match `friendly.rs`. `c::OL_BUSY`/`c::OL_RESPONSE_NONE`/`c::OL_FOLDER_CALENDAR`/`c::OL_NONMEETING` all exist in `constants.rs`. The Task 5 live tests use `EVENT`-shaped ids and the existing `create_event`/`delete_email` signatures.

## Execution Handoff

Plan 6 of 12. Models: T1 sonnet (types + COM reads), T2 sonnet (interface swap), T3 sonnet (filter logic), T4 opus (shared-calendar COM), T5 haiku (live tests). After all five are green with zero warnings, controller pushes `main` → Plan 7 (create_event enhancements). Trait-ripple checklist for the Task 2 signature change: mod.rs + client.rs + fake.rs + server.rs + tests/tools.rs (live_outlook.rs has no `list_events` call site — but Task 1's `EventSummary`/`EventDetail` change is covered because those tests only call `create_event`/`get_event`, which still compile).
