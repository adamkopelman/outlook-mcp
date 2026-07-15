# v2 Plan 9 — Recurrence Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add repeating-event support: a `recurrence` object on `create_event` (daily/weekly/monthly/yearly, with an interval and an end condition) and on `update_event` (set/replace or clear the pattern of an existing event), plus a read-back `recurrence` block on `get_event` so callers can see the pattern of an existing recurring event. Matches the spec's `create_event`/`update_event` rows and the plans index's Plan 9 line: "`recurrence` object on `create_event`/`update_event` via `GetRecurrencePattern()`".

**Architecture:** Outlook exposes recurrence through `AppointmentItem.GetRecurrencePattern()`, which returns a `RecurrencePattern` COM object whose properties (`RecurrenceType`, `Interval`, `DayOfWeekMask`, `DayOfMonth`, `MonthOfYear`, `PatternEndDate`/`Occurrences`/`NoEndDate`) are read/written with the same `get_property`/`put_property`/`call_method` COM primitives every other tool already uses — there is no new COM technique, just a new object to navigate to. The work splits cleanly along the codebase's existing seams:
1. **Pure logic first** (Task 1): enum↔word mappings (`"daily"`/`"weekly"`/`"monthly"`/`"yearly"`) and a weekday-name↔bitmask converter, both fully unit-testable without touching COM — mirrors `friendly.rs`'s existing pattern (`busy_status_word`/`busy_status_to_id`) and gets the same kind of table-driven test coverage `event_matches` and `friendly.rs` already have.
2. **Write path on create** (Task 2): a `RecurrenceInput` field threaded through `CreateEventInput` → fake → tool layer → real COM, following the exact ripple `categories`/`show_as` followed in Plan 7.
3. **Read path** (Task 3): a `recurrence_info()` helper populating a new `EventDetail.recurrence` field, mirroring how `event_summary`/`get_event` already enrich output.
4. **Write path on update** (Task 4): the same `RecurrenceInput` plus a `clear_recurrence` flag (→ `AppointmentItem.ClearRecurrencePattern()`) threaded through `EventUpdate`, following the `update_event` pattern Plan 8 established.
5. **Live verification** (Task 5): round-trip tests against real Outlook for weekly/monthly/yearly patterns, an update that changes the pattern, and one that clears it — all using far-future dates and `send:false`/no-attendee appointments so nothing is ever delivered, matching every prior plan's live-test safety discipline.

Recurring-event edits apply to the **whole series only** — there is no per-occurrence targeting in this API (this was already noted as out of scope back in Plan 8's Global Constraints and in the spec's "Out of scope" section; this plan doesn't change that).

**Tech Stack:** Rust, `windows` 0.62.2 COM, `rmcp` 2.1.0 tool macros, `serde_json`, `chrono`.

## Global Constraints

- **Target crate:** `C:\Users\adamk\projects\outlook-mcp-rs` (the Rust impl, NOT the Python `outlook-mcp`). Edition 2024, rustc 1.95.0.
- **Two implementors per trait-visible struct change.** `OutlookClient` lives in `src/outlook/mod.rs`; adding a field to `CreateEventInput` or `EventUpdate` (both already-existing structs, not new trait methods) means every struct literal of that type across the crate must gain the new field(s) or the build breaks — that's `src/outlook/fake.rs`, `src/server.rs`, and **`tests/live_outlook.rs`**, which has 5 `CreateEventInput { .. }` literals (lines ~84, ~104, ~132, ~182, ~206) and 1 `EventUpdate { .. }` literal (line ~146). Neither struct derives `Default` (intentional — see their doc comments in `mod.rs`), so every one of those 6 call sites needs an explicit `recurrence: None` (plus `clear_recurrence: false` for the `EventUpdate` one) added by Task 2/Task 4. This is the "recurring gotcha" flagged in `V2-RESUME.md`.
- **Reuse existing helpers.** `com::get_property`/`put_property`/`call_method`/`to_disp` (COM primitives), `client.rs::parse_dt` (ISO date/datetime parsing, used for `recurrence.until`), `variant_from_i32`/`variant_from_bool`/`variant_from_datetime`/`variant_to_i32`/`variant_to_bool`/`variant_to_iso_string`. Do not duplicate any of them.
- **No per-occurrence editing.** `update_event`'s existing field edits (subject/time/location/etc., already shipped in Plan 8) already apply to the whole series with no special handling; this plan only adds the ability to set/replace/clear the *recurrence pattern itself*, not to target one occurrence.
- **Zero warnings** on `cargo build` / `cargo test` before the plan is pushed.
- **Model policy:** Task 1 = **sonnet** (pure enum/bitmask logic, mirrors `friendly.rs`). Task 2 = **opus** (new COM object navigation — `GetRecurrencePattern()` — plus multi-branch validation; the project's stated policy is opus for "complex COM (recurrence...)"). Task 3 = **opus** (symmetric read-back logic, same COM object). Task 4 = **opus** (same write logic reused inside `update_event`'s existing attendee/save-vs-send flow, plus the clear branch). Task 5 = **opus** (live COM verification, real recurring appointments).

---

### Task 1: Recurrence enum/bitmask mappings + domain types (pure, no COM)

Add the vocabulary this whole plan is built on: Outlook's `OlRecurrenceType`/`OlDaysOfWeek` constants, friendly-word conversions for both, and the two domain types (`RecurrenceInput` for writes, `RecurrenceInfo` for reads). Nothing here touches the trait, COM, or any existing call site — it's purely additive and independently testable.

**Files:**
- Modify: `src/constants.rs` (add `OlRecurrenceType`/`OlDaysOfWeek` constants after the `OlResponseStatus` block, line ~70)
- Modify: `src/friendly.rs` (add recurrence word/id conversions after `task_status_to_id`, line ~65; needs a new `use crate::error::ToolError;` import)
- Modify: `src/outlook/types.rs` (add `RecurrenceInfo` struct after `EventDetail`, line ~78)
- Modify: `src/outlook/mod.rs` (add `RecurrenceInput` struct after `EventUpdate`, line ~110; add `validate_recurrence` pure function after `create_event_status`, line ~130)

**Interfaces:**
- Produces: `pub fn recurrence_pattern_to_id(name: &str) -> Option<i32>` — `"daily"→0`, `"weekly"→1`, `"monthly"→2`, `"yearly"→5`; anything else `None`.
- Produces: `pub fn recurrence_pattern_word(v: i32) -> &'static str` — inverse (0→"daily", 1→"weekly", 2 or 3→"monthly", 5 or 6→"yearly"; unknown → "daily").
- Produces: `pub fn day_of_week_words_to_mask(days: &[String]) -> Result<i32, ToolError>` — ORs each name's `OlDaysOfWeek` bit; rejects an unknown name.
- Produces: `pub fn day_of_week_mask_to_words(mask: i32) -> Vec<String>` — inverse, in Sunday→Saturday order.
- Produces: `pub struct RecurrenceInput { pub pattern: String, pub interval: Option<i32>, pub days_of_week: Option<Vec<String>>, pub day_of_month: Option<i32>, pub until: Option<String>, pub occurrences: Option<i32> }` (`#[derive(Debug, Clone)]`).
- Produces: `pub struct RecurrenceInfo { pub pattern: String, pub interval: i32, pub days_of_week: Vec<String>, pub day_of_month: Option<i32>, pub until: Option<String>, pub occurrences: Option<i32>, pub no_end: bool }` (`#[derive(Debug, Clone, Serialize)]`).
- Produces: `pub fn validate_recurrence(r: &RecurrenceInput) -> Result<i32, ToolError>` — resolves `r.pattern` to its `OlRecurrenceType` id (via `recurrence_pattern_to_id`) or errors; errors if `pattern == "weekly"` and `days_of_week` is missing/empty; errors if `pattern == "monthly"` and `day_of_month` is `None`; errors if both `occurrences` and `until` are set. Returns the resolved id on success — Task 2/4's COM code calls this first and reuses the id.

- [ ] **Step 1: Write the failing unit tests for the constants/conversions in `src/friendly.rs`**

Add to the existing `#[cfg(test)] mod tests` block at the bottom of `src/friendly.rs` (after `reverse_lookups_are_case_insensitive_and_reject_garbage`):

```rust
    #[test]
    fn recurrence_pattern_round_trips() {
        assert_eq!(recurrence_pattern_to_id("daily"), Some(0));
        assert_eq!(recurrence_pattern_to_id("Weekly"), Some(1));
        assert_eq!(recurrence_pattern_to_id("MONTHLY"), Some(2));
        assert_eq!(recurrence_pattern_to_id("yearly"), Some(5));
        assert_eq!(recurrence_pattern_to_id("biweekly"), None);
        assert_eq!(recurrence_pattern_word(0), "daily");
        assert_eq!(recurrence_pattern_word(1), "weekly");
        assert_eq!(recurrence_pattern_word(2), "monthly");
        assert_eq!(recurrence_pattern_word(3), "monthly"); // olRecursMonthNth, treated as monthly
        assert_eq!(recurrence_pattern_word(5), "yearly");
        assert_eq!(recurrence_pattern_word(6), "yearly"); // olRecursYearNth, treated as yearly
        assert_eq!(recurrence_pattern_word(99), "daily"); // unknown -> default
    }

    #[test]
    fn day_of_week_mask_round_trips() {
        let days = vec!["monday".to_string(), "Wednesday".to_string(), "FRIDAY".to_string()];
        let mask = day_of_week_words_to_mask(&days).unwrap();
        assert_eq!(mask, 2 | 8 | 32); // olMonday | olWednesday | olFriday
        assert_eq!(
            day_of_week_mask_to_words(mask),
            vec!["monday".to_string(), "wednesday".to_string(), "friday".to_string()]
        );
        assert_eq!(day_of_week_mask_to_words(127), vec![
            "sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday",
        ]);
        assert_eq!(day_of_week_mask_to_words(0), Vec::<String>::new());
    }

    #[test]
    fn day_of_week_words_to_mask_rejects_unknown_names() {
        assert!(day_of_week_words_to_mask(&["funday".to_string()]).is_err());
    }
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --lib friendly:: 2>&1 | Select-Object -Last 30`
Expected: FAIL to compile — `recurrence_pattern_to_id`/`recurrence_pattern_word`/`day_of_week_words_to_mask`/`day_of_week_mask_to_words` don't exist yet.

- [ ] **Step 3: Add the `OlRecurrenceType`/`OlDaysOfWeek` constants to `src/constants.rs`**

Add directly after the `OlResponseStatus` block (after `pub const OL_RESPONSE_NOT_RESPONDED: i32 = 5;`, before `pub fn folder_name_to_id`):

```rust
// OlRecurrenceType (RecurrencePattern.RecurrenceType)
pub const OL_RECURS_DAILY: i32 = 0;
pub const OL_RECURS_WEEKLY: i32 = 1;
pub const OL_RECURS_MONTHLY: i32 = 2;
pub const OL_RECURS_MONTH_NTH: i32 = 3;
pub const OL_RECURS_YEARLY: i32 = 5;
pub const OL_RECURS_YEAR_NTH: i32 = 6;

// OlDaysOfWeek (RecurrencePattern.DayOfWeekMask, a bitmask — OR the bits you want)
pub const OL_SUNDAY: i32 = 1;
pub const OL_MONDAY: i32 = 2;
pub const OL_TUESDAY: i32 = 4;
pub const OL_WEDNESDAY: i32 = 8;
pub const OL_THURSDAY: i32 = 16;
pub const OL_FRIDAY: i32 = 32;
pub const OL_SATURDAY: i32 = 64;
```

- [ ] **Step 4: Add the conversions to `src/friendly.rs`**

Add the import at the top of the file (after `use crate::constants as c;`):

```rust
use crate::error::ToolError;
```

Add the functions directly after `task_status_to_id` (after its closing `}`, before the `item_type_from_class` doc comment):

```rust
pub fn recurrence_pattern_to_id(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "daily" => Some(c::OL_RECURS_DAILY),
        "weekly" => Some(c::OL_RECURS_WEEKLY),
        "monthly" => Some(c::OL_RECURS_MONTHLY),
        "yearly" => Some(c::OL_RECURS_YEARLY),
        _ => None,
    }
}

pub fn recurrence_pattern_word(v: i32) -> &'static str {
    match v {
        c::OL_RECURS_WEEKLY => "weekly",
        c::OL_RECURS_MONTHLY | c::OL_RECURS_MONTH_NTH => "monthly",
        c::OL_RECURS_YEARLY | c::OL_RECURS_YEAR_NTH => "yearly",
        _ => "daily",
    }
}

const WEEKDAYS: [(i32, &str); 7] = [
    (c::OL_SUNDAY, "sunday"),
    (c::OL_MONDAY, "monday"),
    (c::OL_TUESDAY, "tuesday"),
    (c::OL_WEDNESDAY, "wednesday"),
    (c::OL_THURSDAY, "thursday"),
    (c::OL_FRIDAY, "friday"),
    (c::OL_SATURDAY, "saturday"),
];

pub fn day_of_week_words_to_mask(days: &[String]) -> Result<i32, ToolError> {
    let mut mask = 0;
    for day in days {
        let bit = WEEKDAYS
            .iter()
            .find(|(_, name)| name.eq_ignore_ascii_case(day))
            .map(|(bit, _)| *bit)
            .ok_or_else(|| {
                ToolError::new(format!(
                    "invalid day_of_week {day:?}: expected a full weekday name like \"monday\""
                ))
            })?;
        mask |= bit;
    }
    Ok(mask)
}

pub fn day_of_week_mask_to_words(mask: i32) -> Vec<String> {
    WEEKDAYS
        .iter()
        .filter(|(bit, _)| mask & bit != 0)
        .map(|(_, name)| name.to_string())
        .collect()
}
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --lib friendly:: 2>&1 | Select-Object -Last 30`
Expected: PASS — all `friendly::tests::*` green, including the 3 new tests.

- [ ] **Step 6: Write the failing unit test for `validate_recurrence` in `src/outlook/mod.rs`**

Add to the existing `#[cfg(test)] mod tests` block at the bottom of `src/outlook/mod.rs` (after `create_event_status_covers_all_three_outcomes`), and change its `use` line to also import `validate_recurrence` and `RecurrenceInput`:

```rust
#[cfg(test)]
mod tests {
    use super::{create_event_status, validate_recurrence, RecurrenceInput};

    #[test]
    fn create_event_status_covers_all_three_outcomes() {
        assert_eq!(create_event_status(true, true), "meeting_sent");
        assert_eq!(create_event_status(true, false), "meeting_saved");
        assert_eq!(create_event_status(false, true), "saved");
        assert_eq!(create_event_status(false, false), "saved");
    }

    fn recurrence(pattern: &str) -> RecurrenceInput {
        RecurrenceInput {
            pattern: pattern.to_string(), interval: None, days_of_week: None,
            day_of_month: None, until: None, occurrences: None,
        }
    }

    #[test]
    fn validate_recurrence_accepts_daily_with_no_extra_fields() {
        assert_eq!(validate_recurrence(&recurrence("daily")).unwrap(), 0);
    }

    #[test]
    fn validate_recurrence_rejects_unknown_pattern() {
        assert!(validate_recurrence(&recurrence("biweekly")).is_err());
    }

    #[test]
    fn validate_recurrence_requires_days_of_week_for_weekly() {
        assert!(validate_recurrence(&recurrence("weekly")).is_err());
        let mut r = recurrence("weekly");
        r.days_of_week = Some(vec!["monday".to_string()]);
        assert_eq!(validate_recurrence(&r).unwrap(), 1);
    }

    #[test]
    fn validate_recurrence_requires_day_of_month_for_monthly() {
        assert!(validate_recurrence(&recurrence("monthly")).is_err());
        let mut r = recurrence("monthly");
        r.day_of_month = Some(15);
        assert_eq!(validate_recurrence(&r).unwrap(), 2);
    }

    #[test]
    fn validate_recurrence_rejects_both_until_and_occurrences() {
        let mut r = recurrence("daily");
        r.until = Some("2099-01-01".to_string());
        r.occurrences = Some(5);
        assert!(validate_recurrence(&r).is_err());
    }
}
```

- [ ] **Step 7: Run the test to verify it fails**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --lib outlook::tests:: 2>&1 | Select-Object -Last 30`
Expected: FAIL to compile — `RecurrenceInput`/`validate_recurrence` don't exist yet.

- [ ] **Step 8: Add `RecurrenceInfo` to `src/outlook/types.rs`**

Add directly after the `EventDetail` struct (after its closing `}`, before `TaskSummary`):

```rust
/// The recurrence pattern of a recurring event, read back via
/// `AppointmentItem.GetRecurrencePattern()`. `None` on `EventDetail` when the
/// event isn't recurring. `until`/`occurrences` are mutually exclusive with
/// each other and with `no_end: true` (exactly one of the three end
/// conditions is populated).
#[derive(Debug, Clone, Serialize)]
pub struct RecurrenceInfo {
    /// "daily" | "weekly" | "monthly" | "yearly".
    pub pattern: String,
    pub interval: i32,
    /// Populated only for "weekly"; e.g. ["monday", "wednesday"].
    pub days_of_week: Vec<String>,
    /// Populated only for "monthly"/"yearly".
    pub day_of_month: Option<i32>,
    /// ISO end date, if the series ends on a date.
    pub until: Option<String>,
    /// Total occurrence count, if the series ends after N occurrences.
    pub occurrences: Option<i32>,
    /// True if the series never ends.
    pub no_end: bool,
}
```

Then add `pub recurrence: Option<RecurrenceInfo>,` as a new field on `EventDetail`, directly after its existing `pub body: String,` field:

```rust
#[derive(Debug, Clone, Serialize)]
pub struct EventDetail {
    #[serde(flatten)]
    pub summary: EventSummary,
    pub body: String,
    pub recurrence: Option<RecurrenceInfo>,
}
```

- [ ] **Step 9: Add `RecurrenceInput` and `validate_recurrence` to `src/outlook/mod.rs`**

Add directly after the `EventUpdate` struct (after its closing `}`, before `pub trait OutlookClient`):

```rust
/// One recurrence pattern for `create_event`/`update_event`. `pattern`
/// selects which of the other fields matter: `"weekly"` requires
/// `days_of_week`; `"monthly"` requires `day_of_month`; `"yearly"` derives
/// its month/day from the event's own start date (no field needed);
/// `"daily"` needs nothing extra. At most one of `until`/`occurrences` may
/// be set; if neither is set the series has no end date.
#[derive(Debug, Clone)]
pub struct RecurrenceInput {
    pub pattern: String,
    pub interval: Option<i32>,
    pub days_of_week: Option<Vec<String>>,
    pub day_of_month: Option<i32>,
    pub until: Option<String>,
    pub occurrences: Option<i32>,
}
```

Also add the new `recurrence` field to both `CreateEventInput` and `EventUpdate` now, so `mod.rs` is internally consistent within this one task (every other struct field it declares already exists). The build stays red until Task 2/4 fix the other call sites that construct these structs — that's expected and confirmed in Step 10 below.

Add `pub recurrence: Option<RecurrenceInput>,` as the last field of `CreateEventInput` (after `pub send: bool,`):

```rust
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
    pub recurrence: Option<RecurrenceInput>,
}
```

Add `pub recurrence: Option<RecurrenceInput>,` and `pub clear_recurrence: bool,` as the last two fields of `EventUpdate` (after `pub send_update: bool,`):

```rust
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
    pub recurrence: Option<RecurrenceInput>,
    pub clear_recurrence: bool,
}
```

Add `validate_recurrence` directly after `create_event_status` (after its closing `}`, before `#[cfg(test)]`):

```rust
/// Resolves `r.pattern` to its `OlRecurrenceType` id and checks the fields
/// each pattern requires. Called first by both `create_event`'s and
/// `update_event`'s real-COM recurrence-writing code, before any COM call is
/// made, so a bad `recurrence` object fails fast with a clear message.
pub fn validate_recurrence(r: &RecurrenceInput) -> Result<i32, ToolError> {
    let recurrence_type = crate::friendly::recurrence_pattern_to_id(&r.pattern).ok_or_else(|| {
        ToolError::new(format!(
            "invalid recurrence.pattern {:?}: expected \"daily\", \"weekly\", \"monthly\", or \"yearly\"",
            r.pattern
        ))
    })?;
    if r.pattern.eq_ignore_ascii_case("weekly")
        && !r.days_of_week.as_ref().is_some_and(|d| !d.is_empty())
    {
        return Err(ToolError::new(
            "recurrence.days_of_week is required for a \"weekly\" pattern",
        ));
    }
    if r.pattern.eq_ignore_ascii_case("monthly") && r.day_of_month.is_none() {
        return Err(ToolError::new(
            "recurrence.day_of_month is required for a \"monthly\" pattern",
        ));
    }
    if r.occurrences.is_some() && r.until.is_some() {
        return Err(ToolError::new(
            "recurrence: specify at most one of \"until\" or \"occurrences\", not both",
        ));
    }
    Ok(recurrence_type)
}
```

- [ ] **Step 10: Fix compile errors from the two new struct fields**

This step doesn't make the build green yet (Task 2/4 do that) — it only confirms the *expected* new failures. Run:

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo build 2>&1 | Select-String "error\[" | Select-Object -First 20`
Expected: a handful of "missing field `recurrence`" errors in `src/outlook/fake.rs`, `src/server.rs`, and `tests/live_outlook.rs` — exactly the call sites Task 2 and Task 4 will fix. No errors anywhere else.

- [ ] **Step 11: Run the new `outlook::tests` unit tests to verify they pass**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --lib outlook::tests:: --no-fail-fast 2>&1 | Select-Object -Last 40`
Expected: the 5 new `validate_recurrence_*`/`create_event_status_*` tests compile and pass in isolation (the crate as a whole won't build yet because of Step 10's known-expected errors — that's fine, Task 2 fixes it next). If your `cargo test` invocation refuses to run with a broken crate, skip this step and fold verification into Task 2's Step 2 instead.

- [ ] **Step 12: Commit**

```bash
git add src/constants.rs src/friendly.rs src/outlook/types.rs src/outlook/mod.rs
git commit -m "Add recurrence enum/bitmask mappings, RecurrenceInput/RecurrenceInfo types, validate_recurrence"
```

---

### Task 2: `recurrence` on `create_event` — fake, tool layer, and real COM write

Wire `CreateEventInput.recurrence` all the way through: the fake client records it, the tool layer accepts a `RecurrenceParams` object and converts it, and the real client actually sets the pattern via `GetRecurrencePattern()`.

**Files:**
- Modify: `src/outlook/fake.rs` (add `recurrence` to the recorded args in `create_event`, line ~171)
- Modify: `src/server.rs` (add `RecurrenceParams` struct before `CreateEventParams`, line ~195; add `recurrence` field to `CreateEventParams`; add `recurrence` conversion in the `create_event` tool method, line ~467)
- Modify: `src/outlook/client.rs` (add `apply_recurrence` helper after `add_meeting_recipient`, line ~218; call it from `create_event`, line ~1002)
- Modify: `tests/live_outlook.rs` (add `recurrence: None` to all 5 `CreateEventInput { .. }` literals: lines ~84, ~104, ~132, ~182, ~206)
- Modify: `tests/tools.rs` (add a `create_event_forwards_recurrence` test)

**Interfaces:**
- Consumes: `RecurrenceInput` (Task 1, `src/outlook/mod.rs`), `validate_recurrence(&RecurrenceInput) -> Result<i32, ToolError>` (Task 1), `day_of_week_words_to_mask` (Task 1, `src/friendly.rs`), `to_disp`/`get_property`/`put_property`/`call_method`/`variant_from_i32`/`variant_from_bool`/`variant_from_datetime`/`variant_to_iso_string` (existing `com.rs` helpers), `parse_dt` (existing `client.rs` helper).
- Produces: `fn apply_recurrence(appt: &IDispatch, r: &RecurrenceInput) -> Result<(), ToolError>` in `client.rs` — sets the appointment's recurrence pattern; called by both `create_event` (this task) and `update_event` (Task 4).
- Produces (server-layer): `pub struct RecurrenceParams { pub pattern: String, pub interval: Option<i32>, pub days_of_week: Option<Vec<String>>, pub day_of_month: Option<i32>, pub until: Option<String>, pub occurrences: Option<i32> }` — the MCP-facing mirror of `RecurrenceInput`, reused as-is by Task 4 for `UpdateEventParams`.

- [ ] **Step 1: Write the failing fake-client tool test in `tests/tools.rs`**

Add directly after `create_event_status_reflects_attendees_and_send` (after its closing `}`):

```rust
#[tokio::test]
async fn create_event_forwards_recurrence() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .create_event(Parameters(CreateEventParams {
            subject: "Standup".to_string(),
            start: "2026-06-12T09:00".to_string(),
            end: "2026-06-12T09:15".to_string(),
            body: None, location: None, attendees: None,
            required_attendees: None, optional_attendees: None,
            all_day: false, reminder_minutes: None, categories: None, show_as: None,
            send: true,
            recurrence: Some(RecurrenceParams {
                pattern: "weekly".to_string(),
                interval: Some(1),
                days_of_week: Some(vec!["monday".to_string(), "wednesday".to_string()]),
                day_of_month: None,
                until: None,
                occurrences: Some(10),
            }),
        }))
        .await
        .unwrap();
    let (_, args) = &fake.calls()[0];
    assert_eq!(args["recurrence"]["pattern"], "weekly");
    assert_eq!(args["recurrence"]["days_of_week"], json!(["monday", "wednesday"]));
    assert_eq!(args["recurrence"]["occurrences"], 10);
}
```

Also add `RecurrenceParams` to the `use outlook_mcp_rs::server::{ ... };` import block at the top of `tests/tools.rs`.

- [ ] **Step 2: Run the test to verify it fails**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --test tools create_event_forwards_recurrence 2>&1 | Select-Object -Last 30`
Expected: FAIL to compile — `RecurrenceParams` doesn't exist, `CreateEventParams`/`CreateEventInput` don't have a `recurrence` field yet.

- [ ] **Step 3: Add `RecurrenceParams` and wire it into `CreateEventParams` in `src/server.rs`**

Add directly before `CreateEventParams` (before its `#[derive(...)]` line):

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct RecurrenceParams {
    /// "daily" | "weekly" | "monthly" | "yearly".
    pub pattern: String,
    /// Repeat every N days/weeks/months/years (default 1).
    #[serde(default)]
    pub interval: Option<i32>,
    /// Required for "weekly": full weekday names, e.g. ["monday", "wednesday"].
    #[serde(default)]
    pub days_of_week: Option<Vec<String>>,
    /// Required for "monthly": day of the month (1-31). Not used for "yearly"
    /// (the event's own start date supplies the month/day).
    #[serde(default)]
    pub day_of_month: Option<i32>,
    /// End date (ISO). At most one of `until`/`occurrences`; neither = no end date.
    #[serde(default)]
    pub until: Option<String>,
    /// Number of occurrences. At most one of `until`/`occurrences`.
    #[serde(default)]
    pub occurrences: Option<i32>,
}
```

Add `pub recurrence: Option<RecurrenceParams>,` as the last field of `CreateEventParams` (after `pub send: bool,`):

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
    /// Repeat this event daily/weekly/monthly/yearly. Omit for a one-off event.
    #[serde(default)]
    pub recurrence: Option<RecurrenceParams>,
}
```

Update the `create_event` tool method's destructuring and body (replace the whole method):

```rust
    #[tool(description = "Create a calendar event. required_attendees/optional_attendees invite two tiers (attendees is a legacy alias merged into required_attendees); any attendee makes it a meeting. categories and show_as (busy status) can be set on creation. recurrence repeats the event (daily/weekly/monthly/yearly, with an interval and an until date or occurrence count). send (default true) controls whether a meeting is actually sent to attendees or just saved for review.")]
    pub async fn create_event(
        &self,
        Parameters(CreateEventParams {
            subject, start, end, body, location, attendees, required_attendees,
            optional_attendees, all_day, reminder_minutes, categories, show_as, send,
            recurrence,
        }): Parameters<CreateEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        // `attendees` is a legacy alias for the required tier; merge it in.
        let mut required = required_attendees.unwrap_or_default();
        required.extend(attendees.unwrap_or_default());
        let required_attendees = (!required.is_empty()).then_some(required);
        let recurrence = recurrence.map(|r| RecurrenceInput {
            pattern: r.pattern, interval: r.interval, days_of_week: r.days_of_week,
            day_of_month: r.day_of_month, until: r.until, occurrences: r.occurrences,
        });
        let input = CreateEventInput {
            subject, start, end, body, location, required_attendees, optional_attendees,
            all_day, reminder_minutes, categories, show_as, send, recurrence,
        };
        let result = run_blocking(move || client.create_event(input)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Add `RecurrenceInput` to the `use super::{...}` import at the top of `src/server.rs` (it currently reads `use super::{CreateEventInput, EmailQuery, EmailUpdate, EventQuery, EventUpdate, OutlookClient};` — change to `use super::{CreateEventInput, EmailQuery, EmailUpdate, EventQuery, EventUpdate, OutlookClient, RecurrenceInput};`).

- [ ] **Step 4: Record `recurrence` in `src/outlook/fake.rs`'s `create_event`**

Replace the `self.record("create_event", json!({ ... }))?;` call in `create_event` with:

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
            "recurrence": input.recurrence.as_ref().map(|r| json!({
                "pattern": r.pattern, "interval": r.interval, "days_of_week": r.days_of_week,
                "day_of_month": r.day_of_month, "until": r.until, "occurrences": r.occurrences,
            })),
        }))?;
        let has_attendees = input.required_attendees.as_ref().is_some_and(|v| !v.is_empty())
            || input.optional_attendees.as_ref().is_some_and(|v| !v.is_empty());
        let status = super::create_event_status(has_attendees, input.send);
        Ok(json!({"status": status, "id": EVENT_ID, "subject": input.subject}))
    }
```

- [ ] **Step 5: Run the tool test to verify it still fails (real client + live tests not yet fixed)**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo build 2>&1 | Select-String "error\[" | Select-Object -First 20`
Expected: FAIL — `src/outlook/client.rs`'s `create_event` and every `CreateEventInput { .. }` literal in `tests/live_outlook.rs` are missing the `recurrence` field. `src/server.rs`/`src/outlook/fake.rs` now compile clean.

- [ ] **Step 6: Add `apply_recurrence` and wire it into `create_event` in `src/outlook/client.rs`**

Add the `use` for `validate_recurrence`/`RecurrenceInput`/`day_of_week_words_to_mask` at the top of `src/outlook/client.rs` — change the existing `use crate::outlook::{ create_event_status, CreateEventInput, EmailQuery, EmailUpdate, EventQuery, EventUpdate, OutlookClient, };` to:

```rust
use crate::outlook::{
    create_event_status, validate_recurrence, CreateEventInput, EmailQuery, EmailUpdate,
    EventQuery, EventUpdate, OutlookClient, RecurrenceInput,
};
```

Add `use chrono::Datelike;` near the top (after the `use crate::outlook::types::*;` line) — needed to pull `.month()`/`.day()` off the `NaiveDateTime` derived from `Start` for the yearly case.

Add `apply_recurrence` directly after `add_meeting_recipient` (after its closing `}`, before the `remove_meeting_recipients` doc comment):

```rust
/// Sets an appointment's recurrence pattern via `GetRecurrencePattern()`.
/// Calling this on a non-recurring appointment converts it into a recurring
/// one (this is how `update_event` adds recurrence to an existing single
/// event, too — see Task 4). `"yearly"` derives its month/day from the
/// appointment's own `Start` property rather than a separate input field, so
/// this must run after `Start` is already set to its final value.
fn apply_recurrence(appt: &IDispatch, r: &RecurrenceInput) -> Result<(), ToolError> {
    let recurrence_type = validate_recurrence(r)?;
    let pattern = to_disp(call_method(appt, "GetRecurrencePattern", &mut [])?)?;
    put_property(&pattern, "RecurrenceType", variant_from_i32(recurrence_type))?;
    put_property(&pattern, "Interval", variant_from_i32(r.interval.unwrap_or(1)))?;
    match r.pattern.to_lowercase().as_str() {
        "weekly" => {
            let mask = crate::friendly::day_of_week_words_to_mask(
                r.days_of_week.as_deref().unwrap_or(&[]),
            )?;
            put_property(&pattern, "DayOfWeekMask", variant_from_i32(mask))?;
        }
        "monthly" => {
            put_property(&pattern, "DayOfMonth", variant_from_i32(r.day_of_month.unwrap()))?;
        }
        "yearly" => {
            let start_iso = variant_to_iso_string(&get_property(appt, "Start")?).ok_or_else(|| {
                ToolError::new("could not read Start to derive the yearly recurrence date")
            })?;
            let start_dt =
                chrono::NaiveDateTime::parse_from_str(&start_iso, "%Y-%m-%dT%H:%M:%S").map_err(|_| {
                    ToolError::new("could not parse Start to derive the yearly recurrence date")
                })?;
            put_property(&pattern, "MonthOfYear", variant_from_i32(start_dt.month() as i32))?;
            put_property(&pattern, "DayOfMonth", variant_from_i32(start_dt.day() as i32))?;
        }
        _ => {}
    }
    match (r.occurrences, r.until.as_deref()) {
        (Some(n), _) => {
            put_property(&pattern, "Occurrences", variant_from_i32(n))?;
        }
        (None, Some(until)) => {
            let until_dt = parse_dt(until, "recurrence.until")?;
            put_property(&pattern, "PatternEndDate", variant_from_datetime(&until_dt)?)?;
        }
        (None, None) => {
            put_property(&pattern, "NoEndDate", variant_from_bool(true))?;
        }
    }
    Ok(())
}
```

No new `use` for `day_of_week_words_to_mask` is needed — the snippet above calls it fully-qualified as `crate::friendly::day_of_week_words_to_mask(...)`, matching the existing fully-qualified-inline style `client.rs` already uses for `crate::friendly::busy_status_to_id(...)`.

Now wire it into `create_event`: add the call directly after the `show_as` block and before the `required`/`optional` attendee block (i.e. right after the closing `}` of the `if let Some(show_as) = ...` block, before `let required = input.required_attendees.unwrap_or_default();`):

```rust
            if let Some(recurrence) = input.recurrence.as_ref() {
                apply_recurrence(&appt, recurrence)?;
            }
```

- [ ] **Step 7: Fix the 5 `CreateEventInput { .. }` literals in `tests/live_outlook.rs`**

Add `recurrence: None,` to each of the 5 literals (after `send: true,`/`send: false,` in each). For example, `create_event_then_delete_it` (line ~84) becomes:

```rust
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs live test event".to_string(),
        start: "2099-01-01T10:00:00".to_string(),
        end: "2099-01-01T10:30:00".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
        recurrence: None,
    }).expect("create_event should succeed");
```

Apply the same one-line addition (`recurrence: None,` right after the `send: ...,` line) to the other 4: `create_event_with_tiers_categories_and_show_as` (~104), `update_event_edits_fields_and_manages_attendees` (~132), `delete_event_removes_a_personal_appointment` (~182), `list_events_filters_by_query_and_category` (~206).

- [ ] **Step 8: Run the full build and test suite**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo build 2>&1 | Select-Object -Last 20 && cargo test 2>&1 | Select-Object -Last 40`
Expected: PASS — clean build, `create_event_forwards_recurrence` and every other test green. `tests/live_outlook.rs` compiles (its tests are `#[ignore]`d, not run here).

- [ ] **Step 9: Commit**

```bash
git add src/server.rs src/outlook/fake.rs src/outlook/client.rs tests/live_outlook.rs tests/tools.rs
git commit -m "Add recurrence to create_event: RecurrenceParams, fake forwarding, real GetRecurrencePattern() write"
```

---

### Task 3: `recurrence` read-back on `get_event`

Add `recurrence_info()`, a real-COM helper that reads an appointment's `RecurrencePattern` back into a `RecurrenceInfo`, and wire it into `get_event`'s `EventDetail`.

**Files:**
- Modify: `src/outlook/client.rs` (add `recurrence_info` after `event_summary`, line ~273; wire into `get_event`, line ~965)
- Modify: `src/outlook/fake.rs` (add `recurrence: None` to the `EventDetail` literal in `get_event`, line ~166)
- Modify: `tests/tools.rs` (add a `get_event_recurrence_is_none_by_default` test using the fake, to lock in the JSON shape)

**Interfaces:**
- Consumes: `RecurrenceInfo` (Task 1, `src/outlook/types.rs`), `recurrence_pattern_word`/`day_of_week_mask_to_words` (Task 1, `src/friendly.rs`), existing `get_property`/`call_method`/`to_disp`/`variant_to_i32`/`variant_to_bool`/`variant_to_iso_string`.
- Produces: `fn recurrence_info(item: &IDispatch) -> Result<Option<RecurrenceInfo>, ToolError>` in `client.rs` — `None` if the item isn't recurring, `Some(..)` otherwise.

- [ ] **Step 1: Write the failing fake-client test in `tests/tools.rs`**

Add directly after `create_event_forwards_recurrence`:

```rust
#[tokio::test]
async fn get_event_recurrence_is_none_by_default() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_event(Parameters(GetEventParams { event_id: EVENT_ID.to_string() }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert!(v["recurrence"].is_null());
}
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --test tools get_event_recurrence_is_none_by_default 2>&1 | Select-Object -Last 30`
Expected: FAIL to compile — `EventDetail` doesn't have a `recurrence` field populated in the fake yet (it's a compile error from Task 1's struct change until fixed here).

- [ ] **Step 3: Fix the fake's `get_event` in `src/outlook/fake.rs`**

Add `recurrence: None,` to the `EventDetail { .. }` literal in `get_event`:

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
            recurrence: None,
        })
    }
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --test tools get_event_recurrence_is_none_by_default 2>&1 | Select-Object -Last 30`
Expected: PASS.

- [ ] **Step 5: Fix the real client's `get_event`/`EventDetail` construction and add `recurrence_info` in `src/outlook/client.rs`**

Add `recurrence_info` directly after `event_summary` (after its closing `}`, before the `event_matches` doc comment):

```rust
/// Reads an appointment's recurrence pattern back via
/// `GetRecurrencePattern()`, or `None` if `IsRecurring` is false. Mirrors
/// `apply_recurrence`'s field set in the opposite direction.
fn recurrence_info(item: &IDispatch) -> Result<Option<RecurrenceInfo>, ToolError> {
    let is_recurring =
        variant_to_bool(&get_property(item, "IsRecurring").unwrap_or_default()).unwrap_or(false);
    if !is_recurring {
        return Ok(None);
    }
    let pattern = to_disp(call_method(item, "GetRecurrencePattern", &mut [])?)?;
    let recurrence_type =
        variant_to_i32(&get_property(&pattern, "RecurrenceType").unwrap_or_default())
            .unwrap_or(c::OL_RECURS_DAILY);
    let interval =
        variant_to_i32(&get_property(&pattern, "Interval").unwrap_or_default()).unwrap_or(1);
    let day_mask =
        variant_to_i32(&get_property(&pattern, "DayOfWeekMask").unwrap_or_default()).unwrap_or(0);
    let day_of_month = variant_to_i32(&get_property(&pattern, "DayOfMonth").unwrap_or_default());
    let no_end =
        variant_to_bool(&get_property(&pattern, "NoEndDate").unwrap_or_default()).unwrap_or(false);
    let occurrences = variant_to_i32(&get_property(&pattern, "Occurrences").unwrap_or_default());
    let until = if no_end {
        None
    } else {
        variant_to_iso_string(&get_property(&pattern, "PatternEndDate").unwrap_or_default())
    };
    Ok(Some(RecurrenceInfo {
        pattern: crate::friendly::recurrence_pattern_word(recurrence_type).to_string(),
        interval,
        days_of_week: crate::friendly::day_of_week_mask_to_words(day_mask),
        day_of_month: day_of_month.filter(|d| *d > 0),
        until,
        occurrences: if no_end { None } else { occurrences },
        no_end,
    }))
}
```

Update `get_event` to populate it:

```rust
    fn get_event(&self, event_id: String) -> Result<EventDetail, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let item = get_item(&ns, &event_id)?;
            let summary = event_summary(&item)?;
            let recurrence = recurrence_info(&item)?;
            Ok(EventDetail {
                summary,
                body: truncate(&variant_to_string(&get_property(&item, "Body").unwrap_or_default())),
                recurrence,
            })
        })
    }
```

- [ ] **Step 6: Run the full build and test suite**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo build 2>&1 | Select-Object -Last 20 && cargo test 2>&1 | Select-Object -Last 40`
Expected: PASS — clean build, all tests green.

- [ ] **Step 7: Commit**

```bash
git add src/outlook/client.rs src/outlook/fake.rs tests/tools.rs
git commit -m "Add recurrence read-back to get_event via GetRecurrencePattern()"
```

---

### Task 4: `recurrence` + `clear_recurrence` on `update_event`

Let an existing event's recurrence pattern be set, replaced, or cleared. Reuses `apply_recurrence` from Task 2 unchanged; adds the one new operation — `AppointmentItem.ClearRecurrencePattern()`.

**Files:**
- Modify: `src/outlook/fake.rs` (add `recurrence`/`clear_recurrence` to `update_event`'s recorded args and `changed` list, line ~193)
- Modify: `src/server.rs` (add `recurrence`/`clear_recurrence` fields to `UpdateEventParams`, line ~235; wire into the `update_event` tool method, line ~486)
- Modify: `src/outlook/client.rs` (wire `apply_recurrence`/`ClearRecurrencePattern` into `update_event`, line ~1075)
- Modify: `tests/live_outlook.rs` (add `recurrence: None, clear_recurrence: false,` to the 1 `EventUpdate { .. }` literal, line ~146)
- Modify: `tests/tools.rs` (add 2 tests: recurrence forwarded + rejecting both `recurrence` and `clear_recurrence` set together)

**Interfaces:**
- Consumes: `apply_recurrence` (Task 2, `client.rs`), `RecurrenceInput`/`RecurrenceParams` (Tasks 1/2).
- Produces (fake/real return contract addition): `changed` may now include `"recurrence"` or `"clear_recurrence"`.

- [ ] **Step 1: Write the failing fake-client tool tests in `tests/tools.rs`**

Add directly after `update_event_remove_attendees_is_tracked`:

```rust
#[tokio::test]
async fn update_event_forwards_recurrence() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_event(Parameters(UpdateEventParams {
            event_id: EVENT_ID.to_string(),
            subject: None, start: None, end: None, location: None, body: None,
            all_day: None, reminder_minutes: None, show_as: None,
            add_categories: None, remove_categories: None,
            add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
            send_update: false,
            recurrence: Some(RecurrenceParams {
                pattern: "daily".to_string(), interval: Some(2), days_of_week: None,
                day_of_month: None, until: Some("2099-06-01".to_string()), occurrences: None,
            }),
            clear_recurrence: false,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["changed"], json!(["recurrence"]));
    let (_, args) = fake.calls().last().unwrap().clone();
    assert_eq!(args["recurrence"]["pattern"], "daily");
    assert_eq!(args["recurrence"]["until"], "2099-06-01");
}

#[tokio::test]
async fn update_event_forwards_clear_recurrence() {
    use outlook_mcp_rs::outlook::fake::EVENT_ID;
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .update_event(Parameters(UpdateEventParams {
            event_id: EVENT_ID.to_string(),
            subject: None, start: None, end: None, location: None, body: None,
            all_day: None, reminder_minutes: None, show_as: None,
            add_categories: None, remove_categories: None,
            add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
            send_update: false,
            recurrence: None,
            clear_recurrence: true,
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["changed"], json!(["clear_recurrence"]));
}
```

Also add `RecurrenceParams` to the `tests/tools.rs` import block if Task 2 didn't already (it did — no change needed here, just confirming).

- [ ] **Step 2: Run the tests to verify they fail**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --test tools update_event_forwards 2>&1 | Select-Object -Last 30`
Expected: FAIL to compile — `UpdateEventParams` has no `recurrence`/`clear_recurrence` fields yet.

- [ ] **Step 3: Add `recurrence`/`clear_recurrence` to `UpdateEventParams` in `src/server.rs`**

Add the two fields as the last two fields of `UpdateEventParams` (after `pub send_update: bool,`):

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
    /// Set or replace the event's recurrence pattern (whole series, not one
    /// occurrence). Mutually exclusive with `clear_recurrence`.
    #[serde(default)]
    pub recurrence: Option<RecurrenceParams>,
    /// Remove the event's recurrence pattern, converting it back to a single
    /// occurrence. Mutually exclusive with `recurrence`.
    #[serde(default)]
    pub clear_recurrence: bool,
}
```

Update the `update_event` tool method:

```rust
    #[tool(description = "Update an existing calendar event: subject, start/end, location, body, show_as, add/remove categories, add/remove attendees, reminder, all_day, recurrence (set/replace) or clear_recurrence (remove it). Adding an attendee converts a personal appointment into a meeting. Recurrence edits apply to the whole series. send_update (default true) notifies attendees if the event is a meeting.")]
    pub async fn update_event(
        &self,
        Parameters(p): Parameters<UpdateEventParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let recurrence = p.recurrence.map(|r| RecurrenceInput {
            pattern: r.pattern, interval: r.interval, days_of_week: r.days_of_week,
            day_of_month: r.day_of_month, until: r.until, occurrences: r.occurrences,
        });
        let u = EventUpdate {
            event_id: p.event_id, subject: p.subject, start: p.start, end: p.end,
            location: p.location, body: p.body, all_day: p.all_day,
            reminder_minutes: p.reminder_minutes, show_as: p.show_as,
            add_categories: p.add_categories, remove_categories: p.remove_categories,
            add_required_attendees: p.add_required_attendees,
            add_optional_attendees: p.add_optional_attendees,
            remove_attendees: p.remove_attendees, send_update: p.send_update,
            recurrence, clear_recurrence: p.clear_recurrence,
        };
        let result = run_blocking(move || client.update_event(u)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

- [ ] **Step 4: Update `src/outlook/fake.rs`'s `update_event`**

Replace the `update_event` implementation:

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
            "recurrence": u.recurrence.as_ref().map(|r| json!({
                "pattern": r.pattern, "interval": r.interval, "days_of_week": r.days_of_week,
                "day_of_month": r.day_of_month, "until": r.until, "occurrences": r.occurrences,
            })),
            "clear_recurrence": u.clear_recurrence,
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
        if u.recurrence.is_some() { changed.push("recurrence"); }
        if u.clear_recurrence { changed.push("clear_recurrence"); }
        Ok(json!({"status": "updated", "id": u.event_id, "changed": changed}))
    }
```

- [ ] **Step 5: Run the tool tests to verify they pass**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --test tools update_event_forwards 2>&1 | Select-Object -Last 30`
Expected: PASS.

- [ ] **Step 6: Wire real COM into `update_event` in `src/outlook/client.rs`**

Add validation plus the two branches directly after the existing `remove_attendees` block (after `changed.push("remove_attendees"); }` closing brace) and before the `// Save vs Send` comment:

```rust
            if u.recurrence.is_some() && u.clear_recurrence {
                return Err(ToolError::new(
                    "cannot set recurrence and clear_recurrence in the same update_event call",
                ));
            }
            if let Some(recurrence) = u.recurrence.as_ref() {
                apply_recurrence(&item, recurrence)?;
                changed.push("recurrence");
            }
            if u.clear_recurrence {
                call_method(&item, "ClearRecurrencePattern", &mut [])?;
                changed.push("clear_recurrence");
            }
```

- [ ] **Step 7: Fix the `EventUpdate { .. }` literal in `tests/live_outlook.rs`**

Add `recurrence: None, clear_recurrence: false,` to the one `EventUpdate { .. }` literal (line ~146, in `update_event_edits_fields_and_manages_attendees`), after `send_update: false,`:

```rust
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
        recurrence: None,
        clear_recurrence: false,
    }).expect("update_event should succeed");
```

- [ ] **Step 8: Run the full build and test suite**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo build 2>&1 | Select-Object -Last 20 && cargo test 2>&1 | Select-Object -Last 40`
Expected: PASS — clean build, all unit + tool tests green, 0 warnings.

- [ ] **Step 9: Commit**

```bash
git add src/server.rs src/outlook/fake.rs src/outlook/client.rs tests/live_outlook.rs tests/tools.rs
git commit -m "Add recurrence set/replace and clear_recurrence to update_event"
```

---

### Task 5: Live recurrence round-trip tests

Verify the whole feature against real Outlook: create weekly/monthly/yearly recurring events and confirm `get_event` reads the pattern back correctly, update a recurring event's pattern, clear recurrence off an event, and confirm `list_events`' existing `is_recurring` flag (already shipped, Plan 6) still reflects a recurring event correctly. Every test uses far-future dates (`2099-...`) and either no attendees or `send:false`, and cleans up via `delete_event`, matching every prior plan's live-test discipline.

**Files:**
- Modify: `tests/live_outlook.rs` (add 4 new `#[ignore]`d tests after `list_events_filters_by_query_and_category`)
- Modify: `TESTING.md` (add a one-line note under the live-test section, if one doesn't already cover calendar round-trips generically — check first; only add if genuinely new ground)

**Interfaces:**
- Consumes: everything from Tasks 1-4 — `CreateEventInput.recurrence`, `EventUpdate.recurrence`/`clear_recurrence`, `EventDetail.recurrence`, `EventSummary.is_recurring` (pre-existing).

- [ ] **Step 1: Write the live weekly-recurrence round-trip test**

Add to `tests/live_outlook.rs`, after the last test in the file (`list_events_filters_by_query_and_category`'s closing `}`):

```rust
#[test]
#[ignore]
fn create_event_weekly_recurrence_round_trips() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P9 weekly recurrence probe".to_string(),
        start: "2099-02-02T09:00".to_string(), // a Monday
        end: "2099-02-02T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
        recurrence: Some(RecurrenceInput {
            pattern: "weekly".to_string(),
            interval: Some(1),
            days_of_week: Some(vec!["monday".to_string(), "wednesday".to_string()]),
            day_of_month: None,
            until: None,
            occurrences: Some(10),
        }),
    }).expect("create_event with weekly recurrence should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    assert!(detail.summary.is_recurring);
    let recurrence = detail.recurrence.expect("recurring event should have a recurrence block");
    assert_eq!(recurrence.pattern, "weekly");
    assert_eq!(recurrence.interval, 1);
    assert_eq!(recurrence.days_of_week, vec!["monday".to_string(), "wednesday".to_string()]);
    assert_eq!(recurrence.occurrences, Some(10));
    assert!(!recurrence.no_end);

    c.delete_event(id, false).expect("cleanup delete_event");
}

#[test]
#[ignore]
fn create_event_monthly_recurrence_with_until_round_trips() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P9 monthly recurrence probe".to_string(),
        start: "2099-02-15T09:00".to_string(),
        end: "2099-02-15T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
        recurrence: Some(RecurrenceInput {
            pattern: "monthly".to_string(),
            interval: Some(2),
            days_of_week: None,
            day_of_month: Some(15),
            until: Some("2099-12-15".to_string()),
            occurrences: None,
        }),
    }).expect("create_event with monthly recurrence should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    let recurrence = detail.recurrence.expect("recurring event should have a recurrence block");
    assert_eq!(recurrence.pattern, "monthly");
    assert_eq!(recurrence.interval, 2);
    assert_eq!(recurrence.day_of_month, Some(15));
    assert!(recurrence.until.is_some());
    assert!(!recurrence.no_end);

    c.delete_event(id, false).expect("cleanup delete_event");
}

#[test]
#[ignore]
fn create_event_yearly_recurrence_with_no_end_round_trips() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P9 yearly recurrence probe".to_string(),
        start: "2099-03-10T09:00".to_string(),
        end: "2099-03-10T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
        recurrence: Some(RecurrenceInput {
            pattern: "yearly".to_string(),
            interval: None,
            days_of_week: None,
            day_of_month: None,
            until: None,
            occurrences: None,
        }),
    }).expect("create_event with yearly recurrence should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    let recurrence = detail.recurrence.expect("recurring event should have a recurrence block");
    assert_eq!(recurrence.pattern, "yearly");
    assert_eq!(recurrence.day_of_month, Some(10)); // derived from the March 10 start date
    assert!(recurrence.no_end);
    assert!(recurrence.until.is_none());
    assert!(recurrence.occurrences.is_none());

    c.delete_event(id, false).expect("cleanup delete_event");
}

#[test]
#[ignore]
fn update_event_changes_then_clears_recurrence() {
    let c = client();
    let created = c.create_event(CreateEventInput {
        subject: "outlook-mcp-rs P9 update recurrence probe".to_string(),
        start: "2099-04-01T09:00".to_string(),
        end: "2099-04-01T09:30".to_string(),
        body: None, location: None, required_attendees: None, optional_attendees: None,
        all_day: false, reminder_minutes: None, categories: None, show_as: None,
        send: true,
        recurrence: Some(RecurrenceInput {
            pattern: "daily".to_string(), interval: Some(1), days_of_week: None,
            day_of_month: None, until: None, occurrences: Some(3),
        }),
    }).expect("create_event should succeed");
    let id = created["id"].as_str().unwrap().to_string();

    // Change the pattern from daily to weekly.
    let updated = c.update_event(EventUpdate {
        event_id: id.clone(),
        subject: None, start: None, end: None, location: None, body: None,
        all_day: None, reminder_minutes: None, show_as: None,
        add_categories: None, remove_categories: None,
        add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
        send_update: false,
        recurrence: Some(RecurrenceInput {
            pattern: "weekly".to_string(), interval: Some(1),
            days_of_week: Some(vec!["tuesday".to_string()]),
            day_of_month: None, until: None, occurrences: Some(4),
        }),
        clear_recurrence: false,
    }).expect("update_event with recurrence should succeed");
    assert!(updated["changed"].as_array().unwrap().iter().any(|v| v == "recurrence"));

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    let recurrence = detail.recurrence.expect("still recurring after the change");
    assert_eq!(recurrence.pattern, "weekly");
    assert_eq!(recurrence.days_of_week, vec!["tuesday".to_string()]);

    // Now clear it entirely.
    let cleared = c.update_event(EventUpdate {
        event_id: id.clone(),
        subject: None, start: None, end: None, location: None, body: None,
        all_day: None, reminder_minutes: None, show_as: None,
        add_categories: None, remove_categories: None,
        add_required_attendees: None, add_optional_attendees: None, remove_attendees: None,
        send_update: false,
        recurrence: None,
        clear_recurrence: true,
    }).expect("update_event with clear_recurrence should succeed");
    assert!(cleared["changed"].as_array().unwrap().iter().any(|v| v == "clear_recurrence"));

    let detail = c.get_event(id.clone()).expect("get_event should succeed");
    assert!(!detail.summary.is_recurring);
    assert!(detail.recurrence.is_none());

    c.delete_event(id, false).expect("cleanup delete_event");
}
```

Add `RecurrenceInput` to the `use outlook_mcp_rs::outlook::{...}` import at the top of `tests/live_outlook.rs` — change `use outlook_mcp_rs::outlook::{CreateEventInput, EmailQuery, EventQuery, OutlookClient, EmailUpdate, EventUpdate};` to `use outlook_mcp_rs::outlook::{CreateEventInput, EmailQuery, EventQuery, OutlookClient, EmailUpdate, EventUpdate, RecurrenceInput};`.

- [ ] **Step 2: Confirm the new tests compile and are ignored by default**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --test live_outlook 2>&1 | Select-Object -Last 20`
Expected: PASS — build succeeds, output shows the 4 new tests (and all existing ones) as `ignored`, 0 run.

- [ ] **Step 3: Run the 4 new live tests against real Outlook**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo test --test live_outlook -- --ignored create_event_weekly_recurrence_round_trips create_event_monthly_recurrence_with_until_round_trips create_event_yearly_recurrence_with_no_end_round_trips update_event_changes_then_clears_recurrence 2>&1 | Select-Object -Last 40`
Expected: PASS — all 4 tests green against the real, running Outlook desktop app. If any fails, use superpowers:systematic-debugging before touching Task 1-4's code (a live COM failure here most likely means a genuine `RecurrencePattern` property-ordering or bitmask bug, not a test bug — trace it against the Microsoft `RecurrencePattern` object model rather than guessing).

- [ ] **Step 4: Run the full non-live suite once more to confirm nothing regressed**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && cargo build 2>&1 | Select-Object -Last 20 && cargo test 2>&1 | Select-Object -Last 40`
Expected: PASS — 0 warnings, every unit + tool test green.

- [ ] **Step 5: Check `TESTING.md` for whether a note is needed**

Run: `cd C:/Users/adamk/projects/outlook-mcp-rs && Get-Content TESTING.md | Select-String -Context 2,2 "live_outlook|calendar"`
If `TESTING.md` already documents "run `cargo test --test live_outlook -- --ignored` for the full live calendar suite" generically (it does, from prior plans), no edit is needed — the 4 new tests are automatically covered by that existing instruction. Only add a note if recurrence introduces a genuinely new manual-only caveat (it doesn't: every recurrence path here is safely automatable with `send:false`/no attendees, unlike `send_email` or a real meeting invite).

- [ ] **Step 6: Commit**

```bash
git add tests/live_outlook.rs
git commit -m "Add live recurrence round-trip tests: weekly/monthly/yearly create, update+clear"
```

---

## After all 5 tasks are green

Run a final whole-branch review (superpowers:requesting-code-review) covering all 5 tasks together — check that `apply_recurrence`/`recurrence_info` stay symmetric (every field `apply_recurrence` writes, `recurrence_info` reads back the same way), that `validate_recurrence` is the single source of pattern-validation truth (grep for any duplicated inline validation), and that all 6 previously-existing `CreateEventInput`/`EventUpdate` call sites plus the 2 new struct fields compile with zero warnings. Then `git push origin main` per the established workflow, and mark Plan 9 shipped in `V2-RESUME.md` and the plans index.
