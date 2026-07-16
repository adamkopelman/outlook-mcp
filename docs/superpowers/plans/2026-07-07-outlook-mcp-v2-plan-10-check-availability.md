# check_availability Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a new `check_availability` tool that returns free/busy availability for a list of people over a time window, plus a computed intersection of when everyone is free — without ever exposing event details.

**Architecture:** A pure-logic layer (parsing Outlook's raw FreeBusy status-code string into timestamped slots, and computing the common-free intersection across people) lives in `src/outlook/mod.rs`, fully unit-testable without COM. The real implementation in `src/outlook/client.rs` resolves each person via `Namespace.CreateRecipient`/`Recipient.Resolve` (the same pattern `list_events`'s `calendar_of` already uses), calls `Recipient.FreeBusy(Start, MinPerChar, CompleteFormat)`, and feeds the raw string through the pure parser. One person failing to resolve is marked `resolved: false` and does not fail the whole call.

**Tech Stack:** Rust, `windows` crate 0.62.2 (Win32 COM/`IDispatch::Invoke`), `rmcp` 2.1.0 tool macros, `chrono`, `serde`/`serde_json`.

## Global Constraints

- `check_availability` parameters (exact names, from the spec): `people` (required, `Vec<String>` of emails/names), `start` (required, ISO datetime string), `end` (required, ISO datetime string), `interval_minutes` (optional, default `30`), `treat_as_free` (optional, default `["free"]`).
- Implementation must use `Recipient.FreeBusy(start, MinPerChar=interval_minutes, CompleteFormat=true)`. `CompleteFormat=true` is required — it's what makes Outlook return the 5-code format (0 free, 1 tentative, 2 busy, 3 out-of-office, 4 working-elsewhere) instead of the legacy 4-code format that omits working-elsewhere.
- Raw `slots` in the output always show the **true** status — never reinterpreted through `treat_as_free`.
- `common_free` = the windows where every **resolved** person's status is in `treat_as_free`. An unresolved person is excluded from the intersection (not counted as unavailable, not counted as available — simply not part of the computation), matching "does not fail the whole call."
- One person failing to resolve is marked `resolved: false` with empty `slots`, not an error for the whole call.
- The friendly-word status vocabulary (`"free"`, `"tentative"`, `"busy"`, `"out_of_office"`, `"working_elsewhere"`) and its underlying `OlBusyStatus` codes (0-4) are **already implemented** in `src/friendly.rs` (`busy_status_word`/`busy_status_to_id`) and `src/constants.rs` (`OL_FREE`/`OL_TENTATIVE`/`OL_BUSY`/`OL_OUT_OF_OFFICE`/`OL_WORKING_ELSEWHERE`) for `create_event`/`update_event`'s `show_as` field — FreeBusy's status codes use the identical numbering, so this plan reuses `busy_status_word` directly rather than adding a second mapping.
- The `OutlookClient` trait (`src/outlook/mod.rs`) has two implementors — `WindowsOutlookClient` (`client.rs`) and `FakeOutlookClient` (`fake.rs`) — plus `src/server.rs` (MCP tool layer) and `tests/tools.rs` (fake-backed tests); every trait change touches all four.
- Per this project's model policy: task 1 (pure logic) is mechanical/cheap-tier; task 2 (COM + trait + fake + server wiring) is moderate/standard-tier; task 3 (live test) is standard-tier.

---

### Task 1: Types and pure logic (`AvailabilityResult`, `parse_freebusy_slots`, `common_free`)

**Files:**
- Modify: `src/outlook/types.rs` (add 4 new `Serialize` structs)
- Modify: `src/outlook/mod.rs` (add `CheckAvailabilityInput`, `parse_freebusy_slots`, `common_free`, unit tests)

**Interfaces:**
- Produces (consumed by Task 2):
  - `types::AvailabilitySlot { pub start: String, pub end: String, pub status: String }`
  - `types::PersonAvailability { pub person: String, pub resolved: bool, pub slots: Vec<AvailabilitySlot> }`
  - `types::FreeWindow { pub start: String, pub end: String }`
  - `types::AvailabilityResult { pub people: Vec<PersonAvailability>, pub common_free: Vec<FreeWindow> }`
  - `CheckAvailabilityInput { pub people: Vec<String>, pub start: String, pub end: String, pub interval_minutes: i32, pub treat_as_free: Vec<String> }` (in `mod.rs`, alongside `CreateEventInput`/`EventQuery`)
  - `pub fn parse_freebusy_slots(raw: &str, start: &chrono::NaiveDateTime, interval_minutes: i32, max_slots: usize) -> Vec<AvailabilitySlot>`
  - `pub fn common_free(people: &[PersonAvailability], treat_as_free: &[String]) -> Vec<FreeWindow>`

- [ ] **Step 1: Add the 4 output structs to `types.rs`**

Open `src/outlook/types.rs` and find the existing `RecurrenceInfo` struct (added in Plan 9) — add the new structs directly after it:

```rust
#[derive(Debug, Clone, Serialize, PartialEq)]
pub struct AvailabilitySlot {
    pub start: String,
    pub end: String,
    pub status: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct PersonAvailability {
    pub person: String,
    pub resolved: bool,
    pub slots: Vec<AvailabilitySlot>,
}

#[derive(Debug, Clone, Serialize, PartialEq)]
pub struct FreeWindow {
    pub start: String,
    pub end: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct AvailabilityResult {
    pub people: Vec<PersonAvailability>,
    pub common_free: Vec<FreeWindow>,
}
```

- [ ] **Step 2: Write the failing tests for `parse_freebusy_slots` and `common_free`**

Open `src/outlook/mod.rs`, find the `#[cfg(test)] mod tests` block (added in Plan 9), and add these tests at the end, right before the closing `}` of the module:

```rust
    fn dt(s: &str) -> chrono::NaiveDateTime {
        chrono::NaiveDateTime::parse_from_str(s, "%Y-%m-%dT%H:%M:%S").unwrap()
    }

    #[test]
    fn parse_freebusy_slots_maps_codes_to_words_and_times() {
        // "02143" = free, busy, tentative, working_elsewhere, out_of_office
        let slots = parse_freebusy_slots("02143", &dt("2099-01-01T09:00:00"), 30, 5);
        assert_eq!(slots.len(), 5);
        assert_eq!(slots[0], AvailabilitySlot {
            start: "2099-01-01T09:00:00".to_string(),
            end: "2099-01-01T09:30:00".to_string(),
            status: "free".to_string(),
        });
        assert_eq!(slots[1].status, "busy");
        assert_eq!(slots[1].start, "2099-01-01T09:30:00");
        assert_eq!(slots[1].end, "2099-01-01T10:00:00");
        assert_eq!(slots[2].status, "tentative");
        assert_eq!(slots[3].status, "working_elsewhere");
        assert_eq!(slots[4].status, "out_of_office");
    }

    #[test]
    fn parse_freebusy_slots_truncates_to_max_slots() {
        // Outlook's raw FreeBusy string commonly covers a much longer range
        // than the caller's requested [start, end) window.
        let slots = parse_freebusy_slots("000000000000", &dt("2099-01-01T09:00:00"), 30, 3);
        assert_eq!(slots.len(), 3);
    }

    #[test]
    fn parse_freebusy_slots_treats_unrecognized_digit_as_busy() {
        let slots = parse_freebusy_slots("9", &dt("2099-01-01T09:00:00"), 30, 1);
        assert_eq!(slots[0].status, "busy");
    }

    fn avail(person: &str, resolved: bool, statuses: &[&str]) -> PersonAvailability {
        let mut slots = Vec::new();
        for (i, s) in statuses.iter().enumerate() {
            let start = dt("2099-01-01T09:00:00") + chrono::Duration::minutes(i as i64 * 30);
            slots.push(AvailabilitySlot {
                start: start.format("%Y-%m-%dT%H:%M:%S").to_string(),
                end: (start + chrono::Duration::minutes(30)).format("%Y-%m-%dT%H:%M:%S").to_string(),
                status: s.to_string(),
            });
        }
        PersonAvailability { person: person.to_string(), resolved, slots }
    }

    #[test]
    fn common_free_intersects_only_where_everyone_is_free() {
        let people = vec![
            avail("alice", true, &["free", "free", "busy"]),
            avail("bob", true, &["free", "busy", "busy"]),
        ];
        let windows = common_free(&people, &["free".to_string()]);
        assert_eq!(windows, vec![FreeWindow {
            start: "2099-01-01T09:00:00".to_string(),
            end: "2099-01-01T09:30:00".to_string(),
        }]);
    }

    #[test]
    fn common_free_merges_contiguous_free_slots_into_one_window() {
        let people = vec![avail("alice", true, &["free", "free", "busy", "free"])];
        let windows = common_free(&people, &["free".to_string()]);
        assert_eq!(windows, vec![
            FreeWindow { start: "2099-01-01T09:00:00".to_string(), end: "2099-01-01T10:00:00".to_string() },
            FreeWindow { start: "2099-01-01T10:30:00".to_string(), end: "2099-01-01T11:00:00".to_string() },
        ]);
    }

    #[test]
    fn common_free_respects_custom_treat_as_free() {
        let people = vec![avail("alice", true, &["tentative"])];
        assert_eq!(common_free(&people, &["free".to_string()]), vec![]);
        assert_eq!(
            common_free(&people, &["free".to_string(), "tentative".to_string()]).len(),
            1
        );
    }

    #[test]
    fn common_free_ignores_unresolved_people() {
        let people = vec![
            avail("alice", true, &["free"]),
            avail("bob", false, &[]),
        ];
        assert_eq!(common_free(&people, &["free".to_string()]).len(), 1);
    }

    #[test]
    fn common_free_empty_when_no_one_resolved() {
        let people = vec![avail("alice", false, &[])];
        assert_eq!(common_free(&people, &["free".to_string()]), vec![]);
    }
```

Add `AvailabilitySlot`, `PersonAvailability`, `FreeWindow` to the test module's imports at the top of `mod tests`: change

```rust
    use super::{
        com_recurrence_interval, create_event_status, friendly_recurrence_interval,
        validate_recurrence, RecurrenceInput,
    };
```

to

```rust
    use super::{
        com_recurrence_interval, common_free, create_event_status, friendly_recurrence_interval,
        parse_freebusy_slots, validate_recurrence, RecurrenceInput,
    };
    use crate::outlook::types::{AvailabilitySlot, FreeWindow, PersonAvailability};
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `cargo test --lib parse_freebusy_slots` and `cargo test --lib common_free`
Expected: FAIL with "cannot find function `parse_freebusy_slots`" / "cannot find function `common_free`" (not yet defined).

- [ ] **Step 4: Implement `CheckAvailabilityInput`, `parse_freebusy_slots`, `common_free`**

In `src/outlook/mod.rs`, add right after `friendly_recurrence_interval` (the last function Plan 9 added) and before `#[cfg(test)] mod tests`:

```rust
/// All inputs for `check_availability`. `treat_as_free` decides which raw
/// statuses count as "free" when computing `common_free` — it never changes
/// what a person's own `slots` report (those always show the true status).
#[derive(Debug, Clone)]
pub struct CheckAvailabilityInput {
    pub people: Vec<String>,
    pub start: String,
    pub end: String,
    pub interval_minutes: i32,
    pub treat_as_free: Vec<String>,
}

/// Parses Outlook's raw `Recipient.FreeBusy` status-code string (one ASCII
/// digit per `interval_minutes`-sized slot: `'0'` free, `'1'` tentative,
/// `'2'` busy, `'3'` out-of-office, `'4'` working-elsewhere — the exact
/// `OlBusyStatus` numbering `friendly::busy_status_word` already maps) into
/// timestamped slots starting at `start`. `FreeBusy` returns a string
/// covering a much longer range than the caller's `[start, end)` window (it
/// has no `end` parameter), so callers must compute `max_slots` themselves
/// — `(end - start) / interval_minutes`, rounded up — and this function
/// truncates to it. Any digit outside 0-4 (Outlook shouldn't produce one,
/// but the string could be malformed) falls back to `"busy"`, the same
/// catch-all `busy_status_word` uses.
pub fn parse_freebusy_slots(
    raw: &str,
    start: &chrono::NaiveDateTime,
    interval_minutes: i32,
    max_slots: usize,
) -> Vec<AvailabilitySlot> {
    raw.chars()
        .take(max_slots)
        .enumerate()
        .map(|(i, ch)| {
            let code = ch.to_digit(10).map(|d| d as i32).unwrap_or(c::OL_BUSY);
            let slot_start = *start + chrono::Duration::minutes(i as i64 * interval_minutes as i64);
            let slot_end = slot_start + chrono::Duration::minutes(interval_minutes as i64);
            AvailabilitySlot {
                start: slot_start.format("%Y-%m-%dT%H:%M:%S").to_string(),
                end: slot_end.format("%Y-%m-%dT%H:%M:%S").to_string(),
                status: crate::friendly::busy_status_word(code).to_string(),
            }
        })
        .collect()
}

/// The windows where every **resolved** person's status is in
/// `treat_as_free` (case-insensitive). Unresolved people are skipped
/// entirely — they neither block nor contribute to a common-free window.
/// Assumes all resolved people's `slots` share the same slot boundaries
/// (true whenever they were built from the same `start`/`interval_minutes`,
/// which `check_availability` always uses); intersects only over the
/// shortest `slots` length present, so a person whose raw string was
/// unexpectedly short doesn't panic the lookup.
pub fn common_free(people: &[PersonAvailability], treat_as_free: &[String]) -> Vec<FreeWindow> {
    let resolved: Vec<&PersonAvailability> = people.iter().filter(|p| p.resolved).collect();
    if resolved.is_empty() {
        return Vec::new();
    }
    let treat_lower: Vec<String> = treat_as_free.iter().map(|s| s.to_lowercase()).collect();
    let min_len = resolved.iter().map(|p| p.slots.len()).min().unwrap_or(0);

    let mut windows = Vec::new();
    let mut run_start: Option<usize> = None;
    for i in 0..min_len {
        let all_free = resolved
            .iter()
            .all(|p| treat_lower.contains(&p.slots[i].status.to_lowercase()));
        if all_free {
            run_start.get_or_insert(i);
        } else if let Some(s) = run_start.take() {
            windows.push(FreeWindow {
                start: resolved[0].slots[s].start.clone(),
                end: resolved[0].slots[i - 1].end.clone(),
            });
        }
    }
    if let Some(s) = run_start {
        windows.push(FreeWindow {
            start: resolved[0].slots[s].start.clone(),
            end: resolved[0].slots[min_len - 1].end.clone(),
        });
    }
    windows
}
```

No new top-level import is needed in `mod.rs` itself: line 8 already has `use types::*;`, a glob import that already brings `AvailabilitySlot`/`PersonAvailability`/`FreeWindow` into scope for `mod.rs`'s own code (including `common_free`'s signature, which references `PersonAvailability`/`FreeWindow` directly). The `mod tests` submodule does **not** inherit that glob import automatically (Rust `use` isn't inherited by child modules) — that's exactly why Step 2 above adds its own explicit `use crate::outlook::types::{AvailabilitySlot, FreeWindow, PersonAvailability};` inside the test module. Do not skip that import.

- [ ] **Step 5: Run the tests to verify they pass**

Run: `cargo test --lib parse_freebusy_slots` and `cargo test --lib common_free`
Expected: all 9 new tests pass.

- [ ] **Step 6: Run the full lib test suite and commit**

Run: `cargo build` (expect 0 warnings) then `cargo test --lib` (expect all passing, count increased by 9).

```bash
git add src/outlook/types.rs src/outlook/mod.rs
git commit -m "Add check_availability types and pure logic: parse_freebusy_slots, common_free"
```

---

### Task 2: Real COM implementation, fake client, and tool wiring

**Files:**
- Modify: `src/outlook/mod.rs` (add `check_availability` to the `OutlookClient` trait)
- Modify: `src/outlook/client.rs` (real COM implementation)
- Modify: `src/outlook/fake.rs` (fake implementation)
- Modify: `src/server.rs` (`CheckAvailabilityParams`, `check_availability` tool method)
- Test: `tests/tools.rs` (fake-backed tests)

**Interfaces:**
- Consumes (from Task 1): `CheckAvailabilityInput`, `types::AvailabilityResult`, `parse_freebusy_slots`, `common_free`.
- Consumes (pre-existing): `com::{call_method, to_disp, variant_from_str, variant_to_bool, variant_from_datetime, variant_from_i32, variant_to_string}`, `parse_dt` (in `client.rs`), `friendly::busy_status_word` (indirectly, via `parse_freebusy_slots`).
- Produces (consumed by Task 3): `OutlookClient::check_availability(&self, input: CheckAvailabilityInput) -> Result<AvailabilityResult, ToolError>` on both implementors; MCP tool `check_availability`.

- [ ] **Step 1: Add the trait method**

In `src/outlook/mod.rs`, find the `OutlookClient` trait's calendar section (`list_events`/`get_event`/`create_event`/.../`delete_event`) and add, right after `delete_event`:

```rust
    fn check_availability(&self, input: CheckAvailabilityInput) -> Result<AvailabilityResult, ToolError>;
```

Run `cargo build` — expect it to fail with "not all trait items implemented" for both `WindowsOutlookClient` and `FakeOutlookClient`. This confirms the trait wiring; Steps 2-3 fix the two implementors.

- [ ] **Step 2: Implement `check_availability` in `client.rs`**

Add the following to the `impl OutlookClient for WindowsOutlookClient` block, right after `delete_event`'s closing `}`:

```rust
    fn check_availability(&self, input: CheckAvailabilityInput) -> Result<AvailabilityResult, ToolError> {
        self.with_com(|| {
            let (_app, ns) = mapi()?;
            let start = parse_dt(&input.start, "start")?;
            let end = parse_dt(&input.end, "end")?;
            let interval = input.interval_minutes.max(1);
            // FreeBusy has no "end" parameter — it returns a string covering a
            // fixed range from `start`. Compute how many of its slots fall
            // within [start, end) and truncate to that.
            let total_minutes = (end - start).num_minutes().max(0);
            let max_slots = ((total_minutes + interval as i64 - 1) / interval as i64) as usize;

            let mut people = Vec::new();
            for person in &input.people {
                let recipient = to_disp(call_method(
                    &ns, "CreateRecipient", &mut [variant_from_str(person)],
                )?)?;
                let resolved = variant_to_bool(&call_method(&recipient, "Resolve", &mut [])?)
                    .unwrap_or(false);
                if !resolved {
                    people.push(PersonAvailability { person: person.clone(), resolved: false, slots: Vec::new() });
                    continue;
                }
                let raw = variant_to_string(&call_method(
                    &recipient,
                    "FreeBusy",
                    &mut [
                        variant_from_datetime(&start)?,
                        variant_from_i32(interval),
                        variant_from_bool(true),
                    ],
                )?);
                let slots = parse_freebusy_slots(&raw, &start, interval, max_slots);
                people.push(PersonAvailability { person: person.clone(), resolved: true, slots });
            }
            let common = common_free(&people, &input.treat_as_free);
            Ok(AvailabilityResult { people, common_free: common })
        })
    }
```

Add `check_availability`'s new dependencies to `client.rs`'s existing `use crate::outlook::{...}` import block: add `check_availability` is the trait method (no import needed, it's `Self`), but add `CheckAvailabilityInput`, `common_free`, `parse_freebusy_slots` to the list already containing `com_recurrence_interval, create_event_status, friendly_recurrence_interval, validate_recurrence, validate_recurrence_update, CreateEventInput, EmailQuery, EmailUpdate, EventQuery, EventUpdate, OutlookClient, RecurrenceInput` — alphabetized into that existing `use` statement.

- [ ] **Step 3: Implement `check_availability` in `fake.rs`**

Open `src/outlook/fake.rs`, find `delete_event`, and add right after its closing `}`:

```rust
    fn check_availability(&self, input: CheckAvailabilityInput) -> Result<AvailabilityResult, ToolError> {
        self.record("check_availability", json!({
            "people": input.people, "start": input.start, "end": input.end,
            "interval_minutes": input.interval_minutes, "treat_as_free": input.treat_as_free,
        }))?;
        // Deterministic fake: every person resolves and is free for the
        // whole requested window (one slot spanning [start, end)), so
        // common_free tests can assert a single window without needing
        // real FreeBusy parsing in the fake.
        let people = input.people.iter().map(|p| PersonAvailability {
            person: p.clone(),
            resolved: true,
            slots: vec![AvailabilitySlot {
                start: input.start.clone(),
                end: input.end.clone(),
                status: "free".to_string(),
            }],
        }).collect::<Vec<_>>();
        let common_free = if people.is_empty() {
            Vec::new()
        } else {
            vec![FreeWindow { start: input.start.clone(), end: input.end.clone() }]
        };
        Ok(AvailabilityResult { people, common_free })
    }
```

`fake.rs` already has `use super::types::*;` at the top, so `AvailabilityResult`/`AvailabilitySlot`/`PersonAvailability`/`FreeWindow` are already in scope — no change needed there. Add `CheckAvailabilityInput` to the existing `use super::{validate_recurrence_update, CreateEventInput, EmailQuery, EmailUpdate, EventQuery, EventUpdate, OutlookClient};` line (alphabetized).

- [ ] **Step 4: Run `cargo build` to confirm both implementors compile**

Run: `cargo build`
Expected: 0 errors, 0 warnings. (This proves both trait implementations are complete and type-correct before wiring the tool layer.)

- [ ] **Step 5: Add `CheckAvailabilityParams` and the `check_availability` tool method to `server.rs`**

Find `DeleteEventParams` in `src/server.rs` (the struct right before `// ---- Attachments ----`) and add, right after it and before the attachments section:

```rust
#[derive(Debug, Deserialize, schemars::JsonSchema)]
pub struct CheckAvailabilityParams {
    pub people: Vec<String>,
    pub start: String,
    pub end: String,
    #[serde(default = "default_interval_minutes")]
    pub interval_minutes: i32,
    #[serde(default = "default_treat_as_free")]
    pub treat_as_free: Vec<String>,
}
fn default_interval_minutes() -> i32 { 30 }
fn default_treat_as_free() -> Vec<String> { vec!["free".to_string()] }
```

Then find `delete_event`'s tool method inside `#[tool_router] impl OutlookMcpServer` and add, right after its closing `}`:

```rust
    #[tool(description = "Check free/busy availability for one or more people over a time window. Returns each person's raw status per time slot (never event details) plus common_free: the windows where everyone is available. treat_as_free (default [\"free\"]) controls which statuses count as available when computing common_free; a person who can't be resolved is marked resolved:false and doesn't fail the call.")]
    pub async fn check_availability(
        &self,
        Parameters(CheckAvailabilityParams { people, start, end, interval_minutes, treat_as_free }):
            Parameters<CheckAvailabilityParams>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let input = CheckAvailabilityInput { people, start, end, interval_minutes, treat_as_free };
        let result = run_blocking(move || client.check_availability(input)).await?;
        Ok(CallToolResult::success(vec![json_content(&result)?]))
    }
```

Add `CheckAvailabilityInput` to `server.rs`'s top-level `use crate::outlook::{...}` import line (currently `CreateEventInput, EmailQuery, EmailUpdate, EventQuery, EventUpdate, OutlookClient, RecurrenceInput`).

- [ ] **Step 6: Run `cargo build` to confirm the tool layer compiles**

Run: `cargo build`
Expected: 0 errors, 0 warnings.

- [ ] **Step 7: Write fake-backed tool tests**

Open `tests/tools.rs`, find `create_event_forwards_recurrence` (or `update_event_forwards_recurrence`), and add these two tests right after it, matching this file's exact existing pattern (`#[tokio::test]`, `Arc::new(FakeOutlookClient::new())`, `fake.calls()`, the local `result_json` helper defined at the top of the file):

```rust
#[tokio::test]
async fn check_availability_forwards_params() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    server
        .check_availability(Parameters(CheckAvailabilityParams {
            people: vec!["alice@example.com".to_string(), "bob@example.com".to_string()],
            start: "2099-01-01T09:00".to_string(),
            end: "2099-01-01T17:00".to_string(),
            interval_minutes: 30,
            treat_as_free: vec!["free".to_string()],
        }))
        .await
        .unwrap();
    let (name, args) = &fake.calls()[0];
    assert_eq!(name, "check_availability");
    assert_eq!(args["people"], json!(["alice@example.com", "bob@example.com"]));
    assert_eq!(args["interval_minutes"], 30);
}

#[tokio::test]
async fn check_availability_returns_people_and_common_free() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .check_availability(Parameters(CheckAvailabilityParams {
            people: vec!["alice@example.com".to_string()],
            start: "2099-01-01T09:00".to_string(),
            end: "2099-01-01T09:30".to_string(),
            interval_minutes: 30,
            treat_as_free: vec!["free".to_string()],
        }))
        .await
        .unwrap();
    let v = result_json(&result);
    assert_eq!(v["people"][0]["person"], "alice@example.com");
    assert_eq!(v["people"][0]["resolved"], true);
    assert_eq!(v["common_free"][0]["start"], "2099-01-01T09:00");
}
```

Add `CheckAvailabilityParams` to this file's `use outlook_mcp_rs::server::{...}` import block (the one currently listing `CompleteTaskParams, CreateDraftParams, CreateEventParams, ...`).

- [ ] **Step 8: Run the new tests to verify they pass**

Run: `cargo test --test tools check_availability`
Expected: 2 tests pass.

- [ ] **Step 9: Run the full non-live suite and commit**

Run: `cargo build` (0 warnings) then `cargo test` (all lib + tool tests passing, count increased by 2 tool tests).

```bash
git add src/outlook/mod.rs src/outlook/client.rs src/outlook/fake.rs src/server.rs tests/tools.rs
git commit -m "Add check_availability: real COM (CreateRecipient/Resolve/FreeBusy), fake, and MCP tool"
```

---

### Task 3: Live test and TESTING.md

**Files:**
- Modify: `tests/live_outlook.rs`
- Modify: `TESTING.md`

**Interfaces:**
- Consumes (from Task 2): `client().check_availability(CheckAvailabilityInput { ... })`.

- [ ] **Step 1: Add the live test**

Open `tests/live_outlook.rs`, add `CheckAvailabilityInput` to its `use outlook_mcp_rs::outlook::{...}` import line, and append this test at the end of the file:

```rust
#[test]
#[ignore]
fn check_availability_against_own_mailbox_returns_free_slots() {
    let c = client();
    // Per the spec's testing strategy: check_availability is tested against
    // the developer's own mailbox where possible; the cross-user sharing
    // path (someone else's calendar) depends on another account having
    // granted access, which can't be set up from a test — see TESTING.md.
    let ns_person = std::env::var("OUTLOOK_TEST_SELF_EMAIL")
        .unwrap_or_else(|_| "adamkopelman@outlook.com".to_string());
    let result = c.check_availability(CheckAvailabilityInput {
        people: vec![ns_person.clone()],
        start: "2099-07-01T09:00".to_string(),
        end: "2099-07-01T11:00".to_string(),
        interval_minutes: 30,
        treat_as_free: vec!["free".to_string()],
    }).expect("check_availability should succeed against a resolvable address");

    assert_eq!(result.people.len(), 1);
    let person = &result.people[0];
    assert_eq!(person.person, ns_person);
    assert!(person.resolved, "self address should always resolve");
    // 2 hours / 30-minute slots = 4 slots.
    assert_eq!(person.slots.len(), 4);
    for slot in &person.slots {
        assert!(["free", "tentative", "busy", "out_of_office", "working_elsewhere"]
            .contains(&slot.status.as_str()));
    }
    // Far-future date with nothing scheduled should read back as free
    // end-to-end (proves common_free's intersection logic against a real
    // FreeBusy string, not just the fake's canned response).
    assert!(!result.common_free.is_empty());
}

#[test]
#[ignore]
fn check_availability_marks_unresolvable_person_without_failing() {
    let c = client();
    let result = c.check_availability(CheckAvailabilityInput {
        people: vec!["this-address-does-not-exist-outlook-mcp-rs-p10@nonexistent-domain-xyz.invalid".to_string()],
        start: "2099-07-01T09:00".to_string(),
        end: "2099-07-01T10:00".to_string(),
        interval_minutes: 30,
        treat_as_free: vec!["free".to_string()],
    }).expect("an unresolvable person should not fail the whole call");
    assert_eq!(result.people.len(), 1);
    assert!(!result.people[0].resolved);
    assert!(result.people[0].slots.is_empty());
    assert!(result.common_free.is_empty());
}
```

- [ ] **Step 2: Run the live tests**

Run: `cargo test --test live_outlook -- --ignored check_availability`
Expected: both tests pass against the real mailbox. If `check_availability_against_own_mailbox_returns_free_slots` fails to resolve the self address, replace the hardcoded fallback email in the test with the actual mailbox address for this environment (check `V2-RESUME.md`/earlier live-test files for the confirmed address already used elsewhere in this suite) and re-run.

- [ ] **Step 3: Update `TESTING.md`**

Open `TESTING.md`, find the section documenting manual-only / cross-user-dependent checks (e.g. `calendar_of`'s shared-calendar path, per the spec's testing strategy), and add a line noting: `check_availability`'s single-mailbox path is covered by the live suite (`cargo test --test live_outlook -- --ignored check_availability`); checking a real second person's free/busy (someone outside this mailbox) is a manual check, since Outlook must actually have published free/busy for that account, which can't be arranged from an automated test.

- [ ] **Step 4: Run the full suite one final time and commit**

Run: `cargo build` (0 warnings) then `cargo test` (all passing) then `cargo test --test live_outlook -- --ignored check_availability` (both live tests passing).

```bash
git add tests/live_outlook.rs TESTING.md
git commit -m "Add live check_availability tests (own mailbox + unresolvable person)"
```

---

## After all 3 tasks are green

Dispatch the final whole-branch review (per `superpowers:subagent-driven-development`), covering all 3 tasks together: confirm `AvailabilityResult`/`PersonAvailability`/`AvailabilitySlot`/`FreeWindow` are threaded consistently through `mod.rs` → `client.rs`/`fake.rs` → `server.rs` → tests; confirm `parse_freebusy_slots`'s `max_slots` truncation and `common_free`'s intersection logic are correctly exercised by both unit and live tests; confirm the fake's deterministic "everyone free" behavior is a reasonable, documented simplification (not silently wrong) compared to the real COM path. Then push to `main`, and update `V2-RESUME.md` / `2026-07-07-outlook-mcp-v2-plans-index.md` to mark Plan 10 shipped, following the exact pattern used for Plans 5-9.
