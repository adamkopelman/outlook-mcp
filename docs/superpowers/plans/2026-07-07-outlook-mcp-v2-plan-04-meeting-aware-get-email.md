# v2 Plan 4 — Meeting-aware get_email Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development or superpowers:executing-plans. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Make `get_email` label every item's type and, when the item is a meeting (invite/update/cancellation/response), surface the embedded meeting details.

**Architecture:** Add `item_type: String`, `is_meeting: bool`, and `meeting: Option<MeetingInfo>` to `EmailDetail`. `item_type` derives from the item's `MessageClass` (pure mapping). When the item exposes `GetAssociatedAppointment` (detected with the existing `has_member` helper — it's a `MeetingItem`), fetch the associated appointment and build a `MeetingInfo` from its properties, with `meeting_type` derived from the message class.

**Tech Stack:** Rust 2024, `windows` 0.62.2 COM, `serde`/`schemars`.

**Depends on:** Plan 1 (Foundations). Plans 1–3 already shipped.

## Global Constraints

- Target crate: `C:\Users\adamk\projects\outlook-mcp-rs`.
- `get_email` is on the trait; both implementors (`client.rs`, `fake.rs`) return `EmailDetail`, so both change, plus tests.
- Tolerance: all new appointment property reads use `.unwrap_or_default()` (never `?`) — a malformed meeting item must not error the whole `get_email`.
- `meeting` is `#[serde(skip_serializing_if = "Option::is_none")]` — absent for normal emails.
- Commit after each task; `cargo test` green before commit. No push (controller pushes at plan end).

---

### Task 1: Types + pure class-mapping helpers

**Files:**
- Modify: `src/friendly.rs` (add `item_type_from_class` + `meeting_type_from_class`, pure, with tests)
- Modify: `src/outlook/types.rs` (add `MeetingInfo` struct; add 3 fields to `EmailDetail`)
- Modify: `src/outlook/fake.rs` (fake `get_email` returns the 3 new fields)
- Modify: `tests/tools.rs` (assert `get_email` now returns `item_type`)

**Interfaces:**
- Produces: `friendly::item_type_from_class(&str) -> &'static str`, `friendly::meeting_type_from_class(&str) -> &'static str`, `types::MeetingInfo`, and `EmailDetail.{item_type: String, is_meeting: bool, meeting: Option<MeetingInfo>}`.

- [ ] **Step 1: Add the pure mappings to `src/friendly.rs`**

```rust
/// Map an Outlook `MessageClass` to a coarse item type.
pub fn item_type_from_class(class: &str) -> &'static str {
    let c = class.to_ascii_uppercase();
    if c.starts_with("IPM.SCHEDULE.MEETING") {
        "meeting"
    } else if c.contains("NDR") || c.starts_with("REPORT.") && c.contains("NDR") {
        "bounce"
    } else if c.contains("RN") && c.starts_with("REPORT.") {
        "read_receipt"
    } else if c.starts_with("IPM.NOTE") {
        "email"
    } else {
        "other"
    }
}

/// Map a meeting-item `MessageClass` to a meeting type. Updates are delivered
/// with the same class as requests, so they map to "request".
pub fn meeting_type_from_class(class: &str) -> &'static str {
    let c = class.to_ascii_uppercase();
    if c.contains("CANCELED") || c.contains("CANCELLED") {
        "cancellation"
    } else if c.contains("RESP") {
        "response"
    } else {
        "request"
    }
}

#[cfg(test)]
mod class_tests {
    use super::*;

    #[test]
    fn item_type_mapping() {
        assert_eq!(item_type_from_class("IPM.Note"), "email");
        assert_eq!(item_type_from_class("IPM.Schedule.Meeting.Request"), "meeting");
        assert_eq!(item_type_from_class("IPM.Schedule.Meeting.Canceled"), "meeting");
        assert_eq!(item_type_from_class("REPORT.IPM.Note.NDR"), "bounce");
        assert_eq!(item_type_from_class("REPORT.IPM.Note.IPNRN"), "read_receipt");
        assert_eq!(item_type_from_class("IPM.Contact"), "other");
    }

    #[test]
    fn meeting_type_mapping() {
        assert_eq!(meeting_type_from_class("IPM.Schedule.Meeting.Request"), "request");
        assert_eq!(meeting_type_from_class("IPM.Schedule.Meeting.Canceled"), "cancellation");
        assert_eq!(meeting_type_from_class("IPM.Schedule.Meeting.Resp.Pos"), "response");
    }
}
```

- [ ] **Step 2: Run the mapping tests**

Run: `cargo test friendly::` → the new `class_tests` pass alongside the existing `tests`.

- [ ] **Step 3: Add `MeetingInfo` + fields to `src/outlook/types.rs`**

```rust
#[derive(Debug, Clone, Serialize)]
pub struct MeetingInfo {
    pub meeting_type: String,
    pub start: Option<String>,
    pub end: Option<String>,
    pub location: String,
    pub organizer: String,
    pub required_attendees: String,
    pub optional_attendees: String,
    pub is_recurring: bool,
}
```
Add to `EmailDetail` (after `attachments`):
```rust
    pub item_type: String,
    pub is_meeting: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub meeting: Option<MeetingInfo>,
```

- [ ] **Step 4: Build to find the construction sites**

Run: `cargo build` → FAIL at `client.rs::get_email` and `fake.rs::get_email` (missing fields). That's the checklist for Steps 5 + Task 2.

- [ ] **Step 5: Update the fake `get_email` in `src/outlook/fake.rs`**

Add to the `EmailDetail { .. }` the fake returns:
```rust
            item_type: "email".to_string(),
            is_meeting: false,
            meeting: None,
```

- [ ] **Step 6: Temporarily satisfy `client.rs::get_email`**

So Task 1 compiles on its own, add to the client's `EmailDetail { .. }` a minimal placeholder (real logic in Task 2):
```rust
                item_type: "email".to_string(),
                is_meeting: false,
                meeting: None,
```

- [ ] **Step 7: Add a fake-backed assertion in `tests/tools.rs`**

```rust
#[tokio::test]
async fn get_email_includes_item_type() {
    let fake = Arc::new(FakeOutlookClient::new());
    let server = OutlookMcpServer::new(fake.clone());
    let result = server
        .get_email(Parameters(GetEmailParams { email_id: EMAIL_ID.to_string(), prefer_html: false }))
        .await
        .unwrap();
    assert_eq!(result_json(&result)["item_type"], "email");
    assert_eq!(result_json(&result)["is_meeting"], false);
}
```

- [ ] **Step 8: Build + test + commit**

Run: `cargo test` → all green. Commit:
```bash
git add src/friendly.rs src/outlook/types.rs src/outlook/fake.rs src/outlook/client.rs tests/tools.rs
git commit -m "Add item_type/is_meeting/meeting fields and class-mapping helpers"
```

---

### Task 2: Real meeting detection in the Windows client

**Files:**
- Modify: `src/outlook/client.rs` (`get_email` — read MessageClass, detect + build MeetingInfo)

**Interfaces:**
- Consumes: `friendly::item_type_from_class`, `friendly::meeting_type_from_class`, `has_member`, `MeetingInfo`, existing COM helpers.

- [ ] **Step 1: Replace the placeholder fields in `client.rs::get_email` with real logic**

Just before building the final `EmailDetail`, compute the type + meeting block. Insert after `attachments` is computed:

```rust
            let message_class = variant_to_string(&get_property(&item, "MessageClass").unwrap_or_default());
            let item_type = crate::friendly::item_type_from_class(&message_class).to_string();

            // A MeetingItem exposes GetAssociatedAppointment; a plain MailItem
            // does not. Build the meeting block from the associated appointment.
            let (is_meeting, meeting) = if has_member(&item, "GetAssociatedAppointment") {
                let appt = to_disp(call_method(
                    &item, "GetAssociatedAppointment", &mut [variant_from_bool(false)],
                )?)?;
                let info = MeetingInfo {
                    meeting_type: crate::friendly::meeting_type_from_class(&message_class).to_string(),
                    start: variant_to_iso_string(&get_property(&appt, "Start").unwrap_or_default()),
                    end: variant_to_iso_string(&get_property(&appt, "End").unwrap_or_default()),
                    location: variant_to_string(&get_property(&appt, "Location").unwrap_or_default()),
                    organizer: variant_to_string(&get_property(&appt, "Organizer").unwrap_or_default()),
                    required_attendees: variant_to_string(&get_property(&appt, "RequiredAttendees").unwrap_or_default()),
                    optional_attendees: variant_to_string(&get_property(&appt, "OptionalAttendees").unwrap_or_default()),
                    is_recurring: variant_to_bool(&get_property(&appt, "IsRecurring").unwrap_or_default()).unwrap_or(false),
                };
                (true, Some(info))
            } else {
                (false, None)
            };
```
Then set the three fields in the returned `EmailDetail`:
```rust
                item_type,
                is_meeting,
                meeting,
```
Ensure `MeetingInfo` is imported (it comes via `use crate::outlook::types::*;` already in client.rs — confirm).

Note on `GetAssociatedAppointment(false)`: the boolean is `AddToCalendar` — pass `false` so merely *reading* a meeting request does NOT silently add it to the user's calendar. This is important: `true` would mutate the calendar as a side effect of a read.

- [ ] **Step 2: Build + test**

Run: `cargo build` (clean, no warnings). Run: `cargo test` (all green — fake-backed tests unaffected; this only changes the real COM path).

- [ ] **Step 3: Commit**

```bash
git add src/outlook/client.rs
git commit -m "Detect meeting items in get_email and surface meeting details"
```

---

### Task 3: Live test

**Files:**
- Modify: `tests/live_outlook.rs`

- [ ] **Step 1: Add an `#[ignore]`d test that get_email returns a valid item_type for a real inbox item**

```rust
#[test]
#[ignore]
fn get_email_reports_item_type_for_real_inbox_item() {
    use outlook_mcp_rs::outlook::EmailQuery;
    let c = WindowsOutlookClient::new();
    let list = c.list_emails(EmailQuery {
        query: None, folder: "inbox".into(), count: 1, unread_only: false,
        from: None, category: None, received_after: None, received_before: None,
        since_days: None, has_attachments: None, flagged: false, high_importance: false,
    }).expect("list");
    if let Some(first) = list.first() {
        let detail = c.get_email(first.id.clone(), false).expect("get_email");
        let v = serde_json::to_value(&detail).unwrap();
        let t = v["item_type"].as_str().unwrap();
        assert!(["email", "meeting", "bounce", "read_receipt", "other"].contains(&t));
        // If it's a meeting, the meeting block must be present.
        if v["is_meeting"].as_bool().unwrap() {
            assert!(v.get("meeting").is_some());
        }
    }
}
```

- [ ] **Step 2: Confirm compile + ignored**

Run: `cargo build --tests` (clean). Run: `cargo test 2>&1 | grep get_email_reports_item_type` → `ignored`.

- [ ] **Step 3: (If Outlook available) run live**

Run: `cargo test --test live_outlook -- --ignored get_email_reports_item_type_for_real_inbox_item` → PASS. Skip if no Outlook.

- [ ] **Step 4: Commit**

```bash
git add tests/live_outlook.rs
git commit -m "Add live get_email item_type test"
```

---

## Self-Review

- **Spec coverage:** `item_type` ✅ (from MessageClass), `is_meeting` ✅, `meeting{}` block with `meeting_type`/start/end/location/organizer/attendees/is_recurring ✅. Tolerance (all `.unwrap_or_default()`) ✅.
- **Placeholder scan:** none.
- **Type consistency:** `MeetingInfo` fields (T1) match the client build (T2); `friendly::item_type_from_class`/`meeting_type_from_class` signatures match their call sites.
- **Side-effect safety:** `GetAssociatedAppointment(false)` — reading a meeting request never adds it to the calendar.

## Execution Handoff

Plan 4 of 12. After green, controller pushes to main → Plan 5 (update_email). Models: T1 sonnet, T2 opus (COM), T3 haiku.
