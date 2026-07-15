# outlook-mcp v2 — plans index (decomposition)

The v2 feature spec (`docs/superpowers/specs/2026-07-07-outlook-mcp-v2-features-design.md`)
is decomposed into **12 small, independently-shippable plans**. Each plan produces
working, tested software on its own and can be released before the next is started.

**Target implementation:** the Rust project at `C:\Users\adamk\projects\outlook-mcp-rs`
(NOT the Python `outlook-mcp`). All plans modify that crate.

## Shared conventions (every plan follows these)

- **Two implementations per trait change.** The `OutlookClient` trait lives in
  `src/outlook/mod.rs`. Every signature change touches BOTH implementors:
  `WindowsOutlookClient` (real COM, `src/outlook/client.rs`) and `FakeOutlookClient`
  (test double, `src/outlook/fake.rs`). Forgetting either = compile error.
- **Tool layer.** MCP tools live in `src/server.rs`: a `#[derive(Debug, Deserialize,
  schemars::JsonSchema)]` params struct + a `#[tool]` async method inside the
  `#[tool_router] impl OutlookMcpServer` block, each wrapping the client call in
  `run_blocking(...)` and returning `Ok(CallToolResult::success(vec![json_content(&result)?]))`.
- **Tests.** Fake-client tool tests go in `tests/tools.rs` (call the `pub async fn`
  tool method directly with `Parameters(XxxParams{..})`, assert on `fake.calls()` /
  return shape). Live COM tests go in `tests/live_outlook.rs` (`#[ignore]`d).
- **Return types.** `list_*` return `Vec<Summary>`; `get_*` return a `Detail`;
  `create_*`/`update_*`/`delete_*` return `serde_json::Value` (a `{"status":...}`
  object).
- **COM helpers** (in `src/outlook/com.rs`, all verified against `windows` 0.62.2):
  `get_property`, `put_property`, `call_method`, `variant_from_str/i32/bool/datetime`,
  `variant_to_string/i32/bool/iso_string`, `make_item_id`, `parse_item_id`,
  `jet_datetime`, `has_member`, `safe_filename`, `ComGuard`.
- **client.rs plumbing helpers:** `with_com`, `mapi()`, `make_id()`, `get_item()`,
  `resolve_folder()`, `to_disp()`, `event_summary()`, `email_summary()`,
  `task_summary()`, `note_summary()`.
- **Safety:** destructive tools (`send_email`, all `delete_*`) stay separate; deletes
  are soft (Deleted Items).
- **Tolerance:** new property reads in summary/detail builders use
  `.unwrap_or_default()` (never `?`) so missing properties on odd item types don't error.

## The 12 plans (in dependency order)

| # | Plan | Scope | Depends on | Ship value |
|---|---|---|---|---|
| 1 | **Foundations** | `friendly.rs` enum↔word module (response/busy/importance/task-status); category read/write COM helpers; add `categories` field to every Summary/Detail type | — | Categories visible everywhere; friendly words replace raw numbers |
| 2 | **Email finder** | Merge `search_emails` into `list_emails`; add filters (`query`, `from`, `category`, `received_after/before`, `since_days`, `has_attachments`, `flagged`, `high_importance`); retire `search_emails` | 1 | One powerful email-finder tool |
| 3 | **Compose attachments** | `attachments: Vec<String>` (file paths) on `send_email`/`create_draft`/`reply_email` via shared compose helper; validate paths exist before send | — | Attach files when composing |
| 4 | **Meeting-aware get_email** | `item_type`, `is_meeting`, `meeting{}` block on `get_email` (via `has_member` + `GetAssociatedAppointment` + `event_summary`) | 1 | Inbox meeting items readable |
| 5 | **update_email** | New `update_email` tool (absorbs `move_email`): `move_to`, `mark_read`, `flag`, `add/remove_categories`, `importance`; retire `move_email` | 1 | Full email state control |
| 6 | **Calendar finder** | `list_events` filters (`query`, `category`, `show_as`, `my_response`, `attendees`, `attendee_role`, `meetings_only`, `all_day`) + `calendar_of`; enrich `get_event` + `EventSummary`/`EventDetail` output | 1 | Rich calendar querying + view others' calendars |
| 7 | **create_event enhancements** | `required_attendees`/`optional_attendees` tiers, `categories`, `show_as`, `send` flag (NO recurrence) | 1 | Meetings with optional attendees, categories, draft-mode |
| 8 | **update_event + delete_event** | New `update_event` (edit fields, `send_update` flag) and `delete_event` (`send_cancellation`, soft-delete) | 1, 6 | Full calendar CRUD |
| 9 | **Recurrence** *(heavy)* | `recurrence` object on `create_event`/`update_event` via `GetRecurrencePattern()` | 7, 8 | Repeating events |
| 10 | **check_availability** | New free/busy tool (`people`, `start`, `end`, `interval_minutes`, `treat_as_free`); `Recipient.FreeBusy`; `common_free` computation | 1 | Scheduling across people |
| 11 | **Tasks CRUD** | `list_tasks` filters (`category`, `importance`, `query`) + output; `create_task` (`categories`, `start_date`, `reminder_time`); new `update_task` (absorbs `complete_task`); new `delete_task` | 1 | Full task CRUD + filters |
| 12 | **Notes CRUD** | `list_notes` filters (`category`, `query`) + output; `get_note` (`categories`, `modified`); `create_note` (`categories`, `color`); new `update_note`; new `delete_note` | 1 | Full notes CRUD + filters + robustness |

## Ordering notes

- **Plan 1 (Foundations) is the prerequisite** for 2, 4, 5, 6, 7, 8, 10, 11, 12 —
  build it first. Plans 3 (compose attachments) is independent and could go anytime.
- Recommended sequence: **1 → (2,3,4,5 email) → (6,7,8 calendar) → 9 (recurrence) →
  10 (availability) → 11 (tasks) → 12 (notes)**.
- Each plan is small enough to implement + review + release in one focused session.
- After each plan ships, the tool count grows toward the final 26; there is never a
  broken intermediate state (retired tools are removed only in the same plan that
  adds their replacement).

## Status

- [x] Plan 1 — Foundations (`2026-07-07-outlook-mcp-v2-plan-01-foundations.md`)
- [x] Plan 2 — Email finder (2026-07-07-outlook-mcp-v2-plan-02-email-finder.md)
- [x] Plan 3 — Compose attachments (plan-03)
- [x] Plan 4 — Meeting-aware get_email (plan-04)
- [x] Plan 5 — update_email (2026-07-07-outlook-mcp-v2-plan-05-update-email.md) — shipped to main
- [x] Plan 6 — Calendar finder (2026-07-07-outlook-mcp-v2-plan-06-calendar-finder.md) — shipped to main
- [x] Plan 7 — create_event enhancements (2026-07-07-outlook-mcp-v2-plan-07-create-event-enhancements.md) — shipped to main
- [ ] Plan 8 — update_event + delete_event
- [ ] Plan 9 — Recurrence
- [ ] Plan 10 — check_availability
- [ ] Plan 11 — Tasks CRUD
- [ ] Plan 12 — Notes CRUD

(Plans are generated on request, one at a time, so each gets full bite-sized-TDD
detail rather than rushed placeholders.)
