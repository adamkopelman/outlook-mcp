# outlook-mcp v2 feature set — design

## Motivation

The v0.1.x `outlook-mcp` / `outlook-mcp-rs` server exposes 21 tools that faithfully
mirror the original Python project. A feature-by-feature review found the surface
is functional but thin in three ways:

1. **Filtering is weak** — the `list_*` tools can barely narrow results (e.g.
   `list_emails` only filters by folder/read-status/count; `list_tasks` only by
   completed-or-not). Categories, which the user relies on heavily, can't be
   filtered on *or even seen* in output.
2. **CRUD is incomplete and asymmetric** — you can create items but mostly can't
   *edit* them after the fact (no way to flag/categorize/mark-read an email, edit
   an event, reopen a task, or edit/delete a note). Calendar, tasks, and notes
   have no delete at all.
3. **Missing real-world capabilities** — no way to attach a file when composing,
   no way to see meeting data embedded in an inbox item, and no way to view
   someone else's calendar or check availability for scheduling.

This spec defines a v2 feature set that closes those gaps while preserving the
architecture and safety properties of the current implementation.

## Scope

This is an evolution of the existing Rust implementation (`outlook-mcp-rs`), not a
rewrite. Every change extends patterns already present in `src/outlook/client.rs`,
`src/outlook/com.rs`, and `src/server.rs`. Windows-only, COM-based, distributed as
a single binary via GitHub Releases — all unchanged.

## Tool inventory: before → after

Tool count **21 → 26** (8 new tools, 3 retired = +5 net). New: `update_email`,
`update_event`, `delete_event`, `check_availability`, `update_task`, `delete_task`,
`update_note`, `delete_note`. Retired (folded into others): `search_emails`,
`move_email`, `complete_task`.

| Category | Before | After |
|---|---|---|
| **Email** | list_folders, list_emails, search_emails, get_email, send_email, create_draft, reply_email, move_email, delete_email | list_folders, **list_emails** (search merged in + filters), get_email (+ meeting/type), send_email (+ attachments), create_draft (+ attachments), reply_email (+ attachments), **update_email** (new; absorbs move + state), delete_email |
| **Calendar** | list_events, get_event, create_event, respond_to_meeting | list_events (+ filters + `calendar_of`), get_event (+ fields), create_event (+ 5 additions), **update_event** (new), respond_to_meeting, **delete_event** (new), **check_availability** (new) |
| **Attachments** | list_attachments, save_attachments | *(unchanged)* |
| **Tasks** | list_tasks, create_task, complete_task | list_tasks (+ filters), create_task (+ 3 additions), **update_task** (new; absorbs complete + edit), **delete_task** (new) |
| **Notes** | list_notes, get_note, create_note | list_notes (+ filters), get_note (+ fields), create_note (+ category/color), **update_note** (new), **delete_note** (new) |

**Retired:** `search_emails` (merged into `list_emails`), `move_email` (absorbed
into `update_email`), `complete_task` (absorbed into `update_task`).

**New tools:** `update_email`, `update_event`, `delete_event`, `check_availability`,
`update_task`, `delete_task`, `update_note`, `delete_note`.

## Cross-cutting principles

These apply across all tools and should be implemented once and reused:

1. **Categories are first-class.** Every `list_*`/`get_*` output includes a
   `categories` field (a list of the item's color-category names). Every relevant
   create/update accepts categories. Every `list_*` can filter by category. Color
   categories are currently completely invisible in the API — this makes them
   visible, filterable, and settable everywhere.

2. **Friendly words, never raw enum numbers.** Every output value that is currently
   a raw Outlook integer is converted to a lowercase string:
   - response status → `"accepted"` / `"declined"` / `"tentative"` /
     `"not_responded"` / `"organizer"`
   - busy status (`show_as`) → `"free"` / `"tentative"` / `"busy"` /
     `"out_of_office"` / `"working_elsewhere"`
   - importance → `"low"` / `"normal"` / `"high"`
   - task status → `"not_started"` / `"in_progress"` / `"complete"` /
     `"waiting"` / `"deferred"`
   Implement a small enum↔string mapping module; use it in every summary builder
   and accept the same strings as input where applicable.

3. **Composable optional filters.** Each `list_*` tool takes many optional filter
   params; supplying none returns the default listing, supplying several ANDs them
   together. Filters are applied server-side via Outlook `Restrict`/DASL where
   cheap, otherwise client-side while iterating.

4. **Safety: destructive actions stay as distinct explicit tools.** Sending
   (`send_email`) and every delete (`delete_email`/`delete_event`/`delete_task`/
   `delete_note`) remain their own named tools, never folded behind an `update`
   flag, so an AI driver cannot trigger an irreversible action by accident. All
   deletes are **soft** (to Deleted Items, recoverable).

5. **Tolerance.** All the missing-property tolerance already applied to summaries
   and detail accessors extends to every new field and tool: a property that
   doesn't exist on a given item type yields an empty/default value, never an
   error. This matters especially for the notes tools (multi-user, must be robust).

6. **Shared implementation.** Tools that differ only by a final action share one
   builder: `send_email`/`create_draft`/`reply_email` share one compose helper
   (Save vs Send only); the `update_*` tools share the property-setting helpers.
   No duplicated compose/update logic.

## Email tools

### list_folders
Unchanged.

### list_emails (absorbs `search_emails`)
`search_emails` is retired; its `query` becomes one optional filter here. Omit
`query` → plain listing; supply it → text search across subject + sender + body.

Parameters (all optional except none required; sensible defaults preserved):
- `query` — text match across subject, sender name, body (the former `search_emails`)
- `folder` (default `"inbox"`)
- `count` (default 10, max 50)
- `unread_only`
- `from` — sender name or email
- `category` — filter to a color category
- `received_after` / `received_before` — ISO date/datetime
- `since_days` — recency shortcut (kept from `search_emails`)
- `has_attachments`
- `flagged`
- `high_importance`

Output summary gains a `categories` field (was absent).

### get_email
Adds meeting awareness and item typing:
- `item_type` — `"email"` / `"meeting"` / `"bounce"` / `"read_receipt"` / other,
  derived from the item's `MessageClass`.
- `is_meeting` — boolean.
- `meeting` — present only when the item is meeting-related; built by detecting
  `GetAssociatedAppointment` (via the existing `has_member` helper) and reusing the
  existing `event_summary` logic on the associated appointment:
  ```json
  "meeting": {
    "meeting_type": "request" | "update" | "cancellation" | "response",
    "start": "...", "end": "...", "location": "...", "organizer": "...",
    "required_attendees": "...", "optional_attendees": "...", "is_recurring": false
  }
  ```
Existing tolerance for non-mail items is retained.

### send_email / create_draft / reply_email
All three gain an `attachments` parameter — a list of local file paths — attached
via `MailItem.Attachments.Add(path)` on the shared compose helper. Every path is
validated to exist **before** Save/Send, so a bad path fails fast and nothing is
half-sent. `send_email` and `create_draft` stay separate tools (send is
irreversible) but both remain thin wrappers over the one compose builder.
`reply_email` keeps its `send: true` default.

### update_email (new; absorbs `move_email`)
Single tool for all non-destructive changes to an existing email:
- `email_id` (required)
- `move_to` — destination folder (replaces standalone `move_email`)
- `mark_read` — `true`/`false` (fills the read/unread gap — no such control existed)
- `flag` — `"follow_up"` / `"complete"` / `"clear"`
- `add_categories` / `remove_categories` — lists; add/remove semantics so tagging
  doesn't wipe existing categories
- `importance` — `"low"` / `"normal"` / `"high"`

Behavior: apply state changes first, then `move_to` last (moving changes the
EntryID); return `{ "status": "updated", "id": "<current-or-new>", "changed": [...] }`.

### delete_email
Unchanged. Stays a separate soft-delete tool.

## Calendar tools

### list_events (+ filters + `calendar_of`)
Keeps `start_date` / `end_date`; adds (all optional, applied client-side during the
existing `GetFirst`/`GetNext` recurrence stream):
- `query` — text match on subject + location
- `category`
- `show_as` — `"free"`/`"tentative"`/`"busy"`/`"out_of_office"`/`"working_elsewhere"`
- `my_response` — `"organizer"`/`"accepted"`/`"declined"`/`"tentative"`/`"not_responded"`
- `attendees` — list of names/emails; match events where ANY listed person
  participates (covers "me" / "my team" [pass the list] / "my manager")
- `attendee_role` — `"required"` / `"optional"` / `"any"` (default `"any"`);
  required = meeting To, optional = meeting CC
- `meetings_only` — only events with other attendees
- `all_day`
- **`calendar_of`** — optional email/name of another person. Omit → your own default
  calendar (current behavior). Provide → resolve the recipient
  (`Namespace.CreateRecipient` + `Resolve`) and open their shared calendar
  (`GetSharedDefaultFolder(recipient, olFolderCalendar=9)`), then apply all the same
  filters. Requires that person to have shared their calendar with you; an
  unresolvable name or missing permission yields a clear error, never a crash.

Output gains `categories`, `show_as`, `my_response`, `required_attendees`,
`optional_attendees`.

### get_event
Output gains `categories`, `show_as` (friendly), `my_response` (friendly word
replacing the raw `ResponseStatus` number). `location` already present. No "flag"
field — appointments have no mail-style follow-up flag.

### create_event (+ 5 additions)
- **Optional attendees:** `required_attendees` + `optional_attendees` (two lists);
  optional attendees are added as the meeting CC tier. `attendees` kept as an alias
  for required. Any attendees → it's a meeting.
- **`categories`** — assign color categories on creation.
- **`show_as`** — busy status (default `"busy"`).
- **`recurrence`** *(heaviest lift)* — an object describing repetition:
  `pattern` (`"daily"`/`"weekly"`/`"monthly"`/`"yearly"`), `interval` (every N),
  `days_of_week` (weekly), `day_of_month` (monthly), and an end condition
  (`until` date OR `occurrences` count). Implemented via
  `AppointmentItem.GetRecurrencePattern()` and setting the `RecurrencePattern`
  fields. Higher implementation and testing effort than everything else in this
  spec.
- **`send`** — default `true` (preserves current auto-send-on-attendees behavior);
  `send: false` saves the meeting without sending invites (draft-style review).

Return status: `"meeting_sent"` (attendees + send) / `"meeting_saved"` (attendees +
no send) / `"saved"` (no attendees).

### update_event (new)
Edit an existing event; all fields optional:
- `event_id` (required)
- `subject`, `start`, `end`, `location`, `body`
- `show_as`
- `add_categories` / `remove_categories`
- `add_required_attendees` / `add_optional_attendees` / `remove_attendees`
- `reminder_minutes`, `all_day`
- `send_update` — default `true`: if the event is a meeting, changing it sends an
  update to attendees; `false` = quiet self-only change. A personal appointment
  sends nothing regardless.

Recurring events: edits apply to the **whole series** (per-occurrence edits are out
of scope for now). Returns `{ "status": "updated", "id", "changed": [...] }`.

### respond_to_meeting
Unchanged behavior; gains friendly-word status via the global principle.

### delete_event (new)
Delete/cancel an event by id:
- `event_id` (required)
- `send_cancellation` — default `true`: if you organize the meeting, deleting it
  cancels and notifies attendees; a personal appointment or one you don't organize
  is simply removed.
Soft-delete (to Deleted Items, recoverable). Separate from `update_event`
(destructive).

### check_availability (new)
Free/busy availability for scheduling. Works without full calendar sharing (as long
as the org publishes free/busy), and returns availability only — never event
details.

Parameters:
- `people` (required) — list of emails/names
- `start` (required) — ISO datetime
- `end` (required) — ISO datetime
- `interval_minutes` (optional, default 30) — slot granularity
- `treat_as_free` (optional, default `["free"]`) — which statuses count as
  available when computing `common_free`

Implementation: `Recipient.FreeBusy(start, MinPerChar=interval_minutes,
CompleteFormat=true)` returns a per-slot status code string
(0 free, 1 tentative, 2 busy, 3 out-of-office, 4 working-elsewhere), converted to
friendly words.

Returns per person plus a computed intersection:
```json
{
  "people": [
    {
      "person": "alice@company.com",
      "resolved": true,
      "slots": [
        { "start": "...", "end": "...", "status": "busy" },
        { "start": "...", "end": "...", "status": "free" }
      ]
    }
  ],
  "common_free": [
    { "start": "...", "end": "..." }
  ]
}
```
- Raw `slots` always show the **true** status — never reinterpreted.
- `common_free` = the windows where everyone is available, where "available" =
  statuses listed in `treat_as_free` (default: only `"free"`). So the caller decides
  whether tentative counts as free, whether working-elsewhere counts as free, etc.;
  out-of-office / busy are unavailable unless explicitly listed.
- One person failing to resolve (unknown, or no published free/busy) is marked
  `resolved: false` and does not fail the whole call.

## Attachments tools

`list_attachments` and `save_attachments` are unchanged.

## Tasks tools

### list_tasks (+ filters)
Keeps `include_completed`; adds:
- `category`
- `importance` — `"low"`/`"normal"`/`"high"`
- `query` — text match on subject + body
Output gains `categories`; `status`/`importance` become friendly words.

### create_task (+ 3 additions)
Keeps `subject`/`body`/`due_date`/`importance`; adds:
- `categories` — assign on creation
- `start_date` — sets `StartDate`
- `reminder_time` — ISO datetime; sets `ReminderSet` + `ReminderTime` (task
  reminders are an absolute time, unlike appointment reminders which are
  minutes-before)

### update_task (new; absorbs `complete_task`)
Edit an existing task; all fields optional:
- `task_id` (required)
- `mark_complete` — `true` (complete/100%) / `false` (reopen); fills the "can't
  reopen" gap
- `subject`, `body`, `due_date`, `start_date`
- `importance`
- `add_categories` / `remove_categories`
- `percent_complete` — 0–100
- `reminder_time` — ISO datetime
Retires standalone `complete_task` (= `update_task` with `mark_complete: true`).
Returns `{ "status": "updated", "id", "changed": [...] }`.

### delete_task (new)
Soft-delete a task by id (to Deleted Items). Separate destructive tool, for
symmetry with the other categories.

## Notes tools

Notes are kept and made robust (a multi-user requirement), with full CRUD.

### list_notes (+ filters)
Adds `category` and `query` (content/text search on the note body). Output gains
`categories`. Keeps the derived-subject behavior (first non-empty line of the body,
truncated to 120 characters).

### get_note (+ fields)
Output gains `categories` and `modified` (`LastModificationTime`) alongside existing
`created`. Full body (100k cap) and derived subject retained.

### create_note (+ category/color)
Keeps `body` (required); adds:
- `categories` — assign on creation
- `color` — Outlook note color (`"blue"`/`"green"`/`"pink"`/`"yellow"`/`"white"`,
  the `OlNoteColor` enum)

### update_note (new)
Edit an existing note: `note_id` (required), `body`, `add_categories` /
`remove_categories`, `color`. Returns `{ "status": "updated", "id", "changed": [...] }`.

### delete_note (new)
Soft-delete a note by id. Separate destructive tool.

## Testing strategy

- **Unit tests** (fake client): every new tool and every new parameter gets a
  fake-client test asserting argument forwarding, defaults, and return shape —
  extending the existing `tests/tools.rs` pattern. The friendly-word conversions
  and the `common_free` computation in `check_availability` are pure logic and get
  direct unit tests.
- **Live tests** (`tests/live_outlook.rs`, `#[ignore]`d): extend the existing live
  suite to exercise the new create/update/delete round-trips against a real
  mailbox, each cleaning up after itself. `check_availability` and `calendar_of`
  are tested against the developer's own mailbox/free-busy where possible; the
  cross-user sharing path is documented as a manual check (it depends on another
  user having granted access, which can't be set up from a test).
- Irreversible paths (`send_email` with attachments, meeting invites/cancellations,
  `update_event` sending updates) remain manual-only checks per `TESTING.md`, as
  now.

## Out of scope (for this spec)

- Per-occurrence editing of recurring events (whole-series only).
- Non-default mail stores / secondary accounts for the mail `list_*` tools (still
  default account only; `calendar_of` is the one cross-account capability, and only
  for calendars).
- Human-readable byte sizes, save-by-index, and inline base64 attachment content
  (considered and declined).
- crates.io publishing, code-signing, non-Windows support (unchanged from v0.1.x).
