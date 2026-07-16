# v2 execution — RESUME POINTER

Read this first to resume the outlook-mcp v2 build after a context/cache clear.

## What's being built
The v2 feature set (26 tools) for the **Rust** impl at `C:\Users\adamk\projects\outlook-mcp-rs`,
decomposed into 12 small plans. Spec: `docs/superpowers/specs/2026-07-07-outlook-mcp-v2-features-design.md`.
Plan map + shared conventions: `docs/superpowers/plans/2026-07-07-outlook-mcp-v2-plans-index.md`.

## How to resume (do these in order)
1. `cat C:/Users/adamk/projects/outlook-mcp-rs/.superpowers/sdd/progress.md` — the ledger; every completed task + commit SHA. Tasks marked complete are DONE; do NOT redo them.
2. `cd C:/Users/adamk/projects/outlook-mcp-rs && git log --oneline -15` — confirm what's on main. Trust the ledger + git over any recollection.
3. Read the plans index (checkboxes show which plan docs exist / are done).
4. Continue from the first unchecked item below.

## Execution method (unchanged)
- One plan at a time: write the plan doc (Plans 5–12 not yet written — regenerate from the spec + index using superpowers:writing-plans), then execute task-by-task via subagent-driven-development.
- Per-task model policy: **haiku** = pure/mechanical, **sonnet** = moderate multi-file, **opus** = complex COM (recurrence, availability, meeting detection).
- Every task is TDD (red→green→commit). Controller verifies each task (spot-check the diff / run tests), logs it to the ledger.
- After a whole plan is green with 0 warnings, `git push origin main`, then start the next plan.
- Pushing directly to `main` on adamkopelman/outlook-mcp-rs is the user's confirmed workflow (solo repo, no branch protection).
- The trait (`src/outlook/mod.rs`) has TWO implementors — `WindowsOutlookClient` (`client.rs`) and `FakeOutlookClient` (`fake.rs`) — plus `server.rs` + `tests/tools.rs`; any trait change touches all of them. Also fix `tests/live_outlook.rs` call sites (a recurring gotcha).

## Progress
- [x] Plan 1 — Foundations (shipped, main)
- [x] Plan 2 — Email finder (shipped, main)
- [x] Plan 3 — Compose attachments (shipped, main)
- [x] Plan 4 — Meeting-aware get_email (shipped, main; commits 1314d2b..ba5145e, live-verified)
- [x] Plan 5 — update_email (absorbs move_email) — shipped, main; commits 40f800a..0824ccf, live-verified (flag is manual-only: MarkAsTask rejects drafts)
- [x] Plan 6 — Calendar finder (list_events filters + calendar_of; enrich get_event) — shipped, main; commits 0824ccf..c662592, live-verified (event_matches unit-tested, 0 warnings)
- [x] Plan 7 — create_event enhancements (attendee tiers, categories, show_as, send flag; NO recurrence) — shipped, main; commits 04838f7..5abd1af, final review Ready-to-merge-Yes (1 known cross-task Minor: duplicate recipients if an address appears in overlapping tiers/alias, not fixed)
- [x] Plan 8 — update_event + delete_event (attendee add/remove, show_as/categories/reminder/all_day edits, organizer-cancel delete) — shipped, main; commits 6c4b7cc..61c65ec, final review Ready-to-merge-Yes (3 known cosmetic Minors, none fixed: changed[] empty-tier sibling dependency, fake delete_event omits subject key, one live-test send_cancellation could be false instead of true)
- [x] Plan 9 — Recurrence (create/update event via GetRecurrencePattern) — shipped, main; commits 61c65ec..7758c9c, live-verified (all 4 recurrence live tests pass: weekly, monthly+until, yearly+no-end, update+clear). 2 real bugs found and fixed in-session: Bug A (Outlook's yearly Interval must be in months, multiple of 12 — client.rs), Bug B (variant_to_iso_string couldn't decode VT_DATE, a pre-existing system-wide bug affecting every date field server-wide — com.rs). Final whole-branch review raised 3 Important findings (yearly interval live coverage, recurrence/clear_recurrence guard not shared with fake client, until+occurrences read-back symmetry) — all fixed and re-reviewed Approved.
- [x] Plan 10 — check_availability (free/busy) — shipped, main; commits 7d59667..0197da3, live-verified (own-mailbox real slots + unresolvable-person-doesn't-abort both pass). Task 3 caught and fixed a real live-COM bug in Task 2's code (FreeBusy() error handling) via a raw COM probe before it was ever pushed.
- [ ] Plan 11 — Tasks CRUD (list filters, create additions, update_task[absorbs complete_task], delete_task) — doc written (plan-11), 5 tasks — **NEXT ACTION: execute task-by-task via subagent-driven-development.**
- [ ] Plan 12 — Notes CRUD (list filters, get fields, create additions, update_note, delete_note)

## Notes captured in durable artifacts (not just chat)
- All brainstorm decisions → the committed spec.
- `friendly.rs` (enum↔word), `com::get_item_categories`/`set_item_categories`, `EmailQuery` struct, `attach_files`, `MeetingInfo` + `item_type_from_class`/`meeting_type_from_class` all exist in the code as of HEAD 502d1a3.
- `EventQuery` struct, `event_matches` client-side filter, `calendar_of` (CreateRecipient/Resolve/GetSharedDefaultFolder) all exist in the code as of HEAD c662592.
- `CreateEventInput` struct, `add_meeting_recipient`, `OL_RECIPIENT_REQUIRED`/`OL_RECIPIENT_OPTIONAL`, `create_event_status` (shared by client.rs + fake.rs) all exist in the code as of HEAD 5abd1af.
- `EventUpdate` struct, `update_event`/`delete_event` trait methods, `remove_meeting_recipients` (reverse-index recipient removal), `OL_MEETING_CANCELED` constant all exist in the code as of HEAD 61c65ec.
