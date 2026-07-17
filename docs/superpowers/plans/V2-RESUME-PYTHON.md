# v2 Python port — RESUME POINTER

Read this first to resume the outlook-mcp **Python** v2 feature port after a
context/cache clear. This is a parallel track to `V2-RESUME.md` (the Rust
build, already fully shipped as of 2026-07-16) — porting the same 12-plan,
26-tool v2 feature set into the original Python implementation at
`C:\Users\adamk\projects\outlook-mcp\outlook_mcp` (pywin32/COM, FastMCP).

## Why this exists

The Rust rewrite (`outlook-mcp-rs`) got the full v2 feature build (12 plans,
recurrence/check_availability/filters/CRUD everywhere) with full rigor:
plan docs, subagent-driven-development, task reviews, live testing against
a real mailbox, and a live-system-test skill. The Python implementation was
never touched — it's still at the original 21-tool v1 baseline. The user
asked (2026-07-17) to bring Python to the same 26-tool parity with the same
rigor.

## What's different from the Rust port

- **Architecture is directly analogous**, just Python idioms instead of
  Rust: `outlook_mcp/outlook/base.py` (`OutlookClientBase` abstract class)
  is the `OutlookClient` trait equivalent; `outlook_mcp/outlook/client.py`
  (`WindowsOutlookClient`, pywin32) is `client.rs`; `tests/conftest.py`'s
  `FakeOutlookClient` fixture is `fake.rs`; `outlook_mcp/server.py`
  (`@mcp.tool()` FastMCP decorators) is `server.rs`; `tests/test_tools.py`
  is `tests/tools.rs`. Every trait-change-touches-all-four-files discipline
  from the Rust build applies identically here, just with these four files.
- **One extra file with no Rust equivalent:** `outlook_mcp/outlook/base.py`
  also defines `UnavailableClient`, a non-Windows fallback (the Rust build
  only ever targets Windows). Don't break this when extending
  `OutlookClientBase`'s abstract methods.
- **COM access differs:** pywin32 (`win32com.client.Dispatch`) with a
  `@_com` decorator (COM init + `pywintypes.com_error` → `ToolError`
  mapping), vs. Rust's raw `windows` crate `IDispatch::Invoke` with a
  `with_com` closure wrapper. The *behavior* Outlook exposes is identical
  (same COM object model) — but pywin32's automatic VARIANT marshalling may
  or may not reproduce bugs the Rust build hit at the raw-`windows`-crate
  level (e.g. the VT_DATE decoding bug, the calendar-enumeration
  `Parent.StoreID` gap). **Don't assume pywin32 is immune to Outlook-side
  quirks just because it dodges the Rust-crate-level ones** — Outlook's own
  object-model behavior (yearly recurrence `Interval` in months, `FlagStatus`
  `Restrict` unreliability, etc.) is a property of Outlook, not of the
  binding layer, and will very likely still need the same fixes. Live-test
  everything; don't skip verification just because "Rust already found this
  bug and it must be a Rust-crate thing."
- **The Rust implementation is a verified reference, not just a spec.** For
  every plan, the corresponding Rust plan doc + shipped `outlook-mcp-rs`
  source is the ground truth for correct behavior (including every live-COM
  bug discovered and fixed during the Rust build — see `V2-RESUME.md` and
  each Rust plan doc's own findings). Cite exact Rust file:line references
  in Python task briefs so implementers port proven-correct logic rather
  than re-deriving it from the spec alone.

## How to resume (do these in order)

1. `cat outlook-mcp/.superpowers/sdd/progress-python.md` — the ledger for
   this track (separate from the Rust ledger, which lives in
   `outlook-mcp-rs/.superpowers/sdd/progress.md`). Tasks marked complete
   are DONE; do NOT redo them.
2. `cd outlook-mcp && git log --oneline -15` — confirm what's on `main`
   for the Python repo. Trust the ledger + git over recollection.
3. Read this doc's Progress checklist below (checkboxes show which Python
   plan docs exist / are done).
4. Continue from the first unchecked item.

## Execution method

- One plan at a time, same order as the Rust build (Plans 1–12), so each
  Python plan can build on conventions the previous one established (e.g.
  Plan 1's friendly-word module is a dependency of nearly everything after
  it, same as in Rust).
- Write each plan doc fresh (`docs/superpowers/plans/YYYY-MM-DD-outlook-mcp-v2-python-planNN-<name>.md`),
  adapting the matching Rust plan doc's task decomposition and citing exact
  Rust source for the proven-correct logic, but writing genuinely
  Python-idiomatic code (not a transliteration) and using pywin32
  conventions already established in this codebase (`@_com`,
  `getattr(item, "X", default) or default`, dict-shaped returns, `ToolError`).
- Execute via `superpowers:subagent-driven-development` (fresh subagent
  per task + task review + final whole-branch review), exactly like Rust.
- Per-task model policy: same as Rust — haiku=mechanical/complete-spec-given,
  sonnet=moderate multi-file, opus=complex COM/design judgment. The final
  whole-branch review for each plan always gets the most capable model.
- Live test every plan's live-COM-touching behavior against the real
  mailbox before considering it shipped — pywin32 will very likely have its
  own surprises even where Rust already solved the underlying Outlook
  behavior. Use the same discipline as
  `outlook-mcp-rs/.claude/skills/live-outlook-system-test/SKILL.md` (that
  skill is Rust-file-path-specific; port its *principles*, not its literal
  paths, when testing Python — a Python-specific version of it can be
  written once this track has enough of its own live-testing experience to
  be worth codifying, same as happened for Rust after Plan 9).
- Push directly to `main` on the Python repo when each plan is green,
  matching the Rust repo's confirmed workflow (solo repo, no branch
  protection) — confirm this is still the user's intent for THIS repo
  before the first push, since it wasn't explicitly re-confirmed for the
  Python track.
- Update this file's Progress checklist after each plan ships.

## Progress

- [x] Plan 1 — Foundations (friendly words + categories) — shipped, main; commits 0c53f34..5ec2a0a, live-verified (categories round-trip through real Outlook via pywin32, passed first try, no VARIANT/BSTR-level quirks for this property). Established the live-test pattern (`tests/test_live.py`, `OUTLOOK_MCP_LIVE_TESTS=1` opt-in) all later plans extend.
- [x] Plan 2 — Email finder — shipped, master; commits 0bc09e2..c1e62df. search_emails fully retired, merged into list_emails with full filter parity (query/sender/category/dates/has_attachments/flagged/high_importance) vs. Rust — Restrict-based for cheap filters, client-side for category/has_attachments. Live-verified: query filter narrows results against the real mailbox.
- [x] Plan 3 — Compose attachments — shipped, master; commits 3051f51..72a768e. attachments param (local file paths) threaded through send_email/create_draft/reply_email via a new module-level _attach_files helper (validate-all-then-attach-all, fail-fast before any send/save). Live-verified: real draft round-trip with an attached temp file, plus a fail-fast missing-path case, both against the real mailbox.
- [x] Plan 4 — Meeting-aware get_email — shipped, master; commits e806411..fca7306. get_email gains item_type/is_meeting/meeting fields. Deliberate mechanism divergence from Rust: is_meeting derived from MessageClass string (not a has_member-style probe, which pywin32 dynamic dispatch can't do reliably), GetAssociatedAppointment(False) called only then, wrapped in try/except so a malformed meeting item degrades to is_meeting=True/meeting absent rather than failing get_email. Live-verified against a real inbox item. Known minor test-robustness gap (documented, not fixed): the live test's meeting-presence assertion is slightly stronger than the code's own tolerance guarantee; low-probability, live-only, does not affect shipped behavior.
- [ ] Plan 5 — update_email (absorbs move_email)
- [ ] Plan 6 — Calendar finder
- [ ] Plan 7 — create_event enhancements
- [ ] Plan 8 — update_event + delete_event
- [ ] Plan 9 — Recurrence
- [ ] Plan 10 — check_availability
- [ ] Plan 11 — Tasks CRUD
- [ ] Plan 12 — Notes CRUD

## Reference: Rust plan docs (source of truth for behavior)

All at `docs/superpowers/plans/2026-07-07-outlook-mcp-v2-plan-NN-*.md` in
this same repo. The shipped Rust source lives at
`C:\Users\adamk\projects\outlook-mcp-rs` — `git log --oneline` there shows
every commit; each plan's exact shipped commit range is recorded in
`V2-RESUME.md`'s Progress section and in `outlook-mcp-rs/.superpowers/sdd/progress.md`.
