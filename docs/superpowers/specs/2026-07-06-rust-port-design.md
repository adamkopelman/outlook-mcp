# outlook-mcp-rs: Rust port design

## Motivation

The existing `outlook-mcp` project (Python, package `outlook-mcp-com`) is an MCP
server that drives classic Outlook desktop via Win32 COM automation. Installing
it requires Python + pip. The goal of this port is a single self-contained
Windows `.exe` that users can download and run with no Python, pip, or Rust
toolchain required — distributed via GitHub Releases.

## Scope

Full parity with the existing Python tool surface: every tool in
`outlook_mcp/server.py` gets a 1:1 Rust equivalent — same tool names, same
parameters, same JSON return shapes, so existing MCP client configs/prompts
work unmodified against either server.

Tool groups (24 tools total, mirroring `outlook_mcp/outlook/client.py`):

- **Email**: list_folders, list_emails, search_emails, get_email, send_email,
  create_draft, reply_email, move_email, delete_email
- **Calendar**: list_events, get_event, create_event, respond_to_meeting
- **Attachments**: list_attachments, save_attachments
- **Tasks**: list_tasks, create_task, complete_task
- **Notes**: list_notes, get_note, create_note

## Repository

New, separate repo: `adamkopelman/outlook-mcp-rs`. Kept separate from the
Python `outlook-mcp` repo (independent versioning/release cadence, since it's
a distinct distribution artifact), MIT licensed to match, with a README that
describes it as the single-binary counterpart to `outlook-mcp` and links back
to it.

## Architecture

Single binary crate (no workspace needed at this scope):

```
outlook-mcp-rs/
  Cargo.toml
  src/
    main.rs          # entry point, wires rmcp server + tool registry
    error.rs          # ToolError type + HRESULT/COM error -> readable message
    constants.rs      # Outlook enum values (OlDefaultFolders, OlImportance, etc.)
    outlook/
      mod.rs
      client.rs       # COM automation (windows-rs), mirrors client.py 1:1
    tools/
      mod.rs
      email.rs
      calendar.rs
      attachments.rs
      tasks.rs
      notes.rs
  tests/
    tools.rs          # trait-based fake client, mirrors test_tools.py
    live_outlook.rs    # #[ignore]d system tests against a real, running Outlook
  TESTING.md
```

### COM interop

`windows-rs` (Microsoft's official crate), late-bound `IDispatch` calls against
`Outlook.Application` / its `Namespace`, the same automation approach as
`win32com` today (typed through `windows-rs`'s `VARIANT` /
`IDispatch::Invoke` rather than Python's dynamic dispatch). Item addressing
keeps the same opaque `"{EntryID}|{StoreID}"` id scheme used by the Python
version, so ids returned by list/search tools behave identically.

### Error handling

A `ToolError(String)` type, with a `From<windows::core::Error>` conversion
that formats the HRESULT + message the same way `format_com_error` does in
Python today — COM failures surface as clean tool errors, not panics.

### MCP layer

The official `rmcp` crate (Rust MCP SDK). Handles the stdio transport,
JSON-RPC framing, and tool JSON-schema generation from typed Rust function
signatures — the direct equivalent of the `@mcp.tool()` decorator in the
Python `mcp` SDK. Chosen over hand-rolling the protocol because the
differentiated work here is the COM/Windows layer, not protocol plumbing.

No crates.io publish for now — GitHub Releases only. Easy to add later as an
independent step once the binary distribution is proven out.

## Testing strategy

**Unit tests** (run in CI on every commit): an `OutlookClient` trait mirroring
`OutlookClientBase`, with a `FakeOutlookClient` mock — a direct port of
`conftest.py`'s fake client, one test per tool covering both success and error
paths. Runs on `windows-latest` in CI.

**Live system tests** (local-only; new — the Python version has no equivalent
today): `tests/live_outlook.rs` exercises the real `WindowsOutlookClient`
against whatever Outlook is actually running on the developer's machine.

- `#[ignore]`d by default — plain `cargo test` never touches live Outlook.
  Run explicitly with `cargo test --test live_outlook -- --ignored`.
- Every test that creates something (draft, task, note, calendar event)
  deletes it at the end, so the mailbox is unchanged after a run.
- `send_email` and `respond_to_meeting` (real send/accept side-effects) are
  excluded from the automatic suite, since a sent email can't be safely
  undone. These get a separate, clearly-documented manual test run by hand
  with a designated test recipient.
- `TESTING.md` documents exactly how to run the live suite and its
  preconditions (Outlook open, signed in, a normal mailbox).

## CI/CD & release

GitHub Actions workflow, modeled on the existing Python project's:

- **`test` job**: `windows-latest`, `cargo test` (unit/fake-client tests
  only — live tests never run in CI; runners have no Outlook installed).
- **`build` job** (tag push `v*` only): `windows-latest`,
  `cargo build --release`, produces `outlook-mcp-rs.exe`.
- **`release` job**: creates a GitHub Release for the tag and attaches the
  `.exe` as a release asset.

## Out of scope (for this design)

- crates.io publishing (may be added later as a follow-up, independent step)
- Code-signing the `.exe` (Windows SmartScreen may warn on first run;
  revisit if it becomes a problem for users)
- Any macOS/Linux support — this is a Windows-only binary, same constraint
  as the Python version's actual COM functionality
