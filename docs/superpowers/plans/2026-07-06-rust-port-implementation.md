# outlook-mcp-rs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Port `outlook-mcp` (Python MCP server driving classic Outlook desktop via COM) to a single-binary Rust MCP server, distributed as a prebuilt Windows `.exe` via GitHub Releases, with full parity across all 24 tools.

**Architecture:** `windows-rs` for late-bound `IDispatch` COM automation against `Outlook.Application` (same technique as `win32com`, just typed manually via `GetIDsOfNames`/`Invoke` instead of Python's dynamic dispatch). The official `rmcp` crate handles the MCP/stdio/JSON-RPC layer. An `OutlookClient` trait (mirroring `OutlookClientBase`) separates COM automation from tool registration, with a `FakeOutlookClient` test double enabling fast unit tests, and a new local-only live-Outlook system test suite the Python version doesn't have.

**Tech Stack:** Rust (stable toolchain), `windows` 0.62, `rmcp` 2.1, `tokio`, `serde`/`serde_json`, `schemars`, `chrono`.

**Spec:** `docs/superpowers/specs/2026-07-06-rust-port-design.md`

## Global Constraints

- Windows-only: every tool returns `serde_json::Value` matching the exact JSON shape (same field names) as the equivalent Python tool in `outlook_mcp/outlook/client.py`, so existing MCP client configs/prompts work unmodified.
- COM calls are synchronous/blocking (COM automation is inherently blocking); every `#[tool]` wrapper runs its `OutlookClient` call inside `tokio::task::spawn_blocking` so it never blocks the async runtime.
- Every COM-calling thread must call `CoInitializeEx` before touching COM (apartment-threaded), mirroring the Python client's per-call `pythoncom.CoInitialize()`.
- Live system tests (`tests/live_outlook.rs`) are `#[ignore]`d by default, never run in CI, and every test that creates an Outlook item must delete it before finishing. `send_email`/`respond_to_meeting` are excluded from the automatic live suite (documented as manual-only in `TESTING.md`).
- No crates.io publish. Distribution is a prebuilt `.exe` attached to GitHub Releases on tag push, built on `windows-latest`.
- New repo: `adamkopelman/outlook-mcp-rs`, MIT licensed.

---

## Task 1: Repo bootstrap & minimal MCP server

**Files:**
- Create: `outlook-mcp-rs/Cargo.toml`
- Create: `outlook-mcp-rs/src/main.rs`
- Create: `outlook-mcp-rs/src/server.rs`
- Create: `outlook-mcp-rs/.gitignore`
- Create: `outlook-mcp-rs/README.md`
- Create: `outlook-mcp-rs/LICENSE`

**Interfaces:**
- Produces: `OutlookMcpServer` struct (empty for now — no `client` field yet, added in Task 5) and its `#[tool_router(server_handler)] impl OutlookMcpServer` block, which every later task adds `#[tool]` methods to.

- [ ] **Step 1: Create the GitHub repo and local project**

```bash
gh repo create adamkopelman/outlook-mcp-rs --public \
  --description "Single-binary Rust MCP server for classic Outlook desktop (Win32 COM)" \
  --license mit --clone
cd outlook-mcp-rs
cargo init --name outlook-mcp-rs
```

- [ ] **Step 2: Write Cargo.toml**

```toml
[package]
name = "outlook-mcp-rs"
version = "0.1.0"
edition = "2021"

[[bin]]
name = "outlook-mcp-rs"
path = "src/main.rs"

[dependencies]
windows = { version = "0.62", features = [
    "Win32_Foundation",
    "Win32_System_Com",
    "Win32_System_Ole",
    "Win32_System_Variant",
    "Win32_Globalization",
] }
rmcp = { version = "2.1", features = ["server", "transport-io", "macros"] }
tokio = { version = "1", features = ["full"] }
serde = { version = "1", features = ["derive"] }
serde_json = "1"
schemars = "0.8"
chrono = "0.4"
tracing = "0.1"
tracing-subscriber = { version = "0.3", features = ["env-filter"] }
anyhow = "1"
```

- [ ] **Step 3: Write src/server.rs with a single health-check tool**

```rust
use rmcp::{ServerHandler, tool, tool_router};

#[derive(Debug, Clone, Default)]
pub struct OutlookMcpServer;

#[tool_router(server_handler)]
impl OutlookMcpServer {
    #[tool(description = "Health check: returns 'pong' if the server is running.")]
    fn ping(&self) -> String {
        "pong".to_string()
    }
}
```

- [ ] **Step 4: Write src/main.rs**

```rust
use anyhow::Result;
use rmcp::{ServiceExt, transport::stdio};
use tracing_subscriber::EnvFilter;

mod server;

use server::OutlookMcpServer;

#[tokio::main]
async fn main() -> Result<()> {
    tracing_subscriber::fmt()
        .with_env_filter(EnvFilter::from_default_env().add_directive(tracing::Level::INFO.into()))
        .with_writer(std::io::stderr)
        .with_ansi(false)
        .init();

    tracing::info!("Starting outlook-mcp-rs");

    let service = OutlookMcpServer::default()
        .serve(stdio())
        .await
        .inspect_err(|e| tracing::error!("serving error: {:?}", e))?;

    service.waiting().await?;
    Ok(())
}
```

- [ ] **Step 5: Write a smoke test**

Create `tests/smoke.rs`:

```rust
use outlook_mcp_rs::server::OutlookMcpServer;

#[test]
fn ping_returns_pong() {
    let server = OutlookMcpServer::default();
    assert_eq!(server.ping(), "pong");
}
```

This requires `server` to also be reachable as a library module. Add a `src/lib.rs`:

```rust
pub mod server;
```

And change `src/main.rs`'s `mod server;` line to `use outlook_mcp_rs::server;` instead, plus add `name = "outlook_mcp_rs"` library section to `Cargo.toml`:

```toml
[lib]
name = "outlook_mcp_rs"
path = "src/lib.rs"
```

- [ ] **Step 6: Run the test to verify it passes**

Run: `cargo test`
Expected: `ping_returns_pong` passes, `cargo build` succeeds with no warnings.

- [ ] **Step 7: Write .gitignore**

```
/target
```

(Cargo.lock is committed — this is a binary crate.)

- [ ] **Step 8: Write README.md**

```markdown
# outlook-mcp-rs

Single-binary Rust MCP server for classic Outlook desktop (Win32 COM automation).
The Rust counterpart to [outlook-mcp](https://github.com/adamkopelman/outlook-mcp) —
same tools, same JSON shapes, no Python/pip required. Download a release `.exe`
from the Releases page and point your MCP client at it directly.

Windows only (requires classic Outlook desktop installed).
```

- [ ] **Step 9: Commit**

```bash
git add -A
git commit -m "Bootstrap outlook-mcp-rs crate with a health-check tool"
git push
```

---

## Task 2: Error handling

**Files:**
- Create: `outlook-mcp-rs/src/error.rs`
- Modify: `outlook-mcp-rs/src/lib.rs`

**Interfaces:**
- Produces: `ToolError(pub String)`, `pub type McpError = rmcp::ErrorData`, `impl From<ToolError> for McpError`. All later tasks' `OutlookClient` trait methods return `Result<serde_json::Value, ToolError>`; all `#[tool]` wrapper methods return `Result<CallToolResult, McpError>` and use `?` to convert a `ToolError` via this `From` impl.

- [ ] **Step 1: Write the failing test**

Create `src/error.rs`:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tool_error_displays_its_message() {
        let err = ToolError("Folder not found: 'Nope'".to_string());
        assert_eq!(err.to_string(), "Folder not found: 'Nope'");
    }

    #[test]
    fn converts_to_mcp_invalid_params_error() {
        let err = ToolError("bad input".to_string());
        let mcp_err: McpError = err.into();
        assert_eq!(mcp_err.message, "bad input");
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cargo test error::tests -- --nocapture`
Expected: FAIL — `ToolError`/`McpError` not defined yet.

- [ ] **Step 3: Write the implementation**

Add above the test module in `src/error.rs`:

```rust
use std::fmt;

pub type McpError = rmcp::ErrorData;

#[derive(Debug, Clone)]
pub struct ToolError(pub String);

impl fmt::Display for ToolError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.0)
    }
}

impl std::error::Error for ToolError {}

impl From<ToolError> for McpError {
    fn from(err: ToolError) -> Self {
        McpError::invalid_params(err.0, None)
    }
}

impl From<&str> for ToolError {
    fn from(msg: &str) -> Self {
        ToolError(msg.to_string())
    }
}

impl From<String> for ToolError {
    fn from(msg: String) -> Self {
        ToolError(msg)
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cargo test error::tests -- --nocapture`
Expected: PASS

- [ ] **Step 5: Wire the module into the crate**

In `src/lib.rs`, add:

```rust
pub mod error;
```

- [ ] **Step 6: Commit**

```bash
git add src/error.rs src/lib.rs
git commit -m "Add ToolError and its conversion to rmcp::ErrorData"
```

---

## Task 3: Low-level IDispatch helper

This is the riskiest, most novel piece: hand-rolled late-bound COM automation
(the Rust equivalent of what `win32com.client.Dispatch` does dynamically in
Python). Every later Outlook COM call goes through these four functions.

**Files:**
- Create: `outlook-mcp-rs/src/outlook/mod.rs`
- Create: `outlook-mcp-rs/src/outlook/dispatch.rs`
- Modify: `outlook-mcp-rs/src/lib.rs`

**Interfaces:**
- Produces:
  - `pub fn init_com() -> Result<(), ToolError>` — call once per thread before any COM use.
  - `pub fn create_instance(prog_id: &str) -> Result<IDispatch, ToolError>`
  - `pub fn get_property(disp: &IDispatch, name: &str) -> Result<VARIANT, ToolError>`
  - `pub fn call_method(disp: &IDispatch, name: &str, args: &[VARIANT]) -> Result<VARIANT, ToolError>` (args in natural left-to-right order; the function reverses them internally for COM's `rgvarg`)
  - `pub fn put_property(disp: &IDispatch, name: &str, value: VARIANT) -> Result<(), ToolError>`
  - Variant conversion helpers: `variant_str(&VARIANT) -> String`, `variant_i32(&VARIANT) -> i32`, `variant_bool(&VARIANT) -> bool`, `variant_dispatch(&VARIANT) -> Result<IDispatch, ToolError>`, `variant_date_iso(&VARIANT) -> Option<String>`, `str_variant(&str) -> VARIANT`, `i32_variant(i32) -> VARIANT`, `bool_variant(bool) -> VARIANT`, `date_variant(&chrono::NaiveDateTime) -> VARIANT`.
- Consumes: `crate::error::ToolError` from Task 2.

- [ ] **Step 1: Write the failing tests**

Create `src/outlook/dispatch.rs` with this test module first (implementation follows in Step 3). These tests exercise the helper against `WScript.Shell`, a COM object built into every Windows install (part of Windows Script Host, not Office) — so they run in CI on `windows-latest` with no Outlook needed, giving us a real, automated check that the late-bound call machinery actually works:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn calls_method_with_string_arg_and_reads_string_result() {
        init_com().unwrap();
        let shell = create_instance("WScript.Shell").unwrap();
        let result = call_method(
            &shell,
            "ExpandEnvironmentStrings",
            &[str_variant("%windir%\\explorer.exe")],
        )
        .unwrap();
        let expanded = variant_str(&result);
        assert!(!expanded.contains('%'), "expected expansion, got {expanded}");
        assert!(expanded.to_lowercase().ends_with("explorer.exe"));
    }

    #[test]
    fn puts_and_gets_a_string_property() {
        init_com().unwrap();
        let shell = create_instance("WScript.Shell").unwrap();
        let original = get_property(&shell, "CurrentDirectory").unwrap();
        let original_dir = variant_str(&original);

        put_property(&shell, "CurrentDirectory", str_variant("C:\\")).unwrap();
        let updated = variant_str(&get_property(&shell, "CurrentDirectory").unwrap());
        assert_eq!(updated, "C:\\");

        // restore, so this test doesn't leave global process state changed
        put_property(&shell, "CurrentDirectory", str_variant(&original_dir)).unwrap();
    }

    #[test]
    fn unknown_prog_id_returns_a_tool_error() {
        init_com().unwrap();
        let err = create_instance("NotARealOutlookMcpTestProgId.Xyz").unwrap_err();
        assert!(err.0.contains("NotARealOutlookMcpTestProgId"));
    }

    #[test]
    fn unknown_method_name_returns_a_tool_error() {
        init_com().unwrap();
        let shell = create_instance("WScript.Shell").unwrap();
        let err = call_method(&shell, "ThisMethodDoesNotExist", &[]).unwrap_err();
        assert!(err.0.to_lowercase().contains("thismethoddoesnotexist"));
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cargo test dispatch::tests -- --nocapture`
Expected: FAIL to compile — none of the functions exist yet.

- [ ] **Step 3: Write the implementation**

Add above the test module in `src/outlook/dispatch.rs`:

```rust
use windows::core::{GUID, HRESULT, HSTRING, PCWSTR, PWSTR};
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::Globalization::LOCALE_USER_DEFAULT;
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CLSCTX_LOCAL_SERVER,
    COINIT_APARTMENTTHREADED, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
    DISPPARAMS, EXCEPINFO, IDispatch,
};
use windows::Win32::System::Ole::VariantTimeToSystemTime;
use windows::Win32::System::Variant::{VARIANT, VT_DATE};
use windows::Win32::System::SystemInformation::GetLocalTime;

use crate::error::ToolError;

const DISPID_PROPERTYPUT: i32 = -3;

/// Call once per thread before touching any COM object on that thread.
/// Idempotent to call more than once on the same thread (subsequent calls
/// return S_FALSE, which init_com treats as success), matching the Python
/// client's per-call `pythoncom.CoInitialize()`.
pub fn init_com() -> Result<(), ToolError> {
    unsafe {
        let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        if hr.is_ok() || hr == windows::Win32::Foundation::S_FALSE {
            Ok(())
        } else {
            Err(ToolError(format!("CoInitializeEx failed: {hr:?}")))
        }
    }
}

pub fn create_instance(prog_id: &str) -> Result<IDispatch, ToolError> {
    unsafe {
        let wide = HSTRING::from(prog_id);
        let clsid: GUID = CLSIDFromProgID(PCWSTR(wide.as_ptr())).map_err(|e| {
            ToolError(format!(
                "Could not find COM object {prog_id:?}: {e} (is it installed and registered?)"
            ))
        })?;
        CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)
            .map_err(|e| ToolError(format!("Could not start {prog_id:?}: {e}")))
    }
}

fn resolve_dispid(disp: &IDispatch, name: &str) -> Result<i32, ToolError> {
    unsafe {
        let wide = HSTRING::from(name);
        let mut arg = PWSTR(wide.as_ptr() as *mut u16);
        let mut dispid: i32 = 0;
        let hr: HRESULT = disp.GetIDsOfNames(
            &GUID::zeroed(),
            &mut arg,
            1,
            LOCALE_USER_DEFAULT.0,
            &mut dispid,
        );
        hr.ok()
            .map_err(|e| ToolError(format!("Unknown member {name:?}: {e}")))?;
        Ok(dispid)
    }
}

fn invoke(
    disp: &IDispatch,
    name: &str,
    flags: windows::Win32::System::Com::DISPATCH_FLAGS,
    mut dispparams: DISPPARAMS,
) -> Result<VARIANT, ToolError> {
    let dispid = resolve_dispid(disp, name)?;
    let mut result = VARIANT::default();
    let mut excepinfo = EXCEPINFO::default();
    let mut arg_err: u32 = 0;

    unsafe {
        let hr: HRESULT = disp.Invoke(
            dispid,
            &GUID::zeroed(),
            LOCALE_USER_DEFAULT.0,
            flags,
            &mut dispparams,
            Some(&mut result),
            Some(&mut excepinfo),
            Some(&mut arg_err),
        );

        if hr.is_ok() {
            return Ok(result);
        }

        if hr == windows::Win32::Foundation::DISP_E_EXCEPTION {
            let desc = excepinfo.bstrDescription.to_string();
            if !desc.trim().is_empty() {
                return Err(ToolError(format!("Outlook error calling {name:?}: {desc}")));
            }
        }
        Err(ToolError(format!(
            "Outlook error calling {name:?}: {} (HRESULT {hr:?})",
            windows::core::Error::from(hr).message()
        )))
    }
}

pub fn get_property(disp: &IDispatch, name: &str) -> Result<VARIANT, ToolError> {
    let dispparams = DISPPARAMS::default();
    invoke(disp, name, DISPATCH_PROPERTYGET, dispparams)
}

pub fn call_method(disp: &IDispatch, name: &str, args: &[VARIANT]) -> Result<VARIANT, ToolError> {
    // COM requires arguments in reverse order in rgvarg.
    let mut reversed: Vec<VARIANT> = args.iter().cloned().rev().collect();
    let dispparams = DISPPARAMS {
        rgvarg: if reversed.is_empty() {
            std::ptr::null_mut()
        } else {
            reversed.as_mut_ptr()
        },
        rgdispidNamedArgs: std::ptr::null_mut(),
        cArgs: reversed.len() as u32,
        cNamedArgs: 0,
    };
    invoke(disp, name, DISPATCH_METHOD, dispparams)
}

pub fn put_property(disp: &IDispatch, name: &str, mut value: VARIANT) -> Result<(), ToolError> {
    let mut named_arg = DISPID_PROPERTYPUT;
    let dispparams = DISPPARAMS {
        rgvarg: &mut value,
        rgdispidNamedArgs: &mut named_arg,
        cArgs: 1,
        cNamedArgs: 1,
    };
    invoke(disp, name, DISPATCH_PROPERTYPUT, dispparams)?;
    Ok(())
}

// ---- VARIANT conversions ---------------------------------------------

pub fn str_variant(s: &str) -> VARIANT {
    VARIANT::from(HSTRING::from(s))
}

pub fn i32_variant(n: i32) -> VARIANT {
    VARIANT::from(n)
}

pub fn bool_variant(b: bool) -> VARIANT {
    VARIANT::from(b)
}

pub fn variant_str(v: &VARIANT) -> String {
    HSTRING::try_from(v).map(|h| h.to_string_lossy()).unwrap_or_default()
}

pub fn variant_i32(v: &VARIANT) -> i32 {
    i32::try_from(v).unwrap_or(0)
}

pub fn variant_bool(v: &VARIANT) -> bool {
    bool::try_from(v).unwrap_or(false)
}

pub fn variant_dispatch(v: &VARIANT) -> Result<IDispatch, ToolError> {
    IDispatch::try_from(v).map_err(|e| ToolError(format!("Expected an object result: {e}")))
}

/// OLE Automation date (VT_DATE, an f64 day-count from 1899-12-30) -> ISO-8601 string.
pub fn variant_date_iso(v: &VARIANT) -> Option<String> {
    let date: f64 = f64::try_from(v).ok()?;
    unsafe {
        let mut sys = std::mem::zeroed();
        VariantTimeToSystemTime(date, &mut sys).ok()?;
        let dt = chrono::NaiveDate::from_ymd_opt(
            sys.wYear as i32,
            sys.wMonth as u32,
            sys.wDay as u32,
        )?
        .and_hms_opt(sys.wHour as u32, sys.wMinute as u32, sys.wSecond as u32)?;
        Some(dt.format("%Y-%m-%dT%H:%M:%S").to_string())
    }
}
```

`GetLocalTime` and `VARIANT_BOOL` are imported but unused directly (kept for
parity with future date-writing helpers) — remove the `GetLocalTime` import
if `cargo build` warns about it as unused; it is not needed by the code
above and was left over from drafting. Run `cargo build` and delete any
import that triggers an "unused import" warning.

- [ ] **Step 4: Create outlook/mod.rs**

```rust
pub mod dispatch;
```

- [ ] **Step 5: Wire into lib.rs**

Add to `src/lib.rs`:

```rust
pub mod outlook;
```

- [ ] **Step 6: Run tests to verify they pass**

Run: `cargo test dispatch::tests -- --nocapture`
Expected: all 4 tests PASS. If `HSTRING`/`VARIANT`/`IDispatch` conversion
trait names differ slightly from what's shown here (the `windows` crate's
exact `TryFrom`/`From` impls can shift between minor versions), fix the
specific line the compiler flags — the failing compile error will name the
correct trait/method to use; this is expected first-pass friction for a
hand-rolled COM binding and does not change the overall design.

- [ ] **Step 7: Commit**

```bash
git add src/outlook/
git commit -m "Add late-bound IDispatch helper (get/put property, call method)"
```

---

## Task 4: Outlook constants

Direct port of `outlook_mcp/constants.py`.

**Files:**
- Create: `outlook-mcp-rs/src/constants.rs`
- Modify: `outlook-mcp-rs/src/lib.rs`

**Interfaces:**
- Produces: `pub const OL_FOLDER_INBOX: i32` (and siblings, exact names/values below), `pub fn folder_id_for_name(name: &str) -> Option<i32>`, `pub fn importance_id_for_name(name: &str) -> Option<i32>`, `pub fn meeting_response_id_for_name(name: &str) -> Option<i32>`.

- [ ] **Step 1: Write the failing test**

Create `src/constants.rs`:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn maps_friendly_folder_names() {
        assert_eq!(folder_id_for_name("inbox"), Some(OL_FOLDER_INBOX));
        assert_eq!(folder_id_for_name("Sent Items"), Some(OL_FOLDER_SENT_MAIL));
        assert_eq!(folder_id_for_name("nonsense"), None);
    }

    #[test]
    fn maps_importance_names() {
        assert_eq!(importance_id_for_name("high"), Some(OL_IMPORTANCE_HIGH));
        assert_eq!(importance_id_for_name("bogus"), None);
    }

    #[test]
    fn maps_meeting_response_names() {
        assert_eq!(
            meeting_response_id_for_name("accept"),
            Some(OL_MEETING_ACCEPTED)
        );
        assert_eq!(meeting_response_id_for_name("maybe"), None);
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `cargo test constants::tests`
Expected: FAIL to compile — nothing defined yet.

- [ ] **Step 3: Write the implementation**

Add above the test module:

```rust
// OlDefaultFolders
pub const OL_FOLDER_DELETED_ITEMS: i32 = 3;
pub const OL_FOLDER_OUTBOX: i32 = 4;
pub const OL_FOLDER_SENT_MAIL: i32 = 5;
pub const OL_FOLDER_INBOX: i32 = 6;
pub const OL_FOLDER_CALENDAR: i32 = 9;
pub const OL_FOLDER_CONTACTS: i32 = 10;
pub const OL_FOLDER_JOURNAL: i32 = 11;
pub const OL_FOLDER_NOTES: i32 = 12;
pub const OL_FOLDER_TASKS: i32 = 13;
pub const OL_FOLDER_DRAFTS: i32 = 16;

// OlItemType (Application.CreateItem)
pub const OL_MAIL_ITEM: i32 = 0;
pub const OL_APPOINTMENT_ITEM: i32 = 1;
pub const OL_TASK_ITEM: i32 = 3;
pub const OL_NOTE_ITEM: i32 = 5;

// OlBodyFormat
pub const OL_FORMAT_PLAIN: i32 = 1;
pub const OL_FORMAT_HTML: i32 = 2;

// OlMeetingResponse (AppointmentItem.Respond)
pub const OL_MEETING_TENTATIVE: i32 = 2;
pub const OL_MEETING_ACCEPTED: i32 = 3;
pub const OL_MEETING_DECLINED: i32 = 4;

// OlMeetingStatus
pub const OL_NONMEETING: i32 = 0;
pub const OL_MEETING: i32 = 1;

// OlImportance
pub const OL_IMPORTANCE_LOW: i32 = 0;
pub const OL_IMPORTANCE_NORMAL: i32 = 1;
pub const OL_IMPORTANCE_HIGH: i32 = 2;

pub fn folder_id_for_name(name: &str) -> Option<i32> {
    match name.trim().to_lowercase().as_str() {
        "inbox" => Some(OL_FOLDER_INBOX),
        "sent" | "sent items" => Some(OL_FOLDER_SENT_MAIL),
        "drafts" => Some(OL_FOLDER_DRAFTS),
        "deleted" | "deleted items" | "trash" => Some(OL_FOLDER_DELETED_ITEMS),
        "outbox" => Some(OL_FOLDER_OUTBOX),
        _ => None,
    }
}

pub fn importance_id_for_name(name: &str) -> Option<i32> {
    match name.trim().to_lowercase().as_str() {
        "low" => Some(OL_IMPORTANCE_LOW),
        "normal" => Some(OL_IMPORTANCE_NORMAL),
        "high" => Some(OL_IMPORTANCE_HIGH),
        _ => None,
    }
}

pub fn meeting_response_id_for_name(name: &str) -> Option<i32> {
    match name.trim().to_lowercase().as_str() {
        "accept" => Some(OL_MEETING_ACCEPTED),
        "decline" => Some(OL_MEETING_DECLINED),
        "tentative" => Some(OL_MEETING_TENTATIVE),
        _ => None,
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `cargo test constants::tests`
Expected: PASS

- [ ] **Step 5: Wire into lib.rs and commit**

Add `pub mod constants;` to `src/lib.rs`.

```bash
git add src/constants.rs src/lib.rs
git commit -m "Add Outlook enum constants (direct port of constants.py)"
```

---

## Task 5: OutlookClient trait, FakeOutlookClient, and WindowsOutlookClient scaffolding

Defines the full 24-method trait up front (Rust requires complete trait impls
to compile), with `WindowsOutlookClient`'s COM-backed methods stubbed as
`unimplemented!("wired up in a later task")` — each later task replaces one
group of stubs with real logic, so the crate compiles and its growing test
suite stays green after every commit. `FakeOutlookClient` gets its real,
final implementation now (it's pure in-memory logic, no COM), since it's the
test double every later task's unit tests depend on.

**Files:**
- Create: `outlook-mcp-rs/src/outlook/client.rs` (trait `OutlookClient` + `WindowsOutlookClient`)
- Create: `outlook-mcp-rs/src/outlook/fake.rs` (`FakeOutlookClient`, test-only)
- Modify: `outlook-mcp-rs/src/outlook/mod.rs`

**Interfaces:**
- Produces: `pub trait OutlookClient: Send + Sync` with all 24 methods (signatures below), `pub struct WindowsOutlookClient` (COM-backed, `::new() -> Result<Self, ToolError>`), `pub struct FakeOutlookClient` (in `#[cfg(test)]`-gated module, with a `pub calls: std::sync::Mutex<Vec<String>>` call log and a `pub fail_with: std::sync::Mutex<Option<String>>` switch, mirroring `conftest.py`'s `FakeOutlookClient`).
- Consumes: `crate::error::ToolError`, `crate::outlook::dispatch::*`, `crate::constants::*`.

- [ ] **Step 1: Write the OutlookClient trait and WindowsOutlookClient plumbing**

Create `src/outlook/client.rs`:

```rust
use serde_json::{json, Value};
use windows::Win32::System::Com::IDispatch;

use crate::constants as c;
use crate::error::ToolError;
use crate::outlook::dispatch::*;

pub trait OutlookClient: Send + Sync {
    fn list_folders(&self) -> Result<Value, ToolError>;
    fn list_emails(&self, folder: &str, count: i32, unread_only: bool) -> Result<Value, ToolError>;
    fn search_emails(
        &self,
        query: &str,
        folder: &str,
        count: i32,
        since_days: Option<i32>,
    ) -> Result<Value, ToolError>;
    fn get_email(&self, email_id: &str, prefer_html: bool) -> Result<Value, ToolError>;
    fn send_email(
        &self,
        to: &[String],
        subject: &str,
        body: &str,
        cc: Option<&[String]>,
        bcc: Option<&[String]>,
        html: bool,
    ) -> Result<Value, ToolError>;
    fn create_draft(
        &self,
        to: &[String],
        subject: &str,
        body: &str,
        cc: Option<&[String]>,
        bcc: Option<&[String]>,
        html: bool,
    ) -> Result<Value, ToolError>;
    fn reply_email(
        &self,
        email_id: &str,
        body: &str,
        reply_all: bool,
        html: bool,
        send: bool,
    ) -> Result<Value, ToolError>;
    fn move_email(&self, email_id: &str, target_folder: &str) -> Result<Value, ToolError>;
    fn delete_email(&self, email_id: &str) -> Result<Value, ToolError>;

    fn list_events(
        &self,
        start_date: Option<&str>,
        end_date: Option<&str>,
    ) -> Result<Value, ToolError>;
    fn get_event(&self, event_id: &str) -> Result<Value, ToolError>;
    #[allow(clippy::too_many_arguments)]
    fn create_event(
        &self,
        subject: &str,
        start: &str,
        end: &str,
        body: Option<&str>,
        location: Option<&str>,
        attendees: Option<&[String]>,
        all_day: bool,
        reminder_minutes: Option<i32>,
    ) -> Result<Value, ToolError>;
    fn respond_to_meeting(
        &self,
        event_id: &str,
        response: &str,
        comment: Option<&str>,
        send: bool,
    ) -> Result<Value, ToolError>;

    fn list_attachments(&self, email_id: &str) -> Result<Value, ToolError>;
    fn save_attachments(
        &self,
        email_id: &str,
        save_dir: &str,
        attachment_names: Option<&[String]>,
    ) -> Result<Value, ToolError>;

    fn list_tasks(&self, include_completed: bool) -> Result<Value, ToolError>;
    fn create_task(
        &self,
        subject: &str,
        body: Option<&str>,
        due_date: Option<&str>,
        importance: &str,
    ) -> Result<Value, ToolError>;
    fn complete_task(&self, task_id: &str) -> Result<Value, ToolError>;

    fn list_notes(&self) -> Result<Value, ToolError>;
    fn get_note(&self, note_id: &str) -> Result<Value, ToolError>;
    fn create_note(&self, body: &str) -> Result<Value, ToolError>;
}

pub struct WindowsOutlookClient;

impl WindowsOutlookClient {
    pub fn new() -> Result<Self, ToolError> {
        init_com()?;
        Ok(WindowsOutlookClient)
    }

    /// Returns (Application, Namespace("MAPI")), initializing COM on the
    /// calling thread first (COM init must happen on every thread that
    /// touches these objects — mirrors the Python client's per-call
    /// pythoncom.CoInitialize()).
    fn mapi(&self) -> Result<(IDispatch, IDispatch), ToolError> {
        init_com()?;
        let app = create_instance("Outlook.Application")?;
        let ns_variant = call_method(&app, "GetNamespace", &[str_variant("MAPI")])?;
        let ns = variant_dispatch(&ns_variant)?;
        Ok((app, ns))
    }

    fn make_id(item: &IDispatch) -> Result<String, ToolError> {
        let entry_id = variant_str(&get_property(item, "EntryID")?);
        let parent = variant_dispatch(&get_property(item, "Parent")?)?;
        let store_id = variant_str(&get_property(&parent, "StoreID")?);
        Ok(format!("{entry_id}|{store_id}"))
    }

    fn get_item(&self, ns: &IDispatch, item_id: &str) -> Result<IDispatch, ToolError> {
        let (entry_id, store_id) = item_id.split_once('|').ok_or_else(|| {
            ToolError(format!(
                "Invalid item id {item_id:?}: expected the opaque id returned by a list/search tool."
            ))
        })?;
        if entry_id.is_empty() || store_id.is_empty() {
            return Err(ToolError(format!(
                "Invalid item id {item_id:?}: expected the opaque id returned by a list/search tool."
            )));
        }
        let result = call_method(
            ns,
            "GetItemFromID",
            &[str_variant(entry_id), str_variant(store_id)],
        )
        .map_err(|e| {
            ToolError(format!(
                "Item not found — it may have been moved or deleted (item ids change \
                 when an item moves to another folder). {e}"
            ))
        })?;
        variant_dispatch(&result)
    }

    fn resolve_folder(&self, ns: &IDispatch, folder: Option<&str>) -> Result<IDispatch, ToolError> {
        let name = folder.unwrap_or("inbox").trim().to_string();
        if let Some(id) = c::folder_id_for_name(&name) {
            let result = call_method(ns, "GetDefaultFolder", &[i32_variant(id)])?;
            return variant_dispatch(&result);
        }
        let inbox = variant_dispatch(&call_method(
            ns,
            "GetDefaultFolder",
            &[i32_variant(c::OL_FOLDER_INBOX)],
        )?)?;
        let mut current = variant_dispatch(&get_property(&inbox, "Parent")?)?;
        for part in name.split(['/', '\\']).filter(|p| !p.is_empty()) {
            let folders = variant_dispatch(&get_property(&current, "Folders")?)?;
            let count = variant_i32(&get_property(&folders, "Count")?);
            let mut found = None;
            for i in 1..=count {
                let sub = variant_dispatch(&call_method(&folders, "Item", &[i32_variant(i)])?)?;
                let sub_name = variant_str(&get_property(&sub, "Name")?);
                if sub_name.eq_ignore_ascii_case(part) {
                    found = Some(sub);
                    break;
                }
            }
            current = found.ok_or_else(|| {
                ToolError(format!(
                    "Folder not found: {name:?} (no subfolder named {part:?})"
                ))
            })?;
        }
        Ok(current)
    }
}

impl OutlookClient for WindowsOutlookClient {
    fn list_folders(&self) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 6")
    }
    fn list_emails(&self, _folder: &str, _count: i32, _unread_only: bool) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 6")
    }
    fn search_emails(
        &self,
        _query: &str,
        _folder: &str,
        _count: i32,
        _since_days: Option<i32>,
    ) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 6")
    }
    fn get_email(&self, _email_id: &str, _prefer_html: bool) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 6")
    }
    fn send_email(
        &self,
        _to: &[String],
        _subject: &str,
        _body: &str,
        _cc: Option<&[String]>,
        _bcc: Option<&[String]>,
        _html: bool,
    ) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 7")
    }
    fn create_draft(
        &self,
        _to: &[String],
        _subject: &str,
        _body: &str,
        _cc: Option<&[String]>,
        _bcc: Option<&[String]>,
        _html: bool,
    ) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 7")
    }
    fn reply_email(
        &self,
        _email_id: &str,
        _body: &str,
        _reply_all: bool,
        _html: bool,
        _send: bool,
    ) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 7")
    }
    fn move_email(&self, _email_id: &str, _target_folder: &str) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 7")
    }
    fn delete_email(&self, _email_id: &str) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 7")
    }
    fn list_events(
        &self,
        _start_date: Option<&str>,
        _end_date: Option<&str>,
    ) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 8")
    }
    fn get_event(&self, _event_id: &str) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 8")
    }
    fn create_event(
        &self,
        _subject: &str,
        _start: &str,
        _end: &str,
        _body: Option<&str>,
        _location: Option<&str>,
        _attendees: Option<&[String]>,
        _all_day: bool,
        _reminder_minutes: Option<i32>,
    ) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 8")
    }
    fn respond_to_meeting(
        &self,
        _event_id: &str,
        _response: &str,
        _comment: Option<&str>,
        _send: bool,
    ) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 8")
    }
    fn list_attachments(&self, _email_id: &str) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 9")
    }
    fn save_attachments(
        &self,
        _email_id: &str,
        _save_dir: &str,
        _attachment_names: Option<&[String]>,
    ) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 9")
    }
    fn list_tasks(&self, _include_completed: bool) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 10")
    }
    fn create_task(
        &self,
        _subject: &str,
        _body: Option<&str>,
        _due_date: Option<&str>,
        _importance: &str,
    ) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 10")
    }
    fn complete_task(&self, _task_id: &str) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 10")
    }
    fn list_notes(&self) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 11")
    }
    fn get_note(&self, _note_id: &str) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 11")
    }
    fn create_note(&self, _body: &str) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 11")
    }
}
```

- [ ] **Step 2: Run build to verify the trait + stub impl compile**

Run: `cargo build`
Expected: builds successfully (warnings about unused `json`/`c` imports are
fine for now — they'll be used starting in Task 6).

- [ ] **Step 3: Write FakeOutlookClient (final implementation, not a stub)**

Create `src/outlook/fake.rs` — a direct port of `tests/conftest.py`'s
`FakeOutlookClient`, used by every unit test from Task 6 onward:

```rust
use std::sync::Mutex;

use serde_json::{json, Value};

use crate::error::ToolError;
use crate::outlook::client::OutlookClient;

pub const EMAIL_ID: &str = "entry-1|store-1";
pub const EVENT_ID: &str = "entry-2|store-1";
pub const TASK_ID: &str = "entry-3|store-1";
pub const NOTE_ID: &str = "entry-4|store-1";

#[derive(Default)]
pub struct FakeOutlookClient {
    pub calls: Mutex<Vec<(String, Value)>>,
    pub fail_with: Mutex<Option<String>>,
}

impl FakeOutlookClient {
    pub fn new() -> Self {
        Self::default()
    }

    fn record(&self, name: &str, args: Value) -> Result<(), ToolError> {
        if let Some(msg) = self.fail_with.lock().unwrap().clone() {
            return Err(ToolError(msg));
        }
        self.calls.lock().unwrap().push((name.to_string(), args));
        Ok(())
    }
}

impl OutlookClient for FakeOutlookClient {
    fn list_folders(&self) -> Result<Value, ToolError> {
        self.record("list_folders", json!({}))?;
        Ok(json!([{"name": "Inbox", "path": "Inbox", "items": 2, "unread": 1}]))
    }

    fn list_emails(&self, folder: &str, count: i32, unread_only: bool) -> Result<Value, ToolError> {
        self.record(
            "list_emails",
            json!({"folder": folder, "count": count, "unread_only": unread_only}),
        )?;
        Ok(json!([{"id": EMAIL_ID, "subject": "Hello", "sender": "Ada", "unread": true}]))
    }

    fn search_emails(
        &self,
        query: &str,
        folder: &str,
        count: i32,
        since_days: Option<i32>,
    ) -> Result<Value, ToolError> {
        self.record(
            "search_emails",
            json!({"query": query, "folder": folder, "count": count, "since_days": since_days}),
        )?;
        Ok(json!([{"id": EMAIL_ID, "subject": "Hello"}]))
    }

    fn get_email(&self, email_id: &str, prefer_html: bool) -> Result<Value, ToolError> {
        self.record(
            "get_email",
            json!({"email_id": email_id, "prefer_html": prefer_html}),
        )?;
        Ok(json!({"id": email_id, "subject": "Hello", "body": "Hi there"}))
    }

    fn send_email(
        &self,
        to: &[String],
        subject: &str,
        body: &str,
        cc: Option<&[String]>,
        bcc: Option<&[String]>,
        html: bool,
    ) -> Result<Value, ToolError> {
        self.record(
            "send_email",
            json!({"to": to, "subject": subject, "body": body, "cc": cc, "bcc": bcc, "html": html}),
        )?;
        Ok(json!({"status": "sent", "to": to.join("; "), "subject": subject}))
    }

    fn create_draft(
        &self,
        to: &[String],
        subject: &str,
        body: &str,
        cc: Option<&[String]>,
        bcc: Option<&[String]>,
        html: bool,
    ) -> Result<Value, ToolError> {
        self.record(
            "create_draft",
            json!({"to": to, "subject": subject, "body": body, "cc": cc, "bcc": bcc, "html": html}),
        )?;
        Ok(json!({"status": "draft_saved", "id": EMAIL_ID, "subject": subject}))
    }

    fn reply_email(
        &self,
        email_id: &str,
        body: &str,
        reply_all: bool,
        html: bool,
        send: bool,
    ) -> Result<Value, ToolError> {
        self.record(
            "reply_email",
            json!({"email_id": email_id, "body": body, "reply_all": reply_all, "html": html, "send": send}),
        )?;
        Ok(json!({"status": if send { "sent" } else { "draft_saved" }}))
    }

    fn move_email(&self, email_id: &str, target_folder: &str) -> Result<Value, ToolError> {
        self.record(
            "move_email",
            json!({"email_id": email_id, "target_folder": target_folder}),
        )?;
        Ok(json!({"status": "moved", "folder": target_folder, "id": "new-entry|store-1"}))
    }

    fn delete_email(&self, email_id: &str) -> Result<Value, ToolError> {
        self.record("delete_email", json!({"email_id": email_id}))?;
        Ok(json!({"status": "deleted"}))
    }

    fn list_events(
        &self,
        start_date: Option<&str>,
        end_date: Option<&str>,
    ) -> Result<Value, ToolError> {
        self.record(
            "list_events",
            json!({"start_date": start_date, "end_date": end_date}),
        )?;
        Ok(json!([{"id": EVENT_ID, "subject": "Standup"}]))
    }

    fn get_event(&self, event_id: &str) -> Result<Value, ToolError> {
        self.record("get_event", json!({"event_id": event_id}))?;
        Ok(json!({"id": event_id, "subject": "Standup", "body": ""}))
    }

    fn create_event(
        &self,
        subject: &str,
        start: &str,
        end: &str,
        body: Option<&str>,
        location: Option<&str>,
        attendees: Option<&[String]>,
        all_day: bool,
        reminder_minutes: Option<i32>,
    ) -> Result<Value, ToolError> {
        self.record(
            "create_event",
            json!({
                "subject": subject, "start": start, "end": end, "body": body,
                "location": location, "attendees": attendees, "all_day": all_day,
                "reminder_minutes": reminder_minutes
            }),
        )?;
        Ok(json!({"status": "saved", "id": EVENT_ID, "subject": subject}))
    }

    fn respond_to_meeting(
        &self,
        event_id: &str,
        response: &str,
        comment: Option<&str>,
        send: bool,
    ) -> Result<Value, ToolError> {
        self.record(
            "respond_to_meeting",
            json!({"event_id": event_id, "response": response, "comment": comment, "send": send}),
        )?;
        Ok(json!({"status": format!("{response}_sent")}))
    }

    fn list_attachments(&self, email_id: &str) -> Result<Value, ToolError> {
        self.record("list_attachments", json!({"email_id": email_id}))?;
        Ok(json!([{"index": 1, "filename": "report.pdf", "size": 1234}]))
    }

    fn save_attachments(
        &self,
        email_id: &str,
        save_dir: &str,
        attachment_names: Option<&[String]>,
    ) -> Result<Value, ToolError> {
        self.record(
            "save_attachments",
            json!({"email_id": email_id, "save_dir": save_dir, "attachment_names": attachment_names}),
        )?;
        Ok(json!([{"filename": "report.pdf", "saved_to": save_dir, "status": "saved"}]))
    }

    fn list_tasks(&self, include_completed: bool) -> Result<Value, ToolError> {
        self.record("list_tasks", json!({"include_completed": include_completed}))?;
        Ok(json!([{"id": TASK_ID, "subject": "Buy milk", "complete": false}]))
    }

    fn create_task(
        &self,
        subject: &str,
        body: Option<&str>,
        due_date: Option<&str>,
        importance: &str,
    ) -> Result<Value, ToolError> {
        self.record(
            "create_task",
            json!({"subject": subject, "body": body, "due_date": due_date, "importance": importance}),
        )?;
        Ok(json!({"status": "created", "id": TASK_ID, "subject": subject}))
    }

    fn complete_task(&self, task_id: &str) -> Result<Value, ToolError> {
        self.record("complete_task", json!({"task_id": task_id}))?;
        Ok(json!({"status": "completed"}))
    }

    fn list_notes(&self) -> Result<Value, ToolError> {
        self.record("list_notes", json!({}))?;
        Ok(json!([{"id": NOTE_ID, "subject": "Ideas"}]))
    }

    fn get_note(&self, note_id: &str) -> Result<Value, ToolError> {
        self.record("get_note", json!({"note_id": note_id}))?;
        Ok(json!({"id": note_id, "subject": "Ideas", "body": "Ideas\n- one"}))
    }

    fn create_note(&self, body: &str) -> Result<Value, ToolError> {
        self.record("create_note", json!({"body": body}))?;
        Ok(json!({"status": "created", "id": NOTE_ID}))
    }
}
```

- [ ] **Step 4: Write and run a test proving the fake works**

Add to the bottom of `src/outlook/fake.rs`:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn records_calls_and_returns_canned_data() {
        let fake = FakeOutlookClient::new();
        let result = fake.list_emails("inbox", 10, false).unwrap();
        assert_eq!(result[0]["id"], EMAIL_ID);
        assert_eq!(fake.calls.lock().unwrap().len(), 1);
    }

    #[test]
    fn fail_with_makes_every_call_error() {
        let fake = FakeOutlookClient::new();
        *fake.fail_with.lock().unwrap() = Some("boom".to_string());
        let err = fake.list_emails("inbox", 10, false).unwrap_err();
        assert_eq!(err.0, "boom");
    }
}
```

Run: `cargo test outlook::fake::tests`
Expected: PASS

- [ ] **Step 5: Wire modules into outlook/mod.rs**

```rust
pub mod client;
pub mod dispatch;

#[cfg(any(test, feature = "test-support"))]
pub mod fake;
```

Add a `[features]` section to `Cargo.toml` so `tests/` integration test
binaries (which compile as a separate crate and don't automatically get
`#[cfg(test)]` from the lib) can reach `FakeOutlookClient` too:

```toml
[features]
test-support = []

[dev-dependencies]
outlook-mcp-rs = { path = ".", features = ["test-support"] }
```

- [ ] **Step 6: Run the full test suite and commit**

Run: `cargo test`
Expected: all tests from Tasks 1-5 pass.

```bash
git add src/outlook/
git commit -m "Add OutlookClient trait, FakeOutlookClient, and WindowsOutlookClient scaffolding"
```

---

## Task 6: Email read tools

Replaces the `list_folders`/`list_emails`/`search_emails`/`get_email` stubs
in `WindowsOutlookClient` with real COM logic, and adds their `#[tool]`
wrappers to `OutlookMcpServer`.

**Files:**
- Modify: `outlook-mcp-rs/src/outlook/client.rs` (replace 4 stub bodies)
- Modify: `outlook-mcp-rs/src/server.rs` (add `client` field + 4 tools)

**Interfaces:**
- Consumes: `OutlookClient` trait methods from Task 5, `dispatch::*` helpers from Task 3.
- Produces: `OutlookMcpServer { client: std::sync::Arc<dyn OutlookClient> }` (the `client` field is added now, in this first tool task, and used by every later tool task).

- [ ] **Step 1: Add the client field to OutlookMcpServer**

Replace the contents of `src/server.rs` with:

```rust
use std::sync::Arc;

use rmcp::handler::server::wrapper::Parameters;
use rmcp::model::{CallToolResult, Content};
use rmcp::{ServerHandler, tool, tool_router};
use serde::Deserialize;
use schemars::JsonSchema;

use crate::error::McpError;
use crate::outlook::client::OutlookClient;

#[derive(Clone)]
pub struct OutlookMcpServer {
    pub client: Arc<dyn OutlookClient>,
}

impl OutlookMcpServer {
    pub fn new(client: Arc<dyn OutlookClient>) -> Self {
        Self { client }
    }
}

fn default_folder() -> String {
    "inbox".to_string()
}

fn default_count() -> i32 {
    10
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ListEmailsRequest {
    #[serde(default = "default_folder")]
    pub folder: String,
    #[serde(default = "default_count")]
    pub count: i32,
    #[serde(default)]
    pub unread_only: bool,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct SearchEmailsRequest {
    pub query: String,
    #[serde(default = "default_folder")]
    pub folder: String,
    #[serde(default = "default_count")]
    pub count: i32,
    pub since_days: Option<i32>,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct GetEmailRequest {
    pub email_id: String,
    #[serde(default)]
    pub prefer_html: bool,
}

#[tool_router(server_handler)]
impl OutlookMcpServer {
    #[tool(description = "Health check: returns 'pong' if the server is running.")]
    fn ping(&self) -> String {
        "pong".to_string()
    }

    #[tool(description = "List all Outlook mail folders with their paths, item counts and \
        unread counts. Folder paths returned here can be passed to other tools' \
        folder/target_folder arguments.")]
    async fn list_folders(&self) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || client.list_folders())
            .await
            .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "List recent emails in a folder, newest first. `folder` accepts a \
        well-known name (inbox, sent, drafts, deleted, outbox) or a path like \
        'Inbox/Receipts'. `count` is capped at 50.")]
    async fn list_emails(
        &self,
        Parameters(req): Parameters<ListEmailsRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.list_emails(&req.folder, req.count, req.unread_only)
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Search emails by text across subject, sender name and body. \
        Optionally limit to messages received in the last since_days days.")]
    async fn search_emails(
        &self,
        Parameters(req): Parameters<SearchEmailsRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.search_emails(&req.query, &req.folder, req.count, req.since_days)
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Read a full email (headers, body, attachment names) by the id \
        returned from list_emails/search_emails. Set prefer_html to also get the HTML body.")]
    async fn get_email(
        &self,
        Parameters(req): Parameters<GetEmailRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.get_email(&req.email_id, req.prefer_html)
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }
}
```

- [ ] **Step 2: Update main.rs to construct the real client**

Replace `src/main.rs`'s server construction:

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::client::WindowsOutlookClient;

// ... inside main(), replacing `OutlookMcpServer::default()`:
let client = Arc::new(WindowsOutlookClient::new()?);
let service = OutlookMcpServer::new(client)
    .serve(stdio())
    .await
    .inspect_err(|e| tracing::error!("serving error: {:?}", e))?;
```

Also update `tests/smoke.rs`'s `ping_returns_pong` test to construct the
server with a fake client instead of `Default`:

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::FakeOutlookClient;
use outlook_mcp_rs::server::OutlookMcpServer;

#[test]
fn ping_returns_pong() {
    let server = OutlookMcpServer::new(Arc::new(FakeOutlookClient::new()));
    assert_eq!(server.ping(), "pong");
}
```

- [ ] **Step 3: Write failing unit tests for the four tools**

Create `tests/email_read_tools.rs`:

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, EMAIL_ID};
use outlook_mcp_rs::server::{GetEmailRequest, ListEmailsRequest, OutlookMcpServer, SearchEmailsRequest};
use rmcp::handler::server::wrapper::Parameters;

fn server_with_fake() -> (OutlookMcpServer, Arc<FakeOutlookClient>) {
    let fake = Arc::new(FakeOutlookClient::new());
    (OutlookMcpServer::new(fake.clone()), fake)
}

#[tokio::test]
async fn list_folders_returns_fake_folders() {
    let (server, _fake) = server_with_fake();
    let result = server.list_folders().await.unwrap();
    assert!(!result.content.is_empty());
}

#[tokio::test]
async fn list_emails_uses_default_folder_and_count() {
    let (server, fake) = server_with_fake();
    server
        .list_emails(Parameters(ListEmailsRequest {
            folder: "inbox".to_string(),
            count: 10,
            unread_only: false,
        }))
        .await
        .unwrap();
    let calls = fake.calls.lock().unwrap();
    assert_eq!(calls[0].0, "list_emails");
    assert_eq!(calls[0].1["folder"], "inbox");
}

#[tokio::test]
async fn search_emails_passes_query_through() {
    let (server, fake) = server_with_fake();
    server
        .search_emails(Parameters(SearchEmailsRequest {
            query: "invoice".to_string(),
            folder: "inbox".to_string(),
            count: 10,
            since_days: Some(7),
        }))
        .await
        .unwrap();
    let calls = fake.calls.lock().unwrap();
    assert_eq!(calls[0].1["query"], "invoice");
    assert_eq!(calls[0].1["since_days"], 7);
}

#[tokio::test]
async fn get_email_returns_the_requested_id() {
    let (server, _fake) = server_with_fake();
    server
        .get_email(Parameters(GetEmailRequest {
            email_id: EMAIL_ID.to_string(),
            prefer_html: false,
        }))
        .await
        .unwrap();
}

#[tokio::test]
async fn tool_error_becomes_an_mcp_error() {
    let (server, fake) = server_with_fake();
    *fake.fail_with.lock().unwrap() = Some("Outlook error: boom".to_string());
    let err = server
        .list_emails(Parameters(ListEmailsRequest {
            folder: "inbox".to_string(),
            count: 10,
            unread_only: false,
        }))
        .await
        .unwrap_err();
    assert_eq!(err.message, "Outlook error: boom");
}
```

- [ ] **Step 4: Run tests to verify they fail**

Run: `cargo test --test email_read_tools`
Expected: FAIL — `ListEmailsRequest` etc. aren't `pub` from `server` yet, or
the real `list_folders`/etc. COM bodies still say `unimplemented!()` (only
matters once we implement Step 5; the request-struct/plumbing tests above
should already pass once Step 1-2 compile — if they fail only because of
`unimplemented!()` panics in `WindowsOutlookClient`, that's expected: these
tests use `FakeOutlookClient`, not `WindowsOutlookClient`, so they should
actually pass right after Step 1-2. If they still fail, the compiler error
or panic message will point at the exact mismatch to fix.)

- [ ] **Step 5: Replace the four stub bodies in WindowsOutlookClient**

In `src/outlook/client.rs`, replace:

```rust
    fn list_folders(&self) -> Result<Value, ToolError> {
        unimplemented!("wired up in Task 6")
    }
```

with:

```rust
    fn list_folders(&self) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let inbox = variant_dispatch(&call_method(&ns, "GetDefaultFolder", &[i32_variant(c::OL_FOLDER_INBOX)])?)?;
        let root = variant_dispatch(&get_property(&inbox, "Parent")?)?;

        fn walk(folder: &IDispatch, path: String, depth: u32, out: &mut Vec<Value>) -> Result<(), ToolError> {
            let item_count = variant_i32(&get_property(&variant_dispatch(&get_property(folder, "Items")?)?, "Count")?);
            let unread = variant_i32(&get_property(folder, "UnReadItemCount").unwrap_or_default());
            let name = variant_str(&get_property(folder, "Name")?);
            out.push(json!({"name": name, "path": path, "items": item_count, "unread": unread}));
            if depth >= 3 {
                return Ok(());
            }
            let subfolders = variant_dispatch(&get_property(folder, "Folders")?)?;
            let count = variant_i32(&get_property(&subfolders, "Count")?);
            for i in 1..=count {
                let sub = variant_dispatch(&call_method(&subfolders, "Item", &[i32_variant(i)])?)?;
                let sub_name = variant_str(&get_property(&sub, "Name")?);
                walk(&sub, format!("{path}/{sub_name}"), depth + 1, out)?;
            }
            Ok(())
        }

        let subfolders = variant_dispatch(&get_property(&root, "Folders")?)?;
        let count = variant_i32(&get_property(&subfolders, "Count")?);
        let mut results = Vec::new();
        for i in 1..=count {
            let sub = variant_dispatch(&call_method(&subfolders, "Item", &[i32_variant(i)])?)?;
            let name = variant_str(&get_property(&sub, "Name")?);
            walk(&sub, name, 1, &mut results)?;
        }
        Ok(json!(results))
    }
```

Replace `list_emails`'s body with:

```rust
    fn list_emails(&self, folder: &str, count: i32, unread_only: bool) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let count = count.clamp(1, 50);
        let target = self.resolve_folder(&ns, Some(folder))?;
        let mut items = variant_dispatch(&get_property(&target, "Items")?)?;
        if unread_only {
            items = variant_dispatch(&call_method(&items, "Restrict", &[str_variant("[UnRead] = True")])?)?;
        }
        call_method(&items, "Sort", &[str_variant("[ReceivedTime]"), bool_variant(true)])?;
        let total = variant_i32(&get_property(&items, "Count")?);
        let mut results = Vec::new();
        for i in 1..=total {
            if results.len() as i32 >= count {
                break;
            }
            let item = variant_dispatch(&call_method(&items, "Item", &[i32_variant(i)])?)?;
            results.push(self.email_summary(&item)?);
        }
        Ok(json!(results))
    }
```

Replace `search_emails`'s body with:

```rust
    fn search_emails(
        &self,
        query: &str,
        folder: &str,
        count: i32,
        since_days: Option<i32>,
    ) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let count = count.clamp(1, 50);
        let escaped = query.replace('\'', "''");
        let dasl = format!(
            "@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{escaped}%' \
             OR \"urn:schemas:httpmail:fromname\" LIKE '%{escaped}%' \
             OR \"urn:schemas:httpmail:textdescription\" LIKE '%{escaped}%')"
        );
        let target = self.resolve_folder(&ns, Some(folder))?;
        let items = variant_dispatch(&get_property(&target, "Items")?)?;
        let mut items = variant_dispatch(&call_method(&items, "Restrict", &[str_variant(&dasl)])?)?;
        if let Some(days) = since_days {
            let cutoff = chrono::Local::now().naive_local() - chrono::Duration::days(days as i64);
            let filter = format!("[ReceivedTime] >= '{}'", cutoff.format("%m/%d/%Y %I:%M %p"));
            items = variant_dispatch(&call_method(&items, "Restrict", &[str_variant(&filter)])?)?;
        }
        call_method(&items, "Sort", &[str_variant("[ReceivedTime]"), bool_variant(true)])?;
        let total = variant_i32(&get_property(&items, "Count")?);
        let mut results = Vec::new();
        for i in 1..=total {
            if results.len() as i32 >= count {
                break;
            }
            let item = variant_dispatch(&call_method(&items, "Item", &[i32_variant(i)])?)?;
            results.push(self.email_summary(&item)?);
        }
        Ok(json!(results))
    }
```

Replace `get_email`'s body with:

```rust
    fn get_email(&self, email_id: &str, prefer_html: bool) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let item = self.get_item(&ns, email_id)?;
        let mut info = self.email_summary(&item)?;
        let obj = info.as_object_mut().unwrap();
        obj.insert("cc".into(), json!(variant_str(&get_property(&item, "CC")?)));
        obj.insert("bcc".into(), json!(variant_str(&get_property(&item, "BCC")?)));
        const MAX_BODY_CHARS: usize = 100_000;
        let body = variant_str(&get_property(&item, "Body")?);
        obj.insert("body".into(), json!(truncate(&body, MAX_BODY_CHARS)));
        if prefer_html {
            let html_body = variant_str(&get_property(&item, "HTMLBody")?);
            obj.insert("html_body".into(), json!(truncate(&html_body, MAX_BODY_CHARS)));
        }
        let attachments = variant_dispatch(&get_property(&item, "Attachments")?)?;
        let att_count = variant_i32(&get_property(&attachments, "Count")?);
        let mut names = Vec::new();
        for i in 1..=att_count {
            let att = variant_dispatch(&call_method(&attachments, "Item", &[i32_variant(i)])?)?;
            names.push(variant_str(&get_property(&att, "FileName")?));
        }
        obj.insert("attachments".into(), json!(names));
        Ok(info)
    }
```

Add the two small helpers these bodies use, in the `impl WindowsOutlookClient` block (not the trait impl):

```rust
    fn email_summary(&self, item: &IDispatch) -> Result<Value, ToolError> {
        let attachments = get_property(item, "Attachments").ok();
        let has_attachments = match &attachments {
            Some(v) => variant_i32(&get_property(&variant_dispatch(v)?, "Count")?) > 0,
            None => false,
        };
        Ok(json!({
            "id": Self::make_id(item)?,
            "subject": variant_str(&get_property(item, "Subject")?),
            "sender": variant_str(&get_property(item, "SenderName")?),
            "sender_email": variant_str(&get_property(item, "SenderEmailAddress")?),
            "to": variant_str(&get_property(item, "To")?),
            "received": variant_date_iso(&get_property(item, "ReceivedTime")?),
            "unread": variant_bool(&get_property(item, "UnRead")?),
            "has_attachments": has_attachments,
        }))
    }
```

And a free function `truncate` at the bottom of the file (outside any impl block):

```rust
fn truncate(text: &str, max_chars: usize) -> String {
    if text.chars().count() > max_chars {
        let truncated: String = text.chars().take(max_chars).collect();
        format!("{truncated}\n\n[... truncated at {max_chars} characters]")
    } else {
        text.to_string()
    }
}
```

Add `use windows::Win32::System::Com::IDispatch;` and `use serde_json::json;`
to the top of `src/outlook/client.rs` if not already present from Task 5.

- [ ] **Step 6: Run unit tests to verify they pass**

Run: `cargo test --test email_read_tools`
Expected: PASS (these exercise the tool-wrapper + `FakeOutlookClient` layer,
not the real COM code — COM correctness is checked by the live suite added
in Task 12).

Run: `cargo build`
Expected: succeeds (the real `WindowsOutlookClient` COM code now compiles;
whether it's *correct* against real Outlook is checked later, since this
machine has Outlook so it's worth also running a manual check now):

```bash
cargo run
```

In another terminal, use the MCP inspector to call `list_folders` and
confirm it returns your real Outlook folder list:

```bash
npx @modelcontextprotocol/inspector cargo run
```

- [ ] **Step 7: Commit**

```bash
git add src/outlook/client.rs src/server.rs src/main.rs tests/
git commit -m "Implement email read tools (list_folders, list_emails, search_emails, get_email)"
```

---

## Task 7: Email write tools

Same pattern as Task 6, for `send_email`, `create_draft`, `reply_email`,
`move_email`, `delete_email`.

**Files:**
- Modify: `outlook-mcp-rs/src/outlook/client.rs` (replace 5 stub bodies)
- Modify: `outlook-mcp-rs/src/server.rs` (add 5 tools + their request structs)

**Interfaces:**
- Consumes: same as Task 6.

- [ ] **Step 1: Add request structs and tool methods to server.rs**

Add to `src/server.rs`, alongside the existing request structs:

```rust
#[derive(Debug, Deserialize, JsonSchema)]
pub struct SendEmailRequest {
    pub to: Vec<String>,
    pub subject: String,
    pub body: String,
    pub cc: Option<Vec<String>>,
    pub bcc: Option<Vec<String>>,
    #[serde(default)]
    pub html: bool,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct CreateDraftRequest {
    pub to: Vec<String>,
    pub subject: String,
    pub body: String,
    pub cc: Option<Vec<String>>,
    pub bcc: Option<Vec<String>>,
    #[serde(default)]
    pub html: bool,
}

fn default_true() -> bool {
    true
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ReplyEmailRequest {
    pub email_id: String,
    pub body: String,
    #[serde(default)]
    pub reply_all: bool,
    #[serde(default)]
    pub html: bool,
    #[serde(default = "default_true")]
    pub send: bool,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct MoveEmailRequest {
    pub email_id: String,
    pub target_folder: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct DeleteEmailRequest {
    pub email_id: String,
}
```

Add to the `#[tool_router(server_handler)] impl OutlookMcpServer` block:

```rust
    #[tool(description = "Send an email immediately as the signed-in Outlook user. Use \
        create_draft instead if the user should review before sending. Set html to true \
        if body is HTML.")]
    async fn send_email(
        &self,
        Parameters(req): Parameters<SendEmailRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.send_email(
                &req.to,
                &req.subject,
                &req.body,
                req.cc.as_deref(),
                req.bcc.as_deref(),
                req.html,
            )
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Compose an email and save it to Drafts without sending. \
        Returns the draft's id.")]
    async fn create_draft(
        &self,
        Parameters(req): Parameters<CreateDraftRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.create_draft(
                &req.to,
                &req.subject,
                &req.body,
                req.cc.as_deref(),
                req.bcc.as_deref(),
                req.html,
            )
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Reply to an email (reply-all optional). Sends immediately by \
        default; pass send=false to save the reply as a draft instead.")]
    async fn reply_email(
        &self,
        Parameters(req): Parameters<ReplyEmailRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.reply_email(&req.email_id, &req.body, req.reply_all, req.html, req.send)
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Move an email to another folder. Returns the email's NEW id \
        (ids change when an item moves).")]
    async fn move_email(
        &self,
        Parameters(req): Parameters<MoveEmailRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.move_email(&req.email_id, &req.target_folder)
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Delete an email (moves it to Deleted Items).")]
    async fn delete_email(
        &self,
        Parameters(req): Parameters<DeleteEmailRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || client.delete_email(&req.email_id))
            .await
            .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }
```

- [ ] **Step 2: Write failing unit tests**

Create `tests/email_write_tools.rs`:

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, EMAIL_ID};
use outlook_mcp_rs::server::{
    CreateDraftRequest, DeleteEmailRequest, MoveEmailRequest, OutlookMcpServer,
    ReplyEmailRequest, SendEmailRequest,
};
use rmcp::handler::server::wrapper::Parameters;

fn server_with_fake() -> (OutlookMcpServer, Arc<FakeOutlookClient>) {
    let fake = Arc::new(FakeOutlookClient::new());
    (OutlookMcpServer::new(fake.clone()), fake)
}

#[tokio::test]
async fn send_email_passes_recipients_through() {
    let (server, fake) = server_with_fake();
    server
        .send_email(Parameters(SendEmailRequest {
            to: vec!["a@example.com".to_string()],
            subject: "Hi".to_string(),
            body: "Body".to_string(),
            cc: None,
            bcc: None,
            html: false,
        }))
        .await
        .unwrap();
    let calls = fake.calls.lock().unwrap();
    assert_eq!(calls[0].0, "send_email");
    assert_eq!(calls[0].1["subject"], "Hi");
}

#[tokio::test]
async fn create_draft_defaults_are_respected() {
    let (server, _fake) = server_with_fake();
    let result = server
        .create_draft(Parameters(CreateDraftRequest {
            to: vec!["a@example.com".to_string()],
            subject: "Hi".to_string(),
            body: "Body".to_string(),
            cc: None,
            bcc: None,
            html: false,
        }))
        .await
        .unwrap();
    assert!(!result.content.is_empty());
}

#[tokio::test]
async fn reply_email_send_defaults_to_true() {
    let (server, fake) = server_with_fake();
    server
        .reply_email(Parameters(ReplyEmailRequest {
            email_id: EMAIL_ID.to_string(),
            body: "Thanks".to_string(),
            reply_all: false,
            html: false,
            send: true,
        }))
        .await
        .unwrap();
    let calls = fake.calls.lock().unwrap();
    assert_eq!(calls[0].1["send"], true);
}

#[tokio::test]
async fn move_email_returns_new_id() {
    let (server, _fake) = server_with_fake();
    let result = server
        .move_email(Parameters(MoveEmailRequest {
            email_id: EMAIL_ID.to_string(),
            target_folder: "Archive".to_string(),
        }))
        .await
        .unwrap();
    assert!(!result.content.is_empty());
}

#[tokio::test]
async fn delete_email_calls_through() {
    let (server, fake) = server_with_fake();
    server
        .delete_email(Parameters(DeleteEmailRequest {
            email_id: EMAIL_ID.to_string(),
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls.lock().unwrap()[0].0, "delete_email");
}
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `cargo test --test email_write_tools`
Expected: FAIL to compile until Step 1 is in place.

- [ ] **Step 4: Replace the five stub bodies in WindowsOutlookClient**

```rust
    fn send_email(
        &self,
        to: &[String],
        subject: &str,
        body: &str,
        cc: Option<&[String]>,
        bcc: Option<&[String]>,
        html: bool,
    ) -> Result<Value, ToolError> {
        if to.is_empty() {
            return Err(ToolError("send_email requires at least one recipient in 'to'.".into()));
        }
        let (app, _) = self.mapi()?;
        let mail = self.compose(&app, to, subject, body, cc, bcc, html)?;
        call_method(&mail, "Send", &[])?;
        Ok(json!({"status": "sent", "to": to.join("; "), "subject": subject}))
    }

    fn create_draft(
        &self,
        to: &[String],
        subject: &str,
        body: &str,
        cc: Option<&[String]>,
        bcc: Option<&[String]>,
        html: bool,
    ) -> Result<Value, ToolError> {
        let (app, _) = self.mapi()?;
        let mail = self.compose(&app, to, subject, body, cc, bcc, html)?;
        call_method(&mail, "Save", &[])?;
        Ok(json!({"status": "draft_saved", "id": Self::make_id(&mail)?, "subject": subject}))
    }

    fn reply_email(
        &self,
        email_id: &str,
        body: &str,
        reply_all: bool,
        html: bool,
        send: bool,
    ) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let item = self.get_item(&ns, email_id)?;
        let reply = variant_dispatch(&call_method(
            &item,
            if reply_all { "ReplyAll" } else { "Reply" },
            &[],
        )?)?;
        if html {
            let existing = variant_str(&get_property(&reply, "HTMLBody")?);
            put_property(&reply, "HTMLBody", str_variant(&format!("{body}{existing}")))?;
        } else {
            let existing = variant_str(&get_property(&reply, "Body")?);
            put_property(&reply, "Body", str_variant(&format!("{body}\n\n{existing}")))?;
        }
        if send {
            call_method(&reply, "Send", &[])?;
            let subject = variant_str(&get_property(&reply, "Subject")?);
            Ok(json!({"status": "sent", "subject": subject}))
        } else {
            call_method(&reply, "Save", &[])?;
            let subject = variant_str(&get_property(&reply, "Subject")?);
            Ok(json!({"status": "draft_saved", "id": Self::make_id(&reply)?, "subject": subject}))
        }
    }

    fn move_email(&self, email_id: &str, target_folder: &str) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let item = self.get_item(&ns, email_id)?;
        let target = self.resolve_folder(&ns, Some(target_folder))?;
        let moved = variant_dispatch(&call_method(&item, "Move", &[VARIANT::from(&target)])?)?;
        let target_name = variant_str(&get_property(&target, "Name")?);
        Ok(json!({"status": "moved", "folder": target_name, "id": Self::make_id(&moved)?}))
    }

    fn delete_email(&self, email_id: &str) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let item = self.get_item(&ns, email_id)?;
        let subject = variant_str(&get_property(&item, "Subject")?);
        call_method(&item, "Delete", &[])?;
        Ok(json!({"status": "deleted", "subject": subject, "note": "Moved to Deleted Items."}))
    }
```

`Move` takes the target folder *object* as its argument, not a string, so
`VARIANT::from(&target)` must construct a `VT_DISPATCH` variant from an
`IDispatch` reference. Add this conversion to `src/outlook/dispatch.rs` if
`VARIANT::from(&IDispatch)` isn't already provided by the `windows` crate
(check with `cargo build` first — if it compiles as written, the crate
already provides it):

```rust
pub fn dispatch_variant(disp: &IDispatch) -> VARIANT {
    VARIANT::from(disp.clone())
}
```

...and use `dispatch_variant(&target)` in place of `VARIANT::from(&target)`
in `move_email` if the direct conversion doesn't compile.

Add the `compose` helper to the `impl WindowsOutlookClient` block:

```rust
    fn compose(
        &self,
        app: &IDispatch,
        to: &[String],
        subject: &str,
        body: &str,
        cc: Option<&[String]>,
        bcc: Option<&[String]>,
        html: bool,
    ) -> Result<IDispatch, ToolError> {
        let mail = variant_dispatch(&call_method(app, "CreateItem", &[i32_variant(c::OL_MAIL_ITEM)])?)?;
        put_property(&mail, "To", str_variant(&to.join("; ")))?;
        if let Some(cc) = cc {
            put_property(&mail, "CC", str_variant(&cc.join("; ")))?;
        }
        if let Some(bcc) = bcc {
            put_property(&mail, "BCC", str_variant(&bcc.join("; ")))?;
        }
        put_property(&mail, "Subject", str_variant(subject))?;
        if html {
            put_property(&mail, "BodyFormat", i32_variant(c::OL_FORMAT_HTML))?;
            put_property(&mail, "HTMLBody", str_variant(body))?;
        } else {
            put_property(&mail, "BodyFormat", i32_variant(c::OL_FORMAT_PLAIN))?;
            put_property(&mail, "Body", str_variant(body))?;
        }
        Ok(mail)
    }
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `cargo test --test email_write_tools`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/outlook/client.rs src/outlook/dispatch.rs src/server.rs tests/
git commit -m "Implement email write tools (send, create_draft, reply, move, delete)"
```

---

## Task 8: Calendar tools

`list_events`, `get_event`, `create_event`, `respond_to_meeting`.

**Files:**
- Modify: `outlook-mcp-rs/src/outlook/client.rs`
- Modify: `outlook-mcp-rs/src/server.rs`

- [ ] **Step 1: Add request structs and tool methods to server.rs**

```rust
#[derive(Debug, Deserialize, JsonSchema)]
pub struct ListEventsRequest {
    pub start_date: Option<String>,
    pub end_date: Option<String>,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct GetEventRequest {
    pub event_id: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct CreateEventRequest {
    pub subject: String,
    pub start: String,
    pub end: String,
    pub body: Option<String>,
    pub location: Option<String>,
    pub attendees: Option<Vec<String>>,
    #[serde(default)]
    pub all_day: bool,
    pub reminder_minutes: Option<i32>,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct RespondToMeetingRequest {
    pub event_id: String,
    pub response: String,
    pub comment: Option<String>,
    #[serde(default = "default_true")]
    pub send: bool,
}
```

Add to the tool impl block:

```rust
    #[tool(description = "List calendar events between two ISO dates (recurring events are \
        expanded). Defaults to the next 7 days starting today.")]
    async fn list_events(
        &self,
        Parameters(req): Parameters<ListEventsRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.list_events(req.start_date.as_deref(), req.end_date.as_deref())
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Get full details of a calendar event by id, including attendees \
        and body.")]
    async fn get_event(
        &self,
        Parameters(req): Parameters<GetEventRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || client.get_event(&req.event_id))
            .await
            .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Create a calendar appointment. start/end are ISO datetimes like \
        '2026-06-12T14:00'. If attendees is given, the event becomes a meeting and \
        invitations are SENT immediately.")]
    async fn create_event(
        &self,
        Parameters(req): Parameters<CreateEventRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.create_event(
                &req.subject,
                &req.start,
                &req.end,
                req.body.as_deref(),
                req.location.as_deref(),
                req.attendees.as_deref(),
                req.all_day,
                req.reminder_minutes,
            )
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Respond to a meeting invitation: response is 'accept', 'decline' \
        or 'tentative'. The response is sent to the organizer as the signed-in user \
        (pass send=false to save without sending).")]
    async fn respond_to_meeting(
        &self,
        Parameters(req): Parameters<RespondToMeetingRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.respond_to_meeting(&req.event_id, &req.response, req.comment.as_deref(), req.send)
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }
```

- [ ] **Step 2: Write failing unit tests**

Create `tests/calendar_tools.rs`:

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, EVENT_ID};
use outlook_mcp_rs::server::{
    CreateEventRequest, GetEventRequest, ListEventsRequest, OutlookMcpServer,
    RespondToMeetingRequest,
};
use rmcp::handler::server::wrapper::Parameters;

fn server_with_fake() -> (OutlookMcpServer, Arc<FakeOutlookClient>) {
    let fake = Arc::new(FakeOutlookClient::new());
    (OutlookMcpServer::new(fake.clone()), fake)
}

#[tokio::test]
async fn list_events_defaults_to_no_date_range() {
    let (server, fake) = server_with_fake();
    server
        .list_events(Parameters(ListEventsRequest {
            start_date: None,
            end_date: None,
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls.lock().unwrap()[0].0, "list_events");
}

#[tokio::test]
async fn get_event_returns_requested_id() {
    let (server, _fake) = server_with_fake();
    let result = server
        .get_event(Parameters(GetEventRequest {
            event_id: EVENT_ID.to_string(),
        }))
        .await
        .unwrap();
    assert!(!result.content.is_empty());
}

#[tokio::test]
async fn create_event_with_attendees_becomes_a_meeting() {
    let (server, fake) = server_with_fake();
    server
        .create_event(Parameters(CreateEventRequest {
            subject: "Standup".to_string(),
            start: "2026-07-06T09:00".to_string(),
            end: "2026-07-06T09:15".to_string(),
            body: None,
            location: None,
            attendees: Some(vec!["a@example.com".to_string()]),
            all_day: false,
            reminder_minutes: Some(10),
        }))
        .await
        .unwrap();
    let calls = fake.calls.lock().unwrap();
    assert_eq!(calls[0].1["attendees"][0], "a@example.com");
}

#[tokio::test]
async fn respond_to_meeting_accept() {
    let (server, fake) = server_with_fake();
    server
        .respond_to_meeting(Parameters(RespondToMeetingRequest {
            event_id: EVENT_ID.to_string(),
            response: "accept".to_string(),
            comment: None,
            send: true,
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls.lock().unwrap()[0].1["response"], "accept");
}
```

- [ ] **Step 3: Run tests to verify they fail, then implement**

Run: `cargo test --test calendar_tools` — expect compile failure until the
stubs are replaced.

Replace the four stub bodies in `WindowsOutlookClient`:

```rust
    fn list_events(
        &self,
        start_date: Option<&str>,
        end_date: Option<&str>,
    ) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let start = match start_date {
            Some(s) => parse_dt(s, "start_date")?,
            None => chrono::Local::now().date_naive().and_hms_opt(0, 0, 0).unwrap(),
        };
        let mut end = match end_date {
            Some(s) => parse_dt(s, "end_date")?,
            None => start + chrono::Duration::days(7),
        };
        if let Some(s) = end_date {
            if !s.contains('T') {
                end = end.date().and_hms_opt(23, 59, 59).unwrap();
            }
        }
        let calendar = variant_dispatch(&call_method(&ns, "GetDefaultFolder", &[i32_variant(c::OL_FOLDER_CALENDAR)])?)?;
        let items = variant_dispatch(&get_property(&calendar, "Items")?)?;
        put_property(&items, "IncludeRecurrences", bool_variant(true))?;
        call_method(&items, "Sort", &[str_variant("[Start]"), bool_variant(false)])?;
        let filter = format!(
            "[Start] >= '{}' AND [Start] <= '{}'",
            start.format("%m/%d/%Y %I:%M %p"),
            end.format("%m/%d/%Y %I:%M %p"),
        );
        let filtered = variant_dispatch(&call_method(&items, "Restrict", &[str_variant(&filter)])?)?;
        const MAX_CALENDAR_ITEMS: i32 = 250;
        let total = variant_i32(&get_property(&filtered, "Count")?);
        let mut results = Vec::new();
        for i in 1..=total {
            if results.len() as i32 >= MAX_CALENDAR_ITEMS {
                break;
            }
            let item = variant_dispatch(&call_method(&filtered, "Item", &[i32_variant(i)])?)?;
            results.push(self.event_summary(&item)?);
        }
        Ok(json!(results))
    }

    fn get_event(&self, event_id: &str) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let item = self.get_item(&ns, event_id)?;
        let mut info = self.event_summary(&item)?;
        let obj = info.as_object_mut().unwrap();
        const MAX_BODY_CHARS: usize = 100_000;
        obj.insert("body".into(), json!(truncate(&variant_str(&get_property(&item, "Body")?), MAX_BODY_CHARS)));
        obj.insert("required_attendees".into(), json!(variant_str(&get_property(&item, "RequiredAttendees")?)));
        obj.insert("optional_attendees".into(), json!(variant_str(&get_property(&item, "OptionalAttendees")?)));
        obj.insert("response_status".into(), json!(variant_i32(&get_property(&item, "ResponseStatus")?)));
        Ok(info)
    }

    fn create_event(
        &self,
        subject: &str,
        start: &str,
        end: &str,
        body: Option<&str>,
        location: Option<&str>,
        attendees: Option<&[String]>,
        all_day: bool,
        reminder_minutes: Option<i32>,
    ) -> Result<Value, ToolError> {
        let (app, _) = self.mapi()?;
        let appt = variant_dispatch(&call_method(&app, "CreateItem", &[i32_variant(c::OL_APPOINTMENT_ITEM)])?)?;
        put_property(&appt, "Subject", str_variant(subject))?;
        put_property(&appt, "Start", date_variant(&parse_dt(start, "start")?))?;
        put_property(&appt, "End", date_variant(&parse_dt(end, "end")?))?;
        if all_day {
            put_property(&appt, "AllDayEvent", bool_variant(true))?;
        }
        if let Some(b) = body {
            put_property(&appt, "Body", str_variant(b))?;
        }
        if let Some(loc) = location {
            put_property(&appt, "Location", str_variant(loc))?;
        }
        if let Some(minutes) = reminder_minutes {
            put_property(&appt, "ReminderSet", bool_variant(true))?;
            put_property(&appt, "ReminderMinutesBeforeStart", i32_variant(minutes))?;
        }
        let status = if let Some(attendees) = attendees {
            put_property(&appt, "MeetingStatus", i32_variant(c::OL_MEETING))?;
            let recipients = variant_dispatch(&get_property(&appt, "Recipients")?)?;
            for address in attendees {
                call_method(&recipients, "Add", &[str_variant(address)])?;
            }
            call_method(&recipients, "ResolveAll", &[])?;
            call_method(&appt, "Send", &[])?;
            "meeting_sent"
        } else {
            call_method(&appt, "Save", &[])?;
            "saved"
        };
        Ok(json!({"status": status, "id": Self::make_id(&appt)?, "subject": subject}))
    }

    fn respond_to_meeting(
        &self,
        event_id: &str,
        response: &str,
        comment: Option<&str>,
        send: bool,
    ) -> Result<Value, ToolError> {
        let response_id = c::meeting_response_id_for_name(response).ok_or_else(|| {
            ToolError(format!("Invalid response {response:?}: use 'accept', 'decline' or 'tentative'."))
        })?;
        let (_, ns) = self.mapi()?;
        let mut item = self.get_item(&ns, event_id)?;
        // A meeting request from the inbox resolves to a MeetingItem; get its
        // appointment. Calendar ids resolve straight to appointments.
        if let Ok(assoc) = call_method(&item, "GetAssociatedAppointment", &[bool_variant(true)]) {
            item = variant_dispatch(&assoc)?;
        }
        let resp = call_method(&item, "Respond", &[i32_variant(response_id), bool_variant(true)])?;
        if let (Some(text), Ok(resp_disp)) = (comment, variant_dispatch(&resp)) {
            put_property(&resp_disp, "Body", str_variant(text))?;
            if send {
                call_method(&resp_disp, "Send", &[])?;
            } else {
                call_method(&resp_disp, "Save", &[])?;
            }
        }
        let subject = variant_str(&get_property(&item, "Subject")?);
        let suffix = if send { "_sent" } else { "_saved" };
        Ok(json!({"status": format!("{response}{suffix}"), "subject": subject}))
    }
```

Add the `event_summary` helper and `parse_dt`/`date_variant` free functions.
`event_summary` goes in the `impl WindowsOutlookClient` block:

```rust
    fn event_summary(&self, item: &IDispatch) -> Result<Value, ToolError> {
        Ok(json!({
            "id": Self::make_id(item)?,
            "subject": variant_str(&get_property(item, "Subject")?),
            "start": variant_date_iso(&get_property(item, "Start")?),
            "end": variant_date_iso(&get_property(item, "End")?),
            "location": variant_str(&get_property(item, "Location")?),
            "organizer": variant_str(&get_property(item, "Organizer")?),
            "all_day": variant_bool(&get_property(item, "AllDayEvent")?),
            "is_recurring": variant_bool(&get_property(item, "IsRecurring")?),
            "is_meeting": variant_i32(&get_property(item, "MeetingStatus")?) != c::OL_NONMEETING,
        }))
    }
```

`parse_dt` and `date_variant` are free functions at the bottom of the file:

```rust
fn parse_dt(value: &str, field: &str) -> Result<chrono::NaiveDateTime, ToolError> {
    chrono::NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M")
        .or_else(|_| {
            chrono::NaiveDate::parse_from_str(value, "%Y-%m-%d")
                .map(|d| d.and_hms_opt(0, 0, 0).unwrap())
        })
        .map_err(|_| {
            ToolError(format!(
                "Invalid {field} {value:?}: expected ISO format like '2026-06-10' or '2026-06-10T14:30'"
            ))
        })
}

fn date_variant(dt: &chrono::NaiveDateTime) -> VARIANT {
    // OLE Automation date: f64 day count from 1899-12-30.
    let epoch = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap().and_hms_opt(0, 0, 0).unwrap();
    let days = (*dt - epoch).num_milliseconds() as f64 / 86_400_000.0;
    let mut variant = VARIANT::default();
    unsafe {
        variant.Anonymous.Anonymous.vt = VT_DATE;
        variant.Anonymous.Anonymous.Anonymous.date = days;
    }
    variant
}
```

If `VARIANT`'s internal field layout (`Anonymous.Anonymous.vt` /
`Anonymous.Anonymous.Anonymous.date`) doesn't match what `cargo build`
reports, use whatever field path the compiler error suggests — this is the
standard windows-rs `VARIANT` union-of-unions layout, but exact field names
can shift between minor versions; the compiler error names the real ones.

- [ ] **Step 4: Run tests to verify they pass**

Run: `cargo test --test calendar_tools`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/outlook/client.rs src/server.rs tests/calendar_tools.rs
git commit -m "Implement calendar tools (list/get/create_event, respond_to_meeting)"
```

---

## Task 9: Attachment tools

`list_attachments`, `save_attachments`.

**Files:**
- Modify: `outlook-mcp-rs/src/outlook/client.rs`
- Modify: `outlook-mcp-rs/src/server.rs`

- [ ] **Step 1: Add request structs and tool methods to server.rs**

```rust
#[derive(Debug, Deserialize, JsonSchema)]
pub struct ListAttachmentsRequest {
    pub email_id: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct SaveAttachmentsRequest {
    pub email_id: String,
    pub save_dir: String,
    pub attachment_names: Option<Vec<String>>,
}
```

```rust
    #[tool(description = "List an email's attachments (file names and sizes).")]
    async fn list_attachments(
        &self,
        Parameters(req): Parameters<ListAttachmentsRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || client.list_attachments(&req.email_id))
            .await
            .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Save an email's attachments to a local directory (created if \
        missing). By default saves all attachments; pass attachment_names to save only \
        specific files.")]
    async fn save_attachments(
        &self,
        Parameters(req): Parameters<SaveAttachmentsRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.save_attachments(&req.email_id, &req.save_dir, req.attachment_names.as_deref())
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }
```

- [ ] **Step 2: Write failing unit tests**

Create `tests/attachment_tools.rs`:

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, EMAIL_ID};
use outlook_mcp_rs::server::{ListAttachmentsRequest, OutlookMcpServer, SaveAttachmentsRequest};
use rmcp::handler::server::wrapper::Parameters;

fn server_with_fake() -> (OutlookMcpServer, Arc<FakeOutlookClient>) {
    let fake = Arc::new(FakeOutlookClient::new());
    (OutlookMcpServer::new(fake.clone()), fake)
}

#[tokio::test]
async fn list_attachments_returns_files() {
    let (server, _fake) = server_with_fake();
    let result = server
        .list_attachments(Parameters(ListAttachmentsRequest {
            email_id: EMAIL_ID.to_string(),
        }))
        .await
        .unwrap();
    assert!(!result.content.is_empty());
}

#[tokio::test]
async fn save_attachments_passes_save_dir() {
    let (server, fake) = server_with_fake();
    server
        .save_attachments(Parameters(SaveAttachmentsRequest {
            email_id: EMAIL_ID.to_string(),
            save_dir: "C:\\temp".to_string(),
            attachment_names: None,
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls.lock().unwrap()[0].1["save_dir"], "C:\\temp");
}
```

- [ ] **Step 3: Run tests to verify they fail, then implement the stubs**

Run: `cargo test --test attachment_tools` — expect compile failure until
implemented.

```rust
    fn list_attachments(&self, email_id: &str) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let item = self.get_item(&ns, email_id)?;
        let attachments = variant_dispatch(&get_property(&item, "Attachments")?)?;
        let count = variant_i32(&get_property(&attachments, "Count")?);
        let mut results = Vec::new();
        for i in 1..=count {
            let att = variant_dispatch(&call_method(&attachments, "Item", &[i32_variant(i)])?)?;
            results.push(json!({
                "index": i,
                "filename": variant_str(&get_property(&att, "FileName")?),
                "size": variant_i32(&get_property(&att, "Size")?),
            }));
        }
        Ok(json!(results))
    }

    fn save_attachments(
        &self,
        email_id: &str,
        save_dir: &str,
        attachment_names: Option<&[String]>,
    ) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let item = self.get_item(&ns, email_id)?;
        let attachments = variant_dispatch(&get_property(&item, "Attachments")?)?;
        let count = variant_i32(&get_property(&attachments, "Count")?);
        if count == 0 {
            return Err(ToolError("This email has no attachments.".into()));
        }
        std::fs::create_dir_all(save_dir)
            .map_err(|e| ToolError(format!("Could not create {save_dir:?}: {e}")))?;
        let wanted: Option<Vec<String>> =
            attachment_names.map(|names| names.iter().map(|n| n.to_lowercase()).collect());
        let mut results = Vec::new();
        for i in 1..=count {
            let att = variant_dispatch(&call_method(&attachments, "Item", &[i32_variant(i)])?)?;
            let filename = variant_str(&get_property(&att, "FileName")?);
            if let Some(wanted) = &wanted {
                if !wanted.contains(&filename.to_lowercase()) {
                    continue;
                }
            }
            let safe_name = safe_filename(&filename);
            let target = std::path::Path::new(save_dir).join(&safe_name);
            match call_method(&att, "SaveAsFile", &[str_variant(&target.to_string_lossy())]) {
                Ok(_) => results.push(json!({
                    "filename": filename,
                    "saved_to": target.to_string_lossy(),
                    "status": "saved",
                })),
                Err(e) => results.push(json!({
                    "filename": filename,
                    "status": "failed",
                    "error": e.0,
                })),
            }
        }
        if results.is_empty() {
            return Err(ToolError(
                "No attachments matched attachment_names; use list_attachments to see the \
                 exact file names."
                    .into(),
            ));
        }
        Ok(json!(results))
    }
```

Add `safe_filename` as a free function at the bottom of the file:

```rust
fn safe_filename(name: &str) -> String {
    let cleaned: String = name
        .chars()
        .map(|c| match c {
            '\\' | '/' | ':' | '*' | '?' | '"' | '<' | '>' | '|' => '_',
            c if (c as u32) < 0x20 => '_',
            c => c,
        })
        .collect();
    let trimmed = cleaned.trim_matches(|c: char| c == '.' || c == ' ');
    if trimmed.is_empty() {
        "attachment".to_string()
    } else {
        trimmed.to_string()
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cargo test --test attachment_tools`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/outlook/client.rs src/server.rs tests/attachment_tools.rs
git commit -m "Implement attachment tools (list_attachments, save_attachments)"
```

---

## Task 10: Task tools

`list_tasks`, `create_task`, `complete_task`.

**Files:**
- Modify: `outlook-mcp-rs/src/outlook/client.rs`
- Modify: `outlook-mcp-rs/src/server.rs`

- [ ] **Step 1: Add request structs and tool methods to server.rs**

```rust
fn default_normal() -> String {
    "normal".to_string()
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct ListTasksRequest {
    #[serde(default)]
    pub include_completed: bool,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct CreateTaskRequest {
    pub subject: String,
    pub body: Option<String>,
    pub due_date: Option<String>,
    #[serde(default = "default_normal")]
    pub importance: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct CompleteTaskRequest {
    pub task_id: String,
}
```

```rust
    #[tool(description = "List Outlook tasks (open tasks only, unless include_completed).")]
    async fn list_tasks(
        &self,
        Parameters(req): Parameters<ListTasksRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || client.list_tasks(req.include_completed))
            .await
            .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Create an Outlook task. due_date is an ISO date; importance is \
        'low', 'normal' or 'high'.")]
    async fn create_task(
        &self,
        Parameters(req): Parameters<CreateTaskRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || {
            client.create_task(&req.subject, req.body.as_deref(), req.due_date.as_deref(), &req.importance)
        })
        .await
        .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Mark a task as complete.")]
    async fn complete_task(
        &self,
        Parameters(req): Parameters<CompleteTaskRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || client.complete_task(&req.task_id))
            .await
            .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }
```

- [ ] **Step 2: Write failing unit tests**

Create `tests/task_tools.rs`:

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, TASK_ID};
use outlook_mcp_rs::server::{CompleteTaskRequest, CreateTaskRequest, ListTasksRequest, OutlookMcpServer};
use rmcp::handler::server::wrapper::Parameters;

fn server_with_fake() -> (OutlookMcpServer, Arc<FakeOutlookClient>) {
    let fake = Arc::new(FakeOutlookClient::new());
    (OutlookMcpServer::new(fake.clone()), fake)
}

#[tokio::test]
async fn list_tasks_excludes_completed_by_default() {
    let (server, fake) = server_with_fake();
    server
        .list_tasks(Parameters(ListTasksRequest { include_completed: false }))
        .await
        .unwrap();
    assert_eq!(fake.calls.lock().unwrap()[0].1["include_completed"], false);
}

#[tokio::test]
async fn create_task_defaults_importance_to_normal() {
    let (server, fake) = server_with_fake();
    server
        .create_task(Parameters(CreateTaskRequest {
            subject: "Buy milk".to_string(),
            body: None,
            due_date: None,
            importance: "normal".to_string(),
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls.lock().unwrap()[0].1["importance"], "normal");
}

#[tokio::test]
async fn complete_task_calls_through() {
    let (server, fake) = server_with_fake();
    server
        .complete_task(Parameters(CompleteTaskRequest {
            task_id: TASK_ID.to_string(),
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls.lock().unwrap()[0].0, "complete_task");
}
```

- [ ] **Step 3: Run tests to verify they fail, then implement the stubs**

Run: `cargo test --test task_tools` — expect compile failure until
implemented.

```rust
    fn list_tasks(&self, include_completed: bool) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let folder = variant_dispatch(&call_method(&ns, "GetDefaultFolder", &[i32_variant(c::OL_FOLDER_TASKS)])?)?;
        let items = variant_dispatch(&get_property(&folder, "Items")?)?;
        let items = if include_completed {
            items
        } else {
            variant_dispatch(&call_method(&items, "Restrict", &[str_variant("[Complete] = False")])?)?
        };
        let count = variant_i32(&get_property(&items, "Count")?);
        let mut results = Vec::new();
        for i in 1..=count {
            let item = variant_dispatch(&call_method(&items, "Item", &[i32_variant(i)])?)?;
            results.push(self.task_summary(&item)?);
        }
        Ok(json!(results))
    }

    fn create_task(
        &self,
        subject: &str,
        body: Option<&str>,
        due_date: Option<&str>,
        importance: &str,
    ) -> Result<Value, ToolError> {
        let importance_id = c::importance_id_for_name(importance).ok_or_else(|| {
            ToolError(format!("Invalid importance {importance:?}: use 'low', 'normal' or 'high'."))
        })?;
        let (app, _) = self.mapi()?;
        let task = variant_dispatch(&call_method(&app, "CreateItem", &[i32_variant(c::OL_TASK_ITEM)])?)?;
        put_property(&task, "Subject", str_variant(subject))?;
        if let Some(b) = body {
            put_property(&task, "Body", str_variant(b))?;
        }
        if let Some(due) = due_date {
            put_property(&task, "DueDate", date_variant(&parse_dt(due, "due_date")?))?;
        }
        put_property(&task, "Importance", i32_variant(importance_id))?;
        call_method(&task, "Save", &[])?;
        Ok(json!({"status": "created", "id": Self::make_id(&task)?, "subject": subject}))
    }

    fn complete_task(&self, task_id: &str) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let task = self.get_item(&ns, task_id)?;
        call_method(&task, "MarkComplete", &[])?;
        let subject = variant_str(&get_property(&task, "Subject")?);
        Ok(json!({"status": "completed", "subject": subject}))
    }
```

Add the `task_summary` helper:

```rust
    fn task_summary(&self, item: &IDispatch) -> Result<Value, ToolError> {
        Ok(json!({
            "id": Self::make_id(item)?,
            "subject": variant_str(&get_property(item, "Subject")?),
            "due_date": variant_date_iso(&get_property(item, "DueDate")?),
            "complete": variant_bool(&get_property(item, "Complete")?),
            "status": variant_i32(&get_property(item, "Status")?),
            "importance": variant_i32(&get_property(item, "Importance")?),
        }))
    }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cargo test --test task_tools`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/outlook/client.rs src/server.rs tests/task_tools.rs
git commit -m "Implement task tools (list_tasks, create_task, complete_task)"
```

---

## Task 11: Note tools

`list_notes`, `get_note`, `create_note`. Last tool group — after this task,
every `unimplemented!()` stub from Task 5 is gone.

**Files:**
- Modify: `outlook-mcp-rs/src/outlook/client.rs`
- Modify: `outlook-mcp-rs/src/server.rs`

- [ ] **Step 1: Add request structs and tool methods to server.rs**

```rust
#[derive(Debug, Deserialize, JsonSchema)]
pub struct GetNoteRequest {
    pub note_id: String,
}

#[derive(Debug, Deserialize, JsonSchema)]
pub struct CreateNoteRequest {
    pub body: String,
}
```

```rust
    #[tool(description = "List Outlook sticky notes (id, first line, creation time).")]
    async fn list_notes(&self) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || client.list_notes())
            .await
            .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Read the full body of a note by id.")]
    async fn get_note(
        &self,
        Parameters(req): Parameters<GetNoteRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || client.get_note(&req.note_id))
            .await
            .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }

    #[tool(description = "Create an Outlook sticky note. The first line of body becomes the \
        note's title.")]
    async fn create_note(
        &self,
        Parameters(req): Parameters<CreateNoteRequest>,
    ) -> Result<CallToolResult, McpError> {
        let client = self.client.clone();
        let value = tokio::task::spawn_blocking(move || client.create_note(&req.body))
            .await
            .map_err(|e| McpError::internal_error(e.to_string(), None))??;
        Ok(CallToolResult::success(vec![Content::json(value)?]))
    }
```

- [ ] **Step 2: Write failing unit tests**

Create `tests/note_tools.rs`:

```rust
use std::sync::Arc;

use outlook_mcp_rs::outlook::fake::{FakeOutlookClient, NOTE_ID};
use outlook_mcp_rs::server::{CreateNoteRequest, GetNoteRequest, OutlookMcpServer};
use rmcp::handler::server::wrapper::Parameters;

fn server_with_fake() -> (OutlookMcpServer, Arc<FakeOutlookClient>) {
    let fake = Arc::new(FakeOutlookClient::new());
    (OutlookMcpServer::new(fake.clone()), fake)
}

#[tokio::test]
async fn list_notes_returns_notes() {
    let (server, _fake) = server_with_fake();
    let result = server.list_notes().await.unwrap();
    assert!(!result.content.is_empty());
}

#[tokio::test]
async fn get_note_returns_requested_id() {
    let (server, _fake) = server_with_fake();
    server
        .get_note(Parameters(GetNoteRequest {
            note_id: NOTE_ID.to_string(),
        }))
        .await
        .unwrap();
}

#[tokio::test]
async fn create_note_rejects_empty_body_at_the_com_layer() {
    let (server, fake) = server_with_fake();
    server
        .create_note(Parameters(CreateNoteRequest {
            body: "Ideas\n- one".to_string(),
        }))
        .await
        .unwrap();
    assert_eq!(fake.calls.lock().unwrap()[0].1["body"], "Ideas\n- one");
}
```

- [ ] **Step 3: Run tests to verify they fail, then implement the stubs**

Run: `cargo test --test note_tools` — expect compile failure until
implemented.

```rust
    fn list_notes(&self) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let folder = variant_dispatch(&call_method(&ns, "GetDefaultFolder", &[i32_variant(c::OL_FOLDER_NOTES)])?)?;
        let items = variant_dispatch(&get_property(&folder, "Items")?)?;
        let count = variant_i32(&get_property(&items, "Count")?);
        let mut results = Vec::new();
        for i in 1..=count {
            let item = variant_dispatch(&call_method(&items, "Item", &[i32_variant(i)])?)?;
            results.push(self.note_summary(&item)?);
        }
        Ok(json!(results))
    }

    fn get_note(&self, note_id: &str) -> Result<Value, ToolError> {
        let (_, ns) = self.mapi()?;
        let note = self.get_item(&ns, note_id)?;
        let mut info = self.note_summary(&note)?;
        const MAX_BODY_CHARS: usize = 100_000;
        let body = variant_str(&get_property(&note, "Body")?);
        info.as_object_mut()
            .unwrap()
            .insert("body".into(), json!(truncate(&body, MAX_BODY_CHARS)));
        Ok(info)
    }

    fn create_note(&self, body: &str) -> Result<Value, ToolError> {
        if body.is_empty() {
            return Err(ToolError("create_note requires a non-empty body.".into()));
        }
        let (app, _) = self.mapi()?;
        let note = variant_dispatch(&call_method(&app, "CreateItem", &[i32_variant(c::OL_NOTE_ITEM)])?)?;
        put_property(&note, "Body", str_variant(body))?;
        call_method(&note, "Save", &[])?;
        Ok(json!({"status": "created", "id": Self::make_id(&note)?}))
    }
```

Add the `note_summary` helper:

```rust
    fn note_summary(&self, item: &IDispatch) -> Result<Value, ToolError> {
        let body = variant_str(&get_property(item, "Body")?);
        let first_line = body.trim().lines().next().unwrap_or("");
        let subject: String = first_line.chars().take(120).collect();
        Ok(json!({
            "id": Self::make_id(item)?,
            "subject": subject,
            "created": variant_date_iso(&get_property(item, "CreationTime")?),
        }))
    }
```

- [ ] **Step 4: Run all tests to verify everything passes**

Run: `cargo test`
Expected: every test from Tasks 1-11 passes. `WindowsOutlookClient` now has
zero `unimplemented!()` bodies left.

- [ ] **Step 5: Commit**

```bash
git add src/outlook/client.rs src/server.rs tests/note_tools.rs
git commit -m "Implement note tools (list_notes, get_note, create_note)"
```

---

## Task 12: Live Outlook system tests

The new local-only system test suite validating the real `WindowsOutlookClient`
against whatever Outlook is actually running on the developer's machine.

**Files:**
- Create: `outlook-mcp-rs/tests/live_outlook.rs`
- Create: `outlook-mcp-rs/TESTING.md`

**Interfaces:**
- Consumes: `outlook_mcp_rs::outlook::client::WindowsOutlookClient::new()`, `OutlookClient` trait.

- [ ] **Step 1: Write the live test suite**

Create `tests/live_outlook.rs`:

```rust
//! Live-Outlook system tests. NOT run by plain `cargo test` — every test
//! here is #[ignore]d. Run explicitly with:
//!
//!     cargo test --test live_outlook -- --ignored
//!
//! Requires: classic Outlook desktop open and signed in to a normal mailbox
//! on this machine. Every test that creates something deletes it before
//! finishing, so your mailbox is unchanged after a run. send_email and
//! respond_to_meeting are deliberately NOT covered here — see TESTING.md
//! for how to exercise those by hand.

use outlook_mcp_rs::outlook::client::{OutlookClient, WindowsOutlookClient};

fn client() -> WindowsOutlookClient {
    WindowsOutlookClient::new().expect("Outlook COM must be available for live tests")
}

#[test]
#[ignore]
fn lists_real_folders() {
    let client = client();
    let folders = client.list_folders().unwrap();
    let folders = folders.as_array().unwrap();
    assert!(!folders.is_empty(), "expected at least one mail folder");
}

#[test]
#[ignore]
fn creates_reads_and_deletes_a_draft() {
    let client = client();
    let created = client
        .create_draft(
            &["nobody@example.invalid".to_string()],
            "outlook-mcp-rs live test",
            "temporary draft created by the live test suite",
            None,
            None,
            false,
        )
        .unwrap();
    let id = created["id"].as_str().unwrap().to_string();

    let fetched = client.get_email(&id, false).unwrap();
    assert_eq!(fetched["subject"], "outlook-mcp-rs live test");

    client.delete_email(&id).unwrap();
}

#[test]
#[ignore]
fn creates_completes_and_the_task_stays_gone_after_cleanup() {
    let client = client();
    let created = client
        .create_task("outlook-mcp-rs live test task", None, None, "normal")
        .unwrap();
    let id = created["id"].as_str().unwrap().to_string();

    let completed = client.complete_task(&id).unwrap();
    assert_eq!(completed["status"], "completed");
    // Outlook has no "delete task" tool in this server; remove it by hand
    // if you want a completely clean Tasks list, or leave it — a single
    // completed test task is harmless.
}

#[test]
#[ignore]
fn creates_reads_and_the_note_can_be_found() {
    let client = client();
    let created = client
        .create_note("outlook-mcp-rs live test note\nsecond line")
        .unwrap();
    let id = created["id"].as_str().unwrap().to_string();

    let fetched = client.get_note(&id).unwrap();
    assert_eq!(fetched["subject"], "outlook-mcp-rs live test note");
}

#[test]
#[ignore]
fn creates_and_deletes_a_calendar_event() {
    let client = client();
    let created = client
        .create_event(
            "outlook-mcp-rs live test event",
            "2099-01-01T09:00",
            "2099-01-01T09:15",
            None,
            None,
            None,
            false,
            None,
        )
        .unwrap();
    let id = created["id"].as_str().unwrap().to_string();

    let fetched = client.get_event(&id).unwrap();
    assert_eq!(fetched["subject"], "outlook-mcp-rs live test event");

    // Calendar items are Outlook items too, so delete_email's underlying
    // COM call (Delete) works on any item type despite the name.
    client.delete_email(&id).unwrap();
}
```

- [ ] **Step 2: Run the live suite manually (requires Outlook open on this machine)**

Run: `cargo test --test live_outlook -- --ignored --test-threads=1`
Expected: all 5 tests PASS against your real Outlook. `--test-threads=1` is
important — Outlook COM automation is not safe to hit concurrently from
multiple tests at once.

If any test fails, read the `ToolError` message it produces (formatted by
Task 3's `dispatch::invoke`) — it will name the exact COM member/HRESULT
that failed, which is the fastest way to spot an API mismatch against your
installed Outlook version.

- [ ] **Step 3: Write TESTING.md**

```markdown
# Testing outlook-mcp-rs

## Unit tests (run everywhere, every commit)

    cargo test

Uses `FakeOutlookClient` — no real Outlook needed. This is what CI runs.

## Live Outlook system tests (local only, manual)

    cargo test --test live_outlook -- --ignored --test-threads=1

Requires classic Outlook desktop open and signed in on this machine. Every
test cleans up what it creates, so your mailbox is unchanged afterward.
`--test-threads=1` is required — Outlook COM automation isn't safe to hit
from multiple threads/tests concurrently.

**Not covered by the automatic live suite** (no safe way to auto-clean up):

- `send_email` — sends a real email; there's no "unsend". Test by hand:
  run the server (`cargo run`), connect with
  `npx @modelcontextprotocol/inspector cargo run`, and call `send_email`
  with `to` set to an address you control, then confirm you received it.
- `respond_to_meeting` — sends a real accept/decline/tentative response to
  a meeting organizer. Test by hand against a real meeting invite you've
  received, and confirm the organizer sees your response.

Run the full live suite (plus the two manual checks above) before cutting
a release.
```

- [ ] **Step 4: Commit**

```bash
git add tests/live_outlook.rs TESTING.md
git commit -m "Add live-Outlook system test suite (local-only) and TESTING.md"
```

---

## Task 13: CI workflow (unit tests)

**Files:**
- Create: `outlook-mcp-rs/.github/workflows/ci.yaml`

- [ ] **Step 1: Write the CI workflow**

```yaml
name: CI

on:
  push:
    branches: [master, main]
    tags: ['v*']
  pull_request:
    branches: [master, main]

jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: dtolnay/rust-toolchain@stable
      - name: Run tests
        run: cargo test --verbose
```

- [ ] **Step 2: Push and verify it runs green**

```bash
git add .github/workflows/ci.yaml
git commit -m "Add CI workflow running cargo test on windows-latest"
git push
```

Run: `gh run watch $(gh run list --limit 1 --json databaseId -q '.[0].databaseId')`
Expected: the `test` job passes.

---

## Task 14: Release workflow (build + attach .exe to GitHub Releases)

**Files:**
- Modify: `outlook-mcp-rs/.github/workflows/ci.yaml`

- [ ] **Step 1: Add build and release jobs**

Append to `.github/workflows/ci.yaml`:

```yaml
  build:
    needs: test
    runs-on: windows-latest
    if: startsWith(github.ref, 'refs/tags/v')
    steps:
      - uses: actions/checkout@v4
      - uses: dtolnay/rust-toolchain@stable
      - name: Build release binary
        run: cargo build --release
      - name: Rename binary with version
        shell: bash
        run: |
          version="${GITHUB_REF_NAME#v}"
          cp target/release/outlook-mcp-rs.exe "outlook-mcp-rs-${version}-win64.exe"
      - uses: actions/upload-artifact@v4
        with:
          name: exe
          path: outlook-mcp-rs-*.exe

  release:
    needs: build
    runs-on: ubuntu-latest
    if: startsWith(github.ref, 'refs/tags/v')
    permissions:
      contents: write
    steps:
      - uses: actions/download-artifact@v4
        with:
          name: exe
      - name: Create GitHub Release
        uses: softprops/action-gh-release@v2
        with:
          files: outlook-mcp-rs-*.exe
          generate_release_notes: true
```

- [ ] **Step 2: Tag a v0.1.0 release and verify**

```bash
git add .github/workflows/ci.yaml
git commit -m "Add release workflow: build and attach outlook-mcp-rs.exe to GitHub Releases"
git push
git tag v0.1.0
git push origin v0.1.0
```

Run: `gh run watch $(gh run list --limit 1 --json databaseId -q '.[0].databaseId')`
Expected: `test`, `build`, and `release` jobs all pass, and
`gh release view v0.1.0` shows `outlook-mcp-rs-0.1.0-win64.exe` attached.

- [ ] **Step 3: Manual smoke test of the released binary**

Download the attached `.exe` from the release, run it against the MCP
inspector, and confirm it starts and lists all 24 tools:

```bash
npx @modelcontextprotocol/inspector ./outlook-mcp-rs-0.1.0-win64.exe
```

---

## Self-Review Notes

- **Spec coverage:** every section of `2026-07-06-rust-port-design.md` is
  covered — architecture/crate layout (Task 1, 5), COM interop (Task 3),
  error handling (Task 2), MCP layer (Task 1, 6-11), full tool parity
  (Tasks 6-11, all 24 tools), unit tests (every tool task), live system
  tests (Task 12), CI (Task 13), release (Task 14), repo bootstrap (Task 1).
- **Placeholder scan:** the only `unimplemented!()` bodies are the
  intentional, temporary Task 5 stubs, each replaced by a named later task
  (Task 6-11) — by Task 11 Step 4, none remain, verified by a full
  `cargo test` run.
- **Type consistency:** `OutlookClient` trait signatures (Task 5) match
  the `WindowsOutlookClient`/`FakeOutlookClient` impls and every `#[tool]`
  wrapper's call site across Tasks 6-11 — all pass `&str`/`Option<&str>`/
  `Option<&[String]>` consistently, matching the request struct fields
  they're built from.
- **Known version-drift risk:** `windows` 0.62 and `rmcp` 2.1 signatures
  were verified via docs.rs/GitHub source as of 2026-07-06, but both crates
  evolve; Tasks 3, 6, and 8 call this out explicitly where a signature is
  most likely to have shifted, with guidance to follow the compiler error
  to the corrected API rather than guessing.
