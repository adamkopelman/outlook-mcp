# outlook-mcp

An [MCP](https://modelcontextprotocol.io) server that lets Claude (or any MCP
client) control **Microsoft Outlook desktop on Windows** through the Win32 COM
API — read and send email, manage your calendar, save attachments, and work
with tasks and notes, all against the Outlook profile you are already signed
in to. No Azure app registration, no Graph API tokens.

## Requirements

- **Windows** with **classic Outlook desktop** installed and a configured
  mail profile.
  > ⚠️ The "new Outlook" (`olk.exe`) does **not** expose a COM API and will
  > not work. You need classic Outlook (Microsoft 365 / Office 2016+).
- **Python 3.10+**
- The server only *runs* on Windows; the test suite runs anywhere (COM access
  is mocked).

## Installation

```bash
git clone https://github.com/adamkopelman/outlook-mcp.git
cd outlook-mcp
pip install .
```

This installs the `outlook-mcp` console command (and `pywin32` on Windows).

## Hooking it up to Claude

**As a Claude Code plugin** (recommended for Claude Code) — this repo is itself a
plugin (`.claude-plugin/plugin.json`) that registers the `outlook` MCP server
for you. After `pip install .` (above), either:

```bash
# Try it locally without installing anything into Claude Code's config:
claude --plugin-dir /path/to/outlook-mcp

# Or install it properly, from a local checkout or directly from GitHub:
claude plugin install outlook-mcp@/path/to/outlook-mcp
claude plugin install outlook-mcp@github.com/adamkopelman/outlook-mcp
```

The plugin launches the server as `python -m outlook_mcp` rather than via the
`outlook-mcp` console script, so it works even if pip's script directory isn't
on `PATH` — it just needs `outlook_mcp` importable by whichever `python` is
first on `PATH`.

**Claude Desktop** — add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "outlook-mcp"
    }
  }
}
```

If `outlook-mcp` isn't on PATH, use the full interpreter instead:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Path\\To\\python.exe",
      "args": ["-m", "outlook_mcp"]
    }
  }
}
```

**Claude Code (manual, without the plugin):**

```bash
claude mcp add outlook -- outlook-mcp
```

## Tools

| Tool | Description |
| --- | --- |
| `list_folders` | List mail folders with item/unread counts |
| `list_emails` | Recent emails in a folder (newest first, `unread_only` option) |
| `search_emails` | Search subject/sender/body, optional `since_days` window |
| `get_email` | Full email by id (body, recipients, attachment names) |
| `send_email` | **Send immediately** (to/cc/bcc, plain or HTML body) |
| `create_draft` | Compose and save to Drafts without sending |
| `reply_email` | Reply / reply-all (send, or save as draft with `send=false`) |
| `move_email` | Move an email to another folder (returns its new id) |
| `delete_email` | Move an email to Deleted Items |
| `list_events` | Calendar events in a date range (recurrences expanded) |
| `get_event` | Full event details including attendees |
| `create_event` | Create an appointment — adding `attendees` sends invites |
| `respond_to_meeting` | Accept / decline / tentative a meeting invite |
| `list_attachments` | List an email's attachments |
| `save_attachments` | Save attachments to a local directory |
| `list_tasks` | List tasks (open only by default) |
| `create_task` | Create a task with due date and importance |
| `complete_task` | Mark a task complete |
| `list_notes` | List sticky notes |
| `get_note` | Read a note's full body |
| `create_note` | Create a sticky note |

Items are addressed by an opaque `id` returned from list/search tools.
**Ids change when an item moves folders** — `move_email` returns the new id,
and a stale id produces a clear "item not found" error.

## Security notes

- `send_email`, `reply_email`, `create_event` (with attendees) and
  `respond_to_meeting` act **immediately as the signed-in Outlook user**,
  with no confirmation step inside the server. If you want a human in the
  loop, prefer `create_draft` / `send=false`, or deny the sending tools in
  your MCP client's permission settings.
- Outlook's object model guard may show a *"A program is trying to send an
  e-mail message on your behalf"* prompt, typically when no up-to-date
  antivirus is registered with Windows or group policy demands it. The
  prompt blocks the tool call until answered. Do **not** disable Outlook
  security to avoid it.
- `save_attachments` writes to any local path the MCP client asks for.

## Troubleshooting

- *"Outlook is not available..."* — you're not on Windows, or classic
  Outlook isn't installed.
- The first tool call may take a few seconds while Outlook launches.
- *"Item not found"* — the id went stale (item moved or was deleted);
  list/search again to get a fresh id.
- Date filters use Outlook's JET format internally; if `search_emails` with
  `since_days` misbehaves on a heavily localized system, try without it and
  filter by eye.

## Development

```bash
pip install -e .[dev]
pytest
```

The COM layer lives in `outlook_mcp/outlook/client.py` behind the
`OutlookClientBase` interface (`outlook_mcp/outlook/base.py`); tests run on
any OS against an in-memory fake (`tests/conftest.py`).

## License

MIT — see [LICENSE](LICENSE).
