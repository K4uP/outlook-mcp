# outlook-mcp (extended fork)

MCP server for Microsoft Outlook on Windows via `win32com` — lets AI assistants
(Claude Code and other MCP clients) work with email, contacts, **calendars, and
tasks** through a locally running Outlook instance. No Azure AD, no Graph API,
no OAuth — it drives the Outlook desktop client over COM.

This is an extended fork of
[lihaokun/outlook-mcp](https://github.com/lihaokun/outlook-mcp) (PyPI:
`outlook-mcp-server` 0.1.1, MIT). Upstream ships 14 tools; this fork ships 21
and fixes three classes of bugs that bite in production. It is **not published
to PyPI** — install from this repository (see below) and do not run
`pip install -U outlook-mcp-server` afterwards, or PyPI will overwrite it.

## What this fork adds over upstream 0.1.1

**New tools (7)** — upstream had only `listCalendars` + `createEvent` for
calendars and nothing for tasks:

- `listEvents`, `getEvent`, `updateEvent`, `deleteEvent`
- `listTasks`, `createTask`, `updateTask`

**Fixes:**

- **Timezone**: event times are read via `StartUTC`/`EndUTC` and converted with
  `astimezone()`. Outlook's COM `Start`/`End` properties report local wall-clock
  time with a bogus `+00:00` offset; naive code shifts every appointment.
- **Locale-safe date filters**: DASL/Jet date literals are formatted with
  `win32api.GetDateFormat(LOCALE_USER_DEFAULT, ...)`. Outlook parses filter
  dates in the *system* locale — a US-formatted `07/02/2026` is read as
  February 7 on a German system.
- **COM apartment discipline**: `CoInitialize`/`CoUninitialize` are balanced in
  a single helper (`_run_com`) instead of being called ad hoc.
- **`--version`**: reports the real installed version (upstream had a typo that
  always printed `dev`).

**Additions:**

- `calendarPath` parameter on all event tools, and emitted in every event dict,
  so events are addressable across multiple/shared calendars.
- `busyStatus` and `meetingStatus` exposed on events.
- `createEvent` with `calendarId` saves, then moves the item to the target
  calendar and displays the moved item.
- Long message bodies are truncated with a hint to call `getMessage`/`getEvent`
  for the full body.

**Safety model (unchanged from upstream):** every outbound action —
`sendMail`, `replyToMessage`, `forwardMessage`, `createEvent`, tasks — ends in
`.Display()`, never `.Send()`. The item opens in Outlook and a human clicks
Send. The server cannot send anything on its own.

## Full tool list (21)

| Area | Tools |
|---|---|
| Accounts/folders | `listAccounts`, `listFolders`, `createFolder` |
| Email | `searchMessages`, `getRecentMessages`, `getMessage`, `sendMail`, `replyToMessage`, `forwardMessage`, `updateMessage`, `deleteMessages` |
| Contacts | `searchContacts` |
| Calendar | `listCalendars`, `createEvent`, `listEvents`, `getEvent`, `updateEvent`, `deleteEvent` |
| Tasks | `listTasks`, `createTask`, `updateTask` |

## Requirements

- Windows with the classic Outlook desktop client installed and configured
  (the server attaches to a running instance or starts one via COM)
- Python 3.10+

## Install

```powershell
pip install git+https://github.com/K4uP/outlook-mcp.git
```

Register with Claude Code (user scope):

```powershell
claude mcp add --scope user outlook -- outlook-mcp-server
```

If `outlook-mcp-server` is not on `PATH`, use the full path to the console
script, e.g.
`%LOCALAPPDATA%\Programs\Python\Python312\Scripts\outlook-mcp-server.exe`.
Restart Claude Code afterwards and verify with `claude mcp list`.

## Known gotchas

- `mcpServers` config belongs in `~/.claude.json` (user scope) — putting it in
  `settings.json` does nothing.
- `listEvents` without `calendarPath` mixes in shared/delegate calendars; pass
  the explicit calendar path when you need one person's availability.
- `createEvent`: omit `calendarId` for the default calendar; some Exchange
  setups fail when the default calendar is addressed by explicit id.
- Mail "sent" via `sendMail` is only displayed — nothing leaves the machine
  until a human presses Send in the Outlook window.

## License

MIT — see [LICENSE](LICENSE). Original work copyright (c) 2026 lihaokun;
extensions and fixes copyright (c) 2026 Rene S. Kaup.
