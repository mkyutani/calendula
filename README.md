# calendula

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![PyPI version](https://badge.fury.io/py/calendula.svg)](https://pypi.org/project/calendula/)

An [MCP](https://modelcontextprotocol.io/) server that lets Claude Code (or any MCP client) read your **Outlook (classic)** calendar from **WSL**.

- Read-only: list the events of a day (or several), or look up the event at a given time with its attendees and body
- Talks to Outlook on Windows through PowerShell and COM: no Microsoft Graph, no admin consent, no tokens
- Keeps what it fetched in memory only; nothing is written to disk (see [Caching](#caching))

## Requirements

- Windows with **Outlook (classic)** running. The new Outlook has no COM interface and does not work
- **WSL**: calendula runs in the Linux distribution and calls `wscript.exe`, `powershell.exe` and `tasklist.exe` through WSL interop
- Python 3.11 or later

## Install

```bash
uv tool install calendula          # or: pipx install calendula
calendula list                     # prints today's events (checks that Outlook answers)
```

Register it with Claude Code for all your projects:

```bash
claude mcp add --scope user calendula -- calendula
```

Other MCP clients: run `calendula` as a stdio server.

Then ask things like "What's on my calendar tomorrow afternoon?" or "Who is in the 14:00 meeting?".

To let Claude read the calendar without asking each time, allow the two read-only tools in `~/.claude/settings.json`:

```json
{
  "permissions": {
    "allow": ["mcp__calendula__list_events", "mcp__calendula__get_event"]
  }
}
```

## Tools

| Tool | Arguments | Returns |
| --- | --- | --- |
| `list_events` | `date` (YYYY-MM-DD, default today), `days` (1-31, default 1), `refresh` | Each day's events: time, subject, location |
| `get_event` | `time` (HH:MM), `date`, `refresh` | The event running at that time (or the one starting nearest to it, with a note): location, organizer, attendees, status, body |
| `restart_outlook` | none | Force-quits Outlook and starts it again, to recover a broken COM connection. Unsaved drafts are lost, so the model is told to ask you first |
| `start_outlook` | none | Starts Outlook when it is closed, and never closes it. It opens a window on your desktop, so the model is told to ask you first |

Every answer starts with the day and the time it was fetched:

```
2026-11-05 (Thu)  fetched 2026-11-05 10:32
09:00-10:00  Weekly sync  [Room A]
13:30-14:30  Customer interview
```

An event that spans several days (a conference, a meeting past midnight) appears on each day it covers. Bodies are kept whole except for blank lines and the Teams join block.

## Caching

Fetching takes a few seconds (PowerShell starts, then Outlook answers over COM), so fetched days are kept in the server's memory:

- **Today and later** are reused for up to 1 hour, then fetched again
- **Past days** fetched after they ended are kept for the life of the server (one Claude Code session)
- `refresh: true` always fetches; ask for it when the calendar has just changed
- A request fetches only the days that are missing or stale, in one call
- When fetching fails, the error is returned as is; older copies are never shown instead

Nothing is written to disk. Set `CALENDULA_TTL_MINUTES` to change the 1 hour.

## Command line

```
calendula                         run the MCP server on stdio (same as `calendula serve`)
calendula list [DATE] [DAYS]      print events, fetched from Outlook now
calendula show TIME [--date DATE] print the event at a time
calendula --version
```

## Troubleshooting

- **`CO_E_SERVER_EXEC_FAILURE` (0x80080005) and similar COM errors**: Outlook's COM connection often breaks after sleep or a network change. Restart Outlook (Claude will ask before calling `restart_outlook`)
- **`Outlook (classic) is not running`**: start Outlook and try again, or let Claude do it (it will ask before calling `start_outlook`). Reading the calendar never starts Outlook by itself: a closed Outlook can stop at its setup screen, and the read would hang
- **`Outlook (classic) has no mail profile`**: `start_outlook` does not open Outlook on a PC where it has not been set up. If you do want to use it there, set it up in Outlook by hand. Only Office 2016 and later are recognized
- **`Outlook gave no answer`**: check that Outlook (classic) is running
- **`calendula needs WSL ...`**: calendula only works inside WSL on a Windows PC with Outlook

## How it works

For each fetch, calendula checks with `tasklist.exe` that Outlook is running, then writes `get_calendar.ps1` (as UTF-16 with a BOM, which Windows PowerShell 5.1 reads correctly) into a temporary directory under `/tmp`, and runs it hidden through `wscript.exe` and `run_hidden.vbs`, using the `\\wsl.localhost\...` path. The script reads the default calendar folder through `Outlook.Application`, including recurring events, and writes `events.json`, which calendula parses.

## Development

```bash
uv sync
uv run pytest                                          # needs neither Windows nor Outlook
CALENDULA_WINDOWS_TESTS=1 uv run pytest                # on WSL: also the PowerShell side, still without Outlook
uv run ruff check && uv run ruff format --check
```

What is tested where, and why, is in [docs/test-design.md](docs/test-design.md); the behavior it checks is specified in [docs/spec.md](docs/spec.md).

Releases: bump `version` in `pyproject.toml`, then push a tag `vX.Y.Z`. GitHub Actions tests, builds, publishes to TestPyPI, installs from there, publishes to PyPI (Trusted Publishing), and creates a GitHub release.

## License

[MIT](LICENSE)
