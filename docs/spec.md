# calendula specification

Status: draft for review (0.1.0)

## How to read this document

The specification is written by stepwise refinement. Each level answers one question and only makes the level above it more precise; nothing at a lower level may contradict a higher one.

| Level | Section | Question |
| --- | --- | --- |
| 0 | [Purpose and scope](#0-purpose-and-scope) | Why does calendula exist, and what is it not for? |
| 1 | [Requirements](#1-requirements) | What must it do? (R1–R9) |
| 2 | [Behavior](#2-behavior) | Exactly how does it behave, case by case? |
| 3 | [Interfaces](#3-interfaces) | What goes in and out, byte for byte? |
| 4 | [Design](#4-design) | Which part of the code does what? |
| — | [Limitations and open questions](#5-limitations-and-open-questions) | What is knowingly left out, and what needs a decision? |

Every requirement and behavior has an ID. [test-design.md](test-design.md) maps each ID to the tests that verify it.

## 0. Purpose and scope

**Purpose.** Let an AI assistant (Claude Code, or any MCP client) read the user's own Outlook calendar, so it can answer questions about the schedule and check a meeting's time, subject and attendees before writing them down (for example, from a skill that keeps meeting notes).

**Setting.** One person, on their own Windows PC, working in WSL. Outlook (classic) runs on Windows. calendula runs in WSL as that person.

**In scope.**
- Reading the events of the default calendar
- Restarting Outlook when its COM connection is broken

**Out of scope.**
- Creating, changing or answering events
- Calendars other than the default one (shared, delegated, subscribed, other accounts)
- The new Outlook, Microsoft Graph, Exchange Web Services
- Running on Windows Python, or on Linux without Windows
- Other people's free/busy information, reminders

**Constraints.** These are facts the design must live with.

| ID | Constraint |
| --- | --- |
| C1 | Outlook (classic) can only be reached through COM on Windows. From WSL that means running `powershell.exe` through WSL interop |
| C2 | No admin rights or consent are available, so Microsoft Graph cannot be used |
| C3 | Calendar events are personal and business data. They must not be stored or sent anywhere except to the MCP client that asked |
| C4 | One fetch takes several seconds (about 5 s: PowerShell starts, then Outlook answers) |
| C5 | Continuous integration runs on Linux without Windows or Outlook |

## 1. Requirements

| ID | Requirement |
| --- | --- |
| R1 | **List events.** For one day, or up to 31 consecutive days, return each day's events with time, subject and location |
| R2 | **Event at a time.** For a day and a time, return the event at that time with all its details |
| R3 | **Freshness.** Answers reflect Outlook within a bounded age, and every answer says when its data was fetched. The caller can force a fresh fetch |
| R4 | **Economy.** Do not fetch from Outlook when a fresh enough copy is held, and fetch at most once per request (C4) |
| R5 | **Privacy.** Keep fetched events in memory only (C3) |
| R6 | **Failures.** When Outlook cannot be read, return a clear error and never old or invented events. Tell the model when restarting Outlook may help, and that it needs the user's consent |
| R7 | **Recovery.** Restart Outlook on request |
| R8 | **Interfaces.** An MCP server on stdio with three tools, and a command line for manual checks. Tool names, arguments and output shape are stable, because skills tell the model to use them |
| R9 | **Platform.** Python 3.11–3.14 in WSL, Windows PowerShell 5.1, Outlook (classic) |

## 2. Behavior

### 2.1 Arguments (R1, R2)

| ID | Rule |
| --- | --- |
| B-DATE-1 | A date is `YYYY-MM-DD`, `YYYY/MM/DD`, `MM-DD` or `MM/DD`. Month and day may have one or two digits. Without a year, the current year is used |
| B-DATE-2 | An omitted or blank date means today (local time) |
| B-DATE-3 | Any other text, or a date that does not exist (`2026-02-30`), is an error: `invalid date: '<text>' ...` |
| B-DAYS-1 | The number of days is an integer from 1 to 31. Anything else is an error |
| B-TIME-1 | A time is `H:MM` or `HH:MM`, hours 0–23, minutes 00–59 (two digits). Anything else is an error: `invalid time: '<text>' ...` |

Relative words ("tomorrow", "next Monday") are not accepted. The model turns them into dates (B-INS-1).

### 2.2 Which events belong to a day (R1)

| ID | Rule |
| --- | --- |
| B-DAY-1 | An event belongs to **every day it overlaps**: start < the next day 00:00 and end > that day 00:00, local time. Occurrences of recurring events are included |
| B-DAY-2 | So an event that spans several days appears on each of them: a three-day all-day event on its three days, a meeting from 23:00 to 01:00 on both days. An all-day event ends at 00:00 of the day after its last day, so it does not appear on that day. A zero-length event at exactly 00:00 overlaps no day (limitation L1) |
| B-DAY-3 | Within a day, all-day events come first, then timed events in order of start (an event that began on an earlier day comes before those that begin on this day) |

### 2.3 Event content (R1, R2)

| ID | Rule |
| --- | --- |
| B-EV-1 | An event has: start, end, all-day flag, subject, location, organizer, required attendees, optional attendees, categories, busy status, body |
| B-EV-2 | A missing text field is empty, never `None` or `null` |
| B-EV-3 | The body is split into lines; each line is trimmed and blank lines are removed. The Teams join block is removed: a line of 10 or more underscores whose next non-blank line contains `Microsoft Teams`, through the next line of 10 or more underscores (or the end of the body). Any other line of underscores is kept, with what follows it |
| B-EV-4 | Busy status is shown by name: 0 Free, 1 Tentative, 2 Busy, 3 Out of office, 4 Working elsewhere. Other values are not shown |

### 2.4 The event at a time (R2)

| ID | Rule |
| --- | --- |
| B-PICK-1 | Only timed events are candidates; all-day events are ignored |
| B-PICK-2 | An event is running at time t of day D when start ≤ (D, t) < end, comparing date and time (the end is exclusive). So a meeting from 23:00 to 01:00 is running at 23:30 on its first day and at 00:30 on the next |
| B-PICK-3 | If several are running, the one that started last is chosen. If they started at the same time, the first in day order (B-DAY-3) |
| B-PICK-4 | If none is running, the event whose start is nearest to (D, t), before or after, is chosen, and the answer says that no event was running. On a tie, the earlier one |
| B-PICK-5 | If the day has no timed events, the answer says so |

### 2.5 Cache (R3, R4, R5)

| ID | Rule |
| --- | --- |
| B-CACHE-1 | The cache holds days. Each entry is the time it was fetched and the day's events |
| B-CACHE-2 | For each requested day, the decision table below says whether to fetch |
| B-CACHE-3 | The time to live (TTL) is 60 minutes. `CALENDULA_TTL_MINUTES` (read when the server starts) changes it; anything but a positive integer means 60 |
| B-CACHE-4 | A request fetches at most once: one fetch covering the earliest to the latest day that needs fetching. Days inside that span are fetched again even if they were fresh |
| B-CACHE-5 | A day missing from what Outlook returned has no events |
| B-CACHE-6 | If a fetch fails, the cache is not changed and the error is returned (no partial update) |
| B-CACHE-7 | Requests are handled one at a time, so a day is never fetched twice at once |
| B-CACHE-8 | The cache lives in memory for the life of the server process (usually one Claude Code session). Nothing is written to disk. `restart_outlook` empties it |
| B-CACHE-9 | Every answer shows, for each day, when that day was fetched |

Decision table for B-CACHE-2. "Fetched after the day ended" means the fetch happened on a later date than the day itself.

| # | `refresh` | Entry | Fetched after the day ended | Age ≥ TTL | Action |
| --- | --- | --- | --- | --- | --- |
| 1 | true | any | any | any | fetch |
| 2 | false | none | — | — | fetch |
| 3 | false | yes | yes | any | reuse |
| 4 | false | yes | no | no | reuse |
| 5 | false | yes | no | yes | fetch |

Why: a day that was over when it was fetched will not change any more (row 3). Today and later days do change, so they are fetched again after an hour (rows 4–5). A past day fetched while it was still today (say, the session ran past midnight) is treated like today (row 5).

### 2.6 Errors (R6)

| ID | Condition | Message | Restart hint |
| --- | --- | --- | --- |
| E1 | `wslpath`, `wscript.exe` or `tasklist.exe` is missing (not WSL) | `calendula needs WSL with Outlook (classic) on Windows; not found: <commands>` | no |
| E2 | The script reported an error that looks like a broken COM connection: it contains `CO_E_SERVER_EXEC_FAILURE`, `80080005`, `800706BA`, `800706BE`, `80010108` or `RPC_E_` | the script's message | yes |
| E3 | The script reported any other error | the script's message | no |
| E4 | The script finished without writing a result | `Outlook gave no answer (is Outlook (classic) running?).` | yes |
| E5 | The script did not finish within 180 seconds | `Outlook did not answer within 180 seconds.` | yes |
| E6 | The result is not valid | `Unexpected output from Outlook: <detail>` | no |
| E7 | An argument is invalid (B-DATE-3, B-DAYS-1, B-TIME-1) | `invalid date ...`, `invalid time ...`, or the range error | no |
| E8 | `OUTLOOK.EXE` is not running when a fetch starts. The script is not run, so a fetch never starts Outlook (creating `Outlook.Application` would, and a PC with no mail profile then stops at the setup screen) | `Outlook (classic) is not running. Start it and try again.` | no |
| E9 | `tasklist.exe` did not finish within 30 seconds | `tasklist.exe did not answer within 30 seconds.` | no |

The restart hint is appended on its own line: `Outlook's COM connection seems broken. Ask the user whether Outlook may be closed and restarted (unsaved drafts are lost); if they agree, call restart_outlook and then retry once.`

| ID | Rule |
| --- | --- |
| B-ERR-1 | A tool that fails returns an MCP tool error (`isError: true`) carrying the message, never a normal result. A failure that is not in the table above (an unexpected exception) carries its type and text, `<type>: <text>`, and its traceback goes to the log (standard error) |
| B-ERR-2 | A failed request never returns events, old or new |

### 2.7 Recovery (R7)

| ID | Rule |
| --- | --- |
| B-RST-1 | If `OUTLOOK.EXE` is running, it is force-quit, and calendula waits up to 10 seconds for it to go away (starting Outlook while the old process is still there hands the start over to it) |
| B-RST-2 | Outlook is started. calendula waits until `OUTLOOK.EXE` appears (checking every 3 seconds, at most 120 seconds, otherwise an error), then 15 more seconds for COM to answer |
| B-RST-3 | The answer says what happened: `Outlook was closed.` or `Outlook was not running.`, then `Outlook was started again.` |
| B-RST-4 | After a successful restart, the cache is emptied |
| B-RST-5 | The running check (B-RST-1, B-RST-2, E8) looks for the image name in the output of `tasklist.exe` as bytes, ignoring case. With no match, `tasklist.exe` prints a message in the console code page (CP932 on Japanese Windows), so the output is not decoded |

### 2.8 Guidance for the model (R6, R8)

| ID | Rule |
| --- | --- |
| B-INS-1 | The server's instructions tell the model: the tools are read-only; how fresh answers are and when to pass `refresh`; dates are `YYYY-MM-DD` and relative words must be turned into dates first; on failure, report the error and never fill in events; on a COM error, ask the user before `restart_outlook`, then retry once and do not restart repeatedly |
| B-ANN-1 | Tool annotations: `list_events` and `get_event` are read-only, idempotent and closed-world. `restart_outlook` is destructive |

## 3. Interfaces

### 3.1 MCP tools (R1, R2, R7, R8)

The server is named `calendula` and speaks MCP over stdio.

| Tool | Arguments | Result |
| --- | --- | --- |
| `list_events` | `date`: string, optional (B-DATE-*). `days`: integer 1–31, default 1. `refresh`: boolean, default false | text, below |
| `get_event` | `time`: string, required (B-TIME-1). `date`: string, optional. `refresh`: boolean, default false | text, below |
| `restart_outlook` | none | text (B-RST-3) |

**`list_events` result.** One block per day, separated by a blank line:

```
2026-11-05 (Thu)  fetched 2026-11-05 10:32
All day      Holiday
09:00-10:00  Weekly sync  [Room A]
13:30-14:30  Customer interview
```

- First line: the date, the weekday (`Mon`…`Sun`), two spaces, `fetched`, the time this day was fetched (B-CACHE-9)
- One line per event: the time span on this day, padded to 11 characters, two spaces, the subject, and, if there is a location, two spaces and `[location]`
- A day without events has the single line `(no events)`

**Time span on day D** (used by both tools):

| Event | Span | Example on 2026-11-05 |
| --- | --- | --- |
| all-day (one day or several) | `All day` | `All day` |
| starts and ends on D | `HH:MM-HH:MM` | `09:00-10:00` |
| starts on D, ends at 00:00 of the next day | `HH:MM-24:00` | `23:00-24:00` |
| starts on D, ends later | `HH:MM-MM-DD HH:MM` | `23:00-11-06 01:00` |
| started earlier, ends on D | `MM-DD HH:MM-HH:MM` | on 2026-11-06: `11-05 23:00-01:00` |
| started earlier, ends after D | `MM-DD HH:MM-MM-DD HH:MM` | `11-04 09:00-11-06 17:00` |

That is: the start is `HH:MM` if it is on D, otherwise `MM-DD HH:MM`; the end is `HH:MM` if it is on D, `24:00` if it is exactly 00:00 of the day after D, otherwise `MM-DD HH:MM`.

**`get_event` result.**

```
2026-11-05 (Thu)  fetched 2026-11-05 10:32
No event is running at 12:00; this is the one starting nearest to it.
13:30-14:30  Customer interview
Location: Room B
Organizer: Host
Required attendees: Member A; Member B
Status: Busy

Agenda
1. Requirements
```

- First line as in `list_events`
- The note line appears only when no event was running (B-PICK-4)
- Then the time span and subject, then these labeled lines in this order, each only if not empty: `Location`, `Organizer`, `Required attendees`, `Optional attendees`, `Categories`, `Status`
- If there is a body: a blank line, then the body lines (B-EV-3)
- If the day has no timed events: the first line, then `No timed events on this day.`

### 3.2 Command line (R8)

| Command | Does |
| --- | --- |
| `calendula`, `calendula serve` | run the MCP server on stdio |
| `calendula list [DATE] [DAYS]` | print events in the `list_events` format |
| `calendula show TIME [--date DATE]` | print the event in the `get_event` format |
| `calendula --version` | print `calendula <version>` |

The command line always fetches (it has no cache). Exit status: 0 success, 1 error (message `calendula: <message>` on standard error), 2 usage error.

### 3.3 PowerShell script `get_calendar.ps1` (C1)

The script is the only code that talks to Outlook. Its contract:

- **Input**, next to the script: `start.txt` (first day, `yyyyMMdd`), `days.txt` (number of days)
- **Output** on success: `events.json`, UTF-8 without BOM:
  ```json
  {"days": [{"date": "2026-11-05", "events": [
    {"start": "2026-11-05T09:00:00", "end": "2026-11-05T10:00:00", "all_day": false,
     "subject": "...", "location": "...", "organizer": "...", "required": "...", "optional": "...",
     "categories": "...", "busy": 2, "body": ["...", "..."]}]}]}
  ```
  One entry per requested day, in order. `events` and `body` are always arrays, even with zero or one element. Text fields are strings, never null. Times are local, `yyyy-MM-ddTHH:mm:ss`
- **Output** on failure: `error.txt` with `ERROR: <message>`, exit code 1
- **Field sources** (AppointmentItem properties): `Start`, `End`, `AllDayEvent`, `Subject`, `Location`, `Organizer`, `RequiredAttendees`, `OptionalAttendees`, `Categories`, `BusyStatus`, `Body` (filtered by B-EV-3)
- **Selection** (B-DAY-1): the default calendar folder (`GetDefaultFolder(9)`), `Items` with `IncludeRecurrences = $true`, sorted by `[Start]`, restricted per day by `[Start] < '<next day 00:00>' AND [End] > '<day 00:00>'`, the dates in the Windows short date-time format (`ToString("g")`), which Outlook parses back
- **Structure:** the parts above are functions (`Get-BodyLines`, `ConvertTo-EventRecord`, `Get-DayFilter`, `Get-DayEvents`, `ConvertTo-CalendarJson`). Only `Invoke-Main` uses COM. Dot-sourcing the script defines the functions without running `Invoke-Main`

### 3.4 Process boundary (C1, R5)

For each fetch, calendula first checks that `OUTLOOK.EXE` is running (E8, E9). Only then does it:

1. creates a temporary directory under `/tmp` (prefix `calendula-`), and writes `start.txt` and `days.txt`
2. copies `get_calendar.ps1` into it as UTF-16LE with a BOM (Windows PowerShell 5.1 reads a BOM-less file in the ANSI code page) and copies `run_hidden.vbs`
3. runs `wscript.exe //B //Nologo <run_hidden.vbs> <get_calendar.ps1>` with their Windows paths (`wslpath -w`), which runs PowerShell without a window and waits for it, with a 180-second limit
4. reads the result (E2–E6)
5. removes the directory, whether the fetch succeeded or not

## 4. Design

```
MCP client ──stdio──> server.py ──> cache.py (Calendar) ──fetch──> outlook.py ──wscript/PowerShell──> get_calendar.ps1 ──COM──> Outlook
                         │                                              │
                         └──── events.py (parse, format, pick) <────────┘
cli.py ──> outlook.py (no cache) and events.py
```

| Module | Responsibility | Spec |
| --- | --- | --- |
| `server.py` | the three tools, argument schemas, annotations, instructions; turns failures into tool errors | 3.1, B-ERR-1, B-INS-1, B-ANN-1, B-RST-4 |
| `cache.py` | the decision table, the fetch span, the lock, the TTL setting | 2.5 |
| `events.py` | parsing `events.json`, ordering, formatting, picking the event at a time, parsing dates and times | 2.1–2.4, 3.1 formats |
| `outlook.py` | the process boundary, reading the result, classifying errors, restarting Outlook | 2.6, 2.7, 3.4 |
| `get_calendar.ps1` | reading Outlook and writing `events.json` | 3.3, B-DAY-1, B-EV-3 |
| `cli.py` | the command line | 3.2 |

`Calendar` takes the fetch function and the clock as parameters, and `build_server` takes the calendar and the restart function. This is what lets the tests replace Outlook and time (see [test-design.md](test-design.md)).

## 5. Limitations and open questions

| ID | Limitation | Why it is accepted (for now) |
| --- | --- | --- |
| L1 | A zero-length event at exactly 00:00 is not listed (B-DAY-2) | It overlaps no day under B-DAY-1; such events are rare, and catching them would complicate the filter |
| L2 | The Restrict filter uses the Windows short date-time format. It has been used only with the Japanese locale | Outlook parses dates in the user's locale; the test W5 checks that the format reads back as the same instant |
| L3 | Only the default calendar is read | Scope |
| L4 | A canceled meeting that is still in the calendar is shown as it is (its subject usually starts with "Canceled:") | It is what Outlook shows |
| L5 | Times are Outlook's local times; there is no time zone handling | One user, one PC |
| L6 | The body is returned whole and may be long | The model may need any part of it; the Teams block, the longest boilerplate, is removed |
| L7 | Outlook left open at its setup screen passes the running check, so the fetch still ends in E5 after 180 seconds | It should be rare: a fetch no longer opens that screen (E8), so a person has to open it and leave it |

| ID | Question for the reviewer | Decision |
| --- | --- | --- |
| Q1 | Should an event that spans several days appear on every day it covers? | Yes (2026-11-05). B-DAY-1, B-DAY-2, B-PICK-2, B-PICK-4 and the time span format were changed accordingly |
| Q2 | Is returning the whole body (L6) acceptable? | Yes (2026-11-05) |
