# calendula test design

Status: draft for review (0.1.0)

This document says how the tests verify [spec.md](spec.md): which behavior is checked where, with what technique, and what is left for a person to check on a real PC. A reviewer should be able to approve it by checking three things:

1. every requirement (R1–R9) is reached by at least one test case ([traceability](#5-traceability));
2. each test case checks the rule it names, including the boundaries ([test cases](#4-test-cases));
3. the residual risks are understood and accepted ([residual risks](#6-residual-risks)).

Test case IDs appear in the test function names (`test_k3_...` is K3), so `pytest -v` reads as this document.

## 1. Test basis

The basis is the ID'd rules of spec.md: R (requirements), B (behavior), E (errors), sections 3.x (interfaces). A test case names the IDs it verifies.

## 2. Test levels

Outlook can only be reached through COM on Windows (C1), CI has no Windows (C5), and the developer's real calendar should not be read from a development session (C3). So the code is split so that **only a few lines touch Outlook** (`Invoke-Main` in the script), and everything else can be tested with stand-ins.

| Level | What it verifies | Stand-ins | Runs where | How |
| --- | --- | --- | --- | --- |
| T1 Unit | parsing, formatting, picking, cache decisions | fake clock, fake fetch | anywhere, CI | `pytest` |
| T2 Component | the bridge to Windows (`outlook.py`), restarting, the command line | fake `subprocess`, `shutil.which`, `time` | anywhere, CI | `pytest` |
| T3 Protocol | the MCP server as a client sees it | fake calendar and restart | anywhere, CI | `pytest` with the `mcp` client, in process and over stdio |
| T4 Windows | the PowerShell script's functions and the wscript/PowerShell plumbing, **without Outlook** | fake appointment items and folder (PowerShell objects) | WSL on Windows, opt-in: `CALENDULA_WINDOWS_TESTS=1 uv run pytest` | `pytest` runs `powershell.exe` on a harness that dot-sources the script |
| T5 Acceptance | everything together, against the real Outlook | none | the user's PC | the manual checklist in section 4.9 |

What T1–T4 cannot reach is listed in [residual risks](#6-residual-risks); T5 covers it.

## 3. Techniques

| Technique | Used for |
| --- | --- |
| Equivalence partitioning | date and time forms (valid forms, invalid forms) |
| Boundary values | days 0/1/31/32, time 0:00/23:59/24:00, age TTL−1/TTL, start and end of an event |
| Decision table | the cache (spec 2.5, rows 1–5, one case per row) |
| State transition | a cached day over time: absent → fresh → stale → fresh, or final (see below) |
| Contract (exact output) | the text formats of spec 3.1, compared character for character |
| Error guessing | PowerShell collapsing one-element arrays, BOMs, Japanese text, null COM properties, the console code page of Windows commands |

State transitions of one cached day (spec 2.5):

```
absent ──fetch──> fresh ──age reaches TTL──> stale ──fetch──> fresh
                    │
                    └── fetched after the day ended ──> final (never fetched again)
any state ──refresh──> fetch;  any state ──restart_outlook──> absent
```

K1 (absent→fresh), K2 (fresh stays), K3 (fresh→stale→fresh), K5 (final), K6 (fresh, then the day ends: still fetched again), K4 (refresh), K12 and S4 (→absent) cover every arrow.

## 4. Test cases

Columns: **ID**, the spec IDs it verifies, the condition and the expected result, the level.

### 4.1 Arguments — `tests/test_arguments.py`

| ID | Spec | Condition → expected | Level |
| --- | --- | --- | --- |
| A1 | B-DATE-2 | `None`, `""`, `"  "` → today | T1 |
| A2 | B-DATE-1 | `2026-12-01`, `2026/12/01` → 2026-12-01 | T1 |
| A3 | B-DATE-1 | `12-1`, `12/01`, `1/5` → that day of the current year | T1 |
| A4 | B-DATE-3 | `tomorrow`, `2026-02-30`, `13/40`, `20261105`, `2026-1-1-1` → error `invalid date` | T1 |
| A5 | B-TIME-1 | `0:00`, `00:00`, `9:05`, `23:59` → accepted | T1 |
| A6 | B-TIME-1 | `24:00`, `23:60`, `9`, `9:5`, `09:00:00` → error `invalid time` | T1 |
| A7 | B-DAYS-1 | the calendar asked for 0 or 32 days → error; 1 and 31 → accepted | T1 |
| A8 | B-DAYS-1, 3.1 | `list_events` with `days` 0 or 32 → tool error; the schema says minimum 1, maximum 31 | T3 |

### 4.2 Event content and formats — `tests/test_events.py`

| ID | Spec | Condition → expected | Level |
| --- | --- | --- | --- |
| V1 | B-EV-1, B-EV-2 | an event with every field, and one with `null` and missing fields → all fields read; missing text is `""`; missing busy is none | T1 |
| V2 | 3.3 | `body` given as one string, `events` given as one object (collapsed arrays) → read as one-element lists | T1 |
| V3 | B-DAY-3 | events out of order with an all-day event last → all-day first, then by start | T1 |
| V4 | B-EV-4 | busy 0–4 → the five names; busy 9 → no `Status` line | T1 |
| V5 | 3.1 | summary lines: timed with location, timed without, all-day → exact strings, span padded to 11 | T1 |
| V6 | 3.1 | detail: labels in the fixed order, only non-empty ones, a blank line before the body; no body → no blank line | T1 |
| V7 | 3.1 | the day heading for seven consecutive days → `Mon`…`Sun` and the fetched time | T1 |
| V8 | B-DAY-2, 3.1 | the time span on each day of: 09:00–10:00; 23:00–24:00; 23:00–01:00 (both days); 11-04 09:00 to 11-06 17:00 (each of the three days); a three-day all-day event → exactly the six forms of the span table | T1 |

### 4.3 The event at a time — `tests/test_pick.py`

| ID | Spec | Condition → expected | Level |
| --- | --- | --- | --- |
| P1 | B-PICK-2 | 09:00–10:00, t = 09:30 → that event, running | T1 |
| P2 | B-PICK-2 | t = 09:00 (start) → running; t = 10:00 (end) → not running | T1 |
| P3 | B-PICK-3 | 09:00–12:00 and 10:00–11:00, t = 10:30 → the 10:00 one | T1 |
| P4 | B-PICK-3 | two events both 10:00–11:00 → the first | T1 |
| P5 | B-PICK-4 | 09:00–10:00 and 15:00–16:00, t = 13:30 → the 15:00 one, not running; t = 10:30 → the 09:00 one | T1 |
| P6 | B-PICK-4 | 10:00–11:00 and 14:00–15:00, t = 12:30 (equally far) → the 10:00 one | T1 |
| P7 | B-PICK-2 | 23:00 to 01:00 the next day: 23:30 on the first day → running; 00:30 on the next day → running; 01:00 on the next day → not running | T1 |
| P8 | B-PICK-1, B-PICK-5 | only an all-day event → none; no events → none | T1 |

### 4.4 Cache — `tests/test_cache.py`

TTL is 1 hour in these cases. The clock starts at 2026-11-05 09:00.

| ID | Spec | Condition → expected | Level |
| --- | --- | --- | --- |
| K1 | B-CACHE-1, B-CACHE-2 row 2, B-CACHE-9 | first request for today → one fetch; the entry's fetched time is the clock | T1 |
| K2 | row 4 | again after 59 minutes → no fetch, same fetched time | T1 |
| K3 | row 5 | again after exactly 60 minutes → fetch | T1 |
| K4 | row 1 | `refresh` right after a fetch → fetch | T1 |
| K5 | row 3 | yesterday, fetched today; again 5 hours later → no fetch | T1 |
| K6 | row 5 | today fetched at 09:00; next morning, the same day (now past) → fetch | T1 |
| K7 | B-CACHE-4 | 2 days, then 4 days from the same start → second fetch covers only days 3–4 | T1 |
| K8 | B-CACHE-4 | days 2–3 fresh, days 1 and 4 not → one fetch covering days 1–4 | T1 |
| K9 | B-CACHE-5 | the fetch returns nothing for a requested day → that day has no events | T1 |
| K10 | B-CACHE-6 | a cached day, then `refresh` with a failing fetch → the error is raised; afterwards the old entry is unchanged and no fetch happens | T1 |
| K11 | B-CACHE-7 | two threads ask for the same day at once (the fetch is slow) → one fetch | T1 |
| K12 | B-CACHE-8 | `clear()` → the next request fetches | T1 |
| K13 | B-CACHE-3 | `CALENDULA_TTL_MINUTES` unset → 60; `5` → 5; `0`, `-3`, `abc`, `""` → 60 | T1 |

### 4.5 Bridge to Windows — `tests/test_outlook.py`

| ID | Spec | Condition → expected | Level |
| --- | --- | --- | --- |
| O1 | E1 | `wslpath`, `wscript.exe` and `tasklist.exe` not found → error naming them; no restart hint | T2 |
| O2 | 3.4 steps 1–3 | a fetch writes `start.txt` (`yyyyMMdd`) and `days.txt`; the script copy is the source as UTF-16LE with a BOM; `run_hidden.vbs` is copied; `wscript.exe //B //Nologo <vbs> <ps1>` gets Windows paths and a 180-second timeout | T2 |
| O3 | 3.4 step 5 | the temporary directory is gone after a successful fetch and after a failed one | T2 |
| O4 | E2 | `error.txt` with each of the six COM patterns → that message plus the restart hint | T2 |
| O5 | E3 | `error.txt` with another message → exactly that message, no hint | T2 |
| O6 | E4 | neither file → "gave no answer" plus the hint | T2 |
| O7 | E5 | the run times out → "did not answer within 180 seconds" plus the hint | T2 |
| O8 | E6 | `events.json` that is not JSON, or lacks `date` → "Unexpected output" | T2 |
| O9 | 3.3 | `events.json` in UTF-8 with and without a BOM, with Japanese → read the same | T2 |
| O10 | E8, 3.4 | `OUTLOOK.EXE` is not running → exactly the E8 message, no restart hint; `wscript.exe` is not run and no temporary directory is created | T2 |
| O11 | E9 | `tasklist.exe` times out → exactly the E9 message, no restart hint; `wscript.exe` is not run. Its limit is 30 seconds | T2 |

### 4.6 Restarting Outlook — `tests/test_restart.py`

| ID | Spec | Condition → expected | Level |
| --- | --- | --- | --- |
| S1 | B-RST-1, B-RST-2, B-RST-3 | Outlook running → `taskkill /f`, waits until it is gone, starts it, waits until it appears, then 15 s; answer `Outlook was closed. Outlook was started again.` | T2 |
| S2 | B-RST-2, B-RST-3 | Outlook not running → no `taskkill`; answer `Outlook was not running. Outlook was started again.` | T2 |
| S3 | B-RST-2 | Outlook never appears → error after 120 s (fake clock) | T2 |
| S4 | B-RST-4, B-ERR-1 | `restart_outlook` succeeds → the next `list_events` fetches; fails → tool error, and the cache is kept | T3 |
| S5 | B-RST-5 | `tasklist.exe` answers its "no match" message in CP932 → the restart succeeds from both states (running, not running); a running image name in lower case counts as running | T2 |

The fake behaves as the real `subprocess.run` does here: it returns bytes, and `text=True` decodes them strictly as UTF-8. So S1–S3 also go through the "no match" message.

### 4.7 MCP server — `tests/test_server.py`

| ID | Spec | Condition → expected | Level |
| --- | --- | --- | --- |
| M1 | R8, B-ANN-1 | tools are exactly `list_events`, `get_event`, `restart_outlook`; the first two read-only, idempotent, closed-world; the third destructive | T3 |
| M2 | 3.1 | schemas: `time` required in `get_event`; `days` default 1, minimum 1, maximum 31; `date` optional; `refresh` default false | T3 |
| M3 | B-INS-1 | the instructions say: read-only; the TTL; `refresh`; `YYYY-MM-DD`; never fill in events; ask the user before `restart_outlook`; retry once | T3 |
| M4 | 3.1 | `list_events` for two days, the second empty → the exact text, blocks separated by a blank line, `(no events)` | T3 |
| M5 | 3.1, B-PICK-4, B-PICK-5 | `get_event` at a running time, between events (note line), and on a day with only an all-day event | T3 |
| M6 | B-ERR-1, B-ERR-2, E2, E7 | a COM error → tool error with the message and the hint, and no events; an invalid time → tool error `invalid time` | T3 |
| M7 | B-CACHE-9 | two `list_events` calls within the TTL report the same fetched time; with `refresh` the time is the new fetch's | T3 |
| M8 | R8 | the real entry point (`python -m calendula.cli`) over stdio: initializes and lists the three tools | T3 |
| M9 | B-ERR-1 | an unexpected exception in each of the three tools → a tool error with its type and text, and the traceback in the log | T3 |

### 4.8 Command line — `tests/test_cli.py`

| ID | Spec | Condition → expected | Level |
| --- | --- | --- | --- |
| L1 | 3.2 | `--version` → `calendula <version>`, exit 0 | T2 |
| L2 | 3.2 | `list 2026-11-05 2` → one fetch of 2 days, the `list_events` format | T2 |
| L3 | 3.2 | `show 12:00 --date 2026-11-05` → the note line and the nearest event | T2 |
| L4 | 3.2, E7 | invalid date, 0 days, a failing fetch → exit 1, `calendula: <message>` on standard error | T2 |
| L5 | 3.2 | no command → the server runs on stdio | T2 |

### 4.9 Windows, without Outlook — `tests/test_windows.py` (opt-in)

| ID | Spec | Condition → expected | Level |
| --- | --- | --- | --- |
| W1 | 3.3 | the PowerShell parser finds no syntax errors in `get_calendar.ps1` | T4 |
| W2 | B-EV-3 | `Get-BodyLines` on: a Teams block in the middle; a Teams block at the end with no closing line; an underscore line not followed by Teams (kept); a 5-underscore line (not a separator); lines with spaces and blank lines; an empty body → the expected lines | T4 |
| W3 | B-EV-2, 3.3 | `ConvertTo-EventRecord` on a fake item with null text fields and a one-line body → `""` for nulls, `yyyy-MM-ddTHH:mm:ss` times, boolean `all_day`, integer `busy`, `body` an array of one | T4 |
| W4 | B-DAY-1, B-DAY-3, 3.3 | `Get-DayEvents` on a fake folder with 0, 1 and 2 (out of order) items → arrays of 0, 1, 2, sorted by start; `IncludeRecurrences` set; the filter of `Get-DayFilter` passed to `Restrict` | T4 |
| W5 | B-DAY-1, L2 | `Get-DayFilter` for 2026-11-05 → exactly `[Start] < '<a>' AND [End] > '<b>'`, where `<a>` and `<b>` read back, in the current Windows locale, as 2026-11-06 00:00 and 2026-11-05 00:00 | T4 |
| W6 | 3.3 | `ConvertTo-CalendarJson` for days with 0, 1 and 2 events, with Japanese, quotes and `<>` → valid JSON; arrays stay arrays; text unchanged | T4 |
| W7 | 3.4 | the plumbing: a probe script run through `wscript.exe` and `run_hidden.vbs`, hidden, writes `events.json` with Japanese → read back by `read_result` | T4 |
| W8 | B-RST-5 | the running check on the real `tasklist.exe`, with an image name that does not exist → not running, in the current Windows locale | T4 |

### 4.10 Acceptance on the real Outlook (T5, manual)

Run on the user's PC, with Outlook (classic) running, after `uv tool install calendula` (or from the source tree with `uv run calendula`). Record pass or fail for each.

| ID | Spec | Do | Expected |
| --- | --- | --- | --- |
| X1 | R1, B-DAY-1, B-DAY-3 | `calendula list` | the same events as Outlook's day view for today, including recurring meetings and all-day events; Japanese subjects intact |
| X2 | R1 | `calendula list <Monday> 7` | a week matches Outlook's week view |
| X3 | R2, B-EV-3 | `calendula show HH:MM` on a Teams meeting with attendees | attendees listed; the Teams join block gone; the rest of the body kept |
| X4 | B-PICK-4 | `calendula show` at a time between meetings | the note line and the nearest meeting |
| X5 | E8 | close Outlook, then `calendula list` | at once, an error saying Outlook (classic) is not running; Outlook stays closed; no events |
| X6 | R8, B-CACHE-2 | in Claude Code, after `claude mcp add`: "What's on my calendar today?", then again, then "check again, it just changed" | the first answer fetches; the second shows the same fetched time; the third a new one |
| X7 | R7, B-INS-1 | when a COM error happens (after sleep), ask Claude for the schedule | Claude asks before restarting; after you agree, Outlook restarts and the retry succeeds |
| X8 | B-DAY-1, B-DAY-2 | `calendula list <first day> <n>` over a multi-day all-day event and a meeting past midnight | each appears on every day it covers, with the spans of spec 3.1; the all-day event not on the day after its last day |

## 5. Traceability

| Requirement | Behavior and interface IDs | Test cases |
| --- | --- | --- |
| R1 List events | B-DATE-1–3, B-DAYS-1, B-DAY-1–3, B-EV-1–4, 3.1 | A1–A4, A7, A8, V1–V5, V7, V8, M4, W2–W6, X1, X2, X8 |
| R2 Event at a time | B-TIME-1, B-PICK-1–5, 3.1 | A5, A6, P1–P8, V6, M5, L3, X3, X4 |
| R3 Freshness | B-CACHE-1–3, B-CACHE-9 | K1–K6, K13, M7, X6 |
| R4 Economy | B-CACHE-2, B-CACHE-4, B-CACHE-7 | K2, K5, K7, K8, K11, X6 |
| R5 Privacy | B-CACHE-8, 3.4 step 5 | K12, O3 (and by inspection: no code writes events to disk) |
| R6 Failures | E1–E9, B-ERR-1, B-ERR-2, B-CACHE-6, B-INS-1 | O1, O4–O8, O10, O11, K10, M3, M6, M9, L4, X5, X7 |
| R7 Recovery | B-RST-1–5 | S1–S5, W8, X7 |
| R8 Interfaces | 3.1, 3.2, B-ANN-1 | M1, M2, M8, L1–L5 |
| R9 Platform | 3.3, 3.4 | O2, O9, W1–W7, CI on Python 3.11–3.14, X1 |

## 6. Residual risks

What the automated levels cannot show, and how it is covered:

| Risk | Why not automated | Covered by |
| --- | --- | --- |
| The COM calls themselves: `New-Object -ComObject Outlook.Application`, `GetDefaultFolder(9)`, `IncludeRecurrences` with `Sort` and `Restrict` | needs Outlook (C1, C5), and reading a real calendar from a development session is avoided (C3) | X1–X4 |
| Outlook parsing the Restrict dates in a locale other than Japanese (L2) | W5 checks the round trip in the developer's locale only | X1 on each new locale |
| How long a real fetch takes, and whether 180 s is enough | timing depends on the PC | X1, X6 |
| Claude following the instructions (asking before a restart, using `refresh`) | model behavior, not code | X6, X7 |

## 7. Approval

The design is approved when the reviewer accepts sections 2–6 and answers the open questions Q1 and Q2 of spec.md. The release 0.1.0 is approved when:

- T1–T3 pass in CI on Python 3.11, 3.12, 3.13 and 3.14, with ruff clean
- T4 passes on a WSL PC (`CALENDULA_WINDOWS_TESTS=1 uv run pytest tests/test_windows.py -v`)
- X1–X7 pass on the user's PC; X8 is recorded with the answer to Q1

## 8. Evidence that the tests can fail

A passing test suite says little unless it fails when the code is wrong. `scripts/mutation_check.py` breaks one rule at a time, runs the tests, and restores the code; for each break, the test case written for that rule must fail.

| Rule | Break | Test cases that failed |
| --- | --- | --- |
| B-DATE-1 | a date without a year gets the year 2000 | A3 |
| B-TIME-1 | hour 24 accepted | A6 |
| B-EV-2 | a null text field becomes `"None"` | V1 |
| B-DAY-3 | all-day events sorted last | V3 |
| B-PICK-2 | the end of an event made inclusive | P2 |
| B-PICK-2 | events compared by time of day only, not date and time | P7 |
| 3.1 span | an event ending at 00:00 of the next day shown as `11-06 00:00` instead of `24:00` | V8 |
| B-DAY-1 (Windows) | the filter's end condition made `[End] >=` (an all-day event would also appear the day after) | W5 |
| B-PICK-3 | the earliest running event chosen | P3 |
| B-PICK-4 | ties go to the later event | P6 |
| B-CACHE-2 row 3 | a day fetched on that same day treated as final | K3, K6, K8 |
| B-CACHE-2 row 3 | past days never treated as final | K5 |
| B-CACHE-2 row 5 | an entry exactly TTL old still reused | K3 |
| B-CACHE-3 | `CALENDULA_TTL_MINUTES=0` accepted | K13 |
| B-CACHE-4 | the fetch covers only as many days as need fetching, from the first | K8 |
| B-CACHE-7 | no lock | K11 |
| E2 | `RPC_E_` not recognized as a COM error | O4 |
| E8 | no running check: a fetch runs the script with Outlook closed | O10, O11 |
| E9 | no limit on `tasklist.exe` | O11 |
| B-ANN-1 | `list_events` and `get_event` not marked read-only | M1 |
| B-RST-4 | the cache kept after a restart | S4 |
| B-RST-5 | the running check decodes the output of `tasklist.exe` as text again | O2, O3, O7, O10, S1, S2, S3, S5, W8 |
| B-RST-5 | the running check made case-sensitive | S5 |
| B-ERR-1 | an unexpected exception is left to the SDK, which answers without its text | M9 |
| B-EV-3 (Windows) | a line of 5 underscores starts a Teams block | W2 |
| 3.3 (Windows) | `Get-DayEvents` returns its array without the comma operator (one event collapses) | W4 |

Result on 2026-11-07: 26 of 26 caught. Rerun after changing the code:

```bash
CALENDULA_WINDOWS_TESTS=1 uv run python scripts/mutation_check.py   # without the variable, the Windows rows are skipped
```

The list covers the decisions most likely to be got wrong (boundaries, tie-breaks, the cache table, error classification, PowerShell's array handling), not every line.
