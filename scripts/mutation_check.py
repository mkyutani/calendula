"""Break one rule at a time and check that the test written for it fails (docs/test-design.md, section 8).

Each mutation replaces one piece of source text, runs the tests, and restores the file. A mutation
"is caught" when at least one of the expected test cases fails. Windows mutations need
CALENDULA_WINDOWS_TESTS=1 on WSL; without it they are skipped.

Usage: uv run python scripts/mutation_check.py
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src" / "calendula"

# (spec rule, file, original text, broken text, test cases expected to fail)
MUTATIONS = [
    ("B-DATE-1", "events.py", "int(m[1] or today.year)", "int(m[1] or 2000)", ["a3"]),
    ("B-TIME-1", "events.py", "int(m[1]) > 23", "int(m[1]) > 24", ["a6"]),
    (
        "B-EV-2",
        "events.py",
        'return "" if v is None else str(v).strip()',
        "return str(v).strip()",
        ["v1"],
    ),
    (
        "B-DAY-3",
        "events.py",
        "key=lambda e: (not e.all_day, e.start)",
        "key=lambda e: (e.all_day, e.start)",
        ["v3"],
    ),
    (
        "B-PICK-2 end",
        "events.py",
        "e.start <= at < e.end",
        "e.start <= at <= e.end",
        ["p2"],
    ),
    (
        "B-PICK-2 next day",
        "events.py",
        "e.start <= at < e.end",
        "e.start.time() <= at.time() < e.end.time()",
        ["p7"],
    ),
    (
        "3.1 span",
        "events.py",
        'end = "24:00"',
        'end = f"{self.end:%m-%d %H:%M}"',
        ["v8"],
    ),
    (
        "B-PICK-3",
        "events.py",
        "max(running, key=lambda e: e.start)",
        "min(running, key=lambda e: e.start)",
        ["p3"],
    ),
    (
        "B-PICK-4",
        "events.py",
        "min(timed, key=lambda e: abs(",
        "min(reversed(timed), key=lambda e: abs(",
        ["p6"],
    ),
    (
        "B-CACHE-2 row 3 bound",
        "cache.py",
        "if day.fetched.date() > d:",
        "if day.fetched.date() >= d:",
        ["k6"],
    ),
    (
        "B-CACHE-2 row 3 gone",
        "cache.py",
        "if day.fetched.date() > d:",
        "if False:",
        ["k5"],
    ),
    (
        "B-CACHE-2 row 5",
        "cache.py",
        "return now - day.fetched >= self._ttl",
        "return now - day.fetched > self._ttl",
        ["k3"],
    ),
    (
        "B-CACHE-3",
        "cache.py",
        "minutes if minutes > 0 else",
        "minutes if minutes >= 0 else",
        ["k13"],
    ),
    (
        "B-CACHE-4",
        "cache.py",
        "span = (last - first).days + 1",
        "span = len(need)",
        ["k8"],
    ),
    (
        "B-CACHE-7",
        "cache.py",
        "self._lock = threading.Lock()",
        "self._lock = type('L', (), {'__enter__': lambda s: s, '__exit__': lambda s, *a: None})()",
        ["k11"],
    ),
    ("E2", "outlook.py", "|80010108|RPC_E_", "|80010108", ["o4"]),
    (
        "E8",
        "outlook.py",
        """    if not _outlook_running():
        raise OutlookError("Outlook (classic) is not running. Start it and try again.")
""",
        "",
        ["o10"],
    ),
    (
        "E9",
        "outlook.py",
        "            timeout=TASKLIST_TIMEOUT,\n",
        "",
        ["o11"],
    ),
    (
        "B-ANN-1",
        "server.py",
        "READ_ONLY = ToolAnnotations(readOnlyHint=True,",
        "READ_ONLY = ToolAnnotations(readOnlyHint=False,",
        ["m1"],
    ),
    ("B-RST-4", "server.py", "            cal.clear()\n", "", ["s4"]),
    (
        "B-RST-5 text",
        "outlook.py",
        """            capture_output=True,
            check=False,
            timeout=TASKLIST_TIMEOUT,
        )
    except subprocess.TimeoutExpired:
        raise OutlookError(
            f"tasklist.exe did not answer within {TASKLIST_TIMEOUT} seconds."
        ) from None
    return b'"outlook.exe"' in r.stdout.lower()""",
        """            capture_output=True,
            text=True,
            check=False,
            timeout=TASKLIST_TIMEOUT,
        )
    except subprocess.TimeoutExpired:
        raise OutlookError(
            f"tasklist.exe did not answer within {TASKLIST_TIMEOUT} seconds."
        ) from None
    return '"outlook.exe"' in r.stdout.lower()""",
        ["s5"],
    ),
    (
        "B-RST-5 case",
        "outlook.py",
        """return b'"outlook.exe"' in r.stdout.lower()""",
        """return b'"OUTLOOK.EXE"' in r.stdout""",
        ["s5"],
    ),
    (
        "B-ERR-1",
        "server.py",
        """    except Exception as e:
        # The SDK would tell the client only "Error executing tool <name>". The client is the
        # user's own, so it gets the text too; the traceback goes to the log, as the SDK's would.
        logger.exception("Unexpected failure in a tool")
        raise ToolError(f"{type(e).__name__}: {e}") from None
""",
        "",
        ["m9"],
    ),
    (
        "B-EV-3 (Windows)",
        "get_calendar.ps1",
        "-match '^_{10,}$'",
        "-match '^_{5,}$'",
        ["w2"],
    ),
    (
        "B-DAY-1 (Windows)",
        "get_calendar.ps1",
        "' AND [End] > '",
        "' AND [End] >= '",
        ["w5"],
    ),
    (
        "3.3 arrays (Windows)",
        "get_calendar.ps1",
        "return ,$events.ToArray()",
        "return $events.ToArray()",
        ["w4"],
    ),
]


def run_tests(windows: bool) -> set[str]:
    """IDs (like 'k3') of the failing test cases."""
    env = dict(os.environ)
    if windows:
        env["CALENDULA_WINDOWS_TESTS"] = "1"
    target = ["tests/test_windows.py"] if windows else ["tests"]
    r = subprocess.run(
        [
            sys.executable,
            "-m",
            "pytest",
            "-q",
            "-p",
            "no:cacheprovider",
            "-rf",
            *target,
        ],
        cwd=ROOT,
        env=env,
        capture_output=True,
        text=True,
    )
    return set(re.findall(r"^FAILED \S+::test_([a-z]\d+)_", r.stdout, re.M))


def main() -> int:
    windows_ok = os.environ.get("CALENDULA_WINDOWS_TESTS") == "1"
    missed = 0
    for rule, name, old, new, expected in MUTATIONS:
        is_windows = name.endswith(".ps1")
        if is_windows and not windows_ok:
            print(f"skip    {rule:<22} (set CALENDULA_WINDOWS_TESTS=1)")
            continue
        path = SRC / name
        original = path.read_text(encoding="utf-8")
        if original.count(old) != 1:
            print(
                f"ERROR   {rule:<22} the original text is not found exactly once in {name}"
            )
            return 2
        try:
            path.write_text(original.replace(old, new), encoding="utf-8")
            failed = run_tests(is_windows)
        finally:
            path.write_text(original, encoding="utf-8")
        caught = sorted(set(expected) & failed)
        status = "caught" if caught else "MISSED"
        missed += not caught
        print(f"{status:<7} {rule:<22} failing: {', '.join(sorted(failed)) or '-'}")
    return 1 if missed else 0


if __name__ == "__main__":
    sys.exit(main())
