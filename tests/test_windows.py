"""4.9 Windows, without Outlook (opt-in: CALENDULA_WINDOWS_TESTS=1, on WSL).

Runs the real Windows PowerShell on tests/windows/harness.ps1, which dot-sources get_calendar.ps1
and exercises its functions on fake items. Outlook is never contacted.
"""

import json
import os
import re
import shutil
import subprocess
from datetime import date
from pathlib import Path

import pytest

from calendula import outlook

pytestmark = pytest.mark.skipif(
    os.environ.get("CALENDULA_WINDOWS_TESTS") != "1"
    or not all(shutil.which(c) for c in ("powershell.exe", "wscript.exe", "wslpath")),
    reason="Windows tests are opt-in: set CALENDULA_WINDOWS_TESTS=1 on WSL",
)

HERE = Path(__file__).parent


def utf16(src: Path, dst: Path) -> None:
    dst.write_bytes(b"\xff\xfe" + src.read_text(encoding="utf-8").encode("utf-16-le"))


def winpath(p: Path) -> str:
    return subprocess.run(
        ["wslpath", "-w", str(p)], check=True, capture_output=True, text=True
    ).stdout.strip()


@pytest.fixture(scope="module")
def harness(tmp_path_factory):
    work = tmp_path_factory.mktemp("harness")
    utf16(outlook.PACKAGE_DIR / "get_calendar.ps1", work / "get_calendar.ps1")
    utf16(HERE / "windows" / "harness.ps1", work / "harness.ps1")
    out = work / "result.json"
    r = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            winpath(work / "harness.ps1"),
            "-Out",
            winpath(out),
        ],
        capture_output=True,
        timeout=120,
    )
    assert out.exists(), r.stderr.decode("utf-8", "replace") + r.stdout.decode(
        "utf-8", "replace"
    )
    return json.loads(out.read_text(encoding="utf-8-sig"))


def test_w1_script_parses(harness):
    assert harness["parse_errors"] == []


def test_w2_body_lines(harness):
    b = harness["bodies"]
    assert b["teams_middle"] == ["Agenda", "1. Budget", "Thanks"]
    assert b["teams_end"] == ["Agenda"]
    assert b["quoted"] == ["Hi", "_" * 80, "From: Someone", "Old text"]
    assert b["short_line"] == ["_____", "Microsoft Teams", "kept"]
    assert b["spaces"] == ["a", "b"]
    assert b["empty"] == []
    # A block with no opening line, and lines of 32 underscores inside a block.
    assert b["no_opening"] == ["Text of the invitation"]
    assert b["no_opening_en"] == ["Text"]
    assert b["inner_lines"] == ["Before", "After"]
    assert b["need_help"] == ["Before", "After"]
    # Not a block.
    assert b["plain_heading"] == ["開催形式", "Microsoft Teams 会議", "議題"]
    assert b["sentence"] == [
        "Microsoft Teams で開催します",
        "https://teams.microsoft.invalid/meet/000?p=x",
        "議題",
    ]
    assert b["quoted_32"] == ["Hi", "_" * 32, "From: Someone", "Old text"]
    # The known limits: these lose text instead of keeping boilerplate (L9, L10).
    assert b["limit_url"] == []
    assert b["limit_end"] == ["Before"]
    # The boundaries of the opening and closing lines.
    assert b["open_9"] == ["Before", "_" * 9, "Microsoft Teams meeting", "After"]
    assert b["open_10"] == ["Before"]
    assert b["close_39"] == ["Before"]
    assert b["close_40"] == ["Before", "After"]


def test_w3_event_record(harness):
    assert harness["record"] == {
        "start": "2026-11-05T09:05:00",
        "end": "2026-11-05T10:00:00",
        "all_day": False,
        "subject": "",
        "location": "",
        "organizer": "",
        "required": "",
        "optional": "",
        "categories": "",
        "busy": 1,
        "body": ["one line"],
    }


def test_w4_day_events(harness):
    d = harness["day_events"]
    for name, count, subjects in (
        ("zero", 0, []),
        ("one", 1, ["solo"]),
        ("two", 2, ["early", "late"]),
    ):
        assert d[name]["is_array"] is True, name
        assert d[name]["count"] == count and d[name]["subjects"] == subjects, name
        assert d[name]["include_recurrences"] is True, name
        assert d[name]["filter"] == harness["filter"], name


def test_w5_filter_is_a_day_wider_on_each_side(harness):
    assert re.fullmatch(r"\[Start\] < '[^']+' AND \[End\] > '[^']+'", harness["filter"])
    # read back in the current Windows locale, for the day 2026-11-05
    assert harness["filter_times"] == ["2026-11-07T00:00:00", "2026-11-04T00:00:00"]


def test_w6_json(harness):
    days = json.loads(harness["json"])["days"]
    assert [len(d["events"]) for d in days] == [0, 1, 2]
    e = days[1]["events"][0]
    assert e["subject"] == '日本語の件名 "Q4" <draft>' and e["body"] == [
        "一行目",
        "2. 二行目",
    ]
    parsed = outlook.parse_days({"days": days})
    assert [len(parsed[date(2026, 11, d)]) for d in (5, 6, 7)] == [0, 1, 2]


def test_w7_plumbing_through_wscript(tmp_path):
    probe = tmp_path / "probe.ps1"
    probe.write_text(
        '$s = (Get-Content "$PSScriptRoot\\start.txt" -Raw).Trim()\n'
        '$d = $s.Substring(0,4) + "-" + $s.Substring(4,2) + "-" + $s.Substring(6,2)\n'
        '$j = \'{"days":[{"date":"\' + $d + \'","events":[{"start":"\' + $d + \'T09:00:00","end":"\' + $d + \'T10:00:00","subject":"日本語の件名"}]}]}\'\n'
        '[System.IO.File]::WriteAllText("$PSScriptRoot\\events.json", $j, [System.Text.UTF8Encoding]::new($false))\n',
        encoding="utf-8",
    )
    work = tmp_path / "work"
    work.mkdir()
    (work / "start.txt").write_text("20261105\n")
    outlook._run_powershell(probe, work)
    assert outlook.read_result(work)[date(2026, 11, 5)][0].subject == "日本語の件名"


def test_w8_running_check_on_the_real_tasklist(monkeypatch):
    # No process has this name, so tasklist.exe answers in the console code page. Outlook is not contacted.
    real_run = subprocess.run

    def other_image(cmd, **kwargs):
        return real_run(
            [a.replace("OUTLOOK.EXE", "calendula-no-such-image.exe") for a in cmd],
            **kwargs,
        )

    monkeypatch.setattr(outlook.subprocess, "run", other_image)
    assert outlook._outlook_running() is False


def test_w9_mail_profile_check_on_the_real_reg():
    # Any user has subkeys under HKCU\\Software. Only the yes or no comes back, never a name.
    assert outlook._has_subkey("HKCU\\Software") is True
    assert outlook._has_subkey("HKCU\\Software\\calendula-no-such-key") is False


def test_w10_only_appointments_that_overlap_the_day_in_local_time(harness):
    # Sorted by start. The fake Restrict returned all six, including those on the days around it.
    assert harness["day_events"]["boundaries"]["subjects"] == [
        "across_the_day",
        "past_midnight",
        "all_day",
    ]
