"""Bridge from WSL to Outlook (classic) on Windows through PowerShell and COM."""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
import time
from datetime import date
from pathlib import Path

from .events import Event, parse_days

PACKAGE_DIR = Path(__file__).resolve().parent
FETCH_TIMEOUT = 180  # seconds
QUERY_TIMEOUT = 30  # seconds, for tasklist.exe and reg.exe
PROFILES_KEY = r"HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles"
RESTART_WAIT = 120  # seconds to wait for OUTLOOK.EXE to appear
RESTART_SETTLE = 15  # seconds to wait after it appears, before COM answers
COM_ERROR_RE = re.compile(
    r"CO_E_SERVER_EXEC_FAILURE|80080005|800706BA|800706BE|80010108|RPC_E_", re.I
)

RESTART_HINT = (
    "Outlook's COM connection seems broken. Ask the user whether Outlook may be closed and restarted "
    "(unsaved drafts are lost); if they agree, call restart_outlook and then retry once."
)


START_HINT = "Ask the user whether Outlook may be started; if they agree, call start_outlook and then retry once."


class OutlookError(Exception):
    """Fetching from Outlook failed. `restart` or `start` is True when doing that to Outlook is likely to help."""

    def __init__(self, message: str, restart: bool = False, start: bool = False):
        super().__init__(message)
        self.restart = restart
        self.start = start

    def __str__(self) -> str:
        msg = super().__str__()
        if self.restart:
            return f"{msg}\n{RESTART_HINT}"
        return f"{msg}\n{START_HINT}" if self.start else msg


def _require(*commands: str) -> None:
    missing = [c for c in commands if shutil.which(c) is None]
    if missing:
        raise OutlookError(
            "calendula needs WSL with Outlook (classic) on Windows; not found: "
            + ", ".join(missing)
        )


def _winpath(path: Path) -> str:
    return subprocess.run(
        ["wslpath", "-w", str(path)], check=True, capture_output=True, text=True
    ).stdout.strip()


def _run_powershell(script: Path, workdir: Path) -> None:
    """Run a PowerShell script hidden (no console window) through wscript and run_hidden.vbs."""
    ps1 = workdir / script.name
    # Windows PowerShell 5.1 reads a BOM-less script as the ANSI code page; UTF-16LE with a BOM is safe.
    ps1.write_bytes(
        b"\xff\xfe" + script.read_text(encoding="utf-8").encode("utf-16-le")
    )
    vbs = workdir / "run_hidden.vbs"
    shutil.copyfile(PACKAGE_DIR / "run_hidden.vbs", vbs)
    try:
        subprocess.run(
            ["wscript.exe", "//B", "//Nologo", _winpath(vbs), _winpath(ps1)],
            check=False,
            timeout=FETCH_TIMEOUT,
            capture_output=True,
        )
    except subprocess.TimeoutExpired:
        raise OutlookError(
            f"Outlook did not answer within {FETCH_TIMEOUT} seconds.", restart=True
        ) from None


def read_result(workdir: Path) -> dict[date, list[Event]]:
    """Read events.json / error.txt that get_calendar.ps1 left in workdir."""
    error = workdir / "error.txt"
    if error.exists():
        text = error.read_text(encoding="utf-8-sig", errors="replace").strip()
        raise OutlookError(
            text or "ERROR: unknown failure", restart=bool(COM_ERROR_RE.search(text))
        )
    out = workdir / "events.json"
    if not out.exists():
        raise OutlookError(
            "Outlook gave no answer (is Outlook (classic) running?).", restart=True
        )
    try:
        return parse_days(json.loads(out.read_text(encoding="utf-8-sig")))
    except (ValueError, KeyError, TypeError) as e:
        raise OutlookError(f"Unexpected output from Outlook: {e}") from None


def fetch(start: date, days: int) -> dict[date, list[Event]]:
    """Events for `days` days from `start`, straight from Outlook."""
    _require("wslpath", "wscript.exe", "tasklist.exe")
    # Creating Outlook.Application would start a closed Outlook, which can stop at its setup screen.
    if not _outlook_running():
        raise OutlookError(
            "Outlook (classic) is not running. Start it and try again.", start=True
        )
    with tempfile.TemporaryDirectory(prefix="calendula-") as tmp:
        workdir = Path(tmp)
        (workdir / "start.txt").write_text(f"{start:%Y%m%d}\n", encoding="ascii")
        (workdir / "days.txt").write_text(f"{days}\n", encoding="ascii")
        _run_powershell(PACKAGE_DIR / "get_calendar.ps1", workdir)
        return read_result(workdir)


def _query(cmd: list[str]) -> bytes:
    """Output of a quick Windows command. Bytes, not text: it is in the console code page (B-RST-5)."""
    try:
        return subprocess.run(
            cmd, capture_output=True, check=False, timeout=QUERY_TIMEOUT
        ).stdout
    except subprocess.TimeoutExpired:
        raise OutlookError(
            f"{cmd[0]} did not answer within {QUERY_TIMEOUT} seconds."
        ) from None


def _outlook_running() -> bool:
    output = _query(
        ["tasklist.exe", "/fi", "IMAGENAME eq OUTLOOK.EXE", "/fo", "csv", "/nh"]
    )
    return b'"outlook.exe"' in output.lower()


def _has_subkey(key: str) -> bool:
    """Whether a registry key has a subkey. The exit code does not say: it can exist with none (B-START-5)."""
    # reg.exe prints each subkey in full on its own line. Only this yes or no leaves here,
    # because a profile name can be a mail address.
    full = key.replace("HKCU\\", "HKEY_CURRENT_USER\\", 1)
    prefix = full.lower().encode("ascii") + b"\\"
    lines = _query(["reg.exe", "query", key]).splitlines()
    return any(line.lower().startswith(prefix) for line in lines)


def _start_and_wait() -> None:
    """Start Outlook, and wait until it has appeared and COM can answer (B-RST-2)."""
    subprocess.run(
        ["powershell.exe", "-NoProfile", "-Command", "Start-Process outlook.exe"],
        capture_output=True,
        check=False,
    )
    deadline = time.monotonic() + RESTART_WAIT
    while not _outlook_running():
        if time.monotonic() > deadline:
            raise OutlookError(
                f"OUTLOOK.EXE did not start within {RESTART_WAIT} seconds."
            )
        time.sleep(3)
    time.sleep(RESTART_SETTLE)


def start() -> str:
    """Start Outlook if it is not running. It never closes Outlook. Returns what happened."""
    _require("tasklist.exe", "reg.exe", "powershell.exe")
    if _outlook_running():
        return "Outlook is already running."
    if not _has_subkey(PROFILES_KEY):
        raise OutlookError(
            "Outlook (classic) has no mail profile on this PC. Set it up in Outlook first."
        )
    _start_and_wait()
    return "Outlook was started."


def restart() -> str:
    """Force-quit Outlook and start it again. Returns what happened."""
    _require("tasklist.exe", "taskkill.exe", "powershell.exe")
    notes = []
    if _outlook_running():
        subprocess.run(
            ["taskkill.exe", "/f", "/im", "OUTLOOK.EXE"],
            capture_output=True,
            check=False,
        )
        notes.append("Outlook was closed.")
        # Starting it before the old process is gone hands the start over to that process.
        for _ in range(10):
            if not _outlook_running():
                break
            time.sleep(1)
    else:
        notes.append("Outlook was not running.")
    _start_and_wait()
    notes.append("Outlook was started again.")
    return " ".join(notes)
