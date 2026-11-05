"""4.6 Restarting Outlook (spec 2.7). Windows commands and time are replaced."""

import subprocess

import pytest
from conftest import FakeOutlook, call_tool, text

from calendula import outlook
from calendula.cache import Calendar
from calendula.server import build_server


class FakeWindows:
    """tasklist answers from `running`; taskkill stops it; Start-Process starts it after `start_delay` seconds."""

    def __init__(self, running: bool, start_delay: float | None = 6):
        self.running = running
        self.start_delay = start_delay
        self.clock = 0.0
        self.started_at = None
        self.commands = []

    def run(self, cmd, **kwargs):
        self.commands.append(cmd[0])
        if cmd[0] == "tasklist.exe":
            if (
                self.started_at is not None
                and self.start_delay is not None
                and self.clock >= self.started_at + self.start_delay
            ):
                self.running = True
            out = (
                '"OUTLOOK.EXE","1234","Console","1","200,000 K"\n'
                if self.running
                else "INFO: No tasks are running.\n"
            )
            return subprocess.CompletedProcess(cmd, 0, stdout=out)
        if cmd[0] == "taskkill.exe":
            self.running = False
        if cmd[0] == "powershell.exe":
            self.started_at = self.clock
        return subprocess.CompletedProcess(cmd, 0, stdout="")

    def sleep(self, seconds):
        self.clock += seconds

    def monotonic(self):
        return self.clock


@pytest.fixture
def fake(monkeypatch):
    def make(**kw):
        w = FakeWindows(**kw)
        monkeypatch.setattr(outlook.shutil, "which", lambda name: f"/usr/bin/{name}")
        monkeypatch.setattr(outlook.subprocess, "run", w.run)
        monkeypatch.setattr(outlook.time, "sleep", w.sleep)
        monkeypatch.setattr(outlook.time, "monotonic", w.monotonic)
        return w

    return make


def test_s1_running_outlook_is_closed_and_started(fake):
    w = fake(running=True)
    assert outlook.restart() == "Outlook was closed. Outlook was started again."
    assert w.commands.index("taskkill.exe") < w.commands.index("powershell.exe")
    assert w.running
    # It waited for the start (6 s, polled every 3 s) and then RESTART_SETTLE seconds more.
    assert w.clock >= 6 + outlook.RESTART_SETTLE


def test_s2_outlook_that_was_not_running_is_started(fake):
    w = fake(running=False)
    assert outlook.restart() == "Outlook was not running. Outlook was started again."
    assert "taskkill.exe" not in w.commands


def test_s3_outlook_that_never_starts_is_an_error(fake):
    w = fake(running=False, start_delay=None)
    with pytest.raises(outlook.OutlookError, match="did not start within 120 seconds"):
        outlook.restart()
    assert w.clock >= outlook.RESTART_WAIT


def test_s4_the_tool_clears_the_cache_only_on_success():
    fetches = FakeOutlook()
    ok = build_server(Calendar(fetches), restart=lambda: "Outlook was started again.")
    call_tool(ok, "list_events", {"date": "2026-11-05"})
    assert text(call_tool(ok, "restart_outlook")) == "Outlook was started again."
    call_tool(ok, "list_events", {"date": "2026-11-05"})
    assert len(fetches.calls) == 2

    def broken():
        raise outlook.OutlookError("OUTLOOK.EXE did not start within 120 seconds.")

    fetches = FakeOutlook()
    bad = build_server(Calendar(fetches), restart=broken)
    call_tool(bad, "list_events", {"date": "2026-11-05"})
    r = call_tool(bad, "restart_outlook")
    assert r.is_error and "did not start" in text(r)
    call_tool(bad, "list_events", {"date": "2026-11-05"})
    assert len(fetches.calls) == 1
