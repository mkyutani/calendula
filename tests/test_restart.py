"""4.6 Starting and restarting Outlook (spec 2.7). Windows commands and time are replaced."""

import subprocess

import pytest
from conftest import NO_MATCH, FakeOutlook, call_tool, text

from calendula import outlook
from calendula.cache import Calendar
from calendula.server import build_server

# The profiles key as reg.exe prints it, with HKCU spelled out.
PRINTED_KEY = outlook.PROFILES_KEY.replace("HKCU", "HKEY_CURRENT_USER")
NO_SUBKEYS = b"\r\n"  # what reg.exe prints for a key that exists but is empty


def subkey(name: str) -> bytes:
    """What reg.exe prints for a subkey of the profiles key. Names come in the console code page."""
    return f"\r\n{PRINTED_KEY}\\{name}\r\n".encode("cp932")


class FakeWindows:
    """tasklist answers from `running`; taskkill stops it; Start-Process starts it after `start_delay` seconds.

    reg.exe answers `reg_output`, or raises it if it is an exception.
    Output is bytes, and text=True decodes it strictly as UTF-8, as the real subprocess.run does here.
    """

    def __init__(
        self,
        running: bool,
        start_delay: float | None = 6,
        image: str = "OUTLOOK.EXE",
        reg_output: bytes | Exception = NO_SUBKEYS,
    ):
        self.running = running
        self.start_delay = start_delay
        self.image = image
        self.reg_output = reg_output
        self.clock = 0.0
        self.started_at = None
        self.commands = []

    def run(self, cmd, text=False, **kwargs):
        self.commands.append(cmd[0])
        out = b""
        if cmd[0] == "tasklist.exe":
            if (
                self.started_at is not None
                and self.start_delay is not None
                and self.clock >= self.started_at + self.start_delay
            ):
                self.running = True
            out = (
                f'"{self.image}","1234","Console","1","200,000 K"\r\n'.encode("ascii")
                if self.running
                else NO_MATCH
            )
        if cmd[0] == "reg.exe":
            if isinstance(self.reg_output, Exception):
                raise self.reg_output
            out = self.reg_output
        if cmd[0] == "taskkill.exe":
            self.running = False
        if cmd[0] == "powershell.exe":
            self.started_at = self.clock
        return subprocess.CompletedProcess(
            cmd, 0, stdout=out.decode("utf-8") if text else out
        )

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


@pytest.mark.parametrize(
    "running, image, answer",
    [
        (True, "OUTLOOK.EXE", "Outlook was closed. Outlook was started again."),
        (True, "outlook.exe", "Outlook was closed. Outlook was started again."),
        (False, "OUTLOOK.EXE", "Outlook was not running. Outlook was started again."),
    ],
)
def test_s5_restart_succeeds_when_tasklist_answers_in_cp932(
    fake, running, image, answer
):
    w = fake(running=running, image=image)
    assert outlook.restart() == answer
    assert w.running


def test_s6_start_leaves_a_running_outlook_alone(fake):
    w = fake(running=True)
    assert outlook.start() == "Outlook is already running."
    assert set(w.commands) == {"tasklist.exe"}


def test_s7_start_refuses_on_a_pc_with_no_mail_profile(fake):
    w = fake(running=False)
    with pytest.raises(outlook.OutlookError) as e:
        outlook.start()
    assert str(e.value) == (
        "Outlook (classic) has no mail profile on this PC. Set it up in Outlook first."
    )
    assert "powershell.exe" not in w.commands
    assert not w.running


def test_s8_start_opens_outlook_and_waits(fake):
    w = fake(running=False, reg_output=subkey("プロファイル"))
    assert outlook.start() == "Outlook was started."
    assert w.running and "taskkill.exe" not in w.commands
    # It waited for the start (6 s, polled every 3 s) and then RESTART_SETTLE seconds more.
    assert w.clock >= 6 + outlook.RESTART_SETTLE


def test_s9_outlook_that_never_starts_is_an_error(fake):
    w = fake(running=False, start_delay=None, reg_output=subkey("Outlook"))
    with pytest.raises(outlook.OutlookError, match="did not start within 120 seconds"):
        outlook.start()
    assert w.clock >= outlook.RESTART_WAIT


def test_s10_the_tool_keeps_the_cache():
    fetches = FakeOutlook()
    ok = build_server(Calendar(fetches), start=lambda: "Outlook was started.")
    call_tool(ok, "list_events", {"date": "2026-11-05"})
    assert text(call_tool(ok, "start_outlook")) == "Outlook was started."
    call_tool(ok, "list_events", {"date": "2026-11-05"})
    assert len(fetches.calls) == 1

    def broken():
        raise outlook.OutlookError("Outlook (classic) has no mail profile on this PC.")

    r = call_tool(build_server(Calendar(FakeOutlook()), start=broken), "start_outlook")
    assert r.is_error and "no mail profile" in text(r)


OWN_LINE = f"\r\n{PRINTED_KEY}\r\n    Some value    REG_SZ    x\r\n".encode("ascii")


@pytest.mark.parametrize(
    "reg_output, expected",
    [
        (subkey("Outlook"), True),
        (subkey("プロファイル"), True),
        (subkey("Outlook").lower(), True),
        (OWN_LINE + subkey("Outlook"), True),
        (b"", False),  # the key does not exist
        (NO_SUBKEYS, False),
        (OWN_LINE, False),
    ],
)
def test_s11_mail_profile_check(fake, reg_output, expected):
    fake(running=False, reg_output=reg_output)
    assert outlook._has_subkey(outlook.PROFILES_KEY) is expected


def test_s11_reg_that_never_answers(fake):
    w = fake(running=False, reg_output=subprocess.TimeoutExpired("reg.exe", 30))
    with pytest.raises(outlook.OutlookError) as e:
        outlook.start()
    assert str(e.value) == "reg.exe did not answer within 30 seconds."
    assert "powershell.exe" not in w.commands
