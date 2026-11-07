"""4.5 Bridge to Windows (spec 2.6, 3.3, 3.4). subprocess, shutil.which and tempfile are replaced."""

import json
import subprocess
from datetime import date
from pathlib import Path

import pytest
from conftest import NO_MATCH

from calendula import outlook

DAY = date(2026, 11, 5)


@pytest.fixture
def windows(monkeypatch, tmp_path):
    """Pretend to be WSL. Each wscript run calls `script(workdir)`, which plays the PowerShell side."""
    monkeypatch.setattr(outlook.shutil, "which", lambda name: f"/usr/bin/{name}")
    state = {"runs": [], "workdirs": [], "script": lambda w: None, "running": True}

    def fake_run(cmd, **kwargs):
        if cmd[0] == "wslpath":
            return subprocess.CompletedProcess(
                cmd, 0, stdout="C:\\fake" + cmd[-1].replace("/", "\\") + "\n"
            )
        if cmd[0] == "tasklist.exe":
            state["tasklist"] = kwargs
            if isinstance(state["running"], Exception):
                raise state["running"]
            out = (
                b'"OUTLOOK.EXE","1234","Console","1","200,000 K"\r\n'
                if state["running"]
                else NO_MATCH
            )
            return subprocess.CompletedProcess(cmd, 0, stdout=out)
        state["runs"].append((cmd, kwargs))
        state["script"](state["workdirs"][-1])
        return subprocess.CompletedProcess(cmd, 0)

    real = outlook.tempfile.TemporaryDirectory

    class Tmp(real):
        def __init__(self, *a, **kw):
            super().__init__(*a, dir=tmp_path, **kw)

        def __enter__(self):
            path = super().__enter__()
            state["workdirs"].append(Path(path))
            return path

    monkeypatch.setattr(outlook.tempfile, "TemporaryDirectory", Tmp)
    monkeypatch.setattr(outlook.subprocess, "run", fake_run)
    return state


def write_events(workdir: Path, days=None):
    data = {
        "days": days if days is not None else [{"date": DAY.isoformat(), "events": []}]
    }
    (workdir / "events.json").write_text(json.dumps(data), encoding="utf-8")


def test_o1_outside_wsl(monkeypatch):
    monkeypatch.setattr(outlook.shutil, "which", lambda name: None)
    with pytest.raises(outlook.OutlookError) as e:
        outlook.fetch(DAY, 1)
    assert (
        "needs WSL" in str(e.value)
        and "wslpath" in str(e.value)
        and "wscript.exe" in str(e.value)
        and "tasklist.exe" in str(e.value)
    )
    assert not e.value.restart


def test_o2_fetch_prepares_and_runs_the_script(windows):
    seen = {}

    def script(w: Path):
        seen["start"] = (w / "start.txt").read_text()
        seen["days"] = (w / "days.txt").read_text()
        seen["ps1"] = (w / "get_calendar.ps1").read_bytes()
        seen["vbs"] = (w / "run_hidden.vbs").read_text()
        write_events(w)

    windows["script"] = script
    assert outlook.fetch(DAY, 3) == {DAY: []}
    assert seen["start"].strip() == "20261105" and seen["days"].strip() == "3"
    source = (outlook.PACKAGE_DIR / "get_calendar.ps1").read_text(encoding="utf-8")
    assert (
        seen["ps1"][:2] == b"\xff\xfe" and seen["ps1"][2:].decode("utf-16-le") == source
    )
    assert seen["vbs"] == (outlook.PACKAGE_DIR / "run_hidden.vbs").read_text()
    ((cmd, kwargs),) = windows["runs"]
    assert cmd[:3] == ["wscript.exe", "//B", "//Nologo"]
    assert cmd[3].startswith("C:\\fake") and cmd[3].endswith("run_hidden.vbs")
    assert cmd[4].startswith("C:\\fake") and cmd[4].endswith("get_calendar.ps1")
    assert kwargs["timeout"] == outlook.FETCH_TIMEOUT == 180


def test_o3_the_temporary_directory_is_removed(windows):
    windows["script"] = write_events
    outlook.fetch(DAY, 1)
    windows["script"] = lambda w: (w / "error.txt").write_text("ERROR: x")
    with pytest.raises(outlook.OutlookError):
        outlook.fetch(DAY, 1)
    assert len(windows["workdirs"]) == 2 and not any(
        w.exists() for w in windows["workdirs"]
    )


@pytest.mark.parametrize(
    "code",
    [
        "CO_E_SERVER_EXEC_FAILURE",
        "80080005",
        "800706BA",
        "800706BE",
        "80010108",
        "RPC_E_DISCONNECTED",
    ],
)
def test_o4_com_errors_suggest_a_restart(tmp_path, code):
    (tmp_path / "error.txt").write_text(f"ERROR: failed with {code}", encoding="utf-8")
    with pytest.raises(outlook.OutlookError) as e:
        outlook.read_result(tmp_path)
    assert e.value.restart
    assert str(e.value) == f"ERROR: failed with {code}\n{outlook.RESTART_HINT}"


def test_o5_other_errors_are_passed_on_as_they_are(tmp_path):
    (tmp_path / "error.txt").write_text("ERROR: something else", encoding="utf-8")
    with pytest.raises(outlook.OutlookError) as e:
        outlook.read_result(tmp_path)
    assert not e.value.restart and str(e.value) == "ERROR: something else"


def test_o6_no_output(tmp_path):
    with pytest.raises(outlook.OutlookError) as e:
        outlook.read_result(tmp_path)
    assert e.value.restart and str(e.value).startswith("Outlook gave no answer")


def test_o7_timeout(windows):
    def hang(workdir):
        raise subprocess.TimeoutExpired("wscript.exe", outlook.FETCH_TIMEOUT)

    windows["script"] = hang
    with pytest.raises(outlook.OutlookError) as e:
        outlook.fetch(DAY, 1)
    assert e.value.restart and str(e.value).startswith(
        "Outlook did not answer within 180 seconds."
    )


@pytest.mark.parametrize(
    "content", ["not json", json.dumps({"days": [{"events": []}]})]
)
def test_o8_unexpected_output(tmp_path, content):
    (tmp_path / "events.json").write_text(content, encoding="utf-8")
    with pytest.raises(outlook.OutlookError, match="Unexpected output from Outlook"):
        outlook.read_result(tmp_path)


@pytest.mark.parametrize("encoding", ["utf-8", "utf-8-sig"])
def test_o9_utf8_with_or_without_a_bom(tmp_path, encoding):
    data = {
        "days": [
            {
                "date": "2026-11-05",
                "events": [
                    {
                        "start": "2026-11-05T09:00:00",
                        "end": "2026-11-05T10:00:00",
                        "subject": "日本語の件名",
                    }
                ],
            }
        ]
    }
    (tmp_path / "events.json").write_text(
        json.dumps(data, ensure_ascii=False), encoding=encoding
    )
    assert outlook.read_result(tmp_path)[DAY][0].subject == "日本語の件名"


def test_o10_outlook_not_running(windows):
    windows["running"] = False
    with pytest.raises(outlook.OutlookError) as e:
        outlook.fetch(DAY, 1)
    assert str(e.value) == (
        "Outlook (classic) is not running. Start it and try again.\n"
        + outlook.START_HINT
    )
    assert not e.value.restart
    assert windows["runs"] == [] and windows["workdirs"] == []


def test_o11_tasklist_that_never_answers(windows):
    windows["running"] = subprocess.TimeoutExpired("tasklist.exe", 30)
    with pytest.raises(outlook.OutlookError) as e:
        outlook.fetch(DAY, 1)
    assert str(e.value) == "tasklist.exe did not answer within 30 seconds."
    assert not e.value.restart
    assert windows["tasklist"]["timeout"] == outlook.QUERY_TIMEOUT == 30
    assert windows["runs"] == [] and windows["workdirs"] == []
