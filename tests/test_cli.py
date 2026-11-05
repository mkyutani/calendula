"""4.8 Command line (spec 3.2). Outlook is replaced."""

from datetime import timedelta

import pytest
from conftest import DAY, ev

from calendula import __version__, cli, outlook


@pytest.fixture
def fetches(monkeypatch):
    calls = []

    def fetch(start, days):
        calls.append((start, days))
        return {start + timedelta(days=i): [] for i in range(days)} | {
            DAY: [
                ev("09:00", "10:00", "Weekly sync", location="Room A"),
                ev("13:30", "14:30", "Interview"),
            ]
        }

    monkeypatch.setattr(cli.outlook, "fetch", fetch)
    return calls


def test_l1_version(capsys):
    with pytest.raises(SystemExit) as e:
        cli.main(["--version"])
    assert (
        e.value.code == 0
        and capsys.readouterr().out.strip() == f"calendula {__version__}"
    )


def test_l2_list(fetches, capsys):
    assert cli.main(["list", "2026-11-05", "2"]) == 0
    out = capsys.readouterr().out.splitlines()
    assert fetches == [(DAY, 2)]
    assert out[0].startswith("2026-11-05 (Thu)  fetched ")
    assert out[1:3] == ["09:00-10:00  Weekly sync  [Room A]", "13:30-14:30  Interview"]
    assert (
        out[3] == ""
        and out[4].startswith("2026-11-06 (Fri)")
        and out[5] == "(no events)"
    )


def test_l3_show(fetches, capsys):
    assert cli.main(["show", "12:00", "--date", "2026-11-05"]) == 0
    out = capsys.readouterr().out.splitlines()
    assert out[1:] == [
        "No event is running at 12:00; this is the one starting nearest to it.",
        "13:30-14:30  Interview",
    ]


@pytest.mark.parametrize(
    ("argv", "message"),
    [
        (["list", "tomorrow"], "invalid date"),
        (["list", "2026-11-05", "0"], "between 1 and 31"),
        (["show", "25:00"], "invalid time"),
    ],
)
def test_l4_errors_exit_1(fetches, capsys, argv, message):
    assert cli.main(argv) == 1
    err = capsys.readouterr().err
    assert err.startswith("calendula: ") and message in err


def test_l4_outlook_errors_exit_1(monkeypatch, capsys):
    def broken(start, days):
        raise outlook.OutlookError(
            "Outlook gave no answer (is Outlook (classic) running?)."
        )

    monkeypatch.setattr(cli.outlook, "fetch", broken)
    assert cli.main(["list"]) == 1
    assert capsys.readouterr().err.startswith("calendula: Outlook gave no answer")


def test_l5_no_command_runs_the_server(monkeypatch):
    ran = []

    class FakeServer:
        def run(self, transport):
            ran.append(transport)

    monkeypatch.setattr("calendula.server.build_server", lambda: FakeServer())
    assert cli.main([]) == 0 and ran == ["stdio"]
