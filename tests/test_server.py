"""4.7 MCP server (spec 3.1, 2.6, 2.8), through the MCP client."""

import asyncio
import sys
from datetime import date, datetime, timedelta

from conftest import DAY, Clock, FakeOutlook, call_tool, ev, text
from mcp import StdioServerParameters
from mcp.client import Client

from calendula.cache import Calendar
from calendula.outlook import RESTART_HINT, OutlookError
from calendula.server import build_server

FETCHED = datetime(2026, 11, 5, 10, 32)


def fixed(days: dict[date, list]):
    """A fetch that returns the given days (and nothing for others)."""
    return lambda start, n: {
        d: days.get(d, []) for d in (start + timedelta(days=i) for i in range(n))
    }


def server_with(days, clock=None):
    return build_server(Calendar(fixed(days), now=clock or Clock(FETCHED)))


def tools_and_instructions():
    async def run():
        async with Client(build_server(Calendar(FakeOutlook()))) as c:
            return {t.name: t for t in (await c.list_tools()).tools}, c.instructions

    return asyncio.run(run())


def test_m1_tools_and_annotations():
    tools, _ = tools_and_instructions()
    assert set(tools) == {"list_events", "get_event", "restart_outlook"}
    for name in ("list_events", "get_event"):
        a = tools[name].annotations
        assert (a.read_only_hint, a.idempotent_hint, a.open_world_hint) == (
            True,
            True,
            False,
        )
    assert tools["restart_outlook"].annotations.destructive_hint is True
    assert tools["restart_outlook"].annotations.read_only_hint is False


def test_m2_argument_schemas():
    tools, _ = tools_and_instructions()
    listing = tools["list_events"].input_schema
    assert listing.get("required", []) == []
    assert {"default": 1, "minimum": 1, "maximum": 31}.items() <= listing["properties"][
        "days"
    ].items()
    assert listing["properties"]["refresh"]["default"] is False
    getting = tools["get_event"].input_schema
    assert getting["required"] == ["time"]
    assert set(getting["properties"]) == {"time", "date", "refresh"}
    assert tools["restart_outlook"].input_schema["properties"] == {}


def test_m3_instructions():
    _, instructions = tools_and_instructions()
    for phrase in (
        "Read-only",
        "60 minutes",
        "refresh=true",
        "YYYY-MM-DD",
        "Never fill in events",
        "ask the user for permission before calling restart_outlook",
        "retry once",
    ):
        assert phrase in instructions, phrase


def test_m4_list_events_text():
    days = {
        DAY: [
            ev("09:00", "10:00", "Weekly sync", location="Room A"),
            ev("13:30", "14:30", "Interview"),
        ]
    }
    r = call_tool(server_with(days), "list_events", {"date": "2026-11-05", "days": 2})
    assert not r.is_error
    assert text(r) == (
        "2026-11-05 (Thu)  fetched 2026-11-05 10:32\n"
        "09:00-10:00  Weekly sync  [Room A]\n"
        "13:30-14:30  Interview\n"
        "\n"
        "2026-11-06 (Fri)  fetched 2026-11-05 10:32\n"
        "(no events)"
    )


def test_m5_get_event_text():
    days = {
        DAY: [
            ev("09:00", "10:00", "Weekly sync"),
            ev("13:30", "14:30", "Interview", required="Member A"),
        ],
        DAY + timedelta(days=1): [
            ev("2026-11-06T00:00", "2026-11-07T00:00", "Holiday", all_day=True)
        ],
    }
    server = server_with(days)
    head = "2026-11-05 (Thu)  fetched 2026-11-05 10:32\n"
    r = call_tool(server, "get_event", {"date": "2026-11-05", "time": "13:45"})
    assert text(r) == head + "13:30-14:30  Interview\nRequired attendees: Member A"
    r = call_tool(server, "get_event", {"date": "2026-11-05", "time": "12:00"})
    assert (
        text(r)
        == head
        + "No event is running at 12:00; this is the one starting nearest to it.\n13:30-14:30  Interview\nRequired attendees: Member A"
    )
    r = call_tool(server, "get_event", {"date": "2026-11-06", "time": "12:00"})
    assert (
        text(r)
        == "2026-11-06 (Fri)  fetched 2026-11-05 10:32\nNo timed events on this day."
    )


def test_m6_errors_are_tool_errors_without_events():
    def broken(start, days):
        raise OutlookError("ERROR: 0x80080005 CO_E_SERVER_EXEC_FAILURE", restart=True)

    r = call_tool(build_server(Calendar(broken)), "list_events", {"date": "2026-11-05"})
    assert r.is_error
    assert "CO_E_SERVER_EXEC_FAILURE" in text(r) and RESTART_HINT in text(r)
    assert "fetched" not in text(r)
    r = call_tool(server_with({}), "get_event", {"date": "2026-11-05", "time": "25:00"})
    assert r.is_error and "invalid time" in text(r)


def test_m9_unexpected_failures_carry_their_type_and_text(caplog):
    def broken_fetch(start, days):
        raise OSError("no such command")

    def broken_restart():
        raise OSError("no such command")

    server = build_server(Calendar(broken_fetch), restart=broken_restart)
    for tool, args in (
        ("list_events", {"date": "2026-11-05"}),
        ("get_event", {"date": "2026-11-05", "time": "10:00"}),
        ("restart_outlook", {}),
    ):
        caplog.clear()
        r = call_tool(server, tool, args)
        assert r.is_error, tool
        # The SDK puts "Error executing tool <name>: " in front of every tool error.
        assert text(r).endswith("OSError: no such command"), tool
        assert any(rec.exc_info for rec in caplog.records), tool


def test_m7_answers_report_the_cached_fetch_time():
    clock = Clock(FETCHED)
    server = server_with({DAY: [ev("09:00", "10:00")]}, clock)
    first = text(call_tool(server, "list_events", {"date": "2026-11-05"})).splitlines()[
        0
    ]
    clock.now += timedelta(minutes=30)
    again = text(call_tool(server, "list_events", {"date": "2026-11-05"})).splitlines()[
        0
    ]
    fresh = text(
        call_tool(server, "list_events", {"date": "2026-11-05", "refresh": True})
    ).splitlines()[0]
    assert first == again == "2026-11-05 (Thu)  fetched 2026-11-05 10:32"
    assert fresh == "2026-11-05 (Thu)  fetched 2026-11-05 11:02"


def test_m8_stdio_entry_point():
    async def run():
        params = StdioServerParameters(
            command=sys.executable, args=["-m", "calendula.cli"]
        )
        async with Client(params) as c:
            return sorted(
                t.name for t in (await c.list_tools()).tools
            ), c.server_info.name

    names, server_name = asyncio.run(run())
    assert (
        names == ["get_event", "list_events", "restart_outlook"]
        and server_name == "calendula"
    )
