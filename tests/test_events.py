"""4.2 Event content and formats (spec 2.3, B-DAY-3, 3.1)."""

from datetime import date, datetime, timedelta

import pytest
from conftest import DAY, ev

from calendula.events import BUSY, day_heading, parse_days


def day_json(events):
    return {"days": [{"date": DAY.isoformat(), "events": events}]}


def test_v1_fields_and_missing_values():
    full = {
        "start": "2026-11-05T13:30:00",
        "end": "2026-11-05T14:30:00",
        "all_day": False,
        "subject": " Interview ",
        "location": "Room B",
        "organizer": "Host",
        "required": "Member A; Member B",
        "optional": "Guest",
        "categories": "Customer",
        "busy": 2,
        "body": ["Agenda", "1. Requirements"],
    }
    sparse = {
        "start": "2026-11-05T15:00:00",
        "end": "2026-11-05T16:00:00",
        "subject": None,
        "location": None,
    }
    a, b = parse_days(day_json([full, sparse]))[DAY]
    assert (
        a.subject,
        a.location,
        a.organizer,
        a.required,
        a.optional,
        a.categories,
        a.busy,
    ) == (
        "Interview",
        "Room B",
        "Host",
        "Member A; Member B",
        "Guest",
        "Customer",
        2,
    )
    assert a.body == ("Agenda", "1. Requirements") and not a.all_day
    assert (
        b.subject,
        b.location,
        b.organizer,
        b.required,
        b.optional,
        b.categories,
    ) == ("",) * 6
    assert b.busy is None and b.body == () and not b.all_day


def test_v2_collapsed_arrays():
    one = {
        "start": "2026-11-05T09:00:00",
        "end": "2026-11-05T10:00:00",
        "subject": "Solo",
        "body": "only line",
    }
    (event,) = parse_days(day_json(one))[DAY]
    assert event.subject == "Solo" and event.body == ("only line",)


def test_v3_order_within_a_day():
    events = [
        {
            "start": "2026-11-05T15:00:00",
            "end": "2026-11-05T16:00:00",
            "subject": "late",
        },
        {
            "start": "2026-11-05T09:00:00",
            "end": "2026-11-05T10:00:00",
            "subject": "early",
        },
        {
            "start": "2026-11-05T00:00:00",
            "end": "2026-11-06T00:00:00",
            "subject": "holiday",
            "all_day": True,
        },
    ]
    assert [e.subject for e in parse_days(day_json(events))[DAY]] == [
        "holiday",
        "early",
        "late",
    ]


def test_v4_busy_status():
    assert BUSY == {
        0: "Free",
        1: "Tentative",
        2: "Busy",
        3: "Out of office",
        4: "Working elsewhere",
    }
    for code, name in BUSY.items():
        assert f"Status: {name}" in ev("09:00", "10:00", busy=code).detail(DAY)
    assert "Status" not in ev("09:00", "10:00", busy=9).detail(DAY)


def test_v5_summary_line():
    assert (
        ev("09:00", "10:00", "Weekly sync", location="Room A").summary(DAY)
        == "09:00-10:00  Weekly sync  [Room A]"
    )
    assert ev("13:30", "14:30", "Interview").summary(DAY) == "13:30-14:30  Interview"
    holiday = ev("2026-11-05T00:00", "2026-11-06T00:00", "Holiday", all_day=True)
    assert holiday.summary(DAY) == "All day      Holiday"


def test_v6_detail_block():
    full = ev(
        "13:30",
        "14:30",
        "Interview",
        location="Room B",
        organizer="Host",
        required="Member A",
        optional="Guest",
        categories="Customer",
        busy=2,
        body=("Agenda", "1. Requirements"),
    )
    assert full.detail(DAY) == (
        "13:30-14:30  Interview\n"
        "Location: Room B\n"
        "Organizer: Host\n"
        "Required attendees: Member A\n"
        "Optional attendees: Guest\n"
        "Categories: Customer\n"
        "Status: Busy\n"
        "\n"
        "Agenda\n"
        "1. Requirements"
    )
    assert (
        ev("09:00", "10:00", "Bare", organizer="Host").detail(DAY)
        == "09:00-10:00  Bare\nOrganizer: Host"
    )


def test_v7_day_heading():
    fetched = datetime(2026, 11, 5, 10, 32)
    monday = date(2026, 11, 2)
    names = [
        day_heading(monday + timedelta(days=i), fetched).split()[1] for i in range(7)
    ]
    assert names == ["(Mon)", "(Tue)", "(Wed)", "(Thu)", "(Fri)", "(Sat)", "(Sun)"]
    assert day_heading(DAY, fetched) == "2026-11-05 (Thu)  fetched 2026-11-05 10:32"


def test_v8_time_span_on_each_day():
    d4, d5, d6 = date(2026, 11, 4), DAY, date(2026, 11, 6)
    cases = [
        (ev("09:00", "10:00"), [(d5, "09:00-10:00")]),
        (ev("2026-11-05T23:00", "2026-11-06T00:00"), [(d5, "23:00-24:00")]),
        (
            ev("2026-11-05T23:00", "2026-11-06T01:00"),
            [(d5, "23:00-11-06 01:00"), (d6, "11-05 23:00-01:00")],
        ),
        (
            ev("2026-11-04T09:00", "2026-11-06T17:00"),
            [
                (d4, "09:00-11-06 17:00"),
                (d5, "11-04 09:00-11-06 17:00"),
                (d6, "11-04 09:00-17:00"),
            ],
        ),
        (
            ev("2026-11-04T00:00", "2026-11-07T00:00", all_day=True),
            [(d4, "All day"), (d5, "All day"), (d6, "All day")],
        ),
    ]
    for event, expected in cases:
        assert [(d, event.span_on(d)) for d, _ in expected] == expected
    # the span is also what summary and detail show
    late = ev("2026-11-05T23:00", "2026-11-06T01:00", "Late")
    assert late.summary(d6) == "11-05 23:00-01:00  Late"
    assert late.detail(d6) == "11-05 23:00-01:00  Late"


SINGLE_LINE_FIELDS = (
    "subject",
    "location",
    "organizer",
    "required",
    "optional",
    "categories",
)


def parse_one(**fields):
    times = {"start": "2026-11-05T10:00:00", "end": "2026-11-05T11:00:00"}
    return parse_days(day_json([{**times, **fields}]))[DAY][0]


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("Weekly\nsync", "Weekly sync"),
        ("Weekly\r\nsync", "Weekly sync"),
        ("Weekly\rsync", "Weekly sync"),
        ("Weekly\x0bsync", "Weekly sync"),
        ("Weekly\u2028sync", "Weekly sync"),
        ("Weekly\n\n\r\nsync", "Weekly sync"),
        ("Weekly  \n  sync", "Weekly sync"),
        ("Weekly\u3000\n\u3000sync", "Weekly sync"),
        ("Weekly\n \nsync", "Weekly sync"),
        ("\nWeekly sync\n", "Weekly sync"),
        ("Weekly  sync", "Weekly  sync"),
        ("週次\u3000定例", "週次\u3000定例"),
        ("\n \n", ""),
    ],
)
def test_v9_line_breaks_in_single_line_fields(raw, expected):
    e = parse_one(**dict.fromkeys(SINGLE_LINE_FIELDS, raw))
    assert [getattr(e, f) for f in SINGLE_LINE_FIELDS] == [expected] * 6


def test_v9_the_summary_stays_on_one_line_and_the_body_is_unchanged():
    e = parse_one(
        subject="Weekly sync\n(round 3)",
        location="Room A",
        body=["First  line", "Second\u3000line"],
    )
    assert e.summary(DAY) == "10:00-11:00  Weekly sync (round 3)  [Room A]"
    assert e.body == ("First  line", "Second\u3000line")
