"""4.1 Arguments (spec 2.1)."""

from datetime import date, time

import pytest
from conftest import DAY, FakeOutlook, call_tool, text

from calendula.cache import Calendar
from calendula.events import parse_date, parse_time
from calendula.server import build_server


def test_a1_blank_date_is_today():
    for blank in (None, "", "  "):
        assert parse_date(blank, DAY) == DAY


def test_a2_full_dates():
    assert parse_date("2026-12-01", DAY) == date(2026, 12, 1)
    assert parse_date("2026/12/01", DAY) == date(2026, 12, 1)


def test_a3_dates_without_a_year_are_this_year():
    assert parse_date("12-1", DAY) == date(2026, 12, 1)
    assert parse_date("12/01", DAY) == date(2026, 12, 1)
    assert parse_date("1/5", DAY) == date(2026, 1, 5)


@pytest.mark.parametrize(
    "bad", ["tomorrow", "2026-02-30", "13/40", "20261105", "2026-1-1-1"]
)
def test_a4_invalid_dates(bad):
    with pytest.raises(ValueError, match="invalid date"):
        parse_date(bad, DAY)


@pytest.mark.parametrize(
    ("good", "expected"),
    [
        ("0:00", time(0, 0)),
        ("00:00", time(0, 0)),
        ("9:05", time(9, 5)),
        ("23:59", time(23, 59)),
    ],
)
def test_a5_valid_times(good, expected):
    assert parse_time(good) == expected


@pytest.mark.parametrize("bad", ["24:00", "23:60", "9", "9:5", "09:00:00"])
def test_a6_invalid_times(bad):
    with pytest.raises(ValueError, match="invalid time"):
        parse_time(bad)


def test_a7_days_range_in_the_calendar(clock):
    cal = Calendar(FakeOutlook(), now=clock)
    for bad in (0, 32):
        with pytest.raises(ValueError, match="between 1 and 31"):
            cal.days(DAY, bad)
    assert len(cal.days(DAY, 1)) == 1
    assert len(cal.days(DAY, 31)) == 31


def test_a8_days_range_in_the_tool_schema():
    server = build_server(Calendar(FakeOutlook()))
    for bad in (0, 32):
        r = call_tool(server, "list_events", {"date": "2026-11-05", "days": bad})
        assert r.is_error, bad
    r = call_tool(server, "list_events", {"date": "2026-11-05", "days": 31})
    assert not r.is_error and text(r).count("fetched") == 31
