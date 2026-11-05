"""4.3 The event at a time (spec 2.4)."""

from datetime import datetime

from conftest import DAY, ev

from calendula.events import pick


def at(hhmm: str, day=DAY) -> datetime:
    return datetime.fromisoformat(f"{day.isoformat()}T{hhmm}")


def test_p1_running_event():
    e = ev("09:00", "10:00")
    assert pick([e], at("09:30")) == (e, True)


def test_p2_start_is_inclusive_end_is_exclusive():
    e = ev("09:00", "10:00")
    assert pick([e], at("09:00")) == (e, True)
    assert pick([e], at("10:00")) == (e, False)


def test_p3_latest_start_wins_when_several_run():
    long, short = ev("09:00", "12:00", "long"), ev("10:00", "11:00", "short")
    assert pick([long, short], at("10:30")) == (short, True)


def test_p4_equal_starts_take_the_first():
    first, second = ev("10:00", "11:00", "first"), ev("10:00", "11:00", "second")
    assert pick([first, second], at("10:30")) == (first, True)


def test_p5_nearest_start_when_none_runs():
    morning, afternoon = ev("09:00", "10:00", "am"), ev("15:00", "16:00", "pm")
    assert pick([morning, afternoon], at("13:30")) == (afternoon, False)
    assert pick([morning, afternoon], at("10:30")) == (morning, False)


def test_p6_tie_takes_the_earlier():
    a, b = ev("10:00", "11:00", "a"), ev("14:00", "15:00", "b")
    assert pick([a, b], at("12:00")) == (a, False)


def test_p7_an_event_past_midnight_runs_on_both_days():
    late = ev("2026-11-05T23:00", "2026-11-06T01:00", "late")
    next_day = DAY.replace(day=6)
    assert pick([late], at("23:30")) == (late, True)
    assert pick([late], at("00:30", next_day)) == (late, True)
    assert pick([late], at("01:00", next_day)) == (late, False)


def test_p8_no_timed_events():
    holiday = ev("2026-11-05T00:00", "2026-11-06T00:00", "holiday", all_day=True)
    assert pick([holiday], at("09:00")) == (None, False)
    assert pick([], at("09:00")) == (None, False)
