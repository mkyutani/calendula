"""4.4 Cache (spec 2.5). TTL 1 hour; the clock starts at 2026-11-05 09:00."""

import threading
import time
from datetime import datetime, timedelta

import pytest
from conftest import DAY, FakeOutlook

from calendula.cache import Calendar, ttl_from_env
from calendula.outlook import OutlookError

HOUR = timedelta(hours=1)


@pytest.fixture
def outlook():
    return FakeOutlook()


@pytest.fixture
def cal(outlook, clock):
    return Calendar(outlook, ttl=HOUR, now=clock)


def test_k1_first_request_fetches(cal, outlook, clock):
    (day,) = cal.days(DAY)
    assert outlook.calls == [(DAY, 1)]
    assert day.fetched == clock.now and day.events[0].subject == "call 1"


def test_k2_reused_within_the_ttl(cal, outlook, clock):
    first = cal.days(DAY)[0].fetched
    clock.now += timedelta(minutes=59)
    (day,) = cal.days(DAY)
    assert len(outlook.calls) == 1 and day.fetched == first


def test_k3_fetched_again_at_exactly_the_ttl(cal, outlook, clock):
    cal.days(DAY)
    clock.now += HOUR
    (day,) = cal.days(DAY)
    assert (
        len(outlook.calls) == 2
        and day.events[0].subject == "call 2"
        and day.fetched == clock.now
    )


def test_k4_refresh_always_fetches(cal, outlook):
    cal.days(DAY)
    cal.days(DAY, refresh=True)
    assert len(outlook.calls) == 2


def test_k5_a_day_fetched_after_it_ended_is_kept(cal, outlook, clock):
    yesterday = DAY - timedelta(days=1)
    cal.days(yesterday)
    clock.now += timedelta(hours=5)
    cal.days(yesterday)
    assert outlook.calls == [(yesterday, 1)]


def test_k6_a_day_fetched_before_it_ended_still_expires(cal, outlook, clock):
    cal.days(DAY)
    clock.now = datetime(
        2026, 11, 6, 8, 0
    )  # next morning: DAY is past, but was fetched while it was today
    cal.days(DAY)
    assert len(outlook.calls) == 2


def test_k7_only_the_missing_tail_is_fetched(cal, outlook):
    cal.days(DAY, 2)
    cal.days(DAY, 4)
    assert outlook.calls == [(DAY, 2), (DAY + timedelta(days=2), 2)]


def test_k8_one_fetch_covers_stale_days_around_fresh_ones(outlook, clock):
    cal = Calendar(outlook, ttl=HOUR, now=clock)
    cal.days(DAY, 1)  # day 1 at 09:00
    clock.now += timedelta(minutes=30)
    cal.days(DAY + timedelta(days=3), 1)  # day 4 at 09:30
    clock.now += timedelta(minutes=20)
    cal.days(DAY + timedelta(days=1), 2)  # days 2-3 at 09:50
    clock.now += timedelta(minutes=45)  # 10:35: days 1 and 4 stale, days 2-3 fresh
    cal.days(DAY, 4)
    assert outlook.calls[-1] == (DAY, 4)
    assert len(outlook.calls) == 4


def test_k9_a_day_missing_from_the_result_has_no_events(clock):
    cal = Calendar(lambda start, days: {}, ttl=HOUR, now=clock)
    (day,) = cal.days(DAY)
    assert day.events == []


def test_k10_a_failed_fetch_changes_nothing(outlook, clock):
    cal = Calendar(outlook, ttl=HOUR, now=clock)
    before = cal.days(DAY)[0]
    outlook.fail = OutlookError("ERROR: broken")
    with pytest.raises(OutlookError):
        cal.days(DAY, refresh=True)
    outlook.fail = None
    (after,) = cal.days(DAY)
    assert after == before and len(outlook.calls) == 1


def test_k11_concurrent_requests_fetch_once(clock):
    calls = []

    def slow(start, days):
        calls.append(start)
        time.sleep(0.2)
        return {start: []}

    cal = Calendar(slow, ttl=HOUR, now=clock)
    threads = [threading.Thread(target=cal.days, args=(DAY,)) for _ in range(2)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert len(calls) == 1


def test_k12_clear_forgets_everything(cal, outlook):
    cal.days(DAY)
    cal.clear()
    cal.days(DAY)
    assert len(outlook.calls) == 2


@pytest.mark.parametrize(
    ("value", "minutes"),
    [(None, 60), ("5", 5), ("0", 60), ("-3", 60), ("abc", 60), ("", 60)],
)
def test_k13_ttl_setting(value, minutes):
    assert ttl_from_env(value) == timedelta(minutes=minutes)
