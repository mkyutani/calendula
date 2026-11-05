"""In-memory cache of fetched days. Nothing is written to disk."""

from __future__ import annotations

import os
import threading
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime, timedelta

from .events import Event

MAX_DAYS = 31
TTL_MINUTES = 60


def ttl_from_env(value: str | None) -> timedelta:
    """CALENDULA_TTL_MINUTES as a timedelta; anything but a positive integer means the default."""
    try:
        minutes = int(value) if value is not None else TTL_MINUTES
    except ValueError:
        minutes = TTL_MINUTES
    return timedelta(minutes=minutes if minutes > 0 else TTL_MINUTES)


DEFAULT_TTL = ttl_from_env(os.environ.get("CALENDULA_TTL_MINUTES"))

Fetcher = Callable[[date, int], dict[date, list[Event]]]


@dataclass(frozen=True)
class Day:
    date: date
    fetched: datetime
    events: list[Event]


class Calendar:
    """Days fetched from Outlook, kept for the life of the server process.

    - A day that was over when it was fetched (yesterday or before) is kept and never fetched again.
    - Any other day (today, the future, or a past day fetched before it ended) is fetched again
      once it is older than `ttl`.
    - `refresh=True` always fetches.
    Missing or stale days in a request are fetched in one call covering them.
    """

    def __init__(
        self,
        fetch: Fetcher,
        ttl: timedelta = DEFAULT_TTL,
        now: Callable[[], datetime] = datetime.now,
    ):
        self._fetch = fetch
        self._ttl = ttl
        self._now = now
        self._days: dict[date, Day] = {}
        self._lock = threading.Lock()

    def _stale(self, d: date, now: datetime) -> bool:
        day = self._days.get(d)
        if day is None:
            return True
        if day.fetched.date() > d:
            return False
        return now - day.fetched >= self._ttl

    def days(self, start: date, count: int = 1, refresh: bool = False) -> list[Day]:
        if not 1 <= count <= MAX_DAYS:
            raise ValueError(f"days must be between 1 and {MAX_DAYS}")
        wanted = [start + timedelta(days=i) for i in range(count)]
        with self._lock:
            now = self._now()
            need = [d for d in wanted if refresh or self._stale(d, now)]
            if need:
                first, last = need[0], need[-1]
                span = (last - first).days + 1
                fetched = self._fetch(first, span)
                at = self._now()
                for i in range(span):
                    d = first + timedelta(days=i)
                    self._days[d] = Day(d, at, fetched.get(d, []))
            return [self._days[d] for d in wanted]

    def clear(self) -> None:
        with self._lock:
            self._days.clear()
