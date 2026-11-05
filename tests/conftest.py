"""Stand-ins shared by the tests (docs/test-design.md, section 2)."""

import asyncio
from datetime import date, datetime, timedelta

import pytest
from mcp.client import Client

from calendula.events import Event

DAY = date(2026, 11, 5)  # a Thursday


def ev(start: str, end: str, subject: str = "x", **kw) -> Event:
    """An event from 'HH:MM' times on DAY, or full ISO datetimes."""

    def dt(s: str) -> datetime:
        return datetime.fromisoformat(s if "T" in s else f"{DAY.isoformat()}T{s}")

    return Event(dt(start), dt(end), subject, **kw)


class Clock:
    """A clock the test moves by hand."""

    def __init__(self, now: datetime):
        self.now = now

    def __call__(self) -> datetime:
        return self.now


class FakeOutlook:
    """A fetch function that records its calls and returns one event per day, named after the call."""

    def __init__(self, fail: Exception | None = None):
        self.calls: list[tuple[date, int]] = []
        self.fail = fail

    def __call__(self, start: date, days: int) -> dict[date, list[Event]]:
        if self.fail:
            raise self.fail
        self.calls.append((start, days))
        out = {}
        for i in range(days):
            d = start + timedelta(days=i)
            base = datetime.combine(d, datetime.min.time())
            out[d] = [
                Event(
                    base + timedelta(hours=9),
                    base + timedelta(hours=10),
                    f"call {len(self.calls)}",
                )
            ]
        return out


def call_tool(server, name: str, args: dict | None = None):
    async def run():
        async with Client(server) as c:
            return await c.call_tool(name, args or {})

    return asyncio.run(run())


def text(result) -> str:
    return result.content[0].text


@pytest.fixture
def clock():
    return Clock(datetime(2026, 11, 5, 9, 0))
