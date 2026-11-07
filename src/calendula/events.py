"""Calendar events: parsing what the PowerShell script writes, and formatting them for the model."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta

# Outlook OlBusyStatus
BUSY = {
    0: "Free",
    1: "Tentative",
    2: "Busy",
    3: "Out of office",
    4: "Working elsewhere",
}
DOW = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")


@dataclass(frozen=True)
class Event:
    start: datetime
    end: datetime
    subject: str
    all_day: bool = False
    location: str = ""
    organizer: str = ""
    required: str = ""
    optional: str = ""
    categories: str = ""
    busy: int | None = None
    body: tuple[str, ...] = field(default_factory=tuple)

    @classmethod
    def from_json(cls, d: dict) -> Event:
        def s(key: str) -> str:
            v = d.get(key)
            if v is None:
                return ""
            # One line, whatever line breaks the field has (B-EV-5).
            lines = (line.strip() for line in str(v).splitlines())
            return " ".join(line for line in lines if line)

        body = d.get("body") or []
        if isinstance(body, str):  # PowerShell may collapse a one-item array
            body = [body]
        busy = d.get("busy")
        return cls(
            start=datetime.fromisoformat(d["start"]),
            end=datetime.fromisoformat(d["end"]),
            subject=s("subject"),
            all_day=bool(d.get("all_day")),
            location=s("location"),
            organizer=s("organizer"),
            required=s("required"),
            optional=s("optional"),
            categories=s("categories"),
            busy=int(busy) if busy is not None else None,
            body=tuple(str(x) for x in body if str(x).strip()),
        )

    def span_on(self, day: date) -> str:
        """The time span as shown on `day` (spec 3.1, "Time span on day D")."""
        if self.all_day:
            return "All day"
        start = (
            f"{self.start:%H:%M}"
            if self.start.date() == day
            else f"{self.start:%m-%d %H:%M}"
        )
        if self.end.date() == day:
            end = f"{self.end:%H:%M}"
        elif self.end == datetime.combine(day + timedelta(days=1), time()):
            end = "24:00"
        else:
            end = f"{self.end:%m-%d %H:%M}"
        return f"{start}-{end}"

    def summary(self, day: date) -> str:
        line = f"{self.span_on(day):<11}  {self.subject}"
        if self.location:
            line += f"  [{self.location}]"
        return line

    def detail(self, day: date) -> str:
        lines = [f"{self.span_on(day)}  {self.subject}"]
        for label, value in (
            ("Location", self.location),
            ("Organizer", self.organizer),
            ("Required attendees", self.required),
            ("Optional attendees", self.optional),
            ("Categories", self.categories),
            ("Status", BUSY.get(self.busy, "") if self.busy is not None else ""),
        ):
            if value:
                lines.append(f"{label}: {value}")
        if self.body:
            lines.append("")
            lines.extend(self.body)
        return "\n".join(lines)


def parse_days(data: dict) -> dict[date, list[Event]]:
    """{date: [events sorted by start]} from the JSON the PowerShell script writes."""
    out: dict[date, list[Event]] = {}
    for day in data.get("days") or []:
        d = date.fromisoformat(day["date"])
        events = day.get("events") or []
        if isinstance(events, dict):  # PowerShell may collapse a one-item array
            events = [events]
        out[d] = sorted(
            (Event.from_json(e) for e in events), key=lambda e: (not e.all_day, e.start)
        )
    return out


def pick(events: list[Event], at: datetime) -> tuple[Event | None, bool]:
    """(event, running) at the date and time `at` (spec 2.4, B-PICK-1 to B-PICK-5).

    Running: start <= at < end; of several, the latest start (ties: the first). Otherwise the event
    whose start is nearest to `at` (ties: the earlier), with running False. All-day events are ignored.
    """
    timed = [e for e in events if not e.all_day]
    running = [e for e in timed if e.start <= at < e.end]
    if running:
        return max(running, key=lambda e: e.start), True
    if timed:
        return min(timed, key=lambda e: abs(e.start - at)), False
    return None, False


def day_heading(d: date, fetched: datetime) -> str:
    return f"{d.isoformat()} ({DOW[d.weekday()]})  fetched {fetched:%Y-%m-%d %H:%M}"


DATE_RE = re.compile(r"^(?:(\d{4})[-/])?(\d{1,2})[-/](\d{1,2})$")
TIME_RE = re.compile(r"^(\d{1,2}):(\d{2})$")


def parse_date(text: str | None, today: date) -> date:
    """YYYY-MM-DD, YYYY/MM/DD, or MM-DD / MM/DD (this year). Empty means today."""
    if not text or not text.strip():
        return today
    m = DATE_RE.match(text.strip())
    if not m:
        raise ValueError(f"invalid date: {text!r} (use YYYY-MM-DD)")
    try:
        return date(int(m[1] or today.year), int(m[2]), int(m[3]))
    except ValueError:
        raise ValueError(f"invalid date: {text!r}") from None


def parse_time(text: str) -> time:
    m = TIME_RE.match(text.strip())
    if not m or int(m[1]) > 23 or int(m[2]) > 59:
        raise ValueError(f"invalid time: {text!r} (use HH:MM)")
    return time(int(m[1]), int(m[2]))
