"""The MCP server: tools that read the Outlook calendar."""

from __future__ import annotations

import logging
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from datetime import date, datetime
from typing import Annotated

from mcp.server import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations
from pydantic import Field

from . import __version__, outlook
from .cache import DEFAULT_TTL, MAX_DAYS, Calendar
from .events import day_heading, parse_date, parse_time, pick

logger = logging.getLogger(__name__)

INSTRUCTIONS = f"""\
Reads the user's Outlook (classic) calendar on Windows, from WSL. Read-only.

- Events for today and later are reused for up to {int(DEFAULT_TTL.total_seconds() // 60)} minutes \
after they were fetched; past days are fetched once per session. Every answer says when it was fetched.
  Pass refresh=true when the user says the calendar has just changed, or when an exact, current answer matters.
- Dates are YYYY-MM-DD in the user's local time. Turn words like "tomorrow" or "next Monday" into dates first.
- If fetching fails, report the error. Never fill in events from memory or from older answers.
- If the error says Outlook's COM connection is broken, ask the user for permission before calling \
restart_outlook (it force-quits Outlook and unsaved drafts are lost), then retry once. Do not restart repeatedly.
- If the error says Outlook is not running, ask the user for permission before calling start_outlook \
(it opens Outlook on their desktop), then retry once.
"""

DateArg = Annotated[
    str | None,
    Field(description="Day as YYYY-MM-DD (or MM-DD for this year). Default: today."),
]
RefreshArg = Annotated[
    bool, Field(description="Fetch from Outlook even if a recent copy is cached.")
]
READ_ONLY = ToolAnnotations(readOnlyHint=True, idempotentHint=True, openWorldHint=False)


@contextmanager
def tool_errors() -> Iterator[None]:
    """Turn every failure inside a tool into a tool error that carries a message (B-ERR-1)."""
    try:
        yield
    except (ValueError, outlook.OutlookError) as e:
        raise ToolError(str(e)) from None
    except Exception as e:
        # The SDK would tell the client only "Error executing tool <name>". The client is the
        # user's own, so it gets the text too; the traceback goes to the log, as the SDK's would.
        logger.exception("Unexpected failure in a tool")
        raise ToolError(f"{type(e).__name__}: {e}") from None


def build_server(
    calendar: Calendar | None = None,
    restart: Callable[[], str] = outlook.restart,
    start: Callable[[], str] = outlook.start,
) -> MCPServer:
    cal = calendar or Calendar(outlook.fetch)
    server = MCPServer("calendula", version=__version__, instructions=INSTRUCTIONS)

    def today() -> date:
        return datetime.now().date()

    @server.tool(annotations=READ_ONLY)
    def list_events(
        date: DateArg = None,
        days: Annotated[
            int, Field(ge=1, le=MAX_DAYS, description="Number of days from `date`.")
        ] = 1,
        refresh: RefreshArg = False,
    ) -> str:
        """List events (time, subject, location) day by day. Use get_event for attendees and the body."""
        with tool_errors():
            fetched = cal.days(parse_date(date, today()), days, refresh)
            blocks = []
            for day in fetched:
                lines = [day_heading(day.date, day.fetched)]
                lines += [e.summary(day.date) for e in day.events] or ["(no events)"]
                blocks.append("\n".join(lines))
            return "\n\n".join(blocks)

    @server.tool(annotations=READ_ONLY)
    def get_event(
        time: Annotated[str, Field(description="Time as HH:MM (24-hour).")],
        date: DateArg = None,
        refresh: RefreshArg = False,
    ) -> str:
        """Details of the event at a time: location, organizer, attendees, status, and the body.

        If no event is running at that time, returns the timed event that starts nearest to it (and says so).
        """
        with tool_errors():
            at = parse_time(time)
            (day,) = cal.days(parse_date(date, today()), 1, refresh)
            event, running = pick(day.events, datetime.combine(day.date, at))
            head = day_heading(day.date, day.fetched)
            if event is None:
                return f"{head}\nNo timed events on this day."
            note = (
                ""
                if running
                else f"No event is running at {at:%H:%M}; this is the one starting nearest to it.\n"
            )
            return f"{head}\n{note}{event.detail(day.date)}"

    @server.tool(
        annotations=ToolAnnotations(
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        )
    )
    def restart_outlook() -> str:
        """Force-quit Outlook (classic) and start it again, to recover a broken COM connection.

        Unsaved drafts are lost. Call it only after the user has agreed, and only after a fetch failed with a COM error.
        """
        with tool_errors():
            result = restart()
            cal.clear()
            return result

    @server.tool(
        annotations=ToolAnnotations(
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        )
    )
    def start_outlook() -> str:
        """Start Outlook (classic) if it is not running. It never closes Outlook.

        It opens a window on the user's desktop. Call it only after the user has agreed, and only after a fetch failed because Outlook is not running.
        """
        with tool_errors():
            return start()

    return server
