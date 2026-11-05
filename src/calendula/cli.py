"""Command line: run the MCP server, or print events for a quick check."""

from __future__ import annotations

import argparse
import sys
from datetime import datetime

from . import __version__, outlook
from .cache import MAX_DAYS
from .events import day_heading, parse_date, parse_time, pick


def _serve(_args: argparse.Namespace) -> int:
    from .server import build_server

    build_server().run("stdio")
    return 0


def _list(args: argparse.Namespace) -> int:
    start = parse_date(args.date, datetime.now().date())
    if not 1 <= args.days <= MAX_DAYS:
        raise ValueError(f"days must be between 1 and {MAX_DAYS}")
    days = outlook.fetch(start, args.days)
    fetched = datetime.now()
    blocks = []
    for d in sorted(days):
        lines = [day_heading(d, fetched)] + (
            [e.summary(d) for e in days[d]] or ["(no events)"]
        )
        blocks.append("\n".join(lines))
    print("\n\n".join(blocks))
    return 0


def _show(args: argparse.Namespace) -> int:
    day = parse_date(args.date, datetime.now().date())
    at = parse_time(args.time)
    events = outlook.fetch(day, 1).get(day, [])
    event, running = pick(events, datetime.combine(day, at))
    print(day_heading(day, datetime.now()))
    if event is None:
        print("No timed events on this day.")
        return 0
    if not running:
        print(
            f"No event is running at {at:%H:%M}; this is the one starting nearest to it."
        )
    print(event.detail(day))
    return 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        prog="calendula",
        description="MCP server that reads your Outlook (classic) calendar from WSL. "
        "With no command, runs the server on stdio.",
    )
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = p.add_subparsers(dest="command")
    sub.add_parser(
        "serve", help="run the MCP server on stdio (the default)"
    ).set_defaults(func=_serve)
    pl = sub.add_parser("list", help="print events, fetched from Outlook now")
    pl.add_argument("date", nargs="?", help="first day, YYYY-MM-DD (default: today)")
    pl.add_argument(
        "days", nargs="?", type=int, default=1, help="number of days (default: 1)"
    )
    pl.set_defaults(func=_list)
    ps = sub.add_parser(
        "show", help="print the event at a time, fetched from Outlook now"
    )
    ps.add_argument("time", help="HH:MM")
    ps.add_argument("--date", help="YYYY-MM-DD (default: today)")
    ps.set_defaults(func=_show)

    args = p.parse_args(argv)
    func = getattr(args, "func", _serve)
    try:
        return func(args)
    except (ValueError, outlook.OutlookError) as e:
        print(f"calendula: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
