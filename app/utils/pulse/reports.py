"""Shared logic for published client reports.

Here rather than in a router because both sides need it and they must agree:
the internal page labels a report when an account manager is choosing a period,
and the client-facing portal labels the same report when they open it. A
heading that differed between the two would read as two different reports.

Deliberately importable by app.routers.pulse_portal, which has no other
dependency on the internal module — the client-facing router is kept free of
staff-side code so that nothing staff-only can leak into it by accident.
"""
from __future__ import annotations

from datetime import date, timedelta


def parse_day(value) -> date | None:
    """A date from whatever PostgREST returned, or None."""
    if isinstance(value, date):
        return value
    text = str(value or "").strip()[:10]
    if not text:
        return None
    try:
        return date.fromisoformat(text)
    except ValueError:
        return None


def period_label(start: date, end: date) -> str:
    """"September 2026" for a whole calendar month, otherwise the dates.

    Named months are what an account manager and a client actually say to each
    other about a reporting cycle; a date pair is only used where the period is
    not one.
    """
    if (start.day == 1 and start.year == end.year and start.month == end.month
            and (end + timedelta(days=1)).day == 1):
        return start.strftime("%B %Y")
    if start.year == end.year:
        # Padded then stripped: "%-d" is glibc-only and "%#d" Windows-only.
        return (f"{start.strftime('%d %b').lstrip('0')}"
                f" – {end.strftime('%d %b').lstrip('0')} {end.year}")
    return (f"{start.strftime('%d %b %Y').lstrip('0')}"
            f" – {end.strftime('%d %b %Y').lstrip('0')}")


def report_heading(report: dict) -> str:
    """What a report is called. An explicit title wins; otherwise the period."""
    title = (report.get("title") or "").strip()
    if title:
        return title
    start = parse_day(report.get("period_start"))
    end = parse_day(report.get("period_end"))
    if start is None or end is None:
        return "Report"
    return period_label(start, end)
