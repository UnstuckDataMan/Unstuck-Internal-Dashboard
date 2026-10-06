"""Figures an account manager types in, for a tool we cannot reach.

NOT the same thing as SOURCE_MANUAL. That one means campaigns run through the
DNC & Merger tool, whose numbers this database already holds and reads in
place. This is the opposite: a person reading a dashboard the app has no
access to and keying the totals. Different provenance, different source_tool,
differently labelled tab.

LINKEDIN ONLY. The gap this fills is Meet Alfred, which is how LinkedIn
outreach is run and whose API access is unconfirmed. Email is covered by
Smartlead and by the DNC & Merger tool, both of which this database already
reads, so there is nothing for a person to type. Offering a channel choice
would only invite figures to be filed where something else already counts them.

GRAIN. One entry per client, channel and month. That is the grain the figures
are read at on the source dashboard, and the grain they get corrected at.

THE COUNTING CONTRACT. normalize.py guarantees that every stage below `sent`
counts each person once, enforced by a unique index on the event dedupe key.
This source cannot honour that — there is no row per person to deduplicate,
only a total somebody read off a screen. The compensation is in the labels:
each field is named at lead grain ("Leads who replied"), so the number typed
is the number the funnel means. Trusted input, not counted input.
"""
from __future__ import annotations

from datetime import date

from app.utils.pulse.normalize import (
    CHANNEL_EMAIL,
    CHANNEL_LINKEDIN,
    EVENT_BOUNCED,
    EVENT_MEETING_BOOKED,
    EVENT_OPENED,
    EVENT_REPLIED,
    EVENT_SENT,
    EVENT_UNSUBSCRIBED,
    SOURCE_HAND_ENTRY,
)

SOURCE = SOURCE_HAND_ENTRY

# The stored columns, and the funnel stage each one becomes.
#
# `opened` carries LinkedIn connection acceptances as well as email opens.
# They are the same funnel stage — the Meet Alfred connector already treats an
# acceptance as the LinkedIn equivalent of an open — and two columns feeding
# one stage would let a single entry silently count both.
STAGE_FOR_METRIC: dict[str, str] = {
    "sent":         EVENT_SENT,
    "opened":       EVENT_OPENED,
    "replied":      EVENT_REPLIED,
    "meetings":     EVENT_MEETING_BOOKED,
    "bounced":      EVENT_BOUNCED,
    "unsubscribed": EVENT_UNSUBSCRIBED,
}

METRICS: tuple[str, ...] = tuple(STAGE_FOR_METRIC)

# Per channel, because the two tools call the same stage different things and
# the label is the only thing stopping the wrong number going in the box.
#
# Every label names people rather than events, because that is what the funnel
# counts everywhere else. "Opens" would invite the raw open count, which for
# most tools counts one person several times.
METRIC_LABELS: dict[str, dict[str, str]] = {
    CHANNEL_EMAIL: {
        "sent":         "Emails sent",
        "opened":       "Leads who opened",
        "replied":      "Leads who replied",
        "meetings":     "Meeting requests / meetings booked",
        "bounced":      "Bounces",
        "unsubscribed": "Unsubscribes",
    },
    CHANNEL_LINKEDIN: {
        "sent":         "Connection requests sent",
        "opened":       "Connection requests accepted",
        "replied":      "Leads who replied",
        "meetings":     "Meeting requests / meetings booked",
        "bounced":      "Failed to deliver",
        "unsubscribed": "Opted out",
    },
}

# The one channel a person enters figures for. The stored column still accepts
# 'email' so any row keyed in before this narrowing still reads, but nothing
# writes one now.
CHANNEL = CHANNEL_LINKEDIN
CHANNELS: tuple[str, ...] = (CHANNEL_LINKEDIN,)


def parse_month(value: str) -> date | None:
    """'2026-09' or any day inside it → date(2026, 9, 1). None if unparseable.

    Any day resolves to the first, so a month input and a date picker both land
    on the one day the funnel reads.
    """
    text = (value or "").strip()
    if not text:
        return None
    if len(text) == 7:          # <input type="month"> gives YYYY-MM
        text += "-01"
    try:
        return date.fromisoformat(text[:10]).replace(day=1)
    except ValueError:
        return None


def month_bounds(month: date) -> tuple[date, date]:
    """First and last day of the month an entry covers."""
    start = month.replace(day=1)
    end = (start.replace(year=start.year + 1, month=1) if start.month == 12
           else start.replace(month=start.month + 1))
    return start, date.fromordinal(end.toordinal() - 1)


def month_label(month: date) -> str:
    return month.strftime("%B %Y")


def clean_metrics(raw: dict) -> tuple[dict[str, int], str]:
    """Parse the form's figures. Returns (metrics, error).

    Blank reads as zero. A negative is refused here as well as by the table's
    CHECK, so the person sees which field is wrong instead of a database error.
    """
    out: dict[str, int] = {}
    for metric in METRICS:
        text = str(raw.get(metric, "") or "").strip().replace(",", "")
        if not text:
            out[metric] = 0
            continue
        try:
            value = int(float(text))
        except ValueError:
            return {}, f"{metric.replace('_', ' ').capitalize()} must be a number."
        if value < 0:
            return {}, f"{metric.replace('_', ' ').capitalize()} cannot be negative."
        out[metric] = value
    return out, ""


def labels_for(channel: str = CHANNEL_LINKEDIN) -> dict[str, str]:
    return METRIC_LABELS.get(channel, METRIC_LABELS[CHANNEL_LINKEDIN])
