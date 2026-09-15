"""
The normalized event contract shared by every Outbound Pulse connector.

Both connectors produce the same five event types, so the funnel view never
branches on which tool a data point came from:

    sent → opened → replied → positive_reply → meeting_booked

Two rules make the funnel a plain COUNT(*) per event_type, which matters
because PostgREST cannot express COUNT(DISTINCT):

  1. `sent` counts sends.  A 4-step sequence to one lead is 4 events, because
     "emails sent" is a volume figure and that is what a client expects to see.
  2. Every qualifying stage counts *leads*.  Their dedupe_key deliberately
     omits the timestamp, so a lead who opens six times still produces one
     `opened` row, and a re-sync that reports a newer open time does not add a
     second.  The unique index on (agency_id, dedupe_key) enforces it in the
     database rather than trusting the connector to behave.

The trade-off in rule 2 is that we keep the FIRST timestamp we ever saw for a
stage, not the latest.  That is the right choice for a funnel: "when did this
lead first reply" is the question a reporting cycle asks, and it makes the
event table genuinely append-only.
"""
from __future__ import annotations

import hashlib
import re
from datetime import datetime, timezone

# Funnel order.  The dashboard renders stages in this sequence and computes
# stage-to-stage conversion against the stage above.
EVENT_SENT           = "sent"
EVENT_OPENED         = "opened"
EVENT_REPLIED        = "replied"
EVENT_POSITIVE_REPLY = "positive_reply"
EVENT_MEETING_BOOKED = "meeting_booked"

FUNNEL_STAGES: tuple[str, ...] = (
    EVENT_SENT,
    EVENT_OPENED,
    EVENT_REPLIED,
    EVENT_POSITIVE_REPLY,
    EVENT_MEETING_BOOKED,
)

# Labels follow the team's own Smartlead categories. The top stage's key stays
# `meeting_booked` because it is stored on every existing event row and in the
# rollup, but what the data actually records is a meeting REQUEST — there is no
# booked-meeting category. Showing clients "Meetings booked" would read as
# confirmed calls on their calendar, which the numbers do not support.
STAGE_LABELS: dict[str, str] = {
    EVENT_SENT:           "Sent",
    EVENT_OPENED:         "Opened",
    EVENT_REPLIED:        "Replied",
    EVENT_POSITIVE_REPLY: "Interested",
    EVENT_MEETING_BOOKED: "Meeting requested",
}

# Stages counted once per lead per campaign (see rule 2 above).
_ONCE_PER_LEAD: frozenset[str] = frozenset(FUNNEL_STAGES) - {EVENT_SENT}

CHANNEL_EMAIL    = "email"
CHANNEL_LINKEDIN = "linkedin"

SOURCE_SMARTLEAD   = "smartlead"
SOURCE_MEET_ALFRED = "meet_alfred"

CHANNEL_FOR_SOURCE: dict[str, str] = {
    SOURCE_SMARTLEAD:   CHANNEL_EMAIL,
    SOURCE_MEET_ALFRED: CHANNEL_LINKEDIN,
}


# ── Timestamp parsing ─────────────────────────────────────────────────────────
# Both APIs return ISO-8601, but inconsistently: with/without a timezone, with
# 'Z', with microseconds, sometimes as a bare date.  A connector that raises on
# an unexpected shape would drop a whole campaign's events, so parse defensively
# and let the caller decide what to do with None.

_ISO_TRAILING_Z = re.compile(r"[Zz]$")


def parse_ts(value) -> datetime | None:
    """Parse an API timestamp into an aware UTC datetime, or None.

    A naive timestamp is assumed to be UTC — both tools report in UTC, and
    guessing a local zone would silently shift events across day boundaries in
    the daily funnel view.
    """
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        dt = value
    elif isinstance(value, (int, float)):
        # Epoch seconds vs milliseconds: anything past ~2001 in seconds is
        # beyond 1e9, so a value past 1e11 can only be milliseconds.
        seconds = float(value) / 1000.0 if float(value) > 1e11 else float(value)
        try:
            dt = datetime.fromtimestamp(seconds, tz=timezone.utc)
        except (OverflowError, OSError, ValueError):
            return None
    else:
        raw = str(value).strip()
        if not raw:
            return None
        raw = _ISO_TRAILING_Z.sub("+00:00", raw)
        # Some payloads use a space separator instead of 'T'.
        if " " in raw and "T" not in raw:
            raw = raw.replace(" ", "T", 1)
        try:
            dt = datetime.fromisoformat(raw)
        except ValueError:
            return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def iso_utc(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).isoformat()


# ── Lead identity ─────────────────────────────────────────────────────────────

def lead_key(*candidates) -> str:
    """First non-empty identifier, normalized for use in a dedupe key.

    Email wins where present; LinkedIn campaigns fall back to the profile URL.
    Lowercased and trimmed so the same lead reported with different casing by
    two endpoints does not become two leads in the funnel.
    """
    for candidate in candidates:
        if candidate is None:
            continue
        value = str(candidate).strip().lower()
        if value:
            # Strip query strings and trailing slashes off LinkedIn URLs, which
            # arrive with tracking parameters attached about half the time.
            if value.startswith("http"):
                value = value.split("?", 1)[0].rstrip("/")
            return value
    return ""


# ── Reply classification ──────────────────────────────────────────────────────
# Smartlead exposes a per-lead category; Meet Alfred does not classify at all.
# Anything we cannot confidently call positive stays a plain `replied`, so the
# stages above it under-report rather than flatter the numbers.
#
# The team's live Smartlead categories, and how they are used:
#   Interested          — asks for more info, or shows low-level interest
#   Information Request — the same intent, so the same stage
#   Meeting Request     — higher interest, or asks for a meeting: the top stage
#   Not Interested / Do Not Contact / Out Of Office / Wrong Person — not positive

_POSITIVE_CATEGORIES = {
    "interested",
    "information request",
    "positive",
    "positive reply",
    "warm",
    "hot lead",
}

# "meeting request" belongs here, not in the positive set: it is the team's
# highest-intent category. It previously sat with Interested, which meant no
# live category ever reached the top stage and it read as zero for every client.
_MEETING_CATEGORIES = {
    "meeting request",
    "meeting requested",
    "meeting booked",
    "meeting completed",
    "meeting scheduled",
    "booked",
    "demo booked",
    "call booked",
}

_NEGATIVE_CATEGORIES = {
    "not interested",
    "negative",
    "do not contact",
    "unsubscribed",
    "out of office",
    "wrong person",
    "bounced",
    "spam",
}


def classify_reply(category) -> str | None:
    """Map a source tool's reply category onto a funnel stage.

    Returns EVENT_MEETING_BOOKED, EVENT_POSITIVE_REPLY, or None for "replied,
    but not positively classified".  Unknown categories return None on purpose:
    a category we have never seen is not evidence of a positive reply.
    """
    if not category:
        return None
    key = str(category).strip().lower()
    if not key:
        return None
    if key in _MEETING_CATEGORIES:
        return EVENT_MEETING_BOOKED
    if key in _POSITIVE_CATEGORIES:
        return EVENT_POSITIVE_REPLY
    if key in _NEGATIVE_CATEGORIES:
        return None
    # Substring fallback for categories the team renames in-tool
    # ("Interested - pricing", "Meeting Request ✅"). Negative is checked FIRST:
    # "Not interested in a meeting request" contains both a negative and a
    # meeting phrase, and when in doubt the funnel must under-report.
    if any(term in key for term in _NEGATIVE_CATEGORIES):
        return None
    if any(term in key for term in _MEETING_CATEGORIES):
        return EVENT_MEETING_BOOKED
    if any(term in key for term in _POSITIVE_CATEGORIES):
        return EVENT_POSITIVE_REPLY
    return None


# ── Event construction ────────────────────────────────────────────────────────

def make_event(
    *,
    agency_id:   str,
    campaign_id: str,
    client_id:   str | None,
    event_type:  str,
    occurred_at: datetime,
    lead:        str,
    source_tool: str,
    sequence_ref: str = "",
    raw:         dict | None = None,
) -> dict:
    """Build one normalized `pulse_campaign_events` row.

    `sequence_ref` distinguishes repeated sends to the same lead (a step number
    or message id).  It is ignored for the once-per-lead stages, which is what
    collapses six opens into one `opened` row.
    """
    if event_type not in FUNNEL_STAGES:
        raise ValueError(f"unknown event_type: {event_type!r}")

    parts = [source_tool, str(campaign_id), event_type, lead]
    if event_type not in _ONCE_PER_LEAD:
        # Sends are per-occurrence: include the step reference and the timestamp
        # so a genuine second send is a second row.
        parts.extend([str(sequence_ref or ""), iso_utc(occurred_at)])

    return {
        "agency_id":   agency_id,
        "campaign_id": campaign_id,
        "client_id":   client_id,
        "event_type":  event_type,
        "occurred_at": iso_utc(occurred_at),
        "lead_key":    lead,
        "dedupe_key":  _hash_key(parts),
        "raw_payload": raw or {},
    }


def _hash_key(parts: list[str]) -> str:
    """Stable, bounded dedupe key.

    Hashed rather than stored as a raw concatenation because lead identifiers
    can be long LinkedIn URLs, and a btree unique index over unbounded text
    risks exceeding the 2704-byte index row limit.
    """
    return hashlib.sha256("\x1f".join(parts).encode("utf-8")).hexdigest()


def empty_funnel() -> dict[str, int]:
    return {stage: 0 for stage in FUNNEL_STAGES}


def opens_are_tracked(counts: dict[str, int]) -> bool:
    """Whether the Opened stage carries real information for these counts.

    The team rarely turns on open tracking, so for most campaigns `opened` is
    zero — and showing it both displays a false "0 opened" and zeroes out every
    reply rate calculated against it. A lead has to open an email before
    replying, so a genuinely tracked funnel always has at least as many opens as
    replies. Fewer opens than replies means tracking was off for some or all of
    the campaigns in view, and the stage is dropped rather than shown wrong.

    Decided per funnel, not globally, so a campaign that does track opens — or a
    LinkedIn funnel, where this stage is a connection acceptance — still shows it.
    """
    opened = counts.get(EVENT_OPENED, 0) or 0
    return opened > 0 and opened >= (counts.get(EVENT_REPLIED, 0) or 0)


def funnel_with_rates(counts: dict[str, int]) -> list[dict]:
    """Funnel stages with both step and top-of-funnel conversion rates.

    Step rate answers "of the leads that reached the previous stage, how many
    got here"; overall answers "what share of sends ended here". Both are shown
    because the first is the one to optimise and the second is the one clients
    quote back.
    """
    top = counts.get(EVENT_SENT, 0) or 0
    rows: list[dict] = []
    previous = None
    stages = [s for s in FUNNEL_STAGES if s != EVENT_OPENED or opens_are_tracked(counts)]
    for index, stage in enumerate(stages):
        value = counts.get(stage, 0) or 0
        rows.append({
            "key":       stage,
            "label":     STAGE_LABELS[stage],
            "value":     value,
            # True only for the first stage — there is nothing above it. A later
            # stage whose parent is zero also has no step rate, but it is not the
            # top of the funnel, and the UI must not label it as one.
            "is_top":    index == 0,
            "step_rate": _rate(value, previous),
            # Suppressed for the stage directly under `sent`, where the two rates
            # are the same number by definition and printing both reads as a bug.
            "overall":   _rate(value, top) if index >= 2 else None,
        })
        previous = value
    return rows


def _rate(value: int, base: int | None) -> float | None:
    if base is None or base <= 0:
        return None
    return round(value * 100.0 / base, 1)
