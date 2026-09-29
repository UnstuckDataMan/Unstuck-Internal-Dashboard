"""
Smartlead connector — email campaigns.

Smartlead's public API authenticates with an `api_key` query parameter against
https://server.smartlead.ai/api/v1.  Two endpoints carry everything the funnel
needs:

  GET /campaigns                     — the campaign list (id, name, status)
  GET /campaigns/{id}/statistics     — one paged row per lead per sequence step,
                                       carrying sent_time / open_time /
                                       reply_time and the lead's category

The statistics endpoint is what makes this a real event stream rather than a
snapshot of counters: each row has timestamps, so re-syncing rebuilds history
instead of overwriting a total.  That is why the funnel can be filtered by date
range at all.

FIELD NAMING: Smartlead has renamed fields across API versions (sent_time vs
sent_at, lead_category vs category), and an account on a different plan can get
a slightly different payload.  Every read goes through `_first()` with a list of
candidate keys instead of a single hard-coded name — a renamed field then costs
one entry in a list, not a silently empty funnel.  This is deliberate: the
connector's shape is not something we can verify without live credentials, and
failing soft on one field beats dropping a campaign.
"""
from __future__ import annotations

import logging
import os
import time

import requests as http_req

from app.utils.pulse.normalize import (
    CHANNEL_EMAIL,
    EVENT_OPENED,
    EVENT_REPLIED,
    EVENT_SENT,
    SOURCE_SMARTLEAD,
    lead_key,
    make_event,
    make_outcome,
    parse_ts,
)

logger = logging.getLogger(__name__)

API_BASE = os.environ.get("SMARTLEAD_API_BASE", "https://server.smartlead.ai/api/v1").rstrip("/")
PAGE_LIMIT = 100
REQUEST_TIMEOUT = 30
# Smartlead rate-limits per API key; a short pause between pages keeps a large
# first sync from tripping it and losing the rest of a campaign's history.
PACING_SECONDS = max(0.0, float(os.environ.get("SMARTLEAD_PACING_MS", "300")) / 1000.0)


class SmartleadError(RuntimeError):
    pass


def api_key() -> str:
    return os.environ.get("SMARTLEAD_API_KEY", "").strip()


def is_configured() -> bool:
    return bool(api_key())


# ── HTTP ──────────────────────────────────────────────────────────────────────

def _request(path: str, params: dict | None = None) -> object:
    key = api_key()
    if not key:
        raise SmartleadError("SMARTLEAD_API_KEY is not set.")
    query = {**(params or {}), "api_key": key}
    try:
        r = http_req.get(f"{API_BASE}{path}", params=query, timeout=REQUEST_TIMEOUT)
    except Exception as exc:
        raise SmartleadError(f"request to {path} failed: {exc}") from exc
    if r.status_code == 401:
        raise SmartleadError("Smartlead rejected the API key (401).")
    if r.status_code == 429:
        raise SmartleadError("Smartlead rate limit hit (429) — try again shortly.")
    if r.status_code >= 400:
        raise SmartleadError(f"{path} returned HTTP {r.status_code}: {r.text[:200]}")
    try:
        return r.json()
    except ValueError as exc:
        raise SmartleadError(f"{path} returned a non-JSON body.") from exc


def _first(row: dict, *names, default=None):
    """First present, non-empty value among candidate field names."""
    for name in names:
        if name in row:
            value = row[name]
            if value not in (None, "", []):
                return value
    return default


# Fields kept in an event's raw_payload. Everything else in a statistics row is
# discarded, for two reasons measured against the live API: a full row is ~2.4KB
# because it embeds `email_message` and `email_subject`, which across millions of
# events is gigabytes of duplicated message bodies; and those bodies plus
# `lead_name` are prospect content that a reporting table has no reason to hold.
# The lead's identity already lives in the event's `lead_key`.
_RAW_KEEP = (
    "stats_id",
    "sequence_number",
    "seq_variant_id",
    "email_campaign_seq_id",
    "lead_category",
    "open_count",
    "click_count",
    "is_bounced",
    "is_unsubscribed",
)


def _trim_row(row: dict) -> dict:
    """Diagnostic subset of a statistics row, safe to store on every event."""
    return {k: row[k] for k in _RAW_KEEP if k in row and row[k] not in (None, "")}


def _rows(payload: object) -> list[dict]:
    """Normalize the several envelope shapes Smartlead returns to a plain list."""
    if isinstance(payload, list):
        return [r for r in payload if isinstance(r, dict)]
    if isinstance(payload, dict):
        for key in ("data", "results", "statistics", "campaigns", "items"):
            value = payload.get(key)
            if isinstance(value, list):
                return [r for r in value if isinstance(r, dict)]
    return []


# ── Campaigns ─────────────────────────────────────────────────────────────────

def fetch_campaigns() -> list[dict]:
    """All campaigns on the account, normalized to {external_id, name, status}."""
    out = []
    for row in _rows(_request("/campaigns")):
        external_id = _first(row, "id", "campaign_id")
        if external_id is None:
            continue
        out.append({
            "external_id": str(external_id),
            "name":        str(_first(row, "name", "campaign_name", default="") or ""),
            "status":      str(_first(row, "status", default="unknown") or "unknown"),
            "raw":         row,
        })
    return out


# ── Statistics → normalized events ────────────────────────────────────────────

def fetch_statistics(external_campaign_id: str) -> list[dict]:
    """Every statistics row for a campaign, following the offset pagination."""
    out: list[dict] = []
    offset = 0
    while True:
        payload = _request(
            f"/campaigns/{external_campaign_id}/statistics",
            {"offset": offset, "limit": PAGE_LIMIT},
        )
        rows = _rows(payload)
        out.extend(rows)
        if len(rows) < PAGE_LIMIT:
            return out
        offset += PAGE_LIMIT
        if PACING_SECONDS:
            time.sleep(PACING_SECONDS)
        # A malformed/looping paginator must not spin forever on a scheduled job.
        if offset > 200_000:
            logger.warning(
                "Smartlead: campaign %s exceeded the pagination ceiling at offset %d.",
                external_campaign_id, offset,
            )
            return out


def events_from_statistics(
    rows: list[dict],
    *,
    agency_id: str,
    campaign_id: str,
    client_id: str | None,
) -> list[dict]:
    """Map Smartlead statistics rows onto normalized funnel events.

    One statistics row is one lead × one sequence step, so:
      • sent_time   → a `sent` event per row (real send volume)
      • open_time   → one `opened` event per lead (the dedupe key collapses the
                      repeats across steps)
      • reply_time  → one `replied` event per lead

    The lead's category is NOT turned into events — see outcomes_from_statistics.
    """
    events: list[dict] = []

    def add(event_type: str, when, lead: str, step: str, raw: dict) -> None:
        occurred = parse_ts(when)
        if occurred is None:
            return
        events.append(make_event(
            agency_id=agency_id,
            campaign_id=campaign_id,
            client_id=client_id,
            event_type=event_type,
            occurred_at=occurred,
            lead=lead,
            source_tool=SOURCE_SMARTLEAD,
            sequence_ref=step,
            raw=raw,
        ))

    for row in rows:
        lead = lead_key(
            _first(row, "lead_email", "email", "to_email"),
            _first(row, "lead_id", "id"),
        )
        if not lead:
            # Without a lead identity the once-per-lead stages cannot dedupe,
            # so counting the row would inflate the funnel on every re-sync.
            continue

        step = str(_first(row, "sequence_number", "step", "email_sequence_number",
                          "sequence_step_id", default="") or "")
        raw = {"source": SOURCE_SMARTLEAD, "row": _trim_row(row)}

        add(EVENT_SENT, _first(row, "sent_time", "sent_at", "email_sent_time"),
            lead, step, raw)
        add(EVENT_OPENED, _first(row, "open_time", "opened_at", "email_open_time"),
            lead, step, raw)

        add(EVENT_REPLIED, _first(row, "reply_time", "replied_at", "email_reply_time"),
            lead, step, raw)

    return events


_CATEGORY_FIELDS = ("lead_category", "category", "lead_category_name", "reply_category")


def outcomes_from_statistics(
    rows: list[dict],
    *,
    agency_id: str,
    campaign_id: str,
) -> list[dict]:
    """One current-outcome row per lead who replied.

    Statistics rows are per sequence step, so a lead appears several times but
    only the step that drew the reply carries reply_time. Rows are collapsed to
    one per lead here, and that is required, not tidiness: a single upsert
    statement containing the same key twice fails in Postgres ("cannot affect
    row a second time"), which would lose the whole batch.

    The category is lead-level in Smartlead, so it is the lead's CURRENT
    category on every sync. The earliest reply fixes the reporting day.
    """
    by_lead: dict[str, dict] = {}
    for row in rows:
        lead = lead_key(
            _first(row, "lead_email", "email", "to_email"),
            _first(row, "lead_id", "id"),
        )
        if not lead:
            continue
        replied_at = parse_ts(_first(row, "reply_time", "replied_at", "email_reply_time"))
        category = _first(row, *_CATEGORY_FIELDS)

        seen = by_lead.get(lead)
        if seen is None:
            by_lead[lead] = {"replied_at": replied_at, "category": category}
            continue
        if replied_at and (seen["replied_at"] is None or replied_at < seen["replied_at"]):
            seen["replied_at"] = replied_at
        if category and not seen["category"]:
            seen["category"] = category

    return [
        make_outcome(
            agency_id=agency_id,
            campaign_id=campaign_id,
            lead=lead,
            category=data["category"],
            replied_at=data["replied_at"],
        )
        for lead, data in by_lead.items()
        if data["replied_at"] is not None
    ]


def sync_campaign(
    *,
    external_campaign_id: str,
    agency_id: str,
    campaign_id: str,
    client_id: str | None,
) -> tuple[list[dict], list[dict]]:
    """Fetch one campaign and return (events, outcomes). Raises SmartleadError.

    One fetch feeds both, so events and outcomes always describe the same
    snapshot of the campaign.
    """
    rows = fetch_statistics(external_campaign_id)
    events = events_from_statistics(
        rows, agency_id=agency_id, campaign_id=campaign_id, client_id=client_id,
    )
    outcomes = outcomes_from_statistics(
        rows, agency_id=agency_id, campaign_id=campaign_id,
    )
    return events, outcomes


CHANNEL = CHANNEL_EMAIL
SOURCE = SOURCE_SMARTLEAD
