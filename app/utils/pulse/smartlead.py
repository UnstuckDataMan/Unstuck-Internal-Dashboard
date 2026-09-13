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
    EVENT_MEETING_BOOKED,
    EVENT_OPENED,
    EVENT_POSITIVE_REPLY,
    EVENT_REPLIED,
    EVENT_SENT,
    SOURCE_SMARTLEAD,
    classify_reply,
    lead_key,
    make_event,
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
      • reply_time  → `replied`, plus `positive_reply` / `meeting_booked` when
                      the lead's category says so
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
        raw = {"source": SOURCE_SMARTLEAD, "row": row}

        add(EVENT_SENT, _first(row, "sent_time", "sent_at", "email_sent_time"),
            lead, step, raw)
        add(EVENT_OPENED, _first(row, "open_time", "opened_at", "email_open_time"),
            lead, step, raw)

        reply_time = _first(row, "reply_time", "replied_at", "email_reply_time")
        if reply_time:
            add(EVENT_REPLIED, reply_time, lead, step, raw)
            stage = classify_reply(_first(row, "lead_category", "category",
                                          "lead_category_name", "reply_category"))
            if stage == EVENT_MEETING_BOOKED:
                # A booked meeting implies a positive reply — record both so the
                # funnel never shows more meetings than positive replies.
                add(EVENT_POSITIVE_REPLY, reply_time, lead, step, raw)
                add(EVENT_MEETING_BOOKED, reply_time, lead, step, raw)
            elif stage == EVENT_POSITIVE_REPLY:
                add(EVENT_POSITIVE_REPLY, reply_time, lead, step, raw)

    return events


def sync_campaign(
    *,
    external_campaign_id: str,
    agency_id: str,
    campaign_id: str,
    client_id: str | None,
) -> list[dict]:
    """Fetch and normalize one campaign's events. Raises SmartleadError."""
    rows = fetch_statistics(external_campaign_id)
    return events_from_statistics(
        rows, agency_id=agency_id, campaign_id=campaign_id, client_id=client_id,
    )


CHANNEL = CHANNEL_EMAIL
SOURCE = SOURCE_SMARTLEAD
