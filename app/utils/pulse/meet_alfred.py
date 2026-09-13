"""
Meet Alfred connector — LinkedIn campaigns.

STAGE MAPPING (read this before trusting a LinkedIn number)
-----------------------------------------------------------
LinkedIn outreach has no "open" event, so the two channels are aligned on what
each stage *means* rather than on what it is called in the source tool:

    sent           ← connection request or message sent
    opened         ← connection request ACCEPTED
    replied        ← replied to a message
    positive_reply ← reply classified positive (tag/label in Meet Alfred)
    meeting_booked ← meeting tracked against the prospect

`opened` = "accepted" is the load-bearing decision.  Both stages answer the same
funnel question — the prospect let the touch through — which is what makes a
combined email+LinkedIn funnel meaningful rather than a stack of two different
rulers.  It does mean a client's blended "open rate" mixes email opens with
LinkedIn acceptances, so the per-channel split in the UI is not a nicety: it is
how anyone checks what a blended number is made of.

API ACCESS
----------
Meet Alfred does not publish a stable public REST API the way Smartlead does,
and access varies by plan.  Rather than hard-code an endpoint shape we cannot
verify, this module:

  * reads its base URL and the three endpoint paths from env vars, so pointing
    it at the real API is configuration rather than a code change;
  * maps fields through candidate-name lists, tolerating the usual naming drift;
  * exposes `events_from_activities()` separately from the HTTP layer, so the
    CSV import path (app/routers/outbound_pulse.py) feeds the exact same
    normalization — a Meet Alfred report export lands in the same funnel as a
    live API sync, with no second code path to keep correct.

Until API access is confirmed, the CSV import is the supported route and
`is_configured()` returns False, which keeps the scheduler from logging a
failed LinkedIn sync every hour.
"""
from __future__ import annotations

import csv
import io
import logging
import os

import requests as http_req

from app.utils.pulse.normalize import (
    CHANNEL_LINKEDIN,
    EVENT_MEETING_BOOKED,
    EVENT_OPENED,
    EVENT_POSITIVE_REPLY,
    EVENT_REPLIED,
    EVENT_SENT,
    SOURCE_MEET_ALFRED,
    classify_reply,
    lead_key,
    make_event,
    parse_ts,
)

logger = logging.getLogger(__name__)

API_BASE = os.environ.get("MEET_ALFRED_API_BASE", "https://api.meetalfred.com/v1").rstrip("/")
CAMPAIGNS_PATH  = os.environ.get("MEET_ALFRED_CAMPAIGNS_PATH", "/campaigns")
ACTIVITIES_PATH = os.environ.get("MEET_ALFRED_ACTIVITIES_PATH", "/campaigns/{id}/activities")
PAGE_LIMIT = 100
REQUEST_TIMEOUT = 30


class MeetAlfredError(RuntimeError):
    pass


def api_key() -> str:
    return os.environ.get("MEET_ALFRED_API_KEY", "").strip()


def is_configured() -> bool:
    return bool(api_key())


# ── HTTP ──────────────────────────────────────────────────────────────────────

def _request(path: str, params: dict | None = None) -> object:
    key = api_key()
    if not key:
        raise MeetAlfredError("MEET_ALFRED_API_KEY is not set.")
    try:
        r = http_req.get(
            f"{API_BASE}{path}",
            params=params or {},
            headers={"Authorization": f"Bearer {key}", "Accept": "application/json"},
            timeout=REQUEST_TIMEOUT,
        )
    except Exception as exc:
        raise MeetAlfredError(f"request to {path} failed: {exc}") from exc
    if r.status_code in (401, 403):
        raise MeetAlfredError(f"Meet Alfred rejected the API key ({r.status_code}).")
    if r.status_code == 404:
        raise MeetAlfredError(
            f"{path} returned 404 — check MEET_ALFRED_API_BASE and the endpoint "
            "path env vars against your account's API documentation."
        )
    if r.status_code >= 400:
        raise MeetAlfredError(f"{path} returned HTTP {r.status_code}: {r.text[:200]}")
    try:
        return r.json()
    except ValueError as exc:
        raise MeetAlfredError(f"{path} returned a non-JSON body.") from exc


def _first(row: dict, *names, default=None):
    for name in names:
        if name in row:
            value = row[name]
            if value not in (None, "", []):
                return value
    return default


def _rows(payload: object) -> list[dict]:
    if isinstance(payload, list):
        return [r for r in payload if isinstance(r, dict)]
    if isinstance(payload, dict):
        for key in ("data", "results", "items", "activities", "campaigns"):
            value = payload.get(key)
            if isinstance(value, list):
                return [r for r in value if isinstance(r, dict)]
    return []


def fetch_campaigns() -> list[dict]:
    out = []
    for row in _rows(_request(CAMPAIGNS_PATH)):
        external_id = _first(row, "id", "campaign_id", "_id")
        if external_id is None:
            continue
        out.append({
            "external_id": str(external_id),
            "name":        str(_first(row, "name", "title", "campaign_name", default="") or ""),
            "status":      str(_first(row, "status", "state", default="unknown") or "unknown"),
            "raw":         row,
        })
    return out


def fetch_activities(external_campaign_id: str) -> list[dict]:
    out: list[dict] = []
    page = 1
    path = ACTIVITIES_PATH.replace("{id}", str(external_campaign_id))
    while True:
        rows = _rows(_request(path, {"page": page, "limit": PAGE_LIMIT}))
        out.extend(rows)
        if len(rows) < PAGE_LIMIT:
            return out
        page += 1
        if page > 2000:
            logger.warning(
                "Meet Alfred: campaign %s exceeded the pagination ceiling.",
                external_campaign_id,
            )
            return out


# ── Activities → normalized events ────────────────────────────────────────────

# Column aliases accepted by the CSV import, so a Meet Alfred report export maps
# without the team having to rename headers first.
_CSV_ALIASES: dict[str, tuple[str, ...]] = {
    "profile":       ("profile url", "profile_url", "linkedin url", "linkedin_url",
                      "profile", "prospect url"),
    "email":         ("email", "email address", "email_address"),
    "name":          ("name", "full name", "full_name", "prospect", "first name"),
    "sent":          ("sent", "sent at", "sent_at", "invite sent", "invited at",
                      "invitation sent", "connection request sent", "message sent"),
    "accepted":      ("accepted", "accepted at", "accepted_at", "connected at",
                      "connection accepted", "connected_at"),
    "replied":       ("replied", "replied at", "replied_at", "reply at", "response at"),
    "meeting":       ("meeting", "meeting booked", "meeting_booked", "meeting at",
                      "booked at", "meeting_date"),
    "category":      ("category", "tag", "label", "status", "outcome", "sentiment"),
}


def events_from_activities(
    rows: list[dict],
    *,
    agency_id: str,
    campaign_id: str,
    client_id: str | None,
) -> list[dict]:
    """Map Meet Alfred activity rows (API or CSV) onto normalized funnel events.

    One row is one prospect, carrying the timestamps they reached.  See the
    stage mapping at the top of this module — in particular `accepted → opened`.
    """
    events: list[dict] = []

    def add(event_type: str, when, lead: str, raw: dict) -> None:
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
            source_tool=SOURCE_MEET_ALFRED,
            sequence_ref=str(_first(raw, "step", "sequence_step", default="") or ""),
            raw={"source": SOURCE_MEET_ALFRED, "row": raw},
        ))

    for row in rows:
        lead = lead_key(
            _first(row, *_CSV_ALIASES["profile"], "profileUrl", "publicIdentifier"),
            _first(row, *_CSV_ALIASES["email"]),
            _first(row, "prospect_id", "lead_id", "id"),
        )
        if not lead:
            continue

        sent_at     = _first(row, *_CSV_ALIASES["sent"], "sentAt", "invitedAt", "created_at")
        accepted_at = _first(row, *_CSV_ALIASES["accepted"], "acceptedAt", "connectedAt")
        replied_at  = _first(row, *_CSV_ALIASES["replied"], "repliedAt", "responseAt")
        meeting_at  = _first(row, *_CSV_ALIASES["meeting"], "meetingAt", "bookedAt")
        category    = _first(row, *_CSV_ALIASES["category"])

        add(EVENT_SENT, sent_at, lead, row)
        add(EVENT_OPENED, accepted_at, lead, row)

        if replied_at:
            add(EVENT_REPLIED, replied_at, lead, row)
            stage = classify_reply(category)
            if stage in (EVENT_POSITIVE_REPLY, EVENT_MEETING_BOOKED):
                add(EVENT_POSITIVE_REPLY, replied_at, lead, row)

        if meeting_at:
            # A tracked meeting implies the two stages above it, which may not
            # have their own timestamps in the export.  Backfill them at the
            # meeting's time so the funnel never narrows then widens.
            if not replied_at:
                add(EVENT_REPLIED, meeting_at, lead, row)
            add(EVENT_POSITIVE_REPLY, meeting_at, lead, row)
            add(EVENT_MEETING_BOOKED, meeting_at, lead, row)
        elif classify_reply(category) == EVENT_MEETING_BOOKED and replied_at:
            add(EVENT_MEETING_BOOKED, replied_at, lead, row)

    return events


def parse_csv(raw: bytes | str) -> list[dict]:
    """Parse a Meet Alfred report export into activity rows.

    Headers are lowercased and trimmed so the alias table matches regardless of
    the export's capitalisation.  Returns rows as-is otherwise — mapping to
    funnel stages is `events_from_activities`' job, so the API and CSV paths
    share one normalization and cannot drift apart.
    """
    if isinstance(raw, bytes):
        # Meet Alfred exports are UTF-8, occasionally with a BOM.
        text = raw.decode("utf-8-sig", errors="replace")
    else:
        text = raw
    reader = csv.DictReader(io.StringIO(text))
    rows: list[dict] = []
    for row in reader:
        rows.append({
            (key or "").strip().lower(): (value or "").strip()
            for key, value in row.items()
            if key
        })
    return rows


def sync_campaign(
    *,
    external_campaign_id: str,
    agency_id: str,
    campaign_id: str,
    client_id: str | None,
) -> list[dict]:
    rows = fetch_activities(external_campaign_id)
    return events_from_activities(
        rows, agency_id=agency_id, campaign_id=campaign_id, client_id=client_id,
    )


CHANNEL = CHANNEL_LINKEDIN
SOURCE = SOURCE_MEET_ALFRED
