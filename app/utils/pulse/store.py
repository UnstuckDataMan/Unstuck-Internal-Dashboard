"""
Supabase access layer for Outbound Pulse.

Every read and write in this module is scoped by agency_id — that is the one
piece of future-proofing the build spec asks for, and it is only worth anything
if nothing bypasses it.  Callers get the agency id from `current_agency_id()`
rather than passing one around, so there is a single place to change when a
second agency row ever exists.

Uses the shared REST helpers in app.utils.supabase (same as every other tool
here) rather than the Supabase Python client — no new dependency.
"""
from __future__ import annotations

import logging
import os
import threading
from datetime import date, datetime, timezone

import requests as http_req

from app.utils.supabase import (
    SUPABASE_URL,
    sb_headers as _sb_headers,
    sb_configured as _sb_configured,
    pg_in_list as _pg_in_list,
)
from app.utils.pulse.normalize import FUNNEL_STAGES, empty_funnel

logger = logging.getLogger(__name__)

# PostgREST caps a single response; page through anything that can exceed it.
PAGE_SIZE = 1000
# Supabase rejects very large single inserts; the connectors can produce tens of
# thousands of events in a first sync, so writes go up in chunks.
INSERT_CHUNK = 500

AGENCY_NAME = os.environ.get("PULSE_AGENCY_NAME", "Unstuck").strip() or "Unstuck"


class PulseNotReady(RuntimeError):
    """Raised when a Supabase read fails.

    `kind` says why, so the UI can give the right advice instead of one generic
    hint for everything:
      * "missing" — the table or view does not exist (migration not run)
      * "timeout" — Postgres cancelled the query (Supabase's anon role stops
        statements after a few seconds)
      * "error"   — anything else
    """

    def __init__(self, message: str, kind: str = "error"):
        super().__init__(message)
        self.kind = kind


# Postgres / PostgREST codes that mean the relation isn't there.
_MISSING_CODES = {"42P01", "PGRST205", "PGRST200"}
# 57014 = query_canceled, which is how a statement timeout surfaces.
_TIMEOUT_CODES = {"57014"}


def _describe_postgrest_error(table: str, response) -> PulseNotReady:
    """Turn a failed PostgREST response into a message naming the real cause."""
    code, detail = "", ""
    try:
        body = response.json()
        if isinstance(body, dict):
            code = str(body.get("code") or "")
            detail = str(body.get("message") or body.get("details") or "")
    except ValueError:
        pass
    if not detail:
        # A gateway error page, or JSON with no message: the raw body is still
        # more useful than an empty explanation.
        detail = (getattr(response, "text", "") or "").strip()[:200]

    lowered = detail.lower()
    if code in _TIMEOUT_CODES or "statement timeout" in lowered:
        kind = "timeout"
    elif code in _MISSING_CODES or response.status_code == 404 or "does not exist" in lowered:
        kind = "missing"
    else:
        kind = "error"

    parts = [f"Query on {table} failed (HTTP {response.status_code}"]
    if code:
        parts.append(f", {code}")
    parts.append(")")
    if detail:
        parts.append(f": {detail}")
    return PulseNotReady("".join(parts), kind=kind)


# ── Agency resolution ─────────────────────────────────────────────────────────
# Cached for the process lifetime: there is exactly one agency row and it never
# changes, so re-fetching it on every request would add a round trip to every
# page load for a value that is effectively a constant.

_agency_id: str | None = None
_agency_lock = threading.Lock()


def current_agency_id() -> str:
    """The agency id every query is scoped by.

    Resolution order: PULSE_AGENCY_ID env var (an explicit pin, useful if the
    row is ever recreated), then a lookup by name, then a create.  The create
    path exists so a fresh Supabase project works without hand-seeding.
    """
    global _agency_id
    if _agency_id:
        return _agency_id

    pinned = os.environ.get("PULSE_AGENCY_ID", "").strip()
    if pinned:
        _agency_id = pinned
        return _agency_id

    with _agency_lock:
        if _agency_id:
            return _agency_id
        if not _sb_configured():
            raise PulseNotReady("Supabase is not configured.")

        try:
            r = http_req.get(
                f"{SUPABASE_URL}/rest/v1/agencies",
                headers=_sb_headers(),
                params={"select": "id", "name": f"eq.{AGENCY_NAME}", "limit": 1},
                timeout=10,
            )
            r.raise_for_status()
            rows = r.json()
        except Exception as exc:
            raise PulseNotReady(
                f"Could not read the agencies table ({exc}). "
                "Has migrations/outbound_pulse_schema.sql been run?"
            ) from exc

        if rows:
            _agency_id = str(rows[0]["id"])
            return _agency_id

        try:
            r = http_req.post(
                f"{SUPABASE_URL}/rest/v1/agencies",
                headers=_sb_headers("return=representation"),
                json={"name": AGENCY_NAME},
                timeout=10,
            )
            r.raise_for_status()
            _agency_id = str(r.json()[0]["id"])
        except Exception as exc:
            raise PulseNotReady(f"Could not create the agency row: {exc}") from exc
        return _agency_id


def reset_agency_cache() -> None:
    """Drop the cached agency id (used by the test suite)."""
    global _agency_id
    _agency_id = None


def _scoped(params: dict) -> dict:
    """Add the agency filter to a PostgREST query. Never build one without it."""
    return {**params, "agency_id": f"eq.{current_agency_id()}"}


# ── Clients ───────────────────────────────────────────────────────────────────

def list_clients(active_only: bool = False) -> list[dict]:
    """Clients belonging to this agency.

    Rows created by the other tools (DNC, Client Profiles) predate the
    agency_id column and are backfilled by the migration; `or` also accepts
    NULL so a client created between the migration and the next backfill still
    shows up rather than vanishing from reporting.
    """
    scope = f"(agency_id.eq.{current_agency_id()},agency_id.is.null)"
    # Same select-fallback as app/routers/profiles.py: the branding columns come
    # from client_profiles_schema.sql, so ask for them but degrade to id,name
    # rather than 500ing on a project where that migration hasn't run.
    rows: list[dict] = []
    last_error: PulseNotReady | None = None
    for select in ("id,name,color,emoji,active", "id,name"):
        try:
            rows = _get("clients", {"select": select, "order": "name.asc", "or": scope})
            break
        except PulseNotReady as exc:
            last_error = exc
            continue
    else:
        # Re-raise the real failure rather than a generic one: it carries the
        # Postgres error and its classification, which decide what the UI tells
        # the user to do. A blanket "could not read clients" hides whether the
        # table is missing or the query simply timed out.
        raise last_error or PulseNotReady("Could not read the clients table.")
    if active_only:
        # Filtered in Python, not PostgREST: `active` is one of the optional
        # columns added by client_profiles_schema.sql, and an `active=eq.true`
        # filter would 400 on a project that hasn't run that migration.
        rows = [c for c in rows if c.get("active") is not False]
    return rows


def backfill_client_agency() -> int:
    """Stamp this agency onto any client row that has no agency_id yet.

    Called after a manual sync so clients created by another tool cannot drift
    out of Pulse's scope between migration runs.  Returns rows patched.
    """
    orphans = _get("clients", {"select": "id", "agency_id": "is.null"})
    if not orphans:
        return 0
    try:
        r = http_req.patch(
            f"{SUPABASE_URL}/rest/v1/clients",
            headers=_sb_headers("return=minimal"),
            params={"agency_id": "is.null"},
            json={"agency_id": current_agency_id()},
            timeout=15,
        )
        r.raise_for_status()
    except Exception as exc:
        logger.warning("Pulse: client agency backfill failed: %s", exc)
        return 0
    return len(orphans)


# ── Campaigns ─────────────────────────────────────────────────────────────────

def upsert_campaign(
    *,
    source_tool: str,
    external_campaign_id: str,
    channel: str,
    name: str,
    status: str,
    client_id: str | None = None,
    raw: dict | None = None,
) -> dict | None:
    """Insert or update one external campaign, returning the stored row.

    client_id is only written when supplied: the connectors cannot always tell
    which of our clients a campaign belongs to, and a sync must never blank a
    mapping a human made in the UI.
    """
    payload = {
        "agency_id":            current_agency_id(),
        "source_tool":          source_tool,
        "external_campaign_id": str(external_campaign_id),
        "channel":              channel,
        "name":                 name or "",
        "status":               status or "unknown",
        "last_synced_at":       datetime.now(timezone.utc).isoformat(),
        "updated_at":           datetime.now(timezone.utc).isoformat(),
        "raw_payload":          raw or {},
    }
    if client_id:
        payload["client_id"] = client_id
    # Single-row path, used by the CSV import. The scheduled sync uses
    # upsert_campaigns() below — one request per campaign does not scale to an
    # account with a thousand of them.

    try:
        r = http_req.post(
            f"{SUPABASE_URL}/rest/v1/pulse_campaigns",
            headers=_sb_headers("resolution=merge-duplicates,return=representation"),
            params={"on_conflict": "agency_id,source_tool,external_campaign_id"},
            json=payload,
            timeout=15,
        )
        r.raise_for_status()
        rows = r.json()
        return rows[0] if rows else None
    except Exception as exc:
        logger.warning(
            "Pulse: campaign upsert failed (%s/%s): %s",
            source_tool, external_campaign_id, exc,
        )
        return None


def upsert_campaigns(entries: list[dict], *, source_tool: str, channel: str) -> int:
    """Bulk-upsert the campaign list from a connector. Returns rows written.

    Deliberately omits BOTH client_id and last_synced_at from the payload:

      * client_id — under resolution=merge-duplicates PostgREST only updates
        the columns actually present, so leaving it out preserves whatever a
        human mapped in the UI. Sending it would mean reading every existing
        mapping first and racing anyone editing one mid-sync.
      * last_synced_at — that means "we synced this campaign's events", not
        "we saw it in the campaign list". Stamping it here would make every
        campaign look freshly synced and break the rolling backfill order in
        app/utils/pulse/sync.py. mark_campaigns_synced() sets it instead.
    """
    if not entries:
        return 0
    now = datetime.now(timezone.utc).isoformat()
    agency = current_agency_id()
    rows = [{
        "agency_id":            agency,
        "source_tool":          source_tool,
        "external_campaign_id": str(e["external_id"]),
        "channel":              channel,
        "name":                 e.get("name") or "",
        "status":               e.get("status") or "unknown",
        "updated_at":           now,
        "raw_payload":          e.get("raw") or {},
    } for e in entries]

    written = 0
    for start in range(0, len(rows), INSERT_CHUNK):
        chunk = rows[start:start + INSERT_CHUNK]
        try:
            r = http_req.post(
                f"{SUPABASE_URL}/rest/v1/pulse_campaigns",
                headers=_sb_headers("resolution=merge-duplicates,return=minimal"),
                params={"on_conflict": "agency_id,source_tool,external_campaign_id"},
                json=chunk,
                timeout=45,
            )
            r.raise_for_status()
            written += len(chunk)
        except Exception as exc:
            logger.warning("Pulse: bulk campaign upsert chunk failed: %s", exc)
    return written


def mark_campaigns_synced(campaign_ids: list[str]) -> None:
    """Stamp last_synced_at on campaigns whose events were just synced.

    One request for the whole batch rather than one per campaign. This is what
    the rolling backfill orders by, so it must only ever be set after a
    successful event sync.
    """
    if not campaign_ids:
        return
    now = datetime.now(timezone.utc).isoformat()
    for start in range(0, len(campaign_ids), INSERT_CHUNK):
        chunk = campaign_ids[start:start + INSERT_CHUNK]
        try:
            http_req.patch(
                f"{SUPABASE_URL}/rest/v1/pulse_campaigns",
                headers=_sb_headers("return=minimal"),
                params=_scoped({"id": f"in.{_pg_in_list(chunk)}"}),
                json={"last_synced_at": now},
                timeout=30,
            )
        except Exception as exc:
            logger.warning("Pulse: could not stamp last_synced_at: %s", exc)


def campaigns_for_sync(source_tool: str) -> list[dict]:
    """Stored campaign rows for a connector, least-recently-synced first.

    nullsfirst puts never-synced campaigns at the head, so a first run starts
    backfilling immediately instead of re-walking whatever happens to sort low.
    """
    return _get("pulse_campaigns", _scoped({
        "select":      "id,client_id,external_campaign_id,name,status,last_synced_at",
        "source_tool": f"eq.{source_tool}",
        "order":       "last_synced_at.asc.nullsfirst",
    }))


def list_campaigns(
    *,
    client_id: str = "",
    channel: str = "",
    source_tool: str = "",
    status: str = "",
    synced: str = "",
    q: str = "",
) -> list[dict]:
    """Campaign rows, optionally filtered.

    Filtering is done here rather than in the browser because a real account
    has ~1000 campaigns: shipping them all as HTML and hiding rows with CSS
    would mean a multi-megabyte partial on every keystroke.

    `synced` accepts "never" or "synced" — whether the campaign's events have
    ever been pulled, which is not the same as its status in the source tool.

    `client_id` accepts a real id, or the literal "none" for campaigns with no
    client mapped. "none" rather than an empty string because empty already
    means "no filter", and the difference between those two matters here.
    """
    params = {
        "select": ("id,client_id,channel,source_tool,external_campaign_id,"
                   "name,status,last_synced_at,created_at"),
        "order":  "name.asc",
    }
    if client_id == "none":
        params["client_id"] = "is.null"
    elif client_id:
        params["client_id"] = f"eq.{client_id}"
    if channel:
        params["channel"] = f"eq.{channel}"
    if source_tool:
        params["source_tool"] = f"eq.{source_tool}"
    if status:
        params["status"] = f"eq.{status}"
    if synced == "never":
        params["last_synced_at"] = "is.null"
    elif synced == "synced":
        params["last_synced_at"] = "not.is.null"
    if q and q.strip():
        params["name"] = f"ilike.{_ilike_value(q.strip())}"
    return _get("pulse_campaigns", _scoped(params))


def _ilike_value(term: str) -> str:
    """Build a PostgREST ilike literal that matches `term` anywhere in the name.

    Two layers of escaping, for two different parsers:
      * `%` and `_` are SQL LIKE wildcards. A campaign name search for "50_off"
        would otherwise match "5000ff", so they are backslash-escaped.
      * the whole value is double-quoted because PostgREST parses commas and
        parentheses out of an unquoted filter value, and campaign names here
        contain both (e.g. "(CURATED) - 11-480 - B2C Health").
    """
    cleaned = (term
               .replace("\\", "\\\\")
               .replace("%", "\\%")
               .replace("_", "\\_")
               .replace('"', '\\"'))
    return f'"*{cleaned}*"'


def set_campaigns_client(campaign_ids: list[str], client_id: str | None) -> int:
    """Map many campaigns to one client (or unmap them). Returns rows attempted.

    Chunked because the ids go into a PostgREST `in.(...)` filter on the URL,
    and a few hundred quoted UUIDs would exceed the server's URL length limit
    and fail the whole batch.
    """
    ids = [str(i) for i in campaign_ids if i]
    if not ids:
        return 0
    now = datetime.now(timezone.utc).isoformat()
    chunk_size = 50
    for start in range(0, len(ids), chunk_size):
        chunk = ids[start:start + chunk_size]
        try:
            r = http_req.patch(
                f"{SUPABASE_URL}/rest/v1/pulse_campaigns",
                headers=_sb_headers("return=minimal"),
                params=_scoped({"id": f"in.{_pg_in_list(chunk)}"}),
                json={"client_id": client_id, "updated_at": now},
                timeout=30,
            )
            r.raise_for_status()
        except Exception as exc:
            logger.warning("Pulse: bulk campaign mapping failed: %s", exc)
            continue
        # Events carry a denormalised client_id so the portal can filter without
        # a join — re-point them too, or the funnel keeps crediting the old
        # client for everything synced before the remap.
        try:
            http_req.patch(
                f"{SUPABASE_URL}/rest/v1/pulse_campaign_events",
                headers=_sb_headers("return=minimal"),
                params=_scoped({"campaign_id": f"in.{_pg_in_list(chunk)}"}),
                json={"client_id": client_id},
                timeout=60,
            )
        except Exception as exc:
            logger.warning("Pulse: bulk event re-point failed: %s", exc)
    return len(ids)


def campaign_filter_options() -> dict:
    """Distinct values (with counts) for the mapping table's filter dropdowns.

    Selects one narrow column across all rows rather than reusing the filtered
    query, so the options always describe the whole set — a status filter that
    only ever offers the status already selected would be a dead end.
    """
    rows = _get("pulse_campaigns", _scoped({
        "select": "status,source_tool,channel,last_synced_at,client_id",
    }))
    statuses: dict[str, int] = {}
    sources: dict[str, int] = {}
    channels: dict[str, int] = {}
    # Per-client counts drive the Client filter, which is how a campaign mapped
    # to the wrong client gets found again — without it, correcting a mistake
    # means scrolling a thousand rows.
    per_client: dict[str, int] = {}
    unmapped = 0
    never = 0
    for row in rows:
        cid = row.get("client_id")
        if cid:
            per_client[str(cid)] = per_client.get(str(cid), 0) + 1
        else:
            unmapped += 1
        statuses[str(row.get("status") or "unknown")] = \
            statuses.get(str(row.get("status") or "unknown"), 0) + 1
        sources[str(row.get("source_tool") or "")] = \
            sources.get(str(row.get("source_tool") or ""), 0) + 1
        channels[str(row.get("channel") or "")] = \
            channels.get(str(row.get("channel") or ""), 0) + 1
        if not row.get("last_synced_at"):
            never += 1
    return {
        # Busiest first: with six statuses the useful one should not be hunted for.
        "statuses": sorted(statuses.items(), key=lambda kv: -kv[1]),
        "sources":    sorted(sources.items()),
        "channels":   sorted(channels.items()),
        "never":      never,
        "synced":     len(rows) - never,
        "total":      len(rows),
        "per_client": per_client,
        "unmapped":   unmapped,
    }


def set_campaign_client(campaign_id: str, client_id: str | None) -> bool:
    """Map (or unmap) a synced campaign to one of our clients."""
    try:
        r = http_req.patch(
            f"{SUPABASE_URL}/rest/v1/pulse_campaigns",
            headers=_sb_headers("return=minimal"),
            params=_scoped({"id": f"eq.{campaign_id}"}),
            json={"client_id": client_id,
                  "updated_at": datetime.now(timezone.utc).isoformat()},
            timeout=15,
        )
        r.raise_for_status()
    except Exception as exc:
        logger.warning("Pulse: could not set client on campaign %s: %s", campaign_id, exc)
        return False

    # Events carry a denormalised client_id so the portal can filter without a
    # join; re-point the existing ones so a remap doesn't strand history.
    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/pulse_campaign_events",
            headers=_sb_headers("return=minimal"),
            params=_scoped({"campaign_id": f"eq.{campaign_id}"}),
            json={"client_id": client_id},
            timeout=30,
        )
    except Exception as exc:
        logger.warning("Pulse: event re-point failed for campaign %s: %s", campaign_id, exc)
    return True


# ── Events ────────────────────────────────────────────────────────────────────

def insert_events(events: list[dict]) -> int:
    """Append normalized events, ignoring ones already stored.

    Relies on the unique index over (agency_id, dedupe_key): re-syncing the same
    window is a no-op rather than a duplicate, which is what makes the scheduled
    job safe to run as often as we like.

    Returns the number of rows actually inserted (PostgREST returns only the
    new rows under return=representation, so duplicates are excluded).
    """
    if not events:
        return 0
    inserted = 0
    for start in range(0, len(events), INSERT_CHUNK):
        chunk = events[start:start + INSERT_CHUNK]
        try:
            r = http_req.post(
                f"{SUPABASE_URL}/rest/v1/pulse_campaign_events",
                headers=_sb_headers("resolution=ignore-duplicates,return=representation"),
                params={"on_conflict": "agency_id,dedupe_key"},
                json=chunk,
                timeout=45,
            )
            r.raise_for_status()
            body = r.json()
            inserted += len(body) if isinstance(body, list) else 0
        except Exception as exc:
            # One bad chunk must not lose the rest of the run — the sync log
            # records the shortfall and the next run re-attempts these events.
            logger.warning("Pulse: event insert chunk failed: %s", exc)
    return inserted


def upsert_outcomes(outcomes: list[dict]) -> int:
    """Write replying leads' current outcomes. Returns rows written.

    An upsert, not an insert: each sync overwrites a lead's stage with its
    current Smartlead category, which is what moves a lead re-marked from
    Interested to Meeting Request, or drops one re-marked Not Interested.

    Returns the number of rows the database accepted, so the sync can tell a
    silent shortfall from success rather than logging a clean run over lost data.
    """
    if not outcomes:
        return 0
    written = 0
    for start in range(0, len(outcomes), INSERT_CHUNK):
        chunk = outcomes[start:start + INSERT_CHUNK]
        try:
            r = http_req.post(
                f"{SUPABASE_URL}/rest/v1/pulse_lead_outcomes",
                headers=_sb_headers("resolution=merge-duplicates,return=minimal"),
                params={"on_conflict": "agency_id,campaign_id,lead_key"},
                json=chunk,
                timeout=45,
            )
            if r.status_code >= 400:
                logger.warning("Pulse: outcome upsert chunk failed: %s",
                               _describe_postgrest_error("pulse_lead_outcomes", r))
                continue
            written += len(chunk)
        except Exception as exc:
            logger.warning("Pulse: outcome upsert chunk failed: %s", exc)
    return written


# ── Funnel aggregation ────────────────────────────────────────────────────────

def funnel(
    *,
    client_id:   str = "",
    channel:     str = "",
    campaign_id: str = "",
    date_from:   date | None = None,
    date_to:     date | None = None,
) -> dict[str, int]:
    """Total funnel counts for a filter set, read from pulse_funnel_daily."""
    totals = empty_funnel()
    for row in _funnel_rows(client_id=client_id, channel=channel,
                            campaign_id=campaign_id,
                            date_from=date_from, date_to=date_to,
                            select="event_type,events"):
        stage = row.get("event_type")
        if stage in totals:
            totals[stage] += int(row.get("events") or 0)
    return totals


def funnel_by_client(
    *,
    channel:   str = "",
    date_from: date | None = None,
    date_to:   date | None = None,
) -> dict[str, dict[str, int]]:
    """Per-client funnel totals in one query — the internal dashboard's grid.

    One request for every client rather than one per client: the daily view is
    already aggregated, so the whole agency's reporting period is a few hundred
    rows at most.
    """
    out: dict[str, dict[str, int]] = {}
    for row in _funnel_rows(channel=channel, date_from=date_from, date_to=date_to,
                            select="client_id,event_type,events"):
        cid = str(row.get("client_id") or "")
        stage = row.get("event_type")
        if stage not in FUNNEL_STAGES:
            continue
        out.setdefault(cid, empty_funnel())[stage] += int(row.get("events") or 0)
    return out


def funnel_by_channel(
    *,
    client_id: str = "",
    date_from: date | None = None,
    date_to:   date | None = None,
) -> dict[str, dict[str, int]]:
    """Funnel split by channel, so email and LinkedIn can be compared."""
    out: dict[str, dict[str, int]] = {}
    for row in _funnel_rows(client_id=client_id, date_from=date_from, date_to=date_to,
                            select="channel,event_type,events"):
        stage = row.get("event_type")
        if stage not in FUNNEL_STAGES:
            continue
        out.setdefault(str(row.get("channel") or "unknown"), empty_funnel())[stage] += \
            int(row.get("events") or 0)
    return out


def funnel_by_campaign(
    *,
    client_id: str = "",
    channel:   str = "",
    date_from: date | None = None,
    date_to:   date | None = None,
) -> dict[str, dict[str, int]]:
    out: dict[str, dict[str, int]] = {}
    for row in _funnel_rows(client_id=client_id, channel=channel,
                            date_from=date_from, date_to=date_to,
                            select="campaign_id,event_type,events"):
        stage = row.get("event_type")
        if stage not in FUNNEL_STAGES:
            continue
        out.setdefault(str(row.get("campaign_id") or ""), empty_funnel())[stage] += \
            int(row.get("events") or 0)
    return out


def funnel_timeseries(
    *,
    client_id: str = "",
    channel:   str = "",
    date_from: date | None = None,
    date_to:   date | None = None,
) -> list[dict]:
    """Daily funnel counts, oldest first — the trend strip on the detail views."""
    by_day: dict[str, dict[str, int]] = {}
    for row in _funnel_rows(client_id=client_id, channel=channel,
                            date_from=date_from, date_to=date_to,
                            select="day,event_type,events"):
        stage = row.get("event_type")
        if stage not in FUNNEL_STAGES:
            continue
        day = str(row.get("day") or "")
        if not day:
            continue
        by_day.setdefault(day, empty_funnel())[stage] += int(row.get("events") or 0)
    return [{"day": day, **counts} for day, counts in sorted(by_day.items())]


def _funnel_rows(
    *,
    select:      str,
    client_id:   str = "",
    channel:     str = "",
    campaign_id: str = "",
    date_from:   date | None = None,
    date_to:     date | None = None,
) -> list[dict]:
    params: dict = {"select": select}
    if client_id:
        params["client_id"] = f"eq.{client_id}"
    if channel:
        params["channel"] = f"eq.{channel}"
    if campaign_id:
        params["campaign_id"] = f"eq.{campaign_id}"
    if date_from:
        params["day"] = f"gte.{date_from.isoformat()}"
    if date_to:
        # PostgREST takes repeated keys as ANDed filters, which a dict cannot
        # express — send the params as a list of pairs when both bounds are set.
        pairs = [(k, v) for k, v in _scoped(params).items()]
        pairs.append(("day", f"lte.{date_to.isoformat()}"))
        return _get("pulse_funnel_daily", pairs)
    return _get("pulse_funnel_daily", _scoped(params))


# ── Sync logs ─────────────────────────────────────────────────────────────────

def start_sync_log(source_tool: str, triggered_by: str = "schedule") -> str | None:
    try:
        r = http_req.post(
            f"{SUPABASE_URL}/rest/v1/pulse_sync_logs",
            headers=_sb_headers("return=representation"),
            json={
                "agency_id":    current_agency_id(),
                "source_tool":  source_tool,
                "status":       "running",
                "triggered_by": triggered_by,
            },
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
        return str(rows[0]["id"]) if rows else None
    except Exception as exc:
        # Never let logging failure abort the sync it was meant to observe.
        logger.warning("Pulse: could not open sync log for %s: %s", source_tool, exc)
        return None


def finish_sync_log(
    log_id: str | None,
    *,
    status: str,
    campaigns_synced: int = 0,
    events_inserted: int = 0,
    duration_s: float | None = None,
    error_message: str = "",
) -> None:
    if not log_id:
        return
    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/pulse_sync_logs",
            headers=_sb_headers("return=minimal"),
            params=_scoped({"id": f"eq.{log_id}"}),
            json={
                "status":           status,
                "finished_at":      datetime.now(timezone.utc).isoformat(),
                "campaigns_synced": campaigns_synced,
                "events_inserted":  events_inserted,
                "duration_s":       round(duration_s, 2) if duration_s is not None else None,
                "error_message":    (error_message or "")[:2000] or None,
            },
            timeout=10,
        )
    except Exception as exc:
        logger.warning("Pulse: could not close sync log %s: %s", log_id, exc)


def recent_sync_logs(limit: int = 20) -> list[dict]:
    return _get("pulse_sync_logs", _scoped({
        "select": ("id,source_tool,run_at,finished_at,status,campaigns_synced,"
                   "events_inserted,duration_s,triggered_by,error_message"),
        "order":  "run_at.desc",
        "limit":  str(limit),
    }))


def latest_sync_per_tool() -> dict[str, dict]:
    """Most recent run per connector — what the sync status pill reads."""
    latest: dict[str, dict] = {}
    for row in recent_sync_logs(limit=60):
        tool = row.get("source_tool") or ""
        if tool and tool not in latest:
            latest[tool] = row
    return latest


# ── Portal access ─────────────────────────────────────────────────────────────

def create_access(client_id: str, token_hash: str, label: str,
                  created_by: str, expires_at: str | None) -> dict | None:
    try:
        r = http_req.post(
            f"{SUPABASE_URL}/rest/v1/pulse_client_access",
            headers=_sb_headers("return=representation"),
            json={
                "agency_id":  current_agency_id(),
                "client_id":  client_id,
                "token_hash": token_hash,
                "label":      label or "",
                "created_by": created_by or "",
                "expires_at": expires_at,
            },
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
        return rows[0] if rows else None
    except Exception as exc:
        logger.warning("Pulse: could not create portal access: %s", exc)
        return None


def access_by_token_hash(token_hash: str) -> dict | None:
    rows = _get("pulse_client_access", _scoped({
        "select":     "id,client_id,label,expires_at,revoked_at,view_count",
        "token_hash": f"eq.{token_hash}",
        "limit":      "1",
    }))
    return rows[0] if rows else None


def list_access(client_id: str) -> list[dict]:
    return _get("pulse_client_access", _scoped({
        "select":    "id,label,created_by,created_at,expires_at,revoked_at,last_used_at,view_count",
        "client_id": f"eq.{client_id}",
        "order":     "created_at.desc",
    }))


def revoke_access(access_id: str) -> bool:
    try:
        r = http_req.patch(
            f"{SUPABASE_URL}/rest/v1/pulse_client_access",
            headers=_sb_headers("return=minimal"),
            params=_scoped({"id": f"eq.{access_id}"}),
            json={"revoked_at": datetime.now(timezone.utc).isoformat()},
            timeout=10,
        )
        r.raise_for_status()
        return True
    except Exception as exc:
        logger.warning("Pulse: could not revoke access %s: %s", access_id, exc)
        return False


def record_visit(access: dict, client_id: str, user_agent: str) -> None:
    """Log a portal view and bump the access counter. Best-effort, never fatal —
    a client must still see their report if the analytics write fails."""
    now = datetime.now(timezone.utc).isoformat()
    try:
        http_req.post(
            f"{SUPABASE_URL}/rest/v1/pulse_portal_visits",
            headers=_sb_headers("return=minimal"),
            json={
                "agency_id":  current_agency_id(),
                "client_id":  client_id,
                "access_id":  access.get("id"),
                "viewed_at":  now,
                "user_agent": (user_agent or "")[:400],
            },
            timeout=10,
        )
    except Exception:
        pass
    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/pulse_client_access",
            headers=_sb_headers("return=minimal"),
            params=_scoped({"id": f"eq.{access.get('id')}"}),
            json={"last_used_at": now,
                  "view_count": int(access.get("view_count") or 0) + 1},
            timeout=10,
        )
    except Exception:
        pass


def visit_summary(client_id: str = "") -> list[dict]:
    """Recent portal visits — feeds the internal engagement panel."""
    params = {
        "select": "client_id,viewed_at",
        "order":  "viewed_at.desc",
        "limit":  "500",
    }
    if client_id:
        params["client_id"] = f"eq.{client_id}"
    return _get("pulse_portal_visits", _scoped(params))


# ── Paged GET ─────────────────────────────────────────────────────────────────

def _get(table: str, params) -> list[dict]:
    """GET every page of a PostgREST query.

    Accepts a dict or a list of (key, value) pairs — repeated keys are how
    PostgREST ANDs two filters on the same column (e.g. a day range), and a
    dict cannot hold them.
    """
    if not _sb_configured():
        raise PulseNotReady("Supabase is not configured.")

    pairs = list(params.items()) if isinstance(params, dict) else list(params)
    has_limit = any(k == "limit" for k, _ in pairs)

    out: list[dict] = []
    offset = 0
    while True:
        page = list(pairs)
        if not has_limit:
            page.append(("limit", str(PAGE_SIZE)))
            page.append(("offset", str(offset)))
        try:
            r = http_req.get(
                f"{SUPABASE_URL}/rest/v1/{table}",
                headers=_sb_headers(),
                params=page,
                timeout=30,
            )
        except Exception as exc:
            raise PulseNotReady(f"Query on {table} failed: {exc}") from exc

        if r.status_code >= 400:
            # raise_for_status() would discard the response body, which is where
            # PostgREST puts the actual Postgres error. Without it a statement
            # timeout and a missing table both read as a bare "500 Server Error".
            raise _describe_postgrest_error(table, r)
        try:
            rows = r.json()
        except ValueError as exc:
            raise PulseNotReady(f"Query on {table} returned a non-JSON body.") from exc

        if not isinstance(rows, list):
            return out
        out.extend(rows)
        if has_limit or len(rows) < PAGE_SIZE:
            return out
        offset += PAGE_SIZE
