"""
Outbound Pulse — internal (team-facing) reporting views.

One funnel per client across both channels, built from the normalized events
the connectors write.  Everything here is agency-scoped through
app.utils.pulse.store; no query in this router builds its own filter.

Follows the house pattern: a Jinja page shell, HTMX partials for the live
regions, JSON only where JavaScript needs to read a value.
"""
from __future__ import annotations

import hashlib
import secrets
from datetime import date, datetime, timedelta, timezone
from html import escape

from fastapi import APIRouter, Depends, File, Form, Query, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse

from app import auth
from app.deps import templates
from app.utils.dates import today_utc
from app.utils.pulse import meet_alfred, store, sync as pulse_sync
from app.utils.pulse.normalize import (
    CHANNEL_EMAIL,
    CHANNEL_LINKEDIN,
    EVENT_MEETING_BOOKED,
    EVENT_SENT,
    FUNNEL_STAGES,
    SOURCE_MEET_ALFRED,
    funnel_with_rates,
)
from app.utils.pulse.store import PulseNotReady

router = APIRouter()

# Portal links are long-lived by default — a client should not have a report
# link die between monthly cycles — but not permanent.
_ACCESS_TTL_DAYS = 180

# A run older than this means the hourly schedule has not fired (a sleeping
# Render instance, a crashed thread), which is the failure mode the sync-status
# panel exists to catch.
_SYNC_STALE_AFTER_MIN = 75

RANGE_PRESETS: dict[str, str] = {
    "7d":   "Last 7 days",
    "30d":  "Last 30 days",
    "mtd":  "This month",
    "90d":  "Last 90 days",
    "all":  "All time",
}


# ── Date range ────────────────────────────────────────────────────────────────

def _resolve_range(preset: str, date_from: str, date_to: str) -> dict:
    """Turn the UI's range controls into concrete bounds plus a label.

    Explicit from/to wins over the preset so a bookmarked custom range keeps
    working.  Bounds are UTC dates via the shared clock in app.utils.dates —
    the same one the campaign stats use, so the two tools never disagree about
    where a day starts.
    """
    parsed_from = _parse_date(date_from)
    parsed_to   = _parse_date(date_to)
    if parsed_from or parsed_to:
        start = parsed_from
        end   = parsed_to or today_utc()
        if start and end and start > end:
            start, end = end, start
        return {
            "preset": "custom",
            "from":   start,
            "to":     end,
            "label":  f"{start.isoformat() if start else 'start'} → {end.isoformat()}",
        }

    today = today_utc()
    key = preset if preset in RANGE_PRESETS else "30d"
    if key == "all":
        return {"preset": "all", "from": None, "to": None, "label": RANGE_PRESETS["all"]}
    if key == "mtd":
        start = today.replace(day=1)
    elif key == "7d":
        start = today - timedelta(days=6)
    elif key == "90d":
        start = today - timedelta(days=89)
    else:
        key = "30d"
        start = today - timedelta(days=29)
    return {"preset": key, "from": start, "to": today, "label": RANGE_PRESETS[key]}


def _parse_date(value: str) -> date | None:
    value = (value or "").strip()
    if not value:
        return None
    try:
        return date.fromisoformat(value)
    except ValueError:
        return None


def _range_query(rng: dict) -> str:
    """Re-encode a resolved range as a query string, for links between views."""
    if rng["preset"] == "custom":
        parts = []
        if rng["from"]:
            parts.append(f"date_from={rng['from'].isoformat()}")
        if rng["to"]:
            parts.append(f"date_to={rng['to'].isoformat()}")
        return "&".join(parts)
    return f"range={rng['preset']}"


# ── Error rendering ───────────────────────────────────────────────────────────

def _error_box(message: str) -> HTMLResponse:
    """Inline error in an HTMX target — house pattern (see profiles.py)."""
    return HTMLResponse(f'<div class="err-box">{escape(message)}</div>')


def _not_ready_box(exc: Exception) -> HTMLResponse:
    return _error_box(
        f"{exc} If this is the first run, apply "
        "migrations/outbound_pulse_schema.sql in the Supabase SQL editor."
    )


# ── Shared view assembly ──────────────────────────────────────────────────────

def _client_index() -> dict[str, dict]:
    return {str(c["id"]): c for c in store.list_clients()}


def _channel_filter(channel: str) -> str:
    return channel if channel in (CHANNEL_EMAIL, CHANNEL_LINKEDIN) else ""


def _overview_context(rng: dict, channel: str) -> dict:
    clients = store.list_clients()
    by_client = store.funnel_by_client(
        channel=channel, date_from=rng["from"], date_to=rng["to"],
    )

    rows = []
    for client in clients:
        cid = str(client["id"])
        counts = by_client.get(cid)
        if not counts:
            continue     # no outbound activity in range — not a reporting row
        rows.append({
            "client":  client,
            "counts":  counts,
            "funnel":  funnel_with_rates(counts),
            "sent":    counts.get(EVENT_SENT, 0),
            "meetings": counts.get(EVENT_MEETING_BOOKED, 0),
        })
    # Busiest clients first: the internal view is a scan for anomalies, and a
    # client with 40 000 sends matters more than one with 12.
    rows.sort(key=lambda r: r["sent"], reverse=True)

    totals = {stage: 0 for stage in FUNNEL_STAGES}
    for row in rows:
        for stage in FUNNEL_STAGES:
            totals[stage] += row["counts"].get(stage, 0)

    unmapped = [
        c for c in store.list_campaigns()
        if not c.get("client_id")
    ]

    return {
        "rows":        rows,
        "totals":      funnel_with_rates(totals),
        "total_sent":  totals[EVENT_SENT],
        "by_channel":  store.funnel_by_channel(date_from=rng["from"], date_to=rng["to"]),
        "unmapped":    unmapped,
        "range":       rng,
        "range_query": _range_query(rng),
        "channel":     channel,
    }


def _client_context(client_id: str, rng: dict, channel: str) -> dict | None:
    clients = _client_index()
    client = clients.get(str(client_id))
    if client is None:
        return None

    counts = store.funnel(
        client_id=client_id, channel=channel,
        date_from=rng["from"], date_to=rng["to"],
    )
    per_campaign = store.funnel_by_campaign(
        client_id=client_id, channel=channel,
        date_from=rng["from"], date_to=rng["to"],
    )
    campaigns = store.list_campaigns(client_id=client_id, channel=channel)

    campaign_rows = []
    for campaign in campaigns:
        campaign_counts = per_campaign.get(str(campaign["id"]))
        if not campaign_counts:
            continue
        campaign_rows.append({
            "campaign": campaign,
            "counts":   campaign_counts,
            "funnel":   funnel_with_rates(campaign_counts),
        })
    campaign_rows.sort(key=lambda r: r["counts"].get(EVENT_SENT, 0), reverse=True)

    return {
        "client":       client,
        "counts":       counts,
        "funnel":       funnel_with_rates(counts),
        "by_channel":   store.funnel_by_channel(
                            client_id=client_id,
                            date_from=rng["from"], date_to=rng["to"]),
        "timeseries":   store.funnel_timeseries(
                            client_id=client_id, channel=channel,
                            date_from=rng["from"], date_to=rng["to"]),
        "campaign_rows": campaign_rows,
        "campaigns":    campaigns,
        "range":        rng,
        "range_query":  _range_query(rng),
        "channel":      channel,
    }


# ── Sync status ───────────────────────────────────────────────────────────────

def _humanize_ago(seconds: float) -> str:
    s = int(seconds)
    if s < 60:
        return "just now" if s < 10 else f"{s}s ago"
    m = s // 60
    if m < 60:
        return f"{m} min ago"
    h = m // 60
    if h < 24:
        rem = m % 60
        return f"{h}h ago" if rem == 0 else f"{h}h {rem}m ago"
    return f"{h // 24}d ago"


def _sync_status_view() -> dict:
    """Per-connector status for the sync panel.

    `state` drives the pill colour and is what someone glances at:
      running | ok | partial | error | stale | never | off
    """
    latest = store.latest_sync_per_tool()
    now = datetime.now(timezone.utc)
    out = []
    for source_tool, connector in pulse_sync.CONNECTORS.items():
        row = latest.get(source_tool)
        entry = {
            "source_tool": source_tool,
            "label":       source_tool.replace("_", " ").title(),
            "configured":  connector.is_configured(),
            "running":     pulse_sync.is_running(source_tool),
            "state":       "never",
            "ago_text":    "",
            "campaigns":   0,
            "events":      0,
            "error":       "",
            "run_at":      "",
        }
        if not entry["configured"]:
            entry["state"] = "off"
            entry["ago_text"] = "no API key set"
        if row:
            entry["campaigns"] = row.get("campaigns_synced") or 0
            entry["events"]    = row.get("events_inserted") or 0
            entry["error"]     = row.get("error_message") or ""
            entry["run_at"]    = row.get("finished_at") or row.get("run_at") or ""
            stamp = _parse_iso(entry["run_at"])
            # A past run must never make an unconfigured connector look healthy.
            # Pulling the key breaks the sync; history from before that says
            # nothing about now, and "ok" here is the exact false all-clear this
            # panel exists to prevent.
            if stamp and entry["configured"]:
                age = (now - stamp).total_seconds()
                entry["ago_text"] = _humanize_ago(age)
                if entry["running"]:
                    entry["state"] = "running"
                elif age > _SYNC_STALE_AFTER_MIN * 60:
                    entry["state"] = "stale"
                else:
                    entry["state"] = row.get("status") or "ok"
        if entry["running"]:
            entry["state"] = "running"
        out.append(entry)
    return {"connectors": out, "logs": store.recent_sync_logs(limit=12)}


def _parse_iso(value: str) -> datetime | None:
    if not value:
        return None
    try:
        stamp = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None
    return stamp if stamp.tzinfo else stamp.replace(tzinfo=timezone.utc)


# ── Pages ─────────────────────────────────────────────────────────────────────

@router.get("/outbound-pulse")
async def pulse_page(request: Request):
    return templates.TemplateResponse("outbound_pulse.html", {
        "request":  request,
        "active":   "outbound_pulse",
        "sop_key":  "outbound_pulse",
        "presets":  RANGE_PRESETS,
    })


@router.get("/outbound-pulse/clients/{client_id}")
async def pulse_client_page(
    request:   Request,
    client_id: str,
    range:     str = Query("30d"),
    date_from: str = Query(""),
    date_to:   str = Query(""),
    channel:   str = Query(""),
):
    rng = _resolve_range(range, date_from, date_to)
    try:
        context = _client_context(client_id, rng, _channel_filter(channel))
    except PulseNotReady as exc:
        return templates.TemplateResponse("outbound_pulse_client.html", {
            "request": request, "active": "outbound_pulse", "sop_key": "outbound_pulse",
            "error": str(exc), "client": None, "presets": RANGE_PRESETS,
            "range": rng, "channel": channel,
        })
    if context is None:
        return templates.TemplateResponse("404.html", {"request": request}, status_code=404)

    return templates.TemplateResponse("outbound_pulse_client.html", {
        "request":  request,
        "active":   "outbound_pulse",
        "sop_key":  "outbound_pulse",
        "presets":  RANGE_PRESETS,
        "error":    "",
        **context,
    })


# ── Partials ──────────────────────────────────────────────────────────────────

@router.get("/api/outbound-pulse/overview")
async def overview(
    request:   Request,
    range:     str = Query("30d"),
    date_from: str = Query(""),
    date_to:   str = Query(""),
    channel:   str = Query(""),
):
    rng = _resolve_range(range, date_from, date_to)
    try:
        context = _overview_context(rng, _channel_filter(channel))
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    return templates.TemplateResponse(
        "partials/pulse_overview.html", {"request": request, **context},
    )


@router.get("/api/outbound-pulse/sync-status")
async def sync_status(request: Request):
    try:
        context = _sync_status_view()
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    return templates.TemplateResponse(
        "partials/pulse_sync_status.html", {"request": request, **context},
    )


def _campaign_table(request: Request, filters: dict):
    """Render the mapping table for a filter set.

    Shared by the list endpoint and the mapping POST so a mapping change
    re-renders the same filtered view the user was working in, rather than
    dumping them back to all ~1000 campaigns after every single mapping.
    """
    campaigns = store.list_campaigns(
        status=filters["status"],
        source_tool=filters["source_tool"],
        channel=filters["channel"],
        synced=filters["synced"],
        client_id=filters["client_id"],
        q=filters["q"],
    )
    clients = store.list_clients()
    options = store.campaign_filter_options()
    # Unmapped first: those are the ones whose events are missing from a funnel.
    campaigns.sort(key=lambda c: (bool(c.get("client_id")), c.get("name") or ""))
    return templates.TemplateResponse("partials/pulse_campaigns.html", {
        "request":      request,
        "campaigns":    campaigns,
        "clients":      clients,
        "client_index": {str(c["id"]): c for c in clients},
        "filters":      filters,
        "options":      options,
    })


def _campaign_filters(status: str, source_tool: str, channel: str,
                      synced: str, client_id: str = "", q: str = "") -> dict:
    return {
        "status":      (status or "").strip(),
        "source_tool": (source_tool or "").strip(),
        "channel":     _channel_filter(channel),
        "synced":      synced if synced in ("never", "synced") else "",
        # "none" means "no client mapped" and is distinct from "" (no filter).
        "client_id":   (client_id or "").strip(),
        # Capped: this becomes a LIKE pattern, and an unbounded one scans the
        # whole table for no benefit — nobody searches on 200 characters.
        "q":           (q or "").strip()[:100],
    }


@router.get("/api/outbound-pulse/campaigns")
async def campaign_mapping(
    request:     Request,
    status:      str = Query(""),
    source_tool: str = Query(""),
    channel:     str = Query(""),
    synced:      str = Query(""),
    client_id:   str = Query(""),
    q:           str = Query(""),
):
    """Campaign → client mapping table, filterable by the source tool's own
    status/source/channel, by whether events have ever been pulled, and by
    which client a campaign is mapped to."""
    try:
        return _campaign_table(
            request,
            _campaign_filters(status, source_tool, channel, synced, client_id, q),
        )
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.get("/api/outbound-pulse/engagement")
async def engagement(request: Request):
    """Portal engagement — success-criteria question #2, "did clients use it"."""
    try:
        visits = store.visit_summary()
        clients = _client_index()
    except PulseNotReady as exc:
        return _not_ready_box(exc)

    per_client: dict[str, dict] = {}
    for visit in visits:
        cid = str(visit.get("client_id") or "")
        entry = per_client.setdefault(cid, {"views": 0, "last": ""})
        entry["views"] += 1
        stamp = str(visit.get("viewed_at") or "")
        if stamp > entry["last"]:
            entry["last"] = stamp

    rows = []
    for cid, entry in per_client.items():
        client = clients.get(cid)
        if not client:
            continue
        stamp = _parse_iso(entry["last"])
        rows.append({
            "client": client,
            "views":  entry["views"],
            "last":   entry["last"],
            "ago":    _humanize_ago((datetime.now(timezone.utc) - stamp).total_seconds())
                      if stamp else "",
        })
    rows.sort(key=lambda r: r["views"], reverse=True)
    return templates.TemplateResponse(
        "partials/pulse_engagement.html", {"request": request, "rows": rows},
    )


# ── Manual sync ───────────────────────────────────────────────────────────────

@router.post("/api/outbound-pulse/sync")
async def manual_sync(
    request: Request,
    source_tool: str = Query(""),
    user: dict = Depends(auth.require_login),
):
    """Re-run a connector now. Safe to press repeatedly — event writes dedupe.

    Runs inline rather than in a background thread so the response reports what
    actually happened; that is the whole point of the control while a connector
    is being validated.
    """
    triggered_by = f"manual:{user.get('email', 'unknown')}"
    try:
        store.backfill_client_agency()
        if source_tool:
            results = [pulse_sync.sync_source(source_tool, triggered_by)]
        else:
            results = pulse_sync.sync_all(triggered_by)
        context = _sync_status_view()
    except PulseNotReady as exc:
        return _not_ready_box(exc)

    response = templates.TemplateResponse(
        "partials/pulse_sync_status.html",
        {"request": request, "results": results, **context},
    )
    # Tells the page to refresh the funnel — new events may have landed.
    response.headers["HX-Trigger"] = "pulseSynced"
    return response


# ── Campaign → client mapping ─────────────────────────────────────────────────

@router.post("/api/outbound-pulse/campaigns/{campaign_id}/client")
async def map_campaign(
    request:     Request,
    campaign_id: str,
    client_id:        str = Form(""),
    status:           str = Form(""),
    source_tool:      str = Form(""),
    channel:          str = Form(""),
    synced:           str = Form(""),
    filter_client_id: str = Form(""),
    q:                str = Form(""),
):
    """Map one campaign, then re-render the table under the SAME filters.

    The filter values ride along with the form post for that reason: mapping is
    done in batches within a filtered view, and resetting to all campaigns after
    each one would make mapping a client's campaigns unworkable.
    """
    filters = _campaign_filters(status, source_tool, channel, synced,
                                filter_client_id, q)
    try:
        store.set_campaign_client(campaign_id, client_id.strip() or None)
        response = _campaign_table(request, filters)
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    # A distinct event from pulseSynced: this response already carries the new
    # table, so the list must not re-fetch itself. Only the funnel needs to know.
    response.headers["HX-Trigger"] = "pulseMappingChanged"
    return response


@router.post("/api/outbound-pulse/campaigns/bulk-map")
async def bulk_map_campaigns(
    request:          Request,
    campaign_ids:     list[str] = Form([]),
    target_client_id: str = Form(""),
    status:           str = Form(""),
    source_tool:      str = Form(""),
    channel:          str = Form(""),
    synced:           str = Form(""),
    client_id:        str = Form(""),
    q:                str = Form(""),
):
    """Map many campaigns to one client at once.

    Two separate client fields, deliberately: `target_client_id` is what the
    selected campaigns get mapped TO, while `client_id` is the table's current
    Client *filter* and only decides what to re-render. Reusing one name would
    make a bulk map performed inside a filtered view silently reassign to the
    filter's client.
    """
    filters = _campaign_filters(status, source_tool, channel, synced, client_id, q)
    ids = [i for i in campaign_ids if i and i.strip()]
    if not ids:
        return _error_box("Select at least one campaign first.")

    try:
        store.set_campaigns_client(ids, target_client_id.strip() or None)
        response = _campaign_table(request, filters)
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    response.headers["HX-Trigger"] = "pulseMappingChanged"
    return response


# ── Meet Alfred CSV import ────────────────────────────────────────────────────

@router.post("/api/outbound-pulse/import/meet-alfred")
async def import_meet_alfred(
    request:       Request,
    file:          UploadFile = File(...),
    client_id:     str = Form(""),
    campaign_name: str = Form(""),
):
    """Import a Meet Alfred report export into the same normalized funnel.

    Exists because Meet Alfred's API access is not confirmed for this account
    (see app/utils/pulse/meet_alfred.py).  It runs the CSV through the exact
    same `events_from_activities` mapping the API path uses, so the two cannot
    drift apart, and the same dedupe keys make a re-import a no-op.
    """
    name = campaign_name.strip()
    if not name:
        return _error_box("Give the campaign a name so its events can be grouped.")

    try:
        raw = await file.read()
    except Exception as exc:
        return _error_box(f"Could not read the upload: {exc}")
    if not raw:
        return _error_box("That file is empty.")

    try:
        rows = meet_alfred.parse_csv(raw)
    except Exception as exc:
        return _error_box(f"Could not parse the CSV: {exc}")
    if not rows:
        return _error_box("No rows found in that CSV.")

    log_id = None
    try:
        agency_id = store.current_agency_id()
        log_id = store.start_sync_log(SOURCE_MEET_ALFRED, "manual:csv-import")
        campaign = store.upsert_campaign(
            source_tool=SOURCE_MEET_ALFRED,
            # Deterministic external id so re-importing an updated export
            # updates the same campaign instead of creating a second one.
            external_campaign_id=f"csv:{name.lower()}",
            channel=CHANNEL_LINKEDIN,
            name=name,
            status="imported",
            client_id=client_id.strip() or None,
            raw={"import": "csv", "filename": file.filename},
        )
        if not campaign:
            store.finish_sync_log(log_id, status="error",
                                  error_message="campaign upsert failed")
            return _error_box("Could not create the campaign record.")

        events = meet_alfred.events_from_activities(
            rows,
            agency_id=agency_id,
            campaign_id=str(campaign["id"]),
            client_id=campaign.get("client_id"),
        )
        inserted = store.insert_events(events)
        store.finish_sync_log(log_id, status="ok", campaigns_synced=1,
                              events_inserted=inserted)
    except PulseNotReady as exc:
        store.finish_sync_log(log_id, status="error", error_message=str(exc))
        return _not_ready_box(exc)
    except Exception as exc:
        store.finish_sync_log(log_id, status="error", error_message=str(exc))
        return _error_box(f"Import failed: {exc}")

    skipped = len(rows) - len({e["lead_key"] for e in events}) if events else len(rows)
    response = HTMLResponse(
        '<div class="result-meta" style="margin:0">'
        f'Imported <strong>{inserted}</strong> new event'
        f'{"" if inserted == 1 else "s"} from '
        f'<strong>{escape(str(len(rows)))}</strong> row'
        f'{"" if len(rows) == 1 else "s"} into '
        f'<strong>{escape(name)}</strong>.'
        + (f' {skipped} row{"" if skipped == 1 else "s"} had no usable prospect '
           'identifier or date and were skipped.' if skipped > 0 else '')
        + '</div>'
    )
    response.headers["HX-Trigger"] = "pulseSynced"
    return response


# ── Client portal access ──────────────────────────────────────────────────────

def _hash_token(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


@router.get("/api/outbound-pulse/clients/{client_id}/access")
async def list_client_access(request: Request, client_id: str):
    try:
        links = store.list_access(client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    return templates.TemplateResponse("partials/pulse_access.html", {
        "request":   request,
        "links":     links,
        "client_id": client_id,
        "new_link":  "",
    })


@router.post("/api/outbound-pulse/clients/{client_id}/access")
async def create_client_access(
    request:   Request,
    client_id: str,
    label:     str = Form(""),
    user:      dict = Depends(auth.require_login),
):
    """Mint a portal link.

    The token is generated here and only its SHA-256 hash is stored, so the
    plaintext link is shown exactly once — it cannot be recovered later, only
    revoked and re-issued.  That is the trade for a database dump not being a
    set of live client report links.
    """
    token = secrets.token_urlsafe(32)
    expires = (datetime.now(timezone.utc) + timedelta(days=_ACCESS_TTL_DAYS)).isoformat()
    try:
        created = store.create_access(
            client_id=client_id,
            token_hash=_hash_token(token),
            label=label.strip(),
            created_by=user.get("email", ""),
            expires_at=expires,
        )
        if not created:
            return _error_box("Could not create the portal link.")
        links = store.list_access(client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)

    base = str(request.base_url).rstrip("/")
    return templates.TemplateResponse("partials/pulse_access.html", {
        "request":   request,
        "links":     links,
        "client_id": client_id,
        "new_link":  f"{base}/portal/{token}",
    })


@router.delete("/api/outbound-pulse/clients/{client_id}/access/{access_id}")
async def revoke_client_access(request: Request, client_id: str, access_id: str):
    try:
        store.revoke_access(access_id)
        links = store.list_access(client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    return templates.TemplateResponse("partials/pulse_access.html", {
        "request":   request,
        "links":     links,
        "client_id": client_id,
        "new_link":  "",
    })


@router.get("/api/outbound-pulse/clients")
async def clients_json():
    """Client list as JSON — used by the import form's client picker."""
    try:
        return JSONResponse(store.list_clients())
    except PulseNotReady as exc:
        return JSONResponse({"error": str(exc)}, status_code=503)
