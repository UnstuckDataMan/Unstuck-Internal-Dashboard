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
import logging
import os
import re
import secrets
from datetime import date, datetime, timedelta, timezone
from html import escape

from fastapi import APIRouter, Depends, File, Form, Query, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse

from app import auth
from app.deps import templates
from app.utils.dates import today_utc
from app.utils.pulse import (
    abtest,
    hand_entry,
    tracker,
    meet_alfred,
    richtext,
    smartlead,
    store,
    sync as pulse_sync,
)
from app.utils import copy_bank as shared_copy_bank
from app.utils import google_sheets
from app.utils.pulse.reports import report_heading
from app.utils.pulse.normalize import (
    CHANNEL_EMAIL,
    empty_funnel,
    CHANNEL_LINKEDIN,
    EVENT_SENT,
    FUNNEL_STAGES,
    SOURCE_HAND_ENTRY,
    SOURCE_LABELS,
    SOURCE_MANUAL,
    SOURCE_MEET_ALFRED,
    SOURCE_SMARTLEAD,
    bucket_timeseries,
    funnel_with_rates,
    opens_are_tracked,
)
from app.utils.pulse import template_filters
from app.utils.pulse.store import PulseNotReady

router = APIRouter()
logger = logging.getLogger(__name__)
template_filters.register(templates.env)

# Portal links are long-lived by default — a client should not have a report
# link die between monthly cycles — but not permanent.
_ACCESS_TTL_DAYS = 180

# 16 bytes is 22 URL-safe characters and 128 bits of entropy. The link was
# built from 32 bytes, which is 43 characters — twice the length for security
# nobody was going to exhaust either way. Guessing one at a thousand attempts
# per second would still take longer than the universe has existed.
_TOKEN_BYTES = 16

# A run older than this means the hourly schedule has not fired (a sleeping
# Render instance, a crashed thread), which is the failure mode the sync-status
# panel exists to catch.
_SYNC_STALE_AFTER_MIN = 75

# Shortest first, so the control reads as a scale. The long ones matter for
# lead quality: a month of outbound is a handful of graded leads, and a shape
# only appears over half a year.
RANGE_PRESETS: dict[str, str] = {
    "7d":   "Last 7 days",
    "30d":  "Last 30 days",
    "mtd":  "This month",
    "90d":  "Last 90 days",
    "6m":   "Last 6 months",
    "12m":  "Last 12 months",
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
    elif key in ("6m", "12m"):
        # Calendar months back, not a fixed number of days: "last 6 months" on
        # the 31st has to mean the same span as on the 1st, and 182 days does
        # not. Clamped to the 1st so a short month cannot overshoot.
        months = 6 if key == "6m" else 12
        year, month = today.year, today.month - months
        while month <= 0:
            month += 12
            year -= 1
        start = date(year, month, 1)
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
    """Explain a failed read with advice that matches its actual cause.

    This used to append "apply the migration" to every failure. When the real
    problem was a query timeout on an already-migrated database, that sent
    people to re-run SQL that was never the issue.
    """
    kind = getattr(exc, "kind", "error")
    if kind == "missing":
        advice = (" A Pulse table or view is missing — apply "
                  "migrations/outbound_pulse_schema.sql in the Supabase SQL editor.")
    elif kind == "timeout":
        advice = (" The database cancelled the query for running too long. If "
                  "migrations/outbound_pulse_rollup.sql has not been applied, apply it — "
                  "it replaces the slow funnel view with a pre-aggregated table.")
    else:
        advice = ""
    return _error_box(f"{exc}{advice}")


# ── Shared view assembly ──────────────────────────────────────────────────────

def _client_index() -> dict[str, dict]:
    return {str(c["id"]): c for c in store.list_clients()}


def _channel_filter(channel: str) -> str:
    """A true channel, for the campaign mapping table.

    Campaigns carry a channel, so the mapping filters stay channel-based. The
    reporting toolbar is a different question — see _report_scope.
    """
    return channel if channel in (CHANNEL_EMAIL, CHANNEL_LINKEDIN) else ""


# The reporting toolbar's "Channel" control, resolved to the source tool that
# produced the numbers rather than to the channel column.
#
# Manual outreach is email, so filtering it by channel would put it inside
# Email as well as inside Manual: two options that overlap, and a Both that is
# not their sum. Every option here is one tool, so they partition the data.
REPORT_SCOPES: dict[str, str] = {
    CHANNEL_EMAIL:     SOURCE_SMARTLEAD,
    CHANNEL_LINKEDIN:  SOURCE_MEET_ALFRED,
    SOURCE_MANUAL:     SOURCE_MANUAL,
    SOURCE_HAND_ENTRY: SOURCE_HAND_ENTRY,
}


def _report_scope(value: str) -> str:
    return value if value in REPORT_SCOPES else ""


def _scope_source(scope: str) -> str:
    return REPORT_SCOPES.get(scope, "")


# Hidden from the overview for anyone who has not said otherwise. Business
# development is our own pipeline rather than a client's, and sorted by send
# volume it sat at the top pushing real clients below the fold.
#
# Matched by name, not by id: the id differs between Supabase projects and this
# has to be right on a fresh database with nothing seeded. Env-overridable,
# comma separated, for the same reason PULSE_AGENCY_NAME is.
DEFAULT_EXCLUDED_CLIENT_NAMES: tuple[str, ...] = tuple(
    name.strip().lower()
    for name in os.environ.get(
        "PULSE_DEFAULT_EXCLUDED_CLIENTS", "Unstuck - Business Development",
    ).split(",")
    if name.strip()
)


def _pref_key(user: dict) -> str:
    return (user.get("email") or "").strip().lower()


def _default_excluded_ids(clients: list[dict]) -> set[str]:
    return {str(c["id"]) for c in clients
            if (c.get("name") or "").strip().lower() in DEFAULT_EXCLUDED_CLIENT_NAMES}


def _hidden_client_ids(user: dict, clients: list[dict]) -> set[str]:
    """Clients this person does not want on the overview.

    Both modes resolve to the same answer — a set to leave out — so everything
    downstream stays as it was. "Show only these" is the complement of the
    selection; "hide these" is the selection itself.

    "Never chosen" and "chose nothing" are different answers, and the
    difference is whether a pulse_user_prefs row exists. Collapsing them would
    make "show me everything" impossible to save: an empty list would read back
    as the default and silently re-hide what the user had just un-hidden.
    """
    saved = store.get_user_filter(_pref_key(user))
    if saved is None:
        return _default_excluded_ids(clients)
    chosen = {str(cid) for cid in saved["ids"]}
    if saved["mode"] == store.FILTER_ONLY:
        return {str(c["id"]) for c in clients} - chosen
    return chosen


def _overview_context(rng: dict, channel: str, hidden: set[str] | None = None,
                      filter_mode: str = "") -> dict:
    """The internal overview.

    `hidden` is resolved by the caller rather than read here, so this function
    knows nothing about sessions and a test can pin the set directly. It
    applies to the client list, and therefore to the totals and the channel
    strip, which are both built from the rows that survive it. It applies
    nowhere else: a hidden client's own page and their portal report do not go
    through this function at all.
    """
    hidden = hidden or set()
    source = _scope_source(channel)
    clients = store.list_clients()
    # One query for both splits. They have to agree: a grid filtered by an
    # exclusion next to a separately-queried channel strip that is not would
    # print a strip that does not add up to the totals above it.
    per_client_channel = store.funnel_by_client_and_channel(
        source_tool=source, date_from=rng["from"], date_to=rng["to"],
    )

    rows: list[dict] = []
    by_channel: dict[str, dict[str, int]] = {}
    hidden_clients: list[dict] = []
    for client in clients:
        cid = str(client["id"])
        channels = per_client_channel.get(cid)
        if not channels:
            continue     # no outbound activity in range — not a reporting row
        if cid in hidden:
            hidden_clients.append(client)   # named on the page, never silently dropped
            continue
        counts = empty_funnel()
        for chan, chan_counts in channels.items():
            target = by_channel.setdefault(chan, empty_funnel())
            for stage in FUNNEL_STAGES:
                counts[stage] += chan_counts.get(stage, 0)
                target[stage] += chan_counts.get(stage, 0)
        rows.append({
            "client":  client,
            "counts":  counts,
            "funnel":  funnel_with_rates(counts),
            "sent":    counts.get(EVENT_SENT, 0),
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
        "total_counts": totals,
        "total_sent":  totals[EVENT_SENT],
        "by_channel":  by_channel,
        "hidden_clients": hidden_clients,
        "shown_clients":  [r["client"] for r in rows],
        "filter_mode":    filter_mode,
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

    source = _scope_source(channel)
    counts = store.funnel(
        client_id=client_id, source_tool=source,
        date_from=rng["from"], date_to=rng["to"],
    )
    per_campaign = store.funnel_by_campaign(
        client_id=client_id, source_tool=source,
        date_from=rng["from"], date_to=rng["to"],
    )
    campaigns = store.list_campaigns(client_id=client_id, source_tool=source)

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

    # Per source, from one query. Every tab is rendered up front because the
    # numbers are already in hand — switching tabs should not cost a request.
    # Filtered by the toolbar too. Left unfiltered, picking Manual in the
    # toolbar and then opening the Smartlead tab showed Smartlead numbers under
    # a page that said Manual — two controls disagreeing about the same view.
    by_source = store.funnel_by_source(
        client_id=client_id, source_tool=source,
        date_from=rng["from"], date_to=rng["to"],
    )
    sources = [{
        "key":    key,
        "label":  SOURCE_LABELS.get(key, key.replace("_", " ").title()),
        "counts": by_source.get(key, {}),
        "funnel": funnel_with_rates(by_source.get(key, {})),
        "active": bool(by_source.get(key)),
    } for key in SOURCE_LABELS]

    return {
        "client":       client,
        "counts":       counts,
        "funnel":       funnel_with_rates(counts),
        "sources":      sources,
        # Same rule the funnel uses to drop its Opened stage, so the campaign
        # table doesn't show a column of zeros the funnel just chose to hide.
        "show_opened":  opens_are_tracked(counts),
        # Filtered like the funnel above it: an unfiltered split under a
        # narrowed funnel reads as a contradiction.
        "by_channel":   store.funnel_by_channel(
                            client_id=client_id, source_tool=source,
                            date_from=rng["from"], date_to=rng["to"]),
        # Bucketed rather than raw days: an all-time range can be years long,
        # and one column per day overflowed the panel instead of fitting it.
        "trend":        bucket_timeseries(store.funnel_timeseries(
                            client_id=client_id, source_tool=source,
                            date_from=rng["from"], date_to=rng["to"])),
        "campaign_rows": campaign_rows,
        "campaigns":    campaigns,
        # Their own list: a manual campaign has no row in the funnel's campaign
        # table, because its events are recorded against the client.
        "manual_campaigns": manual_campaigns(client_id),
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
async def pulse_page(
    request:   Request,
    range:     str = Query("30d"),
    date_from: str = Query(""),
    date_to:   str = Query(""),
):
    """The overview shell. The partial inside it fetches the numbers.

    It takes the range parameters so the control can show the range actually
    in force. Without them the template hardcoded "Last 30 days" as selected,
    so after a custom range was applied the dropdown and the figures on screen
    disagreed with each other.
    """
    return templates.TemplateResponse("outbound_pulse.html", {
        "request":  request,
        "active":   "outbound_pulse",
        "sop_key":  "outbound_pulse",
        "presets":  RANGE_PRESETS,
        "range":    _resolve_range(range, date_from, date_to),
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
        context = _client_context(client_id, rng, _report_scope(channel))
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
    user:      dict = Depends(auth.require_login),
):
    rng = _resolve_range(range, date_from, date_to)
    try:
        saved = store.get_user_filter(_pref_key(user))
        hidden = _hidden_client_ids(user, store.list_clients())
        context = _overview_context(
            rng, _report_scope(channel), hidden,
            filter_mode=(saved or {}).get("mode", store.FILTER_EXCLUDE))
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
        outcomes = meet_alfred.outcomes_from_activities(
            rows, agency_id=agency_id, campaign_id=str(campaign["id"]),
        )
        saved = store.upsert_outcomes(outcomes)
        if saved < len(outcomes):
            message = f"saved {saved} of {len(outcomes)} lead outcomes"
            store.finish_sync_log(log_id, status="partial", campaigns_synced=1,
                                  events_inserted=inserted, error_message=message)
            return _error_box(
                f"Events imported, but {message}. Has "
                "migrations/outbound_pulse_outcomes.sql been applied?"
            )
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

def _portal_base(request: Request) -> str:
    """Where a client report lives, for the link we hand out.

    PORTAL_BASE_URL wins when it is set, so a custom domain can front the
    reports without the app having to know it is behind one. Falls back to
    whatever host this request arrived on, which is the Render URL.
    """
    configured = os.environ.get("PORTAL_BASE_URL", "").strip().rstrip("/")
    if configured:
        return configured if "://" in configured else f"https://{configured}"
    return str(request.base_url).rstrip("/")


def _hash_token(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


@router.get("/api/outbound-pulse/clients/{client_id}/access")
async def list_client_access(request: Request, client_id: str):
    try:
        links = store.list_access(client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    return templates.TemplateResponse("partials/pulse_access.html", {
        "request":     request,
        "links":       links,
        "client_id":   client_id,
        "portal_base": _portal_base(request),
        "new_link":    "",
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
    token = secrets.token_urlsafe(_TOKEN_BYTES)
    expires = (datetime.now(timezone.utc) + timedelta(days=_ACCESS_TTL_DAYS)).isoformat()
    try:
        created = store.create_access(
            client_id=client_id,
            token_hash=_hash_token(token),
            label=label.strip(),
            created_by=user.get("email", ""),
            expires_at=expires,
            token=token,
        )
        if not created:
            return _error_box("Could not create the portal link.")
        links = store.list_access(client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)

    base = _portal_base(request)
    return templates.TemplateResponse("partials/pulse_access.html", {
        "request":     request,
        "links":       links,
        "client_id":   client_id,
        "portal_base": base,
        "new_link":    f"{base}/r/{token}",
    })


@router.delete("/api/outbound-pulse/clients/{client_id}/access/{access_id}")
async def revoke_client_access(request: Request, client_id: str, access_id: str):
    try:
        store.revoke_access(access_id)
        links = store.list_access(client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    return templates.TemplateResponse("partials/pulse_access.html", {
        "request":     request,
        "links":       links,
        "client_id":   client_id,
        "portal_base": _portal_base(request),
        "new_link":    "",
    })


# ── Hand-entered figures ────────────────────────────────────────


def _hand_entry_panel(request: Request, client_id: str, error: str = "",
                      notice: str = "", channel: str = "",
                      month: str = "", values: dict | None = None,
                      source_note: str = ""):
    entries = store.list_hand_entries(client_id)
    for entry in entries:
        parsed = hand_entry.parse_month(str(entry.get("period_month") or ""))
        entry["label"] = hand_entry.month_label(parsed) if parsed else "—"
        entry["channel_label"] = ("LinkedIn" if entry.get("channel") == CHANNEL_LINKEDIN
                                  else "Email")
    # Default to last whole month: the figure you are most likely keying in.
    if not month:
        today = today_utc()
        month = (today.replace(day=1) - timedelta(days=1)).replace(day=1).isoformat()[:7]
    return templates.TemplateResponse("partials/pulse_hand_entry.html", {
        "request":   request,
        "client_id": client_id,
        "entries":   entries,
        "metrics":   hand_entry.METRICS,
        "labels":    hand_entry.labels_for(),
        "channel":   hand_entry.CHANNEL,
        "month":     month,
        # What the person typed, so a re-render gives it back. The overlap
        # warning is a round trip to the server: without this, "save again to
        # confirm" came back to an empty form and confirmed a month of zeros.
        "values":      values or {},
        "source_note": source_note,
        "error":     error,
        "notice":    notice,
    })


@router.get("/api/outbound-pulse/clients/{client_id}/hand-entries")
async def list_client_hand_entries(
    request:   Request,
    client_id: str,
    channel:   str = Query(CHANNEL_EMAIL),
):
    try:
        return _hand_entry_panel(request, client_id,
                                 channel=hand_entry.CHANNEL)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.post("/api/outbound-pulse/clients/{client_id}/hand-entries")
async def save_client_hand_entry(
    request:      Request,
    client_id:    str,
    channel:      str = Form(CHANNEL_EMAIL),
    period_month: str = Form(""),
    source_note:  str = Form(""),
    sent:         str = Form(""),
    opened:       str = Form(""),
    replied:      str = Form(""),
    meetings:     str = Form(""),
    bounced:      str = Form(""),
    unsubscribed: str = Form(""),
    confirm:      str = Form(""),
    user:         dict = Depends(auth.require_login),
):
    """Record a month of figures read off another tool's dashboard.

    Warns, rather than refuses, when a connector already covers the same
    client, channel and month: a partly-synced month is a real situation, and
    only the person keying it in can tell whether the figures overlap.
    """
    # LinkedIn always: Meet Alfred is the only source a person keys in.
    chan = hand_entry.CHANNEL
    month = hand_entry.parse_month(period_month)
    typed = {"sent": sent, "opened": opened, "replied": replied,
             "meetings": meetings, "bounced": bounced,
             "unsubscribed": unsubscribed}
    note = source_note.strip()
    try:
        if month is None:
            return _hand_entry_panel(request, client_id, "Pick a month.",
                                     channel=chan, values=typed,
                                     source_note=note)
        metrics, error = hand_entry.clean_metrics(typed)
        if error:
            return _hand_entry_panel(request, client_id, error, channel=chan,
                                     month=month.isoformat()[:7],
                                     values=typed, source_note=note)
        if not any(metrics.values()):
            # Every field blank or zero. The funnel view skips zero rows, so
            # this would save a record that shows up nowhere and looks exactly
            # like the save having failed.
            return _hand_entry_panel(
                request, client_id,
                "Enter at least one figure. To remove a month you have already "
                "entered, delete it below.",
                channel=chan, month=month.isoformat()[:7],
                values=typed, source_note=note)

        if not confirm:
            start, end = hand_entry.month_bounds(month)
            covered = {
                source: counts for source, counts in store.funnel_by_source(
                    client_id=client_id, date_from=start, date_to=end).items()
                if source != SOURCE_HAND_ENTRY and counts.get(EVENT_SENT)
            }
            if covered:
                named = ", ".join(SOURCE_LABELS.get(s, s) for s in covered)
                return _hand_entry_panel(
                    request, client_id,
                    channel=chan, month=month.isoformat()[:7],
                    values=typed, source_note=note,
                    notice=(f"{named} already reported sends for "
                            f"{hand_entry.month_label(month)}. Entering figures "
                            f"here adds to those rather than replacing them — "
                            f"save again to confirm."))

        saved = store.upsert_hand_entry(
            client_id=client_id, channel=chan,
            period_month=month.isoformat(),
            metrics=metrics,
            source_note=note,
            entered_by=user.get("name", "") or user.get("email", ""),
        )
        if not saved:
            return _hand_entry_panel(
                request, client_id,
                "Could not save that. The details are in the server log.",
                channel=chan, month=month.isoformat()[:7],
                values=typed, source_note=note)
    except PulseNotReady as exc:
        return _not_ready_box(exc)

    response = _hand_entry_panel(request, client_id, channel=chan)
    # A full reload, not a panel swap. This page is server-rendered, so the
    # funnel, the source tabs and the trend above this panel are all stale the
    # moment a figure lands — and a page disagreeing with itself is exactly
    # what this module keeps having to be fixed for.
    response.headers["HX-Refresh"] = "true"
    return response


@router.delete("/api/outbound-pulse/clients/{client_id}/hand-entries/{entry_id}")
async def remove_client_hand_entry(request: Request, client_id: str, entry_id: str):
    try:
        store.delete_hand_entry(entry_id)
        response = _hand_entry_panel(request, client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    response.headers["HX-Refresh"] = "true"
    return response


# ── Which clients show on the overview ─────────────────────────────


def _exclusions_panel(request: Request, user: dict, error: str = "",
                      mode: str = ""):
    clients = store.list_clients()
    saved = store.get_user_filter(_pref_key(user))
    if saved is None:
        chosen = _default_excluded_ids(clients)
        saved_mode = store.FILTER_EXCLUDE
    else:
        chosen = {str(c) for c in saved["ids"]}
        saved_mode = saved["mode"]
    # `mode` is passed only when a rejected save should keep the radio where
    # the user put it rather than snapping back to what is stored.
    shown_mode = mode or saved_mode
    hidden_ids = _hidden_client_ids(user, clients)
    return templates.TemplateResponse("partials/pulse_exclusions.html", {
        "request":    request,
        "clients":    clients,
        "chosen_ids": chosen,
        "mode":       shown_mode,
        "only":       shown_mode == store.FILTER_ONLY,
        "hidden":     [c for c in clients if str(c["id"]) in hidden_ids],
        "shown":      [c for c in clients if str(c["id"]) not in hidden_ids],
        # Shown as a hint, so someone can tell "the house default" from "what I
        # chose" without having to remember whether they ever chose.
        "is_default": saved is None,
        "error":      error,
    })


@router.get("/api/outbound-pulse/exclusions")
async def list_exclusions(request: Request,
                          user: dict = Depends(auth.require_login)):
    try:
        return _exclusions_panel(request, user)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.post("/api/outbound-pulse/exclusions")
async def save_exclusions(
    request:  Request,
    excluded: list[str] = Form([]),
    mode:     str = Form(store.FILTER_EXCLUDE),
    user:     dict = Depends(auth.require_login),
):
    """Save how the overview is narrowed and refresh it.

    No "was this really submitted" marker is needed here, unlike the report
    editor: a POST to this route is itself the deliberate act, so an empty
    selection in exclude mode always writes an empty list — "hide nothing".
    Returning to the default is the DELETE below, a different request.
    """
    chosen_mode = mode if mode in (store.FILTER_EXCLUDE, store.FILTER_ONLY) \
        else store.FILTER_EXCLUDE
    try:
        clients = store.list_clients()
        known = {str(c["id"]) for c in clients}
        # Intersected with the real client list, so a hand-made post cannot put
        # a non-uuid into a uuid[] column and fail the whole save.
        chosen = sorted(known & set(excluded))
        if chosen_mode == store.FILTER_ONLY and not chosen:
            # Refused rather than saved: an overview with nothing in it reads
            # as broken, not as a filter.
            return _exclusions_panel(
                request, user, "Pick at least one client to show.",
                mode=chosen_mode)
        if not store.set_user_filter(_pref_key(user), chosen_mode, chosen):
            return _exclusions_panel(
                request, user,
                "Could not save that. The details are in the server log.",
                mode=chosen_mode)
        response = _exclusions_panel(request, user)
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    # Already in the overview div's hx-trigger list, so the funnel re-fetches
    # with no change to that element.
    response.headers["HX-Trigger"] = "pulseRefresh"
    return response


@router.delete("/api/outbound-pulse/exclusions")
async def reset_exclusions(request: Request,
                           user: dict = Depends(auth.require_login)):
    """Drop the saved preference so the house default applies again.

    Distinct from saving an empty list, which means "show me everything" and
    stays that way.
    """
    try:
        store.clear_user_filter(_pref_key(user))
        response = _exclusions_panel(request, user)
    except PulseNotReady as exc:
        return _not_ready_box(exc)
    response.headers["HX-Trigger"] = "pulseRefresh"
    return response


# ── Copy A/B ─────────────────────────────────────────────────

# Fetched per campaign and never during a sync: this is one HTTP call each,
# and the account has over a thousand campaigns. Only the ones with activity in
# the range are asked for, and only when someone opens the panel.
_AB_MAX_CAMPAIGNS = 8


def manual_campaigns(client_id: str) -> list[dict]:
    """This client's DNC & Merger campaigns, newest first.

    Read straight from the mail-merge `campaigns` table rather than through
    Pulse. Manual campaigns are that tool's records; Pulse only ever reads
    their aggregates, and pulse_manual_daily deliberately carries a NULL
    campaign_id because dnc_entries records a client rather than a campaign.
    So they cannot appear in the funnel's campaign table — they get their own.

    Never raises: this is a supporting list, not the page.
    """
    from app.utils.supabase import SUPABASE_URL, sb_headers

    import requests as http

    if not SUPABASE_URL:
        return []
    try:
        r = http.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            params={
                "select":    ("id,campaign_name,sender_profile_name,sheet_url,"
                              "total_prospects,sent_count,created_at"),
                "client_id": f"eq.{client_id}",
                "order":     "created_at.desc",
                "limit":     "50",
            },
            headers=sb_headers(),
            timeout=15,
        )
        if r.status_code >= 400:
            return []
        return r.json()
    except Exception as exc:
        logger.warning("Pulse: could not list manual campaigns for %s: %s",
                       client_id, exc)
        return []


def _manual_ab_campaigns(client_id: str) -> list[dict]:
    """This client's manual campaigns, as the A/B breakdown needs them.

    Read straight from the mail-merge `campaigns` table rather than through
    Pulse: manual campaigns are that tool's records, and Pulse only ever reads
    their aggregates.

    `ref` is the campaign row's id rather than its name. Copy written down
    against a campaign is keyed by it, and a renamed campaign must not lose
    the copy somebody typed for it.
    """
    from app.utils.supabase import SUPABASE_URL, sb_headers

    import requests as http

    if not SUPABASE_URL:
        return []
    try:
        r = http.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            params={
                "select":    ("id,sheet_id,campaign_name,completed_at,"
                              "copy_territory,copy_industry"),
                "client_id": f"eq.{client_id}",
                "order":     "completed_at.desc.nullslast",
                "limit":     str(_AB_MAX_CAMPAIGNS),
            },
            headers=sb_headers(),
            timeout=15,
        )
        if r.status_code >= 400:
            # The copy-source columns may not be migrated here. Fall back to
            # the ids alone: the breakdown is still worth showing without the
            # copy, and losing the whole panel over it would not be.
            return _manual_ab_campaigns_without_copy(client_id)
        rows = r.json()
    except Exception as exc:
        logger.warning("Pulse A/B: could not list campaigns for %s: %s", client_id, exc)
        return []
    return [{
        "ref":       str(row.get("id") or row.get("sheet_id") or ""),
        "name":      str(row.get("campaign_name") or ""),
        "sheet_id":  str(row.get("sheet_id") or ""),
        "territory": str(row.get("copy_territory") or ""),
        "industry":  str(row.get("copy_industry") or ""),
    } for row in rows if row.get("sheet_id")]


def _manual_ab_campaigns_without_copy(client_id: str) -> list[dict]:
    from app.utils.supabase import SUPABASE_URL, sb_headers

    import requests as http

    try:
        r = http.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            params={
                "select":    "id,sheet_id,campaign_name",
                "client_id": f"eq.{client_id}",
                "limit":     str(_AB_MAX_CAMPAIGNS),
            },
            headers=sb_headers(),
            timeout=15,
        )
        if r.status_code >= 400:
            return []
        rows = r.json()
    except Exception:
        return []
    return [{
        "ref":       str(row.get("id") or row.get("sheet_id") or ""),
        "name":      str(row.get("campaign_name") or ""),
        "sheet_id":  str(row.get("sheet_id") or ""),
        "territory": "", "industry": "",
    } for row in rows if row.get("sheet_id")]


def _copy_index(client_id: str) -> dict[tuple[str, str, str], dict]:
    """Written-down copy, keyed the way a variant identifies itself.

    An empty campaign_ref matches any campaign — copy recorded before the
    per-campaign breakdown existed, or deliberately recorded once for every
    campaign that reused it — so a lookup falls back to it rather than
    demanding the same text be typed against each campaign.
    """
    index: dict[tuple[str, str, str], dict] = {}
    for row in store.list_copy_variants(client_id):
        index[(str(row.get("source_tool") or ""),
               str(row.get("campaign_ref") or ""),
               str(row.get("variant_key") or ""))] = row
    return index


def _copy_for(index: dict, source: str, campaign: str, variant: str) -> dict | None:
    return (index.get((source, campaign, variant))
            or index.get((source, "", variant)))


def _copy_bank_lookup(client_id: str):
    """Resolve manual variant labels against the Copy Bank, entry by entry.

    A manual variant is an index — "S1/B1" is subject 1 with body 1 — so the
    text is already written down and nobody should have to type it again. The
    merge records which entry it pulled, in campaigns.copy_territory /
    copy_industry, which is the piece that makes the index resolvable.

    Returns resolve(territory, industry, label). Each distinct entry is
    fetched once however many labels are looked up against it; an entry with
    no territory or industry recorded resolves to nothing, because a campaign
    run without choosing a copy source has nothing to resolve against.
    """
    cache: dict[tuple[str, str], dict | None] = {}

    def entry(territory: str, industry: str) -> dict | None:
        key = (territory, industry)
        if key not in cache:
            content, found = shared_copy_bank.fetch_content(
                client_id, territory, industry)
            cache[key] = shared_copy_bank.extract(content) if found else None
        return cache[key]

    def resolve(territory: str, industry: str, label: str) -> dict | None:
        if not territory or not industry:
            return None
        copy = entry(territory, industry)
        if copy is None:
            return None
        hit = shared_copy_bank.variant_in(copy, label)
        if hit is None:
            return None
        return {
            "subject": hit["subject"],
            # Copy Bank stores a body as plain text and renders it with
            # textContent, so it has never been escaped — and this body goes
            # on to a client-facing report.
            "body":    richtext.from_text(hit["body"]),
            "from_copy_bank": True,
            "source_label":   f"{territory} · {industry}",
        }

    return resolve


def _attach_manual_copy(client_id: str, breakdown: dict, copy_index: dict) -> str:
    """Put the copy against every manual variant. Returns a note for the view.

    PER CAMPAIGN THERE IS NO AMBIGUITY. A campaign pulled exactly one Copy
    Bank entry, so its labels resolve against that entry and nothing else.
    That is the whole reason the breakdown is worth having: combined across
    campaigns, "S1/B1" can be two different messages and the sheet cannot say
    which one a prospect got, so the combined block only resolves when every
    campaign drew on the same entry.

    Copy typed by hand always wins. It is somebody correcting what the index
    resolves to, and `_copy_for` falls back from this campaign's own entry to
    one recorded against every campaign.
    """
    resolve = _copy_bank_lookup(client_id)

    for block in breakdown["campaigns"]:
        for variant in block["variants"]:
            variant["variant_copy"] = (
                _copy_for(copy_index, SOURCE_MANUAL, block["ref"], variant["variant"])
                or resolve(block["territory"], block["industry"], variant["variant"]))

    combined = breakdown["combined"]
    territory, industry = (combined["copy_sources"] or [("", "")])[0]
    for variant in combined["variants"]:
        variant["variant_copy"] = (
            _copy_for(copy_index, SOURCE_MANUAL, "", variant["variant"])
            or (resolve(territory, industry, variant["variant"])
                if combined["shared_copy"] else None))

    if len(combined["copy_sources"]) > 1:
        return (f"These campaigns pulled {len(combined['copy_sources'])} different "
                "Copy Bank entries, so a label like S1/B1 is not the same message "
                "in each. The per-campaign results above are the ones to read.")
    return ""


def _ab_panel(request: Request, client_id: str, rng: dict, error: str = ""):
    """Winning copy variation, per source.

    Deliberately two separate verdicts. Smartlead decides on its own positive
    replies; a manual campaign decides on the team's Lead and Interested
    statuses. Adding them together would produce a number that means nothing.
    """
    start = rng["from"] or (today_utc() - timedelta(days=29))
    end = rng["to"] or today_utc()
    copy_index = _copy_index(client_id)

    smartlead_steps, smartlead_error = [], ""
    campaigns = store.list_campaigns(client_id=client_id,
                                     source_tool=SOURCE_SMARTLEAD)

    if not smartlead.is_configured():
        smartlead_error = "No Smartlead API key is set, so variants cannot be read."
    else:
        for campaign in campaigns[:_AB_MAX_CAMPAIGNS]:
            external = str(campaign.get("external_campaign_id") or "")
            if not external:
                continue
            try:
                steps = abtest.smartlead_variants(external, start, end)
            except Exception as exc:
                # One campaign failing must not hide the rest.
                logger.warning("Pulse A/B: %s failed: %s", external, exc)
                smartlead_error = smartlead_error or (
                    "Some campaigns could not be read from Smartlead.")
                continue
            for step in steps:
                step["campaign"] = campaign.get("name") or external
                for variant in step["variants"]:
                    variant["variant_copy"] = _copy_for(
                        copy_index, SOURCE_SMARTLEAD, step["campaign"],
                        variant["variant"])
                smartlead_steps.append(step)

    manual = abtest.manual_breakdown(_manual_ab_campaigns(client_id))
    manual["copy_note"] = _attach_manual_copy(client_id, manual, copy_index)

    # Offered so a campaign whose copy source was never recorded can have one
    # set here instead of staying unresolvable forever. Only worth a read when
    # there is actually a campaign missing one.
    copy_options = (
        shared_copy_bank.profile_options(client_id)
        if any(not (b.get("territory") and b.get("industry"))
               for b in manual["campaigns"])
        else {"territories": [], "industries": []}
    )

    return templates.TemplateResponse("partials/pulse_ab.html", {
        "request":          request,
        "client_id":        client_id,
        "smartlead_steps":  smartlead_steps,
        "smartlead_error":  smartlead_error,
        "smartlead_more":   max(0, len(campaigns) - _AB_MAX_CAMPAIGNS),
        "manual":           manual,
        "copy_options":     copy_options,
        "range":            rng,
        "range_key":        rng["preset"],
        "error":            error,
    })


@router.get("/api/outbound-pulse/clients/{client_id}/ab")
async def client_ab_tests(
    request:   Request,
    client_id: str,
    range:     str = Query("30d"),
    date_from: str = Query(""),
    date_to:   str = Query(""),
):
    try:
        return _ab_panel(request, client_id,
                         _resolve_range(range, date_from, date_to))
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.post("/api/outbound-pulse/clients/{client_id}/copy-source")
async def set_campaign_copy_source(
    request:      Request,
    client_id:    str,
    campaign_ref: str = Form(""),
    territory:    str = Form(""),
    industry:     str = Form(""),
    range:        str = Form("30d"),
    user:         dict = Depends(auth.require_login),
):
    """Record which Copy Bank entry a manual campaign was built from.

    The merge is supposed to store this when the sheet is created, and only
    does so when somebody used the Copy Bank loader in that same session — so
    every campaign that predates the capture, and plenty that do not, resolve
    to nothing and the panel can only say "no copy source recorded". The text
    is already written down; what is missing is the pointer to it. This is
    that pointer, set by hand, once per campaign.

    Writes to the mail-merge `campaigns` row rather than anywhere in Pulse,
    because that is where the merge writes it and a second home would drift.
    """
    rng = _resolve_range(range, "", "")
    try:
        if not campaign_ref:
            return _ab_panel(request, client_id, rng,
                             error="That campaign could not be identified.")
        if not territory or not industry:
            return _ab_panel(request, client_id, rng,
                             error="Choose both a territory and an industry — "
                                   "a Copy Bank entry is the pair.")
        if not _save_copy_source(client_id, campaign_ref, territory, industry):
            return _ab_panel(request, client_id, rng,
                             error="Could not record that copy source. The "
                                   "details are in the server log.")
        return _ab_panel(request, client_id, rng)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


def _save_copy_source(client_id: str, campaign_ref: str,
                      territory: str, industry: str) -> bool:
    """PATCH the campaigns row, scoped to this client.

    Scoped on client_id as well as id: the campaign ref arrives from the
    browser, and a copy source belongs to the client whose Copy Bank it names.
    """
    from app.utils.supabase import SUPABASE_URL, sb_headers

    import requests as http

    if not SUPABASE_URL:
        return False
    try:
        r = http.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            params={"id": f"eq.{campaign_ref}", "client_id": f"eq.{client_id}"},
            headers={**sb_headers(), "Prefer": "return=representation"},
            json={"copy_territory": territory, "copy_industry": industry},
            timeout=15,
        )
    except Exception as exc:
        logger.warning("Pulse A/B: could not set copy source on %s: %s",
                       campaign_ref, exc)
        return False
    if r.status_code >= 400:
        logger.warning("Pulse A/B: copy source rejected for %s: %s %s",
                       campaign_ref, r.status_code, r.text[:200])
        return False
    # An empty body means the filter matched nothing — a campaign id that is
    # not this client's, or no longer exists. Reporting that as success would
    # leave the panel unchanged with no explanation.
    try:
        return bool(r.json())
    except Exception:
        return False


@router.post("/api/outbound-pulse/clients/{client_id}/copy")
async def save_variant_copy(
    request:      Request,
    client_id:    str,
    source_tool:  str = Form(""),
    variant_key:  str = Form(""),
    campaign_ref: str = Form(""),
    subject:      str = Form(""),
    body:         str = Form(""),
    range:        str = Form("30d"),
    user:         dict = Depends(auth.require_login),
):
    """Write down what a variant actually said.

    Smartlead reports a subject per sequence STEP rather than per variant, and
    a manual campaign's variant is a Copy Bank index like "S2/B1" — so a winner
    is often known without its text. This is how that gets filled in, and it is
    what lets the client's report show the message rather than a letter.
    """
    rng = _resolve_range(range, "", "")
    try:
        if not variant_key:
            return _ab_panel(request, client_id, rng,
                             error="That variant could not be identified.")
        clean_subject = subject.strip()[:300]
        # Sanitised here, at the one point markup crosses from an author to a
        # reader: this body renders on a client-facing report.
        clean_body = richtext.sanitize(body)
        if not clean_subject and not richtext.to_text(clean_body):
            return _ab_panel(request, client_id, rng,
                             error="Add a subject line or some copy.")
        if not store.upsert_copy_variant(
                client_id=client_id,
                source_tool=source_tool or SOURCE_SMARTLEAD,
                variant_key=variant_key,
                campaign_ref=campaign_ref,
                subject=clean_subject,
                body=clean_body,
                entered_by=user.get("name", "") or user.get("email", "")):
            return _ab_panel(request, client_id, rng,
                             error="Could not save that copy. The details are "
                                   "in the server log.")
        return _ab_panel(request, client_id, rng)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


# ── The ICP Performance Tracker ────────────────────────────────

def _tracker_link_panel(request: Request, client_id: str, error: str = "",
                        notice: str = ""):
    """The link form and its status. Cheap — one Postgres read, no sheet."""
    tracker_row = store.get_tracker(client_id)
    return templates.TemplateResponse("partials/pulse_tracker_link.html", {
        "request":   request,
        "client_id": client_id,
        "tracker":   tracker_row,
        "error":     error,
        "notice":    notice,
    })


@router.get("/api/outbound-pulse/clients/{client_id}/tracker")
async def tracker_link(request: Request, client_id: str):
    try:
        return _tracker_link_panel(request, client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.post("/api/outbound-pulse/clients/{client_id}/tracker")
async def link_tracker(
    request:   Request,
    client_id: str,
    sheet_url: str = Form(""),
    tab_title: str = Form(""),
    user:      dict = Depends(auth.require_login),
):
    """Point this client at their Performance Tracker sheet.

    The sheet is READ BEFORE IT IS STORED. extract_sheet_id falls through to
    returning whatever it was given, so a pasted Notion link becomes a
    "sheet id" that 404s on every later read with no clue why. Reading the
    header row first costs one request and turns "the panel is broken" into
    a message naming the actual problem, at the moment it can be fixed.
    """
    pasted = (sheet_url or "").strip()
    tab = (tab_title or "").strip() or tracker.DEFAULT_TAB
    try:
        if not pasted:
            return _tracker_link_panel(request, client_id,
                                       "Paste the Performance Tracker sheet's link.")
        if tab.lower() in tracker.FORBIDDEN_TABS:
            # Locked, different schema, and not ours. Refuse by name rather
            # than letting it fail later as an unreadable column layout.
            return _tracker_link_panel(
                request, client_id,
                "“Raw Leads” is locked and has a different layout. "
                "Point this at the “All leads” tab.")

        sheet_id = google_sheets.extract_sheet_id(pasted)
        if not _looks_like_a_sheet_id(sheet_id):
            return _tracker_link_panel(
                request, client_id,
                "That does not look like a Google Sheets link. Copy the URL "
                "from the sheet's address bar.")
        if not google_sheets.is_configured():
            return _tracker_link_panel(
                request, client_id,
                "Google Sheets is not configured on this server, so the link "
                "cannot be checked. Ask an administrator to set "
                "GOOGLE_SHEETS_SA_JSON.")

        try:
            values = google_sheets.read_tracker_values(sheet_id, tab)
            _index, headers = tracker.find_header(values)
        except google_sheets.TrackerUnavailable as exc:
            return _tracker_link_panel(request, client_id, str(exc))
        except Exception as exc:
            logger.warning("Pulse: tracker probe failed for %s: %s", client_id, exc)
            return _tracker_link_panel(
                request, client_id,
                "That sheet could not be read. The details are in the server log.")

        columns, warnings = tracker.resolve_columns(headers)
        if "date" not in columns:
            # Without a date nothing can be tied to a report period, and a
            # period figure built from undated rows would be fiction.
            return _tracker_link_panel(
                request, client_id,
                f"The “{tab}” tab has no date column, so leads cannot be "
                "matched to a reporting period. Expected a column ending in "
                "“Date/Time”.")

        if not store.set_tracker(client_id=client_id, sheet_id=sheet_id,
                                 sheet_url=pasted, tab_title=tab,
                                 linked_by=user.get("name", "") or user.get("email", "")):
            return _tracker_link_panel(
                request, client_id,
                "Could not save that link. The details are in the server log.")

        notice = "Sheet linked."
        if warnings:
            notice += " " + " ".join(warnings)
        return _tracker_link_panel(request, client_id, notice=notice)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.delete("/api/outbound-pulse/clients/{client_id}/tracker")
async def unlink_tracker(request: Request, client_id: str):
    try:
        store.clear_tracker(client_id)
        return _tracker_link_panel(request, client_id, notice="Sheet unlinked.")
    except PulseNotReady as exc:
        return _not_ready_box(exc)


def _looks_like_a_sheet_id(value: str) -> bool:
    """Google keys are a long run of url-safe characters and nothing else.

    extract_sheet_id returns its input when the URL does not match, so this is
    what stops a pasted Notion page being stored as a spreadsheet key.
    """
    return bool(re.fullmatch(r"[A-Za-z0-9_\-]{20,}", value or ""))


def _tracker_panel(request: Request, client_id: str, rng: dict):
    """The lead-quality breakdown. One sheet read, then everything else is
    filtered in the browser from the rows this ships with it."""
    tracker_row = store.get_tracker(client_id)
    context = {
        "request":   request,
        "client_id": client_id,
        "tracker":   tracker_row,
        "range":     rng,
        "rows":      [],
        "summary":   None,
        "meta":      {},
        "error":     "",
        "read_at":   datetime.now(timezone.utc).strftime("%H:%M"),
    }
    if not tracker_row:
        context["error"] = "No Performance Tracker sheet is linked for this client."
        return templates.TemplateResponse("partials/pulse_tracker.html", context)

    sheet_id = str(tracker_row.get("sheet_id") or "")
    tab = str(tracker_row.get("tab_title") or tracker.DEFAULT_TAB)
    try:
        values = google_sheets.read_tracker_values(sheet_id, tab)
    except google_sheets.TrackerUnavailable as exc:
        store.set_tracker_status(client_id, str(exc))
        context["error"] = str(exc)
        return templates.TemplateResponse("partials/pulse_tracker.html", context)
    except Exception as exc:
        logger.warning("Pulse: tracker read failed for %s: %s", client_id, exc)
        store.set_tracker_status(client_id, "The sheet could not be read.")
        context["error"] = ("That sheet could not be read. The details are in "
                            "the server log.")
        return templates.TemplateResponse("partials/pulse_tracker.html", context)

    rows, meta = tracker.parse_rows(values)
    store.set_tracker_status(client_id, "")

    # The initial view matches the range the page is showing, so the panel
    # agrees with the funnel above it. Every row ships regardless, so widening
    # to all time is a click in the browser rather than another sheet read.
    scoped = tracker.in_period(rows, rng.get("from"), rng.get("to"))
    context.update({
        "rows":     rows,
        "summary":  tracker.summarise(scoped),
        "meta":     meta,
        # Orderings and labels travel WITH the rows, so the browser's filtered
        # recount hardcodes no business rule and cannot drift from Python's.
        "config":   {
            "seniority": [[k, tracker.SENIORITY_LABELS[k]]
                          for k in tracker.SENIORITY_ORDER],
            "bands":     [[k, l] for k, l, _lo, _hi in tracker.ICP_BANDS]
                         + [list(tracker.UNRATED_BAND)],
            "top_n":     tracker.TOP_N,
            "top_rated_from": tracker.TOP_RATED_FROM,
        },
        "in_range": len(scoped),
        "from":     rng["from"].isoformat() if rng.get("from") else "",
        "to":       rng["to"].isoformat() if rng.get("to") else "",
    })
    return templates.TemplateResponse("partials/pulse_tracker.html", context)


@router.get("/api/outbound-pulse/clients/{client_id}/tracker/leads")
async def tracker_leads(
    request:   Request,
    client_id: str,
    range:     str = Query("30d"),
    date_from: str = Query(""),
    date_to:   str = Query(""),
):
    try:
        return _tracker_panel(request, client_id,
                              _resolve_range(range, date_from, date_to))
    except PulseNotReady as exc:
        return _not_ready_box(exc)


# ── Client reports ─────────────────────────────────────────────


def report_snapshot(client_id: str, start: date, end: date) -> dict:
    """The figures a report freezes at publish.

    Everything the client-facing page renders, resolved once. A published
    report reads only from this — it never re-queries — so that the numbers a
    client was sent are the numbers they see when they open it again.
    """
    counts = store.funnel(client_id=client_id, date_from=start, date_to=end)
    by_source = store.funnel_by_source(client_id=client_id, date_from=start, date_to=end)
    return {
        "version":    1,
        "counts":     counts,
        "by_channel": store.funnel_by_channel(
                          client_id=client_id, date_from=start, date_to=end),
        "by_source":  by_source,
        "trend":      bucket_timeseries(store.funnel_timeseries(
                          client_id=client_id, date_from=start, date_to=end)),
        # Frozen with everything else. Reading it live on the portal would put
        # a figure that moves next to a column of figures that cannot.
        "ab":         _report_ab_winners(client_id, start, end),
        "icp":        _report_icp(client_id, start, end),
        "taken_at":   datetime.now(timezone.utc).isoformat(),
    }


def _report_icp(client_id: str, start: date, end: date) -> dict:
    """Lead quality for the report's period, frozen at publish.

    Never raises, for the same reason _report_ab_winners does not: this reaches
    out to Google, and a slow or revoked sheet is not worth failing a publish
    over. A missing section is a missing section; a publish that dies at the
    last step loses the account manager's write-up with it.

    Returns {} — which renders nothing at all — when no sheet is linked, when
    the sheet cannot be read, when it has no date column (a period figure built
    from undated rows would be fiction), and when the period contains no leads.

    CARRIES ALL-TIME ALONGSIDE THE PERIOD. A month of outbound is a handful of
    graded leads, so a monthly average on its own is noise. The sheet's own
    summary makes the same comparison, which is a fair sign of how it gets read.
    """
    try:
        linked = store.get_tracker(client_id)
        if not linked:
            return {}
        values = google_sheets.read_tracker_values(
            str(linked.get("sheet_id") or ""),
            str(linked.get("tab_title") or tracker.DEFAULT_TAB))
        rows, meta = tracker.parse_rows(values)
        if not meta.get("has_dates") or not rows:
            return {}

        scoped = tracker.in_period(rows, start, end)
        if not scoped:
            return {}

        summary = tracker.summarise(scoped)
        everything = tracker.summarise(rows)
        summary.update({
            "version": 1,
            "top_rated_from": tracker.TOP_RATED_FROM,
            # Frozen with the figures, so the report links to the sheet those
            # figures came from rather than to wherever the client is pointed
            # later. The client opens it to grade leads we have not graded.
            "sheet_url": str(linked.get("sheet_url") or ""),
            "all_time": {
                "total":   everything["total"],
                "rated":   everything["rated"],
                "average": everything["average"],
            },
        })
        # The ranked job titles stay internal: a client's report wants the
        # shape of who we reached, not a list of individual prospects.
        summary.pop("roles", None)
        return summary
    except Exception as exc:
        logger.warning("Pulse: lead quality unavailable for the report on %s: %s",
                       client_id, exc)
        return {}


def _report_ab_winners(client_id: str, start: date, end: date) -> list[dict]:
    """Winning copy for the report, or an empty list.

    Never raises. This reaches out to Smartlead and to Google Sheets, both of
    which can be slow or down, and neither is worth failing a publish over —
    the report's actual numbers come from our own database. A missing A/B
    section is a missing section; a publish that dies at the last step loses
    the write-up with it.
    """
    steps: list[dict] = []
    copy_index = _copy_index(client_id)
    try:
        if smartlead.is_configured():
            for campaign in store.list_campaigns(
                    client_id=client_id,
                    source_tool=SOURCE_SMARTLEAD)[:_AB_MAX_CAMPAIGNS]:
                external = str(campaign.get("external_campaign_id") or "")
                if not external:
                    continue
                for step in abtest.smartlead_variants(external, start, end):
                    step["campaign"] = campaign.get("name") or external
                    for variant in step["variants"]:
                        variant["variant_copy"] = _copy_for(
                            copy_index, SOURCE_SMARTLEAD, step["campaign"],
                            variant["variant"])
                    steps.append(step)
    except Exception as exc:
        logger.warning("Pulse: A/B unavailable for the report on %s: %s",
                       client_id, exc)
    try:
        breakdown = abtest.manual_breakdown(_manual_ab_campaigns(client_id))
        _attach_manual_copy(client_id, breakdown, copy_index)
        # Per campaign, not combined. A campaign is one send of one Copy Bank
        # entry, so its winner is a verdict the client can act on; the
        # combined figure can span two different messages under one label.
        manual = breakdown["campaigns"]
    except Exception as exc:
        logger.warning("Pulse: manual A/B unavailable for %s: %s", client_id, exc)
        manual = []
    return abtest.winners_for_report(steps, manual)


def _default_period() -> tuple[date, date]:
    """Last whole calendar month — the cycle these reports follow."""
    today = today_utc()
    end = today.replace(day=1) - timedelta(days=1)
    return end.replace(day=1), end


# Typed, not clicked. Short enough to type without copying, specific enough
# that it cannot be the result of hitting return on a focused field.
_CLEAR_REPORTS_WORD = "DELETE"


def _reports_panel(request: Request, client_id: str, error: str = "",
                   editing: str = "", notice: str = "", clearing: bool = False):
    reports = store.list_reports(client_id)
    start, end = _default_period()
    taken = {(str(r.get("period_start")), str(r.get("period_end"))) for r in reports}
    # Offer a period that is not already taken, so the obvious first click does
    # not hit the one-report-per-period rule.
    while (start.isoformat(), end.isoformat()) in taken:
        end = start - timedelta(days=1)
        start = end.replace(day=1)
    for report in reports:
        report["heading"] = report_heading(report)
        report["summary"] = richtext.to_text(report.get("body") or "")[:180]
    return templates.TemplateResponse("partials/pulse_reports.html", {
        "request":     request,
        "reports":     reports,
        "client_id":   client_id,
        "error":       error,
        "notice":      notice,
        "editing":     editing,
        # Kept open across a refused confirmation, so a mistyped word does not
        # also collapse the form it was typed into.
        "clearing":    clearing,
        "clear_word":  _CLEAR_REPORTS_WORD,
        "next_start":  start.isoformat(),
        "next_end":    end.isoformat(),
    })


@router.get("/api/outbound-pulse/clients/{client_id}/reports")
async def list_client_reports(request: Request, client_id: str):
    try:
        return _reports_panel(request, client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.post("/api/outbound-pulse/clients/{client_id}/reports")
async def create_client_report(
    request:      Request,
    client_id:    str,
    period_start: str = Form(""),
    period_end:   str = Form(""),
    title:        str = Form(""),
    user:         dict = Depends(auth.require_login),
):
    start = _parse_date(period_start)
    end = _parse_date(period_end)
    try:
        if start is None or end is None:
            return _reports_panel(request, client_id, "Pick a start and an end date.")
        if end < start:
            start, end = end, start
        created = store.create_report(
            client_id,
            period_start=start.isoformat(),
            period_end=end.isoformat(),
            title=title.strip()[:120],
            created_by=user.get("name", "") or user.get("email", ""),
        )
        if not created:
            # The unique index is the likely cause and the only one worth
            # naming: everything else is already in the log.
            return _reports_panel(
                request, client_id,
                "Could not create that report. There may already be one for "
                "that exact period.")
        return _reports_panel(request, client_id, editing=str(created["id"]))
    except PulseNotReady as exc:
        return _not_ready_box(exc)


# Registered BEFORE the /reports/{report_id} routes. FastAPI matches in
# definition order, so declared after them "clear" binds as a report id
# and this never runs.
@router.post("/api/outbound-pulse/clients/{client_id}/reports/clear")
async def clear_client_reports(
    request:   Request,
    client_id: str,
    confirm:   str = Form(""),
    user:      dict = Depends(auth.require_login),
):
    """Wipe this client's whole report history. Nothing comes back.

    Guarded by a typed word rather than a confirm dialog. The capability
    already exists one row at a time, so this is not about who may do it — it
    is about not doing it by reflex on a live client, where the thing lost is
    every published report they can currently open.

    Logged with the name of whoever did it, because the client noticing their
    history is gone is a question somebody has to be able to answer.
    """
    try:
        if confirm.strip().upper() != _CLEAR_REPORTS_WORD:
            return _reports_panel(
                request, client_id,
                f"Type {_CLEAR_REPORTS_WORD} to confirm. Nothing has been deleted.",
                clearing=True)
        removed = store.delete_all_reports(client_id)
        if removed is None:
            return _reports_panel(
                request, client_id,
                "Could not clear the reports. The details are in the server log.")
        logger.warning("Pulse: %s cleared %d report(s) for client %s",
                       user.get("email") or user.get("name") or "someone",
                       removed, client_id)
        if not removed:
            # Somebody else got there first, most likely. Saying the link now
            # has no history would be claiming credit for a delete that did
            # nothing, and the reader cannot tell the two apart otherwise.
            return _reports_panel(request, client_id,
                                  notice="There was nothing left to clear.")
        return _reports_panel(
            request, client_id,
            notice=(f"Cleared {removed} report{'' if removed == 1 else 's'}. "
                    "The client's link now has no history on it."))
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.post("/api/outbound-pulse/clients/{client_id}/reports/{report_id}")
async def save_client_report(
    request:   Request,
    client_id: str,
    report_id: str,
    title:     str = Form(""),
    body:      str = Form(""),
):
    """Save the write-up. Does not publish — a saved draft stays invisible to
    the client until it is published explicitly."""
    try:
        saved = store.update_report(report_id, {
            "title": title.strip()[:120],
            # Sanitised here, at the one point markup crosses from an author to
            # a reader, rather than trusted on the way out to a client's page.
            "body":  richtext.sanitize(body),
        })
        if not saved:
            return _reports_panel(
                request, client_id,
                "Could not save that write-up. The details are in the server log.")
        return _reports_panel(request, client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.post("/api/outbound-pulse/clients/{client_id}/reports/{report_id}/publish")
async def publish_client_report(
    request:   Request,
    client_id: str,
    report_id: str,
    title:     str = Form(""),
    body:      str = Form(""),
    # Sent by the editor form and by nothing else. FastAPI collapses an empty
    # form value to None, so "" and "not sent" are indistinguishable from the
    # values alone — and the two mean opposite things here: clear the write-up,
    # or leave the saved one exactly as it is.
    from_editor: str = Form(""),
):
    """Save the write-up, freeze the figures, and put it on the client's link.

    One request, in that order. Publishing used to only snapshot, so whether
    the write-up made it depended on a separate save request that had no
    ordering against this one.

    The write-up is only touched when the request came from the editor. The
    Publish button on a collapsed row carries no editor content, and treating
    that as an empty write-up would publish the report with its text erased.

    The snapshot is taken now, not when the draft was opened, and re-publishing
    retakes it — that is how a report is corrected after a late sync.
    """
    try:
        report = store.get_report(report_id)
        if report is None:
            return _reports_panel(request, client_id, "That report no longer exists.")
        start = _parse_date(str(report.get("period_start") or ""))
        end = _parse_date(str(report.get("period_end") or ""))
        if start is None or end is None:
            return _reports_panel(request, client_id, "That report has no period.")
        patch = {
            "snapshot":     report_snapshot(client_id, start, end),
            "status":       "published",
            "published_at": datetime.now(timezone.utc).isoformat(),
        }
        if from_editor:
            patch["title"] = title.strip()[:120]
            patch["body"] = richtext.sanitize(body)
        if not store.update_report(report_id, patch):
            # Swallowing this is what made a broken publish look like a
            # working one that changed nothing.
            return _reports_panel(
                request, client_id,
                "Could not publish that report. The details are in the server log.")
        return _reports_panel(request, client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.post("/api/outbound-pulse/clients/{client_id}/reports/{report_id}/unpublish")
async def unpublish_client_report(request: Request, client_id: str, report_id: str):
    """Pull a report back off the client's link. The snapshot is kept, so
    re-publishing without editing puts back exactly what was there."""
    try:
        store.update_report(report_id, {"status": "draft", "published_at": None})
        return _reports_panel(request, client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.delete("/api/outbound-pulse/clients/{client_id}/reports/{report_id}")
async def remove_client_report(request: Request, client_id: str, report_id: str):
    try:
        store.delete_report(report_id)
        return _reports_panel(request, client_id)
    except PulseNotReady as exc:
        return _not_ready_box(exc)


@router.get("/api/outbound-pulse/clients")
async def clients_json():
    """Client list as JSON — used by the import form's client picker."""
    try:
        return JSONResponse(store.list_clients())
    except PulseNotReady as exc:
        return JSONResponse({"error": str(exc)}, status_code=503)
