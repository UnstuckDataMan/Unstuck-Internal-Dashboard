"""
Campaigns router — tracks mail-merge campaigns linked to Google Sheets.
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from datetime import date as _date, datetime as _datetime, timedelta, timezone as _tz

import requests as http_req
from fastapi import APIRouter, Depends, Query, Request
from fastapi.responses import JSONResponse, HTMLResponse

from app import auth
from app.deps import templates
from app.utils.supabase import (
    SUPABASE_URL,
    sb_headers as _sb_headers,
    sb_configured as _sb_configured,
    parse_total as _parse_total,
)

router = APIRouter()

CHUNK_SIZE = 500


def _pg_in_list(values: list[str]) -> str:
    """Build a PostgREST in.(...) literal with each value double-quoted.

    Unquoted values break on commas/parens; double-quoting (with backslash
    escapes) makes campaign names with any punctuation filter correctly.
    """
    quoted = ",".join(
        '"' + v.replace("\\", "\\\\").replace('"', '\\"') + '"' for v in values
    )
    return f"({quoted})"


def _count_contacted(
    client_id: str,
    date_eq:  str = "",
    date_gte: str = "",
    date_lte: str = "",
    campaign_names: list[str] | None = None,
) -> int:
    """
    Count contacted_prospects rows that represent actual sends (source = auto_sync),
    optionally filtered by date and/or a list of campaign names.

    Only rows with source = 'auto_sync' are counted.  This excludes:
      • scrub_upload  — prospects saved when a list is scrubbed
      • manual        — manually added contacted entries
      • csv_upload    — bulk-uploaded contacted lists
      • dnc_manual    — auto-logged contacts from DNC entries

    Scrub uploads deliberately share the same table so the "recently contacted"
    removal feature in the scrubber still works; they must not appear in the
    campaign send stats or they would show 804 sends the moment a list is scrubbed.
    """
    base_params: list[tuple[str, str]] = [
        ("select",    "id"),
        ("client_id", f"eq.{client_id}"),
        ("source",    "eq.auto_sync"),   # ← only count genuine sheet sends
        ("limit",     "1"),
    ]
    if campaign_names:
        base_params.append(("campaign_name", f"in.{_pg_in_list(campaign_names)}"))
    # (no fallback filter needed — source=auto_sync already excludes non-send rows)
    if date_eq:
        base_params.append(("contacted_at", f"eq.{date_eq}"))
    if date_gte:
        base_params.append(("contacted_at", f"gte.{date_gte}"))
    if date_lte:
        base_params.append(("contacted_at", f"lte.{date_lte}"))
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/contacted_prospects",
            headers={**_sb_headers(), "Prefer": "count=exact"},
            params=base_params,
            timeout=10,
        )
        return _parse_total(r.headers)
    except Exception:
        return 0


def _count_chaser_contacted(
    client_id: str,
    date_eq:  str = "",
    date_gte: str = "",
    date_lte: str = "",
    campaign_names: list[str] | None = None,
) -> int:
    """
    Count contacted_prospects rows by chaser_contacted_at date — i.e. the number
    of chaser emails sent in the given period.

    Returns 0 gracefully if the chaser_contacted_at column does not yet exist
    (Supabase returns HTTP 400 for unknown column names).
    """
    base_params: list[tuple[str, str]] = [
        ("select",              "id"),
        ("client_id",           f"eq.{client_id}"),
        ("source",              "eq.auto_sync"),
        ("chaser_contacted_at", "not.is.null"),   # only rows that have been chased
        ("limit",               "1"),
    ]
    if campaign_names:
        base_params.append(("campaign_name", f"in.{_pg_in_list(campaign_names)}"))
    if date_eq:
        base_params.append(("chaser_contacted_at", f"eq.{date_eq}"))
    if date_gte:
        base_params.append(("chaser_contacted_at", f"gte.{date_gte}"))
    if date_lte:
        base_params.append(("chaser_contacted_at", f"lte.{date_lte}"))
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/contacted_prospects",
            headers={**_sb_headers(), "Prefer": "count=exact"},
            params=base_params,
            timeout=10,
        )
        if r.status_code >= 400:
            return 0   # column doesn't exist yet — graceful zero
        return _parse_total(r.headers)
    except Exception:
        return 0


def _time_breakdown(
    client_id: str,
    campaign_names: list[str] | None = None,
) -> dict:
    """Return today / this-week / this-month / all-time send counts plus labels.

    Each bucket includes both initial sends (contacted_at) and chaser sends
    (chaser_contacted_at) so the stats reflect total email touches per period.
    """
    today       = _date.today()
    monday      = today - timedelta(days=today.weekday())
    sunday      = monday + timedelta(days=6)
    month_start = today.replace(day=1)
    # Day numbers formatted manually — strftime('%-d') is Linux-only and
    # raises ValueError on Windows, which would break the whole breakdown.
    week_label  = (
        f"{monday.strftime('%b')} {monday.day}–{sunday.day}"
        if monday.month == sunday.month
        else f"{monday.strftime('%b')} {monday.day}–{sunday.strftime('%b')} {sunday.day}"
    )
    kw = {"campaign_names": campaign_names}

    # Each bucket needs two count queries (initial + chaser) — eight requests
    # total.  They are independent, so run them concurrently instead of
    # sequentially: the panel load drops from ~8× to ~1× round-trip latency.
    buckets = {
        "today":    {"date_eq": today.isoformat()},
        "week":     {"date_gte": monday.isoformat(), "date_lte": sunday.isoformat()},
        "month":    {"date_gte": month_start.isoformat(), "date_lte": today.isoformat()},
        "all_time": {},
    }
    with ThreadPoolExecutor(max_workers=8) as ex:
        futures = {
            name: (
                ex.submit(_count_contacted, client_id, **dates, **kw),
                ex.submit(_count_chaser_contacted, client_id, **dates, **kw),
            )
            for name, dates in buckets.items()
        }
        totals = {
            name: initial.result() + chasers.result()
            for name, (initial, chasers) in futures.items()
        }

    return {
        **totals,
        "week_label":  week_label,
        "month_label": today.strftime("%B %Y"),
    }


def _fetch_send_counts(client_id: str) -> dict:
    """Return calendar-accurate send counts for the initial page load."""
    return _time_breakdown(client_id)


# ── List campaigns + dashboard stats ──────────────────────────────────────────

@router.get("/api/campaigns")
async def list_campaigns(request: Request, client_id: str = Query("")):
    """Return the campaigns partial HTML (with stats dashboard when client selected)."""
    if not client_id:
        return templates.TemplateResponse(
            "partials/campaigns_list.html",
            {"request": request, "campaigns": [], "error": "", "client_id": "",
             "no_client": True, "stats": {}, "send_counts": {}},
        )

    campaigns: list[dict] = []
    error: str = ""
    stats:  dict = {}
    send_counts: dict = {}

    # Column sets — try with newer optional columns first; fall back if they
    # don't yet exist in the client's Supabase schema.
    _SELECT_FULL = (
        "id,created_at,campaign_name,sender_profile_name,"
        "client_id,client_name,sheet_id,sheet_url,"
        "total_prospects,sent_count,completed,completed_at,tags,"
        "lead_count,reply_count,interested_count,unsubscribe_count,"
        "paused,chaser_count"
    )
    _SELECT_BASE = (
        "id,created_at,campaign_name,sender_profile_name,"
        "client_id,client_name,sheet_id,sheet_url,"
        "total_prospects,sent_count,completed,completed_at,tags,"
        "lead_count,reply_count,interested_count,unsubscribe_count"
    )
    # Middle tiers: paused and chaser_count are independent optional columns.
    # If only one is missing, the other must still be selected — otherwise a
    # missing 'paused' would also hide the chaser progress bar (and vice
    # versa), since a PostgREST select fails entirely on any unknown column.
    _SELECT_CHASER = _SELECT_BASE + ",chaser_count"
    _SELECT_PAUSED = _SELECT_BASE + ",paused"

    if _sb_configured():
        for _sel in (_SELECT_FULL, _SELECT_CHASER, _SELECT_PAUSED, _SELECT_BASE):
            try:
                r = http_req.get(
                    f"{SUPABASE_URL}/rest/v1/campaigns",
                    headers=_sb_headers(),
                    params={"select": _sel, "order": "created_at.desc",
                            "client_id": f"eq.{client_id}"},
                    timeout=10,
                )
                r.raise_for_status()
                campaigns = r.json()
                break
            except Exception as exc:
                if _sel == _SELECT_BASE:
                    error = str(exc)
                campaigns = []

        if campaigns and not error:
            def _s(key: str) -> int:
                return sum(c.get(key) or 0 for c in campaigns)

            def _is_past(c: dict) -> bool:
                if c.get("paused"):
                    return False
                return bool(c.get("completed")) or (
                    (c.get("total_prospects") or 0) > 0
                    and (c.get("sent_count") or 0) >= (c.get("total_prospects") or 0)
                )

            active  = [c for c in campaigns if not c.get("paused") and not _is_past(c)]
            paused_ = [c for c in campaigns if c.get("paused")]
            past    = [c for c in campaigns if not c.get("paused") and _is_past(c)]

            prospects_in_pipeline = sum(
                max(0, (c.get("total_prospects") or 0) - (c.get("sent_count") or 0))
                for c in active
            )

            stats = {
                "total_campaigns":       len(campaigns),
                "active_count":          len(active),
                "paused_count":          len(paused_),
                "past_count":            len(past),
                "prospects_in_pipeline": prospects_in_pipeline,
                "total_prospects":       _s("total_prospects"),
                "total_sent":            _s("sent_count"),
                "total_leads":           _s("lead_count"),
                "total_replies":         _s("reply_count"),
                "total_interested":      _s("interested_count"),
                "total_unsubs":          _s("unsubscribe_count"),
            }
            send_counts = _fetch_send_counts(client_id)

    return templates.TemplateResponse(
        "partials/campaigns_list.html",
        {
            "request":     request,
            "campaigns":   campaigns,
            "error":       error,
            "client_id":   client_id,
            "no_client":   False,
            "stats":       stats,
            "send_counts": send_counts,
        },
    )


# ── Send-stats endpoint (custom range + tag-filter + drill-down) ──────────────

@router.get("/api/campaigns/send-stats")
async def campaign_send_stats(
    client_id:      str = Query(""),
    date_from:      str = Query(""),
    date_to:        str = Query(""),
    campaign_names: str = Query(""),   # comma-separated; empty = all campaigns
):
    """
    Flexible send-stats endpoint used by three features:

    1. Custom date-range picker  → provide date_from + date_to
       Returns {count, date_from, date_to}

    2. Tag-filter stats update   → provide campaign_names (no dates)
       Returns full today/week/month/all_time breakdown for those campaigns.

    3. Campaign drill-down       → provide campaign_names (no dates)
       Same as above but for a single campaign name.
    """
    if not client_id:
        return JSONResponse({"error": "client_id is required."}, status_code=400)
    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)

    names_list = (
        [n.strip() for n in campaign_names.split(",") if n.strip()]
        if campaign_names else None
    )

    # Custom date range → return a single count (initial sends + chasers)
    if date_from and date_to:
        count = (
            _count_contacted(client_id, date_gte=date_from, date_lte=date_to,
                             campaign_names=names_list)
            + _count_chaser_contacted(client_id, date_gte=date_from, date_lte=date_to,
                                      campaign_names=names_list)
        )
        return JSONResponse({"count": count, "date_from": date_from, "date_to": date_to})

    # No date range → return full time breakdown (today / week / month / all-time)
    return JSONResponse(_time_breakdown(client_id, campaign_names=names_list))


# ── Reset client stats ─────────────────────────────────────────────────────────

def _reset_stats_error(detail: str) -> HTMLResponse:
    """Visible error for reset-stats failures.

    Returned with HTTP 200 because HTMX does not swap error-status responses —
    the user must see why the reset did not (fully) apply.
    """
    return HTMLResponse(
        '<div style="padding:16px;border:1px solid #ef5350;border-radius:8px;'
        'color:#ef9a9a;font-size:.85rem;line-height:1.6">'
        f'<strong>Reset failed:</strong> {detail}<br>'
        'Reload the page and try again — nothing is lost by retrying, the '
        'reset is safe to run twice.</div>'
    )


@router.post("/api/campaigns/reset-stats", dependencies=[Depends(auth.require_admin)])
async def reset_client_stats(request: Request, client_id: str = Query("")):
    """
    Start stats tracking fresh for one client:

      1. Delete the client's synced send-history rows
         (contacted_prospects WHERE source = 'auto_sync') — these power the
         Today / This Week / This Month / All-Time send counts.
      2. Zero every campaign counter (sent / lead / reply / interested /
         unsub / chaser) for the client.

    Deliberately untouched:
      • dnc_entries — compliance data, never part of stats
      • scrub_upload / manual / csv_upload / dnc_manual contacted rows —
        they power the scrubber's "recently contacted" removal, not stats

    Campaigns with a linked sheet repopulate accurate numbers on their next
    sync (the sheet is the source of truth); rows from renamed or unlinked
    campaigns — the usual source of stale, inaccurate stats — stay gone.
    """
    if not client_id:
        return JSONResponse({"error": "client_id is required."}, status_code=400)
    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)

    # 1. Delete synced send history (stats source) — auto_sync rows ONLY.
    try:
        dr = http_req.delete(
            f"{SUPABASE_URL}/rest/v1/contacted_prospects",
            headers=_sb_headers(),
            params={
                "client_id": f"eq.{client_id}",
                "source":    "eq.auto_sync",
            },
            timeout=60,
        )
        dr.raise_for_status()
    except Exception as exc:
        return _reset_stats_error(f"could not clear send history: {exc}")

    # 2. Zero campaign counters.  chaser_count is optional (added by a manual
    # ALTER TABLE) and PostgREST rejects the whole PATCH on unknown columns,
    # so retry without it — same fallback pattern as the sync counter update.
    zero_counters = {
        "sent_count": 0, "lead_count": 0, "reply_count": 0,
        "interested_count": 0, "unsubscribe_count": 0,
    }
    patched = False
    try:
        pr = http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"client_id": f"eq.{client_id}"},
            json={**zero_counters, "chaser_count": 0},
            timeout=30,
        )
        patched = pr.status_code < 400
    except Exception:
        patched = False

    if not patched:
        try:
            pr = http_req.patch(
                f"{SUPABASE_URL}/rest/v1/campaigns",
                headers=_sb_headers("return=minimal"),
                params={"client_id": f"eq.{client_id}"},
                json=zero_counters,
                timeout=30,
            )
            pr.raise_for_status()
        except Exception as exc:
            return _reset_stats_error(
                f"send history was cleared, but campaign counters could not "
                f"be zeroed: {exc}"
            )

    return await list_campaigns(request, client_id=client_id)


# ── Sync + refresh (combined) ──────────────────────────────────────────────────
# Single endpoint used by the per-campaign "Sync" button.  Runs the full
# sync pipeline (DNC blocking, contacted_prospects sync, stat update) via
# sync_campaign_core, then returns the refreshed campaigns panel HTML.

@router.post("/api/campaigns/{campaign_id}/refresh")
async def refresh_campaign(request: Request, campaign_id: str):
    from app.utils.google_sheets import is_configured
    from app.utils.auto_sync import sync_campaign_core

    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)

    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={"id": f"eq.{campaign_id}",
                    "select": "id,sheet_id,client_id,campaign_name,total_prospects,completed_at"},
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)

    if not rows:
        return JSONResponse({"error": "Campaign not found."}, status_code=404)

    campaign         = rows[0]
    sheet_id         = campaign.get("sheet_id", "")
    client_id        = campaign.get("client_id", "")
    campaign_name    = campaign.get("campaign_name", "") or ""
    total_prospects  = campaign.get("total_prospects") or 0
    completed_at     = campaign.get("completed_at")

    if not sheet_id:
        return JSONResponse({"error": "No sheet linked."}, status_code=400)
    if not is_configured():
        return JSONResponse({"error": "Google Sheets not configured."}, status_code=503)

    # Run the full sync pipeline:
    #   • Back-fills Sent Date column for any rows missing it
    #   • Updates contacted_prospects (deletes old rows, inserts fresh auto_sync rows
    #     with the correct contacted_at dates so Today/Week/Month stats stay accurate)
    #   • Adds new DNC entries for leads / interested / unsubscribes
    #   • Updates campaign counters in Supabase
    result = sync_campaign_core(
        campaign_id, sheet_id, client_id, campaign_name,
        total_prospects=total_prospects,
        completed_at=completed_at,
    )

    if result["error"]:
        return JSONResponse({"error": result["error"]}, status_code=500)

    return await list_campaigns(request, client_id=client_id)


# ── Sync All campaigns for a client ───────────────────────────────────────────

@router.post("/api/campaigns/sync-all")
async def sync_all_campaigns(request: Request, client_id: str = Query("")):
    """
    Sync every campaign with a linked sheet for the given client.
    Runs sync_campaign_core for each campaign (sequentially) then returns the
    refreshed campaigns panel HTML.  Per-campaign errors are non-fatal.
    """
    from app.utils.google_sheets import is_configured
    from app.utils.auto_sync import sync_campaign_core

    if not client_id:
        return JSONResponse({"error": "client_id is required."}, status_code=400)
    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)
    if not is_configured():
        return JSONResponse({"error": "Google Sheets not configured."}, status_code=503)

    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={
                "select":    "id,sheet_id,client_id,campaign_name,total_prospects,completed_at",
                "client_id": f"eq.{client_id}",
                "sheet_id":  "not.is.null",
            },
            timeout=10,
        )
        r.raise_for_status()
        campaigns = [c for c in r.json() if (c.get("sheet_id") or "").strip()]
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)

    for c in campaigns:
        try:
            sync_campaign_core(
                c["id"],
                c["sheet_id"],
                c.get("client_id", ""),
                c.get("campaign_name", "") or "",
                total_prospects=c.get("total_prospects") or 0,
                completed_at=c.get("completed_at"),
            )
        except Exception:
            pass   # Non-fatal — continue with remaining campaigns

    return await list_campaigns(request, client_id=client_id)


# ── Sync from Google Sheet ─────────────────────────────────────────────────────

@router.post("/api/campaigns/{campaign_id}/sync")
async def sync_campaign(request: Request, campaign_id: str):
    """
    Read the linked Google Sheet and sync statuses:
      Send Status = TRUE/Sent  → add to contacted_prospects (campaign-tagged)
      Lead Status = Lead        → block DOMAIN in DNC (reason: lead)
      Lead Status = Interested  → block EMAIL in DNC  (reason: interested)
      Lead Status = Unsubscribe → block EMAIL in DNC  (reason: opt_out)
      Lead Status = Reply       → count only, no block

    Returns an HTML snippet for the campaign card's result div.
    """
    from app.utils.google_sheets import is_configured
    from app.utils.auto_sync import sync_campaign_core

    if not _sb_configured():
        return HTMLResponse('<span class="camp-sync-err">Supabase not configured.</span>')
    if not is_configured():
        return HTMLResponse('<span class="camp-sync-err">Google Sheets not configured.</span>')

    # ── Fetch campaign ────────────────────────────────────────────────────
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={"id": f"eq.{campaign_id}",
                    "select": "id,sheet_id,client_id,campaign_name,total_prospects,completed_at"},
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
    except Exception as exc:
        return HTMLResponse(f'<span class="camp-sync-err">{exc}</span>')

    if not rows:
        return HTMLResponse('<span class="camp-sync-err">Campaign not found.</span>')

    campaign         = rows[0]
    sheet_id         = campaign.get("sheet_id", "")
    client_id        = campaign.get("client_id", "")
    campaign_name    = campaign.get("campaign_name", "")
    total_prospects  = campaign.get("total_prospects") or 0
    completed_at     = campaign.get("completed_at")

    if not sheet_id:
        return HTMLResponse('<span class="camp-sync-err">No sheet linked.</span>')
    if not client_id:
        return HTMLResponse('<span class="camp-sync-err">Campaign has no client ID.</span>')

    # ── Delegate to shared core ───────────────────────────────────────────
    result = sync_campaign_core(
        campaign_id, sheet_id, client_id, campaign_name,
        total_prospects=total_prospects,
        completed_at=completed_at,
    )

    if result["error"]:
        return HTMLResponse(f'<span class="camp-sync-err">{result["error"]}</span>')

    leads_added       = result["leads_added"]
    interested_added  = result["interested_added"]
    unsubscribes_added = result["unsubscribes_added"]
    reply_count       = result["reply_count"]
    contacted_added   = result["contacted_added"]

    # ── Result snippet ────────────────────────────────────────────────────
    parts: list[str] = []
    if leads_added:
        parts.append(f"<strong>{leads_added}</strong> domain(s) blocked (Lead)")
    if interested_added:
        parts.append(f"<strong>{interested_added}</strong> email(s) blocked (Interested)")
    if unsubscribes_added:
        parts.append(f"<strong>{unsubscribes_added}</strong> email(s) blocked (Unsub)")
    if reply_count:
        parts.append(f"<strong>{reply_count}</strong> repl{'y' if reply_count == 1 else 'ies'} noted")
    if contacted_added:
        parts.append(f"<strong>{contacted_added}</strong> added to Contacted")
    if not parts:
        parts = ["No new entries to sync"]

    summary = " &nbsp;·&nbsp; ".join(parts)
    return HTMLResponse(f'<span class="camp-sync-ok">✓ {summary}</span>')


# ── A/B stats for a single campaign ───────────────────────────────────────────

@router.get("/api/campaigns/{campaign_id}/ab-stats")
async def campaign_ab_stats(campaign_id: str):
    """
    Read the linked Google Sheet and return per-variant response counts.
    Used by the campaign drill-down dashboard to render A/B winner + breakdown.
    """
    from app.utils.google_sheets import is_configured, read_ab_stats

    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)

    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={"id": f"eq.{campaign_id}", "select": "sheet_id"},
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)

    if not rows:
        return JSONResponse({"error": "Campaign not found."}, status_code=404)

    sheet_id = rows[0].get("sheet_id", "")
    if not sheet_id:
        return JSONResponse({"error": "No sheet linked to this campaign."}, status_code=400)
    if not is_configured():
        return JSONResponse({"error": "Google Sheets not configured."}, status_code=503)

    try:
        variants = read_ab_stats(sheet_id)
    except Exception as exc:
        return JSONResponse({"error": f"Sheet read error: {exc}"}, status_code=500)

    return JSONResponse({"variants": variants})


# ── Pause / Resume a campaign ─────────────────────────────────────────────────
# Requires a boolean `paused` column on the campaigns table (default false).

def _paused_patch_error(action: str, detail: str) -> HTMLResponse:
    """Visible error for pause/resume failures.

    Returned with HTTP 200 because HTMX does not swap error-status responses —
    a 500 here made the Pause button appear to silently do nothing.
    """
    return HTMLResponse(
        '<div style="padding:16px;border:1px solid #ef5350;border-radius:8px;'
        'color:#ef9a9a;font-size:.85rem;line-height:1.6">'
        f'<strong>{action} failed:</strong> {detail}<br>'
        "If this mentions a missing 'paused' column, run this in the "
        'Supabase SQL editor, then reload:<br>'
        '<code>ALTER TABLE campaigns ADD COLUMN IF NOT EXISTS paused '
        'boolean DEFAULT false;</code></div>'
    )


@router.post("/api/campaigns/{campaign_id}/pause")
async def pause_campaign(
    request:     Request,
    campaign_id: str,
    client_id:   str = Query(""),
):
    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)
    try:
        pr = http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json={"paused": True},
            timeout=10,
        )
        if pr.status_code >= 400:
            return _paused_patch_error("Pause", pr.text[:300])
    except Exception as exc:
        return _paused_patch_error("Pause", str(exc))
    return await list_campaigns(request, client_id=client_id)


@router.post("/api/campaigns/{campaign_id}/resume")
async def resume_campaign(
    request:     Request,
    campaign_id: str,
    client_id:   str = Query(""),
):
    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)
    try:
        pr = http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json={"paused": False},
            timeout=10,
        )
        if pr.status_code >= 400:
            return _paused_patch_error("Resume", pr.text[:300])
    except Exception as exc:
        return _paused_patch_error("Resume", str(exc))
    return await list_campaigns(request, client_id=client_id)


# ── Delete a campaign ─────────────────────────────────────────────────────────

@router.delete("/api/campaigns/{campaign_id}", dependencies=[Depends(auth.require_admin)])
async def delete_campaign(
    request:     Request,
    campaign_id: str,
    client_id:   str = Query(""),
):
    if _sb_configured():
        try:
            http_req.delete(
                f"{SUPABASE_URL}/rest/v1/campaigns",
                headers=_sb_headers(),
                params={"id": f"eq.{campaign_id}"},
                timeout=10,
            ).raise_for_status()
        except Exception:
            pass
    return await list_campaigns(request, client_id=client_id)


# ── Mark campaign as completed ─────────────────────────────────────────────────

@router.post("/api/campaigns/{campaign_id}/complete")
async def complete_campaign(
    request:     Request,
    campaign_id: str,
    client_id:   str = Query(""),
):
    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)
    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json={
                "completed":    True,
                "completed_at": _datetime.now(_tz.utc).isoformat(),
            },
            timeout=10,
        ).raise_for_status()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)
    return await list_campaigns(request, client_id=client_id)


# ── Update campaign tags ───────────────────────────────────────────────────────

@router.patch("/api/campaigns/{campaign_id}/tags")
async def update_campaign_tags(
    request:     Request,
    campaign_id: str,
    client_id:   str = Query(""),
):
    """Receive {tags: [...]} and persist to Supabase. Returns refreshed campaign list."""
    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)
    try:
        body = await request.json()
        tags = [t.strip() for t in (body.get("tags") or []) if str(t).strip()]
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json={"tags": tags},
            timeout=10,
        ).raise_for_status()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)
    return await list_campaigns(request, client_id=client_id)
