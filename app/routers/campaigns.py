"""
Campaigns router — tracks mail-merge campaigns linked to Google Sheets.
"""
from __future__ import annotations

import os
from datetime import date as _date, datetime as _datetime, timedelta, timezone as _tz

import requests as http_req
from fastapi import APIRouter, Query, Request
from fastapi.responses import JSONResponse, HTMLResponse

from app.deps import templates

router = APIRouter()

SUPABASE_URL      = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")

CHUNK_SIZE = 500


def _sb_headers(prefer: str = "") -> dict:
    h = {
        "apikey":        SUPABASE_ANON_KEY,
        "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
        "Content-Type":  "application/json",
    }
    if prefer:
        h["Prefer"] = prefer
    return h


def _sb_configured() -> bool:
    return bool(SUPABASE_URL and SUPABASE_ANON_KEY)


def _parse_total(headers: dict) -> int:
    cr = headers.get("Content-Range", "")
    if "/" in cr:
        try:
            return int(cr.split("/")[1])
        except ValueError:
            pass
    return 0


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
        val = "(" + ",".join(campaign_names) + ")"
        base_params.append(("campaign_name", f"in.{val}"))
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


def _time_breakdown(
    client_id: str,
    campaign_names: list[str] | None = None,
) -> dict:
    """Return today / this-week / this-month / all-time send counts plus labels."""
    today       = _date.today()
    monday      = today - timedelta(days=today.weekday())
    sunday      = monday + timedelta(days=6)
    month_start = today.replace(day=1)
    week_label  = (
        f"{monday.strftime('%b %-d')}–{sunday.strftime('%-d')}"
        if monday.month == sunday.month
        else f"{monday.strftime('%b %-d')}–{sunday.strftime('%b %-d')}"
    )
    kw = {"campaign_names": campaign_names}
    return {
        "today":       _count_contacted(client_id, date_eq=today.isoformat(), **kw),
        "week":        _count_contacted(client_id, date_gte=monday.isoformat(), date_lte=sunday.isoformat(), **kw),
        "month":       _count_contacted(client_id, date_gte=month_start.isoformat(), date_lte=today.isoformat(), **kw),
        "all_time":    _count_contacted(client_id, **kw),
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

    if _sb_configured():
        for _sel in (_SELECT_FULL, _SELECT_BASE):
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

    # Custom date range → return a single count
    if date_from and date_to:
        count = _count_contacted(client_id, date_gte=date_from, date_lte=date_to,
                                 campaign_names=names_list)
        return JSONResponse({"count": count, "date_from": date_from, "date_to": date_to})

    # No date range → return full time breakdown (today / week / month / all-time)
    return JSONResponse(_time_breakdown(client_id, campaign_names=names_list))


# ── Refresh sent count ─────────────────────────────────────────────────────────

@router.post("/api/campaigns/{campaign_id}/refresh")
async def refresh_campaign(request: Request, campaign_id: str):
    from app.utils.google_sheets import is_configured, read_sheet_status, read_leads

    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)

    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={"id": f"eq.{campaign_id}",
                    "select": "id,sheet_id,client_id,campaign_name,total_prospects"},
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)

    if not rows:
        return JSONResponse({"error": "Campaign not found."}, status_code=404)

    campaign      = rows[0]
    sheet_id      = campaign.get("sheet_id", "")
    client_id     = campaign.get("client_id", "")
    campaign_name = campaign.get("campaign_name", "") or ""

    if not sheet_id:
        return JSONResponse({"error": "No sheet linked."}, status_code=400)
    if not is_configured():
        return JSONResponse({"error": "Google Sheets not configured."}, status_code=503)

    try:
        status = read_sheet_status(sheet_id)
    except Exception as exc:
        return JSONResponse({"error": f"Sheet read error: {exc}"}, status_code=500)

    # Read Lead Status counts from the sheet so all counter columns stay in sync
    lead_count = reply_count = interested_count = unsub_count = chaser_count = 0
    try:
        leads_data       = read_leads(sheet_id)
        lead_count       = sum(1 for e in leads_data if e["status"] == "Lead")
        reply_count      = sum(1 for e in leads_data if e["status"] == "Reply")
        interested_count = sum(1 for e in leads_data if e["status"] == "Interested")
        unsub_count      = sum(1 for e in leads_data if e["status"] == "Unsubscribe")
    except Exception:
        pass   # Non-fatal — sent_count will still be updated below

    sent_data: list[dict] = []
    chaser_count = 0
    try:
        from app.utils.google_sheets import read_sent_with_dates, write_sent_dates
        sent_data    = read_sent_with_dates(sheet_id)
        chaser_count = sum(1 for e in sent_data if e.get("chaser_sent"))

        # ── Back-fill missing Sent Date cells ────────────────────────────
        # Mirror the same logic as auto_sync so Refresh gives an accurate
        # contacted_at date for the Today stat rather than always writing today.
        today_str  = _date.today().isoformat()
        to_write: list[tuple[int, str]] = []
        for entry in sent_data:
            if not entry["sent_date"]:
                entry["sent_date"] = today_str
                to_write.append((entry["row_num"], today_str))
        if to_write:
            try:
                write_sent_dates(sheet_id, to_write)
            except Exception:
                pass

        # ── Sync contacted_prospects so the Today stat is current ─────────
        # The auto-sync runs twice daily; clicking Refresh brings the count
        # up to date immediately so users don't wait for the next scheduled run.
        if sent_data:
            cp_rows = [
                {
                    "client_id":    client_id,
                    "email":        e["email"],
                    "contacted_at": e["sent_date"],
                    "source":       "auto_sync",
                    **({"campaign_name": campaign_name} if campaign_name else {}),
                }
                for e in sent_data
            ]
            for i in range(0, len(cp_rows), CHUNK_SIZE):
                try:
                    http_req.post(
                        f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                        headers=_sb_headers("resolution=merge-duplicates,return=minimal"),
                        json=cp_rows[i : i + CHUNK_SIZE],
                        timeout=30,
                    )
                except Exception:
                    pass
    except Exception:
        pass   # Non-fatal — counters below will still update

    patch_body: dict = {
        "sent_count":        status["sent"],
        "lead_count":        lead_count,
        "reply_count":       reply_count,
        "interested_count":  interested_count,
        "unsubscribe_count": unsub_count,
    }
    # chaser_count requires the column to exist in Supabase; skip gracefully if not
    if chaser_count:
        patch_body["chaser_count"] = chaser_count

    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json=patch_body,
            timeout=10,
        ).raise_for_status()
    except Exception as exc:
        return JSONResponse({"error": f"Supabase update failed: {exc}"}, status_code=500)

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

@router.post("/api/campaigns/{campaign_id}/pause")
async def pause_campaign(
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
            json={"paused": True},
            timeout=10,
        ).raise_for_status()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)
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
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json={"paused": False},
            timeout=10,
        ).raise_for_status()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)
    return await list_campaigns(request, client_id=client_id)


# ── Delete a campaign ─────────────────────────────────────────────────────────

@router.delete("/api/campaigns/{campaign_id}")
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
