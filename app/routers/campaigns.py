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


def _fetch_send_counts(client_id: str) -> dict:
    """Return Today / This Week / This Month / All Time send counts from contacted_prospects."""
    today     = _date.today()
    week_ago  = (today - timedelta(days=7)).isoformat()
    month_ago = (today - timedelta(days=30)).isoformat()
    today_str = today.isoformat()

    counts = {"today": 0, "week": 0, "month": 0, "all_time": 0}
    for key, cutoff in (("month", month_ago), ("week", week_ago), ("today", today_str)):
        try:
            r = http_req.get(
                f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                headers={**_sb_headers(), "Prefer": "count=exact"},
                params={
                    "select":        "id",
                    "client_id":     f"eq.{client_id}",
                    "contacted_at":  f"gte.{cutoff}",
                    "campaign_name": "not.is.null",
                    "limit":         "1",
                },
                timeout=10,
            )
            counts[key] = _parse_total(r.headers)
        except Exception:
            pass

    # All-time: no date filter — source of truth from contacted_prospects
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/contacted_prospects",
            headers={**_sb_headers(), "Prefer": "count=exact"},
            params={
                "select":        "id",
                "client_id":     f"eq.{client_id}",
                "campaign_name": "not.is.null",
                "limit":         "1",
            },
            timeout=10,
        )
        counts["all_time"] = _parse_total(r.headers)
    except Exception:
        pass

    return counts


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

    if _sb_configured():
        try:
            r = http_req.get(
                f"{SUPABASE_URL}/rest/v1/campaigns",
                headers=_sb_headers(),
                params={
                    "select":    "id,created_at,campaign_name,sender_profile_name,"
                                 "client_id,client_name,sheet_id,sheet_url,"
                                 "total_prospects,sent_count,completed,completed_at,tags,"
                                 "lead_count,reply_count,interested_count,unsubscribe_count",
                    "order":     "created_at.desc",
                    "client_id": f"eq.{client_id}",
                },
                timeout=10,
            )
            r.raise_for_status()
            campaigns = r.json()
        except Exception as exc:
            error = str(exc)

        if campaigns and not error:
            def _s(key: str) -> int:
                return sum(c.get(key) or 0 for c in campaigns)

            def _is_past(c: dict) -> bool:
                return bool(c.get("completed")) or (
                    (c.get("total_prospects") or 0) > 0
                    and (c.get("sent_count") or 0) >= (c.get("total_prospects") or 0)
                )

            active = [c for c in campaigns if not _is_past(c)]
            past   = [c for c in campaigns if _is_past(c)]

            prospects_in_pipeline = sum(
                max(0, (c.get("total_prospects") or 0) - (c.get("sent_count") or 0))
                for c in active
            )

            stats = {
                "total_campaigns":       len(campaigns),
                "active_count":          len(active),
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


# ── Refresh sent count ─────────────────────────────────────────────────────────

@router.post("/api/campaigns/{campaign_id}/refresh")
async def refresh_campaign(request: Request, campaign_id: str):
    from app.utils.google_sheets import is_configured, read_sheet_status

    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)

    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={"id": f"eq.{campaign_id}",
                    "select": "id,sheet_id,client_id,total_prospects"},
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)

    if not rows:
        return JSONResponse({"error": "Campaign not found."}, status_code=404)

    campaign  = rows[0]
    sheet_id  = campaign.get("sheet_id", "")
    client_id = campaign.get("client_id", "")

    if not sheet_id:
        return JSONResponse({"error": "No sheet linked."}, status_code=400)
    if not is_configured():
        return JSONResponse({"error": "Google Sheets not configured."}, status_code=503)

    try:
        status = read_sheet_status(sheet_id)
    except Exception as exc:
        return JSONResponse({"error": f"Sheet read error: {exc}"}, status_code=500)

    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json={"sent_count": status["sent"]},
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
