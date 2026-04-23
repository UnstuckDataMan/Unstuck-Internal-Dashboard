"""
Campaigns router — tracks mail-merge campaigns linked to Google Sheets.
"""
from __future__ import annotations

import os

import requests as http_req
from fastapi import APIRouter, Query, Request
from fastapi.responses import JSONResponse

from app.deps import templates

router = APIRouter()

SUPABASE_URL      = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")


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


# ── List campaigns ─────────────────────────────────────────────────────────────

@router.get("/api/campaigns")
async def list_campaigns(request: Request, client_id: str = Query("")):
    """Return the campaigns partial HTML."""
    campaigns: list[dict] = []
    error: str = ""

    if _sb_configured():
        params: dict = {
            "select": "id,created_at,campaign_name,sender_profile_name,client_id,client_name,sheet_id,sheet_url,total_prospects,sent_count",
            "order": "created_at.desc",
        }
        if client_id:
            params["client_id"] = f"eq.{client_id}"

        try:
            r = http_req.get(
                f"{SUPABASE_URL}/rest/v1/campaigns",
                headers=_sb_headers(),
                params=params,
                timeout=10,
            )
            r.raise_for_status()
            campaigns = r.json()
        except Exception as exc:
            error = str(exc)

    return templates.TemplateResponse(
        "partials/campaigns_list.html",
        {
            "request":    request,
            "campaigns":  campaigns,
            "error":      error,
            "client_id":  client_id,
        },
    )


# ── Refresh a campaign's sent count from Google Sheet ─────────────────────────

@router.post("/api/campaigns/{campaign_id}/refresh")
async def refresh_campaign(request: Request, campaign_id: str):
    """Re-read the Google Sheet and update sent_count in Supabase."""
    from app.utils.google_sheets import is_configured, read_sheet_status

    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured."}, status_code=503)

    # Fetch the campaign record to get sheet_id
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={"id": f"eq.{campaign_id}", "select": "id,sheet_id,total_prospects"},
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)

    if not rows:
        return JSONResponse({"error": "Campaign not found."}, status_code=404)

    campaign = rows[0]
    sheet_id = campaign.get("sheet_id", "")

    if not sheet_id:
        return JSONResponse({"error": "No Google Sheet linked to this campaign."}, status_code=400)

    if not is_configured():
        return JSONResponse({"error": "Google Sheets not configured."}, status_code=503)

    try:
        status = read_sheet_status(sheet_id)
    except Exception as exc:
        return JSONResponse({"error": f"Could not read sheet: {exc}"}, status_code=500)

    # Update sent_count in Supabase
    try:
        r = http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json={"sent_count": status["sent"]},
            timeout=10,
        )
        r.raise_for_status()
    except Exception as exc:
        return JSONResponse({"error": f"Supabase update failed: {exc}"}, status_code=500)

    # Re-render the partial for this campaign's client
    client_id = campaign.get("client_id", "")
    return await list_campaigns(request, client_id=client_id)
