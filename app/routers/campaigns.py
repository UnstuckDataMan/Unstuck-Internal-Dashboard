"""
Campaigns router — tracks mail-merge campaigns linked to Google Sheets.
"""
from __future__ import annotations

import os
from datetime import date as _date

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


# ── List campaigns ─────────────────────────────────────────────────────────────

@router.get("/api/campaigns")
async def list_campaigns(request: Request, client_id: str = Query("")):
    """Return the campaigns partial HTML."""
    # Require a client to be selected
    if not client_id:
        return templates.TemplateResponse(
            "partials/campaigns_list.html",
            {"request": request, "campaigns": [], "error": "", "client_id": "",
             "no_client": True},
        )

    campaigns: list[dict] = []
    error: str = ""

    if _sb_configured():
        params: dict = {
            "select": "id,created_at,campaign_name,sender_profile_name,client_id,client_name,sheet_id,sheet_url,total_prospects,sent_count",
            "order":     "created_at.desc",
            "client_id": f"eq.{client_id}",
        }
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
            "request":   request,
            "campaigns": campaigns,
            "error":     error,
            "client_id": client_id,
            "no_client": False,
        },
    )


# ── Refresh a campaign's sent count from Google Sheet ─────────────────────────

@router.post("/api/campaigns/{campaign_id}/refresh")
async def refresh_campaign(request: Request, campaign_id: str):
    """Re-read the Google Sheet and update sent_count in Supabase."""
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
        return JSONResponse({"error": "No Google Sheet linked to this campaign."}, status_code=400)
    if not is_configured():
        return JSONResponse({"error": "Google Sheets not configured."}, status_code=503)

    try:
        status = read_sheet_status(sheet_id)
    except Exception as exc:
        return JSONResponse({"error": f"Could not read sheet: {exc}"}, status_code=500)

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


# ── Sync a campaign from its Google Sheet ─────────────────────────────────────

@router.post("/api/campaigns/{campaign_id}/sync")
async def sync_campaign(request: Request, campaign_id: str):
    """
    Read the linked Google Sheet and:
      • Send Status = TRUE/Sent → add email to contacted_prospects (campaign-tagged)
      • Lead Status = "Lead"    → add DOMAIN to dnc_entries (reason: lead)
      • Lead Status = "Unsubscribe" → add EMAIL to dnc_entries (reason: opt_out)

    Returns an HTML snippet rendered into the campaign card's result div.
    """
    from app.utils.google_sheets import is_configured, read_leads, read_sent_emails

    if not _sb_configured():
        return HTMLResponse('<span class="camp-sync-err">Supabase not configured.</span>')
    if not is_configured():
        return HTMLResponse('<span class="camp-sync-err">Google Sheets not configured.</span>')

    # ── Fetch campaign record ─────────────────────────────────────────────
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={"id": f"eq.{campaign_id}",
                    "select": "id,sheet_id,client_id,campaign_name"},
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
    except Exception as exc:
        return HTMLResponse(f'<span class="camp-sync-err">{exc}</span>')

    if not rows:
        return HTMLResponse('<span class="camp-sync-err">Campaign not found.</span>')

    campaign      = rows[0]
    sheet_id      = campaign.get("sheet_id", "")
    client_id     = campaign.get("client_id", "")
    campaign_name = campaign.get("campaign_name", "")

    if not sheet_id:
        return HTMLResponse('<span class="camp-sync-err">No sheet linked.</span>')
    if not client_id:
        return HTMLResponse('<span class="camp-sync-err">Campaign has no client ID.</span>')

    # ── Read sheet ────────────────────────────────────────────────────────
    try:
        leads = read_leads(sheet_id)
    except Exception as exc:
        return HTMLResponse(f'<span class="camp-sync-err">Sheet read error: {exc}</span>')

    try:
        sent_emails = read_sent_emails(sheet_id)
    except Exception:
        sent_emails = []

    leads_added = unsubscribes_added = contacted_added = 0

    # ── DNC: Lead → domain block, Unsubscribe → email block ──────────────
    if leads:
        dnc_rows: list[dict] = []
        for entry in leads:
            email  = entry["email"]
            status = entry["status"]
            if status == "Lead" and "@" in email:
                domain = email.split("@")[1]
                dnc_rows.append({
                    "client_id": client_id,
                    "email":     domain,
                    "reason":    "lead",
                    "added_by":  "google_sheets_sync",
                })
                leads_added += 1
            elif status == "Unsubscribe":
                dnc_rows.append({
                    "client_id": client_id,
                    "email":     email,
                    "reason":    "opt_out",
                    "added_by":  "google_sheets_sync",
                })
                unsubscribes_added += 1

        for i in range(0, len(dnc_rows), CHUNK_SIZE):
            chunk = dnc_rows[i : i + CHUNK_SIZE]
            try:
                r = http_req.post(
                    f"{SUPABASE_URL}/rest/v1/dnc_entries",
                    headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
                    json=chunk,
                    timeout=30,
                )
                # 409 = entries already exist in DNC — treat as success
                if r.status_code not in (200, 201, 204, 409):
                    r.raise_for_status()
            except Exception as exc:
                return HTMLResponse(
                    f'<span class="camp-sync-err">DNC write error: {exc}</span>'
                )

    # ── Contacted: Sent rows → contacted_prospects ────────────────────────
    if sent_emails:
        today = _date.today().isoformat()
        for i in range(0, len(sent_emails), CHUNK_SIZE):
            chunk = sent_emails[i : i + CHUNK_SIZE]
            rows_to_insert = [
                {
                    "client_id":    client_id,
                    "email":        e,
                    "contacted_at": today,
                    "source":       "google_sheets_sync",
                    **({"campaign_name": campaign_name} if campaign_name else {}),
                }
                for e in chunk
            ]
            try:
                http_req.post(
                    f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                    headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
                    json=rows_to_insert,
                    timeout=30,
                ).raise_for_status()
                contacted_added += len(chunk)
            except Exception:
                pass   # non-fatal

    # ── Update sent_count on campaign record ──────────────────────────────
    if sent_emails:
        try:
            http_req.patch(
                f"{SUPABASE_URL}/rest/v1/campaigns",
                headers=_sb_headers("return=minimal"),
                params={"id": f"eq.{campaign_id}"},
                json={"sent_count": len(sent_emails)},
                timeout=10,
            )
        except Exception:
            pass

    # ── Return result snippet ─────────────────────────────────────────────
    parts: list[str] = []
    if leads_added:
        parts.append(f"<strong>{leads_added}</strong> domain(s) blocked (Lead)")
    if unsubscribes_added:
        parts.append(f"<strong>{unsubscribes_added}</strong> email(s) blocked (Unsub)")
    if contacted_added:
        parts.append(f"<strong>{contacted_added}</strong> added to Contacted")
    if not parts:
        parts = ["No new entries to sync"]

    summary = " &nbsp;·&nbsp; ".join(parts)
    return HTMLResponse(f'<span class="camp-sync-ok">✓ {summary}</span>')


# ── Delete a campaign ─────────────────────────────────────────────────────────

@router.delete("/api/campaigns/{campaign_id}")
async def delete_campaign(
    request:     Request,
    campaign_id: str,
    client_id:   str = Query(""),
):
    """Delete a campaign record and re-render the campaigns list."""
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
