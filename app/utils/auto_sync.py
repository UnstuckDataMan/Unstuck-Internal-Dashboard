"""
Scheduled twice-daily sync for all campaigns that have a linked Google Sheet.

Runs for every client automatically — no manual trigger needed.
Times are controlled by the AUTO_SYNC_HOURS env var (comma-separated UTC hours,
default "9,21" → 09:00 and 21:00 UTC).
"""
from __future__ import annotations

import logging
import os
from datetime import date as _date

import requests as http_req

logger = logging.getLogger(__name__)

SUPABASE_URL      = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")
CHUNK_SIZE        = 500


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


# ── Core per-campaign sync (shared with the manual endpoint) ──────────────────

def sync_campaign_core(
    campaign_id:   str,
    sheet_id:      str,
    client_id:     str,
    campaign_name: str,
) -> dict:
    """
    Sync one campaign's Google Sheet into Supabase.
    Returns a summary dict: {leads_added, interested_added, unsubscribes_added,
                              reply_count, contacted_added, error}.
    Does NOT raise — errors are returned in the dict so callers can log/continue.
    """
    from app.utils.google_sheets import read_leads, read_sent_emails

    result = {
        "leads_added": 0, "interested_added": 0,
        "unsubscribes_added": 0, "reply_count": 0,
        "contacted_added": 0, "error": "",
    }

    # ── Read sheet ────────────────────────────────────────────────────
    try:
        leads = read_leads(sheet_id)
    except Exception as exc:
        result["error"] = f"Sheet read error: {exc}"
        return result

    try:
        sent_emails = read_sent_emails(sheet_id)
    except Exception:
        sent_emails = []

    # ── Build DNC rows ────────────────────────────────────────────────
    if leads:
        dnc_rows: list[dict] = []
        for entry in leads:
            email  = entry["email"]
            status = entry["status"]

            if status == "Lead" and "@" in email:
                domain = email.split("@")[1]
                dnc_rows.append({
                    "client_id": client_id, "email": domain,
                    "reason": "lead", "added_by": "auto_sync",
                })
                result["leads_added"] += 1

            elif status == "Interested":
                dnc_rows.append({
                    "client_id": client_id, "email": email,
                    "reason": "interested", "added_by": "auto_sync",
                })
                result["interested_added"] += 1

            elif status == "Unsubscribe":
                dnc_rows.append({
                    "client_id": client_id, "email": email,
                    "reason": "opt_out", "added_by": "auto_sync",
                })
                result["unsubscribes_added"] += 1

            elif status == "Reply":
                result["reply_count"] += 1

        for i in range(0, len(dnc_rows), CHUNK_SIZE):
            chunk = dnc_rows[i : i + CHUNK_SIZE]
            try:
                r = http_req.post(
                    f"{SUPABASE_URL}/rest/v1/dnc_entries",
                    headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
                    json=chunk,
                    timeout=30,
                )
                if r.status_code not in (200, 201, 204, 409):
                    r.raise_for_status()
            except Exception as exc:
                result["error"] = f"DNC write error: {exc}"
                return result

    # ── Add sent rows to contacted_prospects ──────────────────────────
    if sent_emails:
        today = _date.today().isoformat()
        for i in range(0, len(sent_emails), CHUNK_SIZE):
            chunk = sent_emails[i : i + CHUNK_SIZE]
            rows_to_insert = [
                {
                    "client_id":    client_id,
                    "email":        e,
                    "contacted_at": today,
                    "source":       "auto_sync",
                    **({"campaign_name": campaign_name} if campaign_name else {}),
                }
                for e in chunk
            ]
            try:
                r = http_req.post(
                    f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                    headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
                    json=rows_to_insert,
                    timeout=30,
                )
                if r.status_code not in (200, 201, 204, 409):
                    r.raise_for_status()
                result["contacted_added"] += len(chunk)
            except Exception:
                pass

    # ── Update campaign counters ──────────────────────────────────────
    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json={
                "sent_count":        len(sent_emails),
                "lead_count":        result["leads_added"],
                "reply_count":       result["reply_count"],
                "interested_count":  result["interested_added"],
                "unsubscribe_count": result["unsubscribes_added"],
            },
            timeout=10,
        )
    except Exception:
        pass

    return result


# ── Scheduled job ─────────────────────────────────────────────────────────────

def run_auto_sync() -> None:
    """
    Called by APScheduler twice a day.
    Fetches every campaign that has a sheet_id and syncs it.
    """
    from app.utils.google_sheets import is_configured

    if not _sb_configured():
        logger.warning("Auto-sync skipped: Supabase not configured.")
        return
    if not is_configured():
        logger.warning("Auto-sync skipped: Google Sheets not configured.")
        return

    # Fetch all campaigns that have a linked sheet
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={
                "select":   "id,sheet_id,client_id,campaign_name",
                "sheet_id": "not.is.null",
            },
            timeout=15,
        )
        r.raise_for_status()
        campaigns = [c for c in r.json() if c.get("sheet_id", "").strip()]
    except Exception as exc:
        logger.error("Auto-sync: failed to fetch campaigns: %s", exc)
        return

    if not campaigns:
        logger.info("Auto-sync: no campaigns with linked sheets — nothing to do.")
        return

    logger.info("Auto-sync: starting sync for %d campaign(s) across all clients.", len(campaigns))

    ok = failed = 0
    for c in campaigns:
        cid  = c.get("id", "")
        name = c.get("campaign_name", "") or ""
        try:
            summary = sync_campaign_core(
                campaign_id=cid,
                sheet_id=c["sheet_id"],
                client_id=c.get("client_id", ""),
                campaign_name=name,
            )
            if summary["error"]:
                logger.warning("Auto-sync: campaign %s (%s) error — %s", cid, name, summary["error"])
                failed += 1
            else:
                logger.info(
                    "Auto-sync: campaign %s (%s) — leads=%d interested=%d unsubs=%d contacted=%d",
                    cid, name,
                    summary["leads_added"], summary["interested_added"],
                    summary["unsubscribes_added"], summary["contacted_added"],
                )
                ok += 1
        except Exception as exc:
            logger.error("Auto-sync: campaign %s (%s) unexpected error: %s", cid, name, exc)
            failed += 1

    logger.info("Auto-sync complete: %d succeeded, %d failed.", ok, failed)
