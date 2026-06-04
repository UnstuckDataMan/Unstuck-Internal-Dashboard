"""
Scheduled twice-daily sync for all campaigns that have a linked Google Sheet.

Runs for every client automatically — no manual trigger needed.
Times are controlled by the AUTO_SYNC_HOURS env var (comma-separated UTC hours,
default "9,21" → 09:00 and 21:00 UTC).
"""
from __future__ import annotations

import logging
import os
from datetime import date as _date, datetime as _datetime, timezone as _tz

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
    campaign_id:     str,
    sheet_id:        str,
    client_id:       str,
    campaign_name:   str,
    total_prospects: int = 0,
    completed_at:    str | None = None,
) -> dict:
    """
    Sync one campaign's Google Sheet into Supabase.
    Returns a summary dict: {leads_added, interested_added, unsubscribes_added,
                              reply_count, contacted_added, error}.
    Pass total_prospects + completed_at so the function can auto-stamp
    completed_at the first time sent_count reaches total_prospects.
    Does NOT raise — errors are returned in the dict so callers can log/continue.
    """
    from app.utils.google_sheets import read_leads, read_sent_with_dates, write_sent_dates

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
        sent_data = read_sent_with_dates(sheet_id)
    except Exception:
        sent_data = []

    # ── Back-fill missing Sent Date cells in the sheet ────────────────
    # For every sent row that has no Sent Date yet, stamp today.
    # Rows with an existing date keep it, so contacted_at always reflects
    # the real send date rather than the date the sync job happened to run.
    if sent_data:
        today_str = _date.today().isoformat()
        to_write: list[tuple[int, str]] = []
        for entry in sent_data:
            if not entry["sent_date"]:
                entry["sent_date"] = today_str
                to_write.append((entry["row_num"], today_str))
        if to_write:
            try:
                write_sent_dates(sheet_id, to_write)
            except Exception:
                pass   # Non-fatal — stats will still be correct

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
    # Use the actual Sent Date from the sheet so the date-bucketed stats
    # (Today / This Week / This Month) accurately reflect when emails were sent.
    #
    # Strategy: DELETE any existing scrub_upload rows for these emails first,
    # then INSERT fresh auto_sync rows.  scrub_upload rows are created when a
    # list is scrubbed with "Save to contacted history" and would otherwise
    # block the auto_sync insert (unique constraint on client_id + email).
    # Using merge-duplicates is unreliable because PostgREST defaults to
    # conflicting on the primary key (UUID), not the email-level unique
    # constraint, so new rows are appended rather than replacing existing ones.
    if sent_data:
        for i in range(0, len(sent_data), CHUNK_SIZE):
            chunk = sent_data[i : i + CHUNK_SIZE]
            email_list = "(" + ",".join(e["email"] for e in chunk) + ")"

            # Step 1: Remove any scrub_upload rows for these emails.
            # Scrub uploads block the upsert because the unique constraint
            # (client_id, email) would conflict, and merge-duplicates can't
            # overwrite a scrub_upload row with an auto_sync row via a simple
            # POST — we must explicitly delete them first.
            try:
                http_req.delete(
                    f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                    headers=_sb_headers(),
                    params={
                        "client_id": f"eq.{client_id}",
                        "source":    "eq.scrub_upload",
                        "email":     f"in.{email_list}",
                    },
                    timeout=30,
                )
            except Exception:
                pass   # Non-fatal

            # Step 2: Upsert auto_sync rows using PostgREST's on_conflict param.
            # ?on_conflict=client_id,email tells PostgREST to use the (client_id,
            # email) unique constraint for conflict detection — NOT the PK UUID.
            # merge-duplicates then updates existing rows (so contacted_at is
            # refreshed on every sync), and inserts new rows for emails not yet
            # present.  This correctly handles re-syncs: each sent email ends up
            # with the date from the sheet's Sent Date column, updated in-place.
            rows_to_insert = [
                {
                    "client_id":    client_id,
                    "email":        entry["email"],
                    "contacted_at": entry["sent_date"],
                    "source":       "auto_sync",
                    **({"campaign_name": campaign_name} if campaign_name else {}),
                }
                for entry in chunk
            ]
            try:
                r = http_req.post(
                    f"{SUPABASE_URL}/rest/v1/contacted_prospects"
                    f"?on_conflict=client_id,email",
                    headers=_sb_headers("resolution=merge-duplicates,return=minimal"),
                    json=rows_to_insert,
                    timeout=30,
                )
                if r.status_code not in (200, 201, 204, 409):
                    r.raise_for_status()
                result["contacted_added"] += len(chunk)
            except Exception:
                pass

    # ── Update campaign counters ──────────────────────────────────────
    new_sent     = len(sent_data)
    chaser_count = sum(1 for e in sent_data if e.get("chaser_sent"))
    patch_body: dict = {
        "sent_count":        new_sent,
        "lead_count":        result["leads_added"],
        "reply_count":       result["reply_count"],
        "interested_count":  result["interested_added"],
        "unsubscribe_count": result["unsubscribes_added"],
    }
    if chaser_count:
        patch_body["chaser_count"] = chaser_count

    # Auto-stamp completed_at the first time this campaign hits 100 %
    if total_prospects and new_sent >= total_prospects and not completed_at:
        patch_body["completed_at"] = _datetime.now(_tz.utc).isoformat()

    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json=patch_body,
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
                "select":   "id,sheet_id,client_id,campaign_name,total_prospects,completed_at",
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
                total_prospects=c.get("total_prospects") or 0,
                completed_at=c.get("completed_at"),
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
