"""
Scheduled twice-daily sync for all campaigns that have a linked Google Sheet.

Runs for every client automatically — no manual trigger needed.
Times are controlled by the AUTO_SYNC_HOURS env var (comma-separated UTC hours,
default "9,21" → 09:00 and 21:00 UTC).
"""
from __future__ import annotations

import logging
from datetime import date as _date, datetime as _datetime, timezone as _tz

import requests as http_req

from app.utils.supabase import (
    SUPABASE_URL,
    sb_headers as _sb_headers,
    sb_configured as _sb_configured,
)

logger = logging.getLogger(__name__)

CHUNK_SIZE = 500


def _next_working_day(from_date: _date) -> _date:
    """Return the first Monday–Friday after from_date (skips weekends)."""
    from datetime import timedelta
    d = from_date + timedelta(days=1)
    while d.weekday() >= 5:   # 5 = Saturday, 6 = Sunday
        d += timedelta(days=1)
    return d


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
    from app.utils.google_sheets import (
        read_leads, read_sent_with_dates, write_sent_dates, write_chaser_dates,
    )

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

    # sent_read_ok distinguishes "sheet genuinely has no ticked rows" from a
    # transient read failure — the stale-row cleanup below must never run
    # after a failed read or it would wipe the campaign's send history.
    sent_read_ok = True
    try:
        sent_data = read_sent_with_dates(sheet_id)
    except Exception:
        sent_data = []
        sent_read_ok = False

    # ── Back-fill missing Sent Date and Chaser Date cells ────────────
    # For every sent row that has no Sent Date yet, stamp today.
    # For every chased row (Chaser Sent? = TRUE) with no Chaser Date, stamp today.
    # Existing dates are kept so contacted_at always reflects the real send date.
    if sent_data:
        today_str      = _date.today().isoformat()
        send_date_str  = _next_working_day(_date.today()).isoformat()
        sent_to_write:   list[tuple[int, str]] = []
        chaser_to_write: list[tuple[int, str]] = []
        for entry in sent_data:
            if not entry["sent_date"]:
                entry["sent_date"] = send_date_str
                sent_to_write.append((entry["row_num"], send_date_str))
            if entry.get("chaser_sent") and not entry.get("chaser_date"):
                entry["chaser_date"] = send_date_str
                chaser_to_write.append((entry["row_num"], send_date_str))
        if sent_to_write:
            try:
                write_sent_dates(sheet_id, sent_to_write)
            except Exception as _exc:
                logger.warning("write_sent_dates failed for sheet %s: %s", sheet_id, _exc)
                result["error"] = f"Sent Date write failed: {_exc}"
        if chaser_to_write:
            try:
                write_chaser_dates(sheet_id, chaser_to_write)
            except Exception as _exc:
                logger.warning("write_chaser_dates failed for sheet %s: %s", sheet_id, _exc)

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

        # Deduplicate within the batch.  Two Leads at the same company both
        # yield the same domain row, and duplicate values inside one INSERT
        # make Postgres reject the whole statement.
        seen_dnc: set[str] = set()
        unique_dnc: list[dict] = []
        for row in dnc_rows:
            key = row["email"].lower()
            if key not in seen_dnc:
                seen_dnc.add(key)
                unique_dnc.append(row)

        for i in range(0, len(unique_dnc), CHUNK_SIZE):
            chunk = unique_dnc[i : i + CHUNK_SIZE]
            # on_conflict=client_id,email targets the table's real unique
            # constraint.  Without it, ignore-duplicates conflicts on the PK
            # (a fresh UUID that never collides), so any row already in the
            # DNC failed the ENTIRE chunk with 409 — and the old code treated
            # 409 as success, silently dropping every new lead/unsub after
            # the first sync.
            ok = False
            try:
                r = http_req.post(
                    f"{SUPABASE_URL}/rest/v1/dnc_entries",
                    headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
                    params={"on_conflict": "client_id,email"},
                    json=chunk,
                    timeout=30,
                )
                ok = r.status_code in (200, 201, 204)
            except Exception:
                ok = False

            if not ok:
                # Fallback: insert rows one at a time so a single conflicting
                # row (or an on_conflict/constraint mismatch) can't block the
                # rest.  409 = already in DNC = fine.
                failures = 0
                for row in chunk:
                    try:
                        r1 = http_req.post(
                            f"{SUPABASE_URL}/rest/v1/dnc_entries",
                            headers=_sb_headers("return=minimal"),
                            json=[row],
                            timeout=15,
                        )
                        if r1.status_code not in (200, 201, 204, 409):
                            failures += 1
                    except Exception:
                        failures += 1
                if failures:
                    # Record the problem but keep syncing — contacted data and
                    # counters should still update even if some DNC rows failed.
                    result["error"] = f"DNC write error: {failures} entr{'y' if failures == 1 else 'ies'} failed"

    # ── Add sent rows to contacted_prospects ──────────────────────────
    # Use the actual Sent Date from the sheet so the date-bucketed stats
    # (Today / This Week / This Month) accurately reflect when emails were sent.
    #
    # Strategy: DELETE all existing rows for these emails first, then INSERT
    # fresh auto_sync rows with the correct contacted_at dates.
    # This ensures re-syncs always reflect updated send dates (e.g. when more
    # prospects are marked as sent between syncs).

    # Deduplicate by email first.  Sheets can contain duplicate prospect rows;
    # contacted_prospects is unique on (client_id, email) so without this the
    # ignore-duplicates resolution decides arbitrarily which row wins.  Keep
    # the earliest sent date and merge chaser info from any duplicate so the
    # stats count each prospect exactly once per period.
    unique_sent: list[dict] = []
    if sent_data:
        by_email: dict[str, dict] = {}
        for entry in sent_data:
            cur = by_email.get(entry["email"])
            if cur is None:
                by_email[entry["email"]] = dict(entry)
            else:
                if entry.get("sent_date") and (
                    not cur.get("sent_date") or entry["sent_date"] < cur["sent_date"]
                ):
                    cur["sent_date"] = entry["sent_date"]
                if entry.get("chaser_sent"):
                    cur["chaser_sent"] = True
                if entry.get("chaser_date") and (
                    not cur.get("chaser_date") or entry["chaser_date"] < cur["chaser_date"]
                ):
                    cur["chaser_date"] = entry["chaser_date"]
        unique_sent = list(by_email.values())

    if unique_sent:
        for i in range(0, len(unique_sent), CHUNK_SIZE):
            chunk = unique_sent[i : i + CHUNK_SIZE]
            email_list = "(" + ",".join(f'"{e["email"]}"' for e in chunk) + ")"

            # Remove all existing rows for these emails so the fresh auto_sync
            # rows carry the correct contacted_at date on every sync.
            try:
                http_req.delete(
                    f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                    headers=_sb_headers(),
                    params={
                        "client_id": f"eq.{client_id}",
                        "email":     f"in.{email_list}",
                    },
                    timeout=30,
                )
            except Exception:
                pass   # Non-fatal

            rows_to_insert = [
                {
                    "client_id":    client_id,
                    "email":        entry["email"],
                    "contacted_at": entry["sent_date"],
                    "source":       "auto_sync",
                    # chaser_contacted_at must be present on EVERY row (null when
                    # not chased): PostgREST rejects bulk inserts whose rows have
                    # differing keys, so omitting it on unchased rows failed any
                    # mixed batch — the fallback then stripped the column from all
                    # rows and chaser dates were silently never stored.
                    "chaser_contacted_at": entry.get("chaser_date") or None,
                    **({"campaign_name": campaign_name} if campaign_name else {}),
                }
                for entry in chunk
            ]
            inserted = False
            try:
                r = http_req.post(
                    f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                    headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
                    json=rows_to_insert,
                    timeout=30,
                )
                if r.status_code not in (200, 201, 204, 409):
                    r.raise_for_status()
                inserted = True
            except Exception:
                pass

            if not inserted:
                # Retry without chaser_contacted_at — the column may not exist yet
                # (requires: ALTER TABLE contacted_prospects ADD COLUMN IF NOT EXISTS
                # chaser_contacted_at date).  Base columns always exist so this
                # fallback guarantees the send data is stored regardless.
                try:
                    base_rows = [
                        {k: v for k, v in row.items() if k != "chaser_contacted_at"}
                        for row in rows_to_insert
                    ]
                    r2 = http_req.post(
                        f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                        headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
                        json=base_rows,
                        timeout=30,
                    )
                    if r2.status_code not in (200, 201, 204, 409):
                        r2.raise_for_status()
                    inserted = True
                except Exception:
                    pass

            if inserted:
                result["contacted_added"] += len(chunk)

    # ── Remove stale auto_sync rows for this campaign ─────────────────
    # If a prospect was previously synced as sent but the tick has since been
    # removed (or the row deleted from the sheet), its old auto_sync row would
    # keep inflating the send stats forever.  Compare the campaign's existing
    # rows against the current sent set and delete anything no longer ticked.
    # Guarded by sent_read_ok so a transient read failure never wipes history.
    if sent_read_ok and campaign_name:
        current_emails = {e["email"] for e in unique_sent}
        try:
            existing: list[str] = []
            offset = 0
            while True:   # paginate — Supabase caps responses at 1000 rows
                rr = http_req.get(
                    f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                    headers=_sb_headers(),
                    params={
                        "select":        "email",
                        "client_id":     f"eq.{client_id}",
                        "source":        "eq.auto_sync",
                        "campaign_name": f"eq.{campaign_name}",
                        "limit":         "1000",
                        "offset":        str(offset),
                    },
                    timeout=30,
                )
                rr.raise_for_status()
                page = [row.get("email", "") for row in rr.json()]
                existing.extend(p for p in page if p)
                if len(page) < 1000:
                    break
                offset += 1000

            stale = [e for e in existing if e not in current_emails]
            for i in range(0, len(stale), CHUNK_SIZE):
                stale_chunk = stale[i : i + CHUNK_SIZE]
                stale_list = "(" + ",".join(f'"{e}"' for e in stale_chunk) + ")"
                http_req.delete(
                    f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                    headers=_sb_headers(),
                    params={
                        "client_id":     f"eq.{client_id}",
                        "source":        "eq.auto_sync",
                        "campaign_name": f"eq.{campaign_name}",
                        "email":         f"in.{stale_list}",
                    },
                    timeout=30,
                )
        except Exception:
            pass   # Non-fatal — stale rows are retried on the next sync

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

    # Auto-stamp completed_at the first time this campaign hits 100 %
    if total_prospects and new_sent >= total_prospects and not completed_at:
        patch_body["completed_at"] = _datetime.now(_tz.utc).isoformat()

    # chaser_count is always included (so unticking every chaser resets it to
    # 0), but the column may not exist yet — and PostgREST rejects the ENTIRE
    # patch when any column is unknown, which would silently freeze sent_count
    # and every other counter.  On failure, retry with the base body alone.
    try:
        pr = http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"id": f"eq.{campaign_id}"},
            json={**patch_body, "chaser_count": chaser_count},
            timeout=10,
        )
        patch_ok = pr.status_code < 400
    except Exception:
        patch_ok = False

    if not patch_ok:
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
