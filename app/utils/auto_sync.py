"""
Scheduled twice-daily sync for all campaigns that have a linked Google Sheet.

Runs for every client automatically — no manual trigger needed.
Times are controlled by the AUTO_SYNC_HOURS env var (comma-separated UTC hours,
default "9,21" → 09:00 and 21:00 UTC).
"""
from __future__ import annotations

import logging
import os
import threading
import time
from datetime import date as _date, datetime as _datetime, timedelta as _timedelta, timezone as _tz

import requests as http_req

from app.utils.supabase import (
    SUPABASE_URL,
    sb_headers as _sb_headers,
    sb_configured as _sb_configured,
    pg_in_list as _pg_in_list,
)

logger = logging.getLogger(__name__)

CHUNK_SIZE = 500

# Pause between campaigns in the sequential auto-sync loop so Sheets reads spread
# across the minute instead of bursting past the 60-reads/min-per-user quota
# (the whole app shares one service account = one "user").  ~1 campaign ≈ 1–2
# reads, so ~1.5s/campaign keeps a run comfortably under the limit. Tune via
# AUTO_SYNC_PACING_MS; 0 disables pacing.
try:
    _PACING_SECONDS = max(0.0, int(os.environ.get("AUTO_SYNC_PACING_MS", "1500")) / 1000.0)
except ValueError:
    _PACING_SECONDS = 1.5


# ── Auto-sync run status ──────────────────────────────────────────────────────
# Records the most recent scheduled run so the Campaigns panel can show that the
# hourly sync actually fired (and how long ago).  In-memory + a lock: this is a
# single-worker deployment, and the value is display-only — it is recomputed on
# every run, never read for correctness.
_last_run: dict = {
    "started_at":  None,   # ISO-8601 UTC of the most recent run start
    "finished_at": None,   # ISO-8601 UTC of completion (None while running)
    "running":     False,
    "ok":          0,
    "failed":      0,
    "total":       0,
    "duration_s":  None,
    "failures":    [],     # [{name, client, status, reason}] for the last run
    "sheet_reads": None,   # approx Sheets read API calls consumed by the run
}
_last_run_lock = threading.Lock()
_MAX_FAILURES_KEPT = 100


def get_sync_status() -> dict:
    """Return a copy of the most recent auto-sync run status (display-only)."""
    with _last_run_lock:
        snap = dict(_last_run)
        snap["failures"] = list(_last_run["failures"])
        return snap


def _next_working_day(from_date: _date) -> _date:
    """Return the first Monday–Friday after from_date (skips weekends)."""
    from datetime import timedelta
    d = from_date + timedelta(days=1)
    while d.weekday() >= 5:   # 5 = Saturday, 6 = Sunday
        d += timedelta(days=1)
    return d


# ── Core per-campaign sync (shared with the manual endpoint) ──────────────────

# One lock per campaign: the sync below is a non-atomic delete-then-insert on
# contacted_prospects, so two overlapping syncs of the same campaign (scheduled
# job + a user's Sync click, or a double-click) could interleave the delete and
# insert phases and drop rows.  Overlapping callers skip instead of waiting —
# the running sync already does the work they wanted.
_sync_locks: dict[str, threading.Lock] = {}
_sync_locks_guard = threading.Lock()


def _campaign_lock(campaign_id: str) -> threading.Lock:
    with _sync_locks_guard:
        lock = _sync_locks.get(campaign_id)
        if lock is None:
            lock = _sync_locks[campaign_id] = threading.Lock()
        return lock


def sync_campaign_core(
    campaign_id:     str,
    sheet_id:        str,
    client_id:       str,
    campaign_name:   str,
    client_name:     str = "",
    total_prospects: int = 0,
    completed_at:    str | None = None,
) -> dict:
    """Locked wrapper around _sync_campaign_core_unlocked — see its docstring."""
    lock = _campaign_lock(str(campaign_id))
    if not lock.acquire(blocking=False):
        return {
            "leads_added": 0, "interested_added": 0,
            "unsubscribes_added": 0, "reply_count": 0,
            "contacted_added": 0,
            "error": "Sync already running for this campaign — skipped.",
        }
    try:
        return _sync_campaign_core_unlocked(
            campaign_id, sheet_id, client_id, campaign_name,
            client_name=client_name,
            total_prospects=total_prospects, completed_at=completed_at,
        )
    finally:
        lock.release()


def _sync_campaign_core_unlocked(
    campaign_id:     str,
    sheet_id:        str,
    client_id:       str,
    campaign_name:   str,
    client_name:     str = "",
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
        read_stats_cells,
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
            # Normalise BEFORE storing: scrub lookups compare lowercased
            # values, so a mixed-case DNC entry would never match a scrub.
            email  = str(entry["email"] or "").strip().lower()
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
            # Normalise so contacted rows match the lowercased scrub lookups
            # and the (client_id, email) unique constraint can't case-dupe.
            norm = str(entry.get("email") or "").strip().lower()
            if not norm:
                continue
            entry = dict(entry)
            entry["email"] = norm
            cur = by_email.get(norm)
            if cur is None:
                by_email[norm] = entry
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

    # ── Diff against existing rows: only rewrite what changed ─────────
    # The delete-all-then-reinsert of EVERY sent row each run was the main
    # Supabase churn (it scaled with a campaign's lifetime volume, not with
    # what's new).  The stale-cleanup fetch below already pages through the
    # campaign's existing auto_sync rows every run — extend it to include the
    # stored dates and use it as a diff base, so unchanged rows are skipped
    # entirely.  On any fetch failure we fall back to rewriting everything
    # (the old, always-correct behaviour).
    #
    # NOTE: an on_conflict upsert is NOT an option here — the unique index is
    # an expression index (client_id, email, COALESCE(campaign_name,'')) which
    # PostgREST's on_conflict parameter cannot target (see the 8ce9e75 →
    # d20c389 incident where a mismatched on_conflict silently stopped sends
    # from recording).
    existing_rows: dict[str, tuple] | None = None   # email → (contacted_at, chaser); None = unknown
    if sent_read_ok and campaign_name:
        def _fetch_existing(select: str) -> dict[str, tuple]:
            found: dict[str, tuple] = {}
            offset = 0
            while True:   # paginate — Supabase caps responses at 1000 rows
                rr = http_req.get(
                    f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                    headers=_sb_headers(),
                    params={
                        "select":        select,
                        "client_id":     f"eq.{client_id}",
                        "source":        "eq.auto_sync",
                        "campaign_name": f"eq.{campaign_name}",
                        "limit":         "1000",
                        "offset":        str(offset),
                    },
                    timeout=30,
                )
                rr.raise_for_status()
                page = rr.json()
                for row in page:
                    em = (row.get("email") or "").strip()
                    if em:
                        found[em] = (row.get("contacted_at") or "",
                                     row.get("chaser_contacted_at") or None)
                if len(page) < 1000:
                    return found
                offset += 1000

        try:
            existing_rows = _fetch_existing("email,contacted_at,chaser_contacted_at")
        except Exception:
            try:   # chaser column may not exist yet — retry without it
                existing_rows = _fetch_existing("email,contacted_at")
            except Exception:
                existing_rows = None   # unknown — rewrite everything below

    if existing_rows is None:
        to_write = unique_sent
    else:
        to_write = [
            e for e in unique_sent
            if e["email"] not in existing_rows
            or existing_rows[e["email"]][0] != (e.get("sent_date") or "")
            or existing_rows[e["email"]][1] != (e.get("chaser_date") or None)
        ]

    if to_write:
        for i in range(0, len(to_write), CHUNK_SIZE):
            chunk = to_write[i : i + CHUNK_SIZE]
            email_list = _pg_in_list([e["email"] for e in chunk])

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
    # keep inflating the send stats forever.  The diff fetch above already
    # holds every existing row for this campaign, so stale detection is free —
    # no second paginated fetch.  Guarded by sent_read_ok (via existing_rows
    # being None on read failure) so a transient failure never wipes history.
    if sent_read_ok and campaign_name and existing_rows is not None:
        current_emails = {e["email"] for e in unique_sent}
        try:
            stale = [e for e in existing_rows if e not in current_emails]
            for i in range(0, len(stale), CHUNK_SIZE):
                stale_chunk = stale[i : i + CHUNK_SIZE]
                stale_list = _pg_in_list(stale_chunk)
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

    # ── Cache stats-tab cell values (Test client only for initial rollout) ──
    # Remove the client_name guard to roll out to all clients.
    if client_name == "Test":
        stats = read_stats_cells(sheet_id)
        if stats:
            try:
                http_req.patch(
                    f"{SUPABASE_URL}/rest/v1/campaigns",
                    headers=_sb_headers("return=minimal"),
                    params={"id": f"eq.{campaign_id}"},
                    json=stats,
                    timeout=10,
                )
            except Exception as exc:
                logger.warning("stats_cells PATCH failed for %s: %s", campaign_id, exc)

    return result


# ── Scheduled job ─────────────────────────────────────────────────────────────

def _mark_run_start() -> _datetime:
    from app.utils.google_sheets import reset_sheet_reads
    reset_sheet_reads()   # zero the Sheets-read counter for this run
    now = _datetime.now(_tz.utc)
    with _last_run_lock:
        _last_run.update(started_at=now.isoformat(), finished_at=None,
                         running=True, ok=0, failed=0, total=0, duration_s=None,
                         failures=[], sheet_reads=None)
    return now


def _mark_run_end(started: _datetime, ok: int, failed: int, total: int,
                  failures: list | None = None) -> None:
    from app.utils.google_sheets import sheet_reads_count
    end = _datetime.now(_tz.utc)
    with _last_run_lock:
        _last_run.update(finished_at=end.isoformat(), running=False,
                         ok=ok, failed=failed, total=total,
                         duration_s=round((end - started).total_seconds(), 1),
                         failures=(failures or [])[:_MAX_FAILURES_KEPT],
                         sheet_reads=sheet_reads_count())


# Select tiers for the campaigns fetch — try the richest first and fall back so
# a schema missing the newer `paused`/`completed` columns still works (matches
# the resilience pattern in campaigns.list_campaigns).
_CAMPAIGN_SELECT_TIERS = (
    "id,sheet_id,client_id,client_name,campaign_name,total_prospects,completed_at,sent_count,completed,paused",
    "id,sheet_id,client_id,client_name,campaign_name,total_prospects,completed_at,sent_count",
    "id,sheet_id,client_id,client_name,campaign_name,total_prospects,completed_at",
)


def _campaign_status(c: dict) -> str:
    """Classify a campaign row as active / paused / past (best-effort — falls
    back gracefully when the status columns aren't in the fetched select)."""
    if c.get("paused"):
        return "paused"
    total = c.get("total_prospects") or 0
    sent  = c.get("sent_count") or 0
    if c.get("completed") or c.get("completed_at") or (total > 0 and sent >= total):
        return "past"
    return "active"


# Only one full run at a time — a trigger (hourly tick or a manual "Run now")
# that arrives while a run is in progress is skipped, since the running pass
# already does the same work.
_run_lock = threading.Lock()


def run_auto_sync() -> None:
    """
    Runs the hourly campaign sync (see the scheduler thread below, and the
    manual /api/campaigns/sync-now trigger).
    Fetches every campaign that has a sheet_id and syncs them one at a time —
    sequential on purpose, to stay under the Google Sheets read quota and avoid
    hammering Supabase with a burst.  Records a run-status snapshot (start/end,
    counts) so the Campaigns panel can show the sync actually ran.
    """
    if not _run_lock.acquire(blocking=False):
        logger.info("Auto-sync: a run is already in progress — skipping this trigger.")
        return
    try:
        _run_auto_sync_locked()
    finally:
        _run_lock.release()


def _run_auto_sync_locked() -> None:
    from app.utils.google_sheets import is_configured

    started = _mark_run_start()
    ok = failed = 0
    failures: list[dict] = []
    campaigns: list[dict] = []
    try:
        if not _sb_configured():
            logger.warning("Auto-sync skipped: Supabase not configured.")
            return
        if not is_configured():
            logger.warning("Auto-sync skipped: Google Sheets not configured.")
            return

        # Fetch all campaigns that have a linked sheet (tiered select so a schema
        # missing the newer status columns still works).
        last_exc = None
        for _sel in _CAMPAIGN_SELECT_TIERS:
            try:
                r = http_req.get(
                    f"{SUPABASE_URL}/rest/v1/campaigns",
                    headers=_sb_headers(),
                    params={"select": _sel, "sheet_id": "not.is.null"},
                    timeout=15,
                )
                r.raise_for_status()
                campaigns = [c for c in r.json() if (c.get("sheet_id") or "").strip()]
                break
            except Exception as exc:
                last_exc = exc
                campaigns = []
        else:
            logger.error("Auto-sync: failed to fetch campaigns: %s", last_exc)
            return

        if not campaigns:
            logger.info("Auto-sync: no campaigns with linked sheets — nothing to do.")
            return

        logger.info("Auto-sync: starting sync for %d campaign(s) across all clients (pacing %.1fs).",
                    len(campaigns), _PACING_SECONDS)

        # Sequential by design — one campaign finishes before the next starts,
        # with a pause between them so Sheets reads stay under the per-user quota.
        for i, c in enumerate(campaigns):
            if i > 0 and _PACING_SECONDS > 0:
                time.sleep(_PACING_SECONDS)

            cid    = c.get("id", "")
            name   = c.get("campaign_name", "") or "(unnamed)"
            client = c.get("client_name", "") or ""
            status = _campaign_status(c)

            def _record_failure(reason: str) -> None:
                failures.append({"name": name, "client": client,
                                 "status": status, "reason": str(reason)[:200]})

            try:
                summary = sync_campaign_core(
                    campaign_id=cid,
                    sheet_id=c["sheet_id"],
                    client_id=c.get("client_id", ""),
                    campaign_name=name,
                    client_name=client,
                    total_prospects=c.get("total_prospects") or 0,
                    completed_at=c.get("completed_at"),
                )
                if summary["error"]:
                    logger.warning("Auto-sync: campaign %s (%s) error — %s", cid, name, summary["error"])
                    failed += 1
                    _record_failure(summary["error"])
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
                _record_failure(exc)

        logger.info("Auto-sync complete: %d succeeded, %d failed.", ok, failed)
    finally:
        _mark_run_end(started, ok, failed, len(campaigns), failures)


# ── In-process hourly scheduler ───────────────────────────────────────────────
# A plain daemon thread rather than APScheduler: APScheduler's AsyncIOScheduler
# fires jobs off the uvicorn event loop, which proved unreliable here (the first
# run fired but hourly ticks silently stopped).  A thread that simply sleeps
# until the next top-of-hour has no dependency on the event loop and fires
# reliably for as long as the web process is alive.
_scheduler_thread: threading.Thread | None = None
_scheduler_stop = threading.Event()
_STARTUP_DELAY_S = 20   # let the app finish booting before the first run


def _seconds_to_next_hour() -> float:
    now = _datetime.now(_tz.utc)
    nxt = now.replace(minute=0, second=0, microsecond=0) + _timedelta(hours=1)
    return max(1.0, (nxt - now).total_seconds())


def _scheduler_loop() -> None:
    # Run once shortly after boot so a fresh deploy reflects immediately instead
    # of waiting up to an hour for the first top-of-hour tick.
    if _scheduler_stop.wait(_STARTUP_DELAY_S):
        return
    while not _scheduler_stop.is_set():
        try:
            run_auto_sync()
        except Exception as exc:   # never let the loop die on one bad run
            logger.error("Auto-sync scheduler loop error: %s", exc)
        # Interruptible sleep until the next top of the hour.
        if _scheduler_stop.wait(_seconds_to_next_hour()):
            break


def start_scheduler() -> None:
    """Start the hourly auto-sync thread (idempotent)."""
    global _scheduler_thread
    if _scheduler_thread and _scheduler_thread.is_alive():
        return
    _scheduler_stop.clear()
    _scheduler_thread = threading.Thread(
        target=_scheduler_loop, name="auto-sync-scheduler", daemon=True,
    )
    _scheduler_thread.start()


def stop_scheduler() -> None:
    """Signal the scheduler thread to exit (used on app shutdown)."""
    _scheduler_stop.set()
