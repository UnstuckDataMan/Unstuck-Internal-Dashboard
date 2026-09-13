"""
Outbound Pulse sync orchestration.

Runs each connector, writes normalized events, and records every run to
pulse_sync_logs so a broken connector is visible internally before a client
notices a stale dashboard.

The schedule is a plain daemon thread, matching app/utils/auto_sync.py — that
file documents why (APScheduler's AsyncIOScheduler fired once under uvicorn and
then silently stopped).  Using the same mechanism here keeps one scheduling
pattern in the codebase rather than two.

Sync is idempotent by construction: events carry a dedupe key with a unique
index behind it, so re-running over a window already synced inserts nothing.
That is what lets the manual trigger be safe to hammer while validating a
connector, and what lets a failed run simply be retried.
"""
from __future__ import annotations

import logging
import os
import threading
import time
from datetime import datetime, timedelta, timezone

from app.utils.pulse import meet_alfred, smartlead, store
from app.utils.pulse.normalize import SOURCE_MEET_ALFRED, SOURCE_SMARTLEAD

logger = logging.getLogger(__name__)

# Connector registry. Adding a third tool means adding an entry here and a
# module with the same three names — nothing else in the module knows the
# difference between an email tool and a LinkedIn one.
CONNECTORS = {
    SOURCE_SMARTLEAD:   smartlead,
    SOURCE_MEET_ALFRED: meet_alfred,
}

# Same cadence as the existing campaign auto-sync: hourly, on the hour.
_SYNC_INTERVAL_HOURS = 1
_STARTUP_DELAY_S = 45   # after auto_sync's 20s, so the two don't boot together
_PACING_SECONDS = max(0.0, float(os.environ.get("PULSE_PACING_MS", "500")) / 1000.0)

# In-memory view of the run in progress, for the status pill.  Display-only and
# recomputed every run — the durable record is pulse_sync_logs.
_running: dict[str, bool] = {}
_running_lock = threading.Lock()


def is_running(source_tool: str = "") -> bool:
    with _running_lock:
        if source_tool:
            return bool(_running.get(source_tool))
        return any(_running.values())


def _mark(source_tool: str, running: bool) -> None:
    with _running_lock:
        _running[source_tool] = running


# ── One connector run ─────────────────────────────────────────────────────────

def sync_source(source_tool: str, triggered_by: str = "schedule") -> dict:
    """Sync every campaign for one connector.

    Returns a summary dict.  Never raises: a connector failure is data the
    dashboard needs to show, not an exception for the caller to handle.
    """
    connector = CONNECTORS.get(source_tool)
    if connector is None:
        return {"source_tool": source_tool, "status": "error",
                "error": f"unknown connector: {source_tool}",
                "campaigns": 0, "events": 0}

    if not connector.is_configured():
        return {"source_tool": source_tool, "status": "skipped",
                "error": f"{source_tool} is not configured (API key missing).",
                "campaigns": 0, "events": 0}

    if is_running(source_tool):
        # Overlapping runs would do identical work; the one already going gets it.
        return {"source_tool": source_tool, "status": "skipped",
                "error": "A sync for this connector is already running.",
                "campaigns": 0, "events": 0}

    _mark(source_tool, True)
    started = time.monotonic()
    log_id = store.start_sync_log(source_tool, triggered_by)

    campaigns_synced = 0
    events_inserted = 0
    failures: list[str] = []

    try:
        agency_id = store.current_agency_id()
        remote = connector.fetch_campaigns()

        # Existing mappings, so a sync never blanks a client a human assigned.
        known = {
            str(c.get("external_campaign_id")): c
            for c in store.list_campaigns(source_tool=source_tool)
        }

        for entry in remote:
            external_id = entry["external_id"]
            existing = known.get(external_id)
            row = store.upsert_campaign(
                source_tool=source_tool,
                external_campaign_id=external_id,
                channel=connector.CHANNEL,
                name=entry["name"],
                status=entry["status"],
                client_id=(existing or {}).get("client_id"),
                raw=entry.get("raw") or {},
            )
            if not row:
                failures.append(f"{entry['name'] or external_id}: campaign upsert failed")
                continue

            try:
                events = connector.sync_campaign(
                    external_campaign_id=external_id,
                    agency_id=agency_id,
                    campaign_id=str(row["id"]),
                    client_id=row.get("client_id"),
                )
            except Exception as exc:
                logger.warning("Pulse %s: campaign %s failed: %s",
                               source_tool, external_id, exc)
                failures.append(f"{entry['name'] or external_id}: {exc}")
                continue

            events_inserted += store.insert_events(events)
            campaigns_synced += 1
            if _PACING_SECONDS:
                time.sleep(_PACING_SECONDS)

        if failures and campaigns_synced:
            status = "partial"
        elif failures:
            status = "error"
        else:
            status = "ok"
        error_message = "; ".join(failures[:20])

    except Exception as exc:
        logger.error("Pulse %s: sync failed: %s", source_tool, exc)
        status = "error"
        error_message = str(exc)
    finally:
        _mark(source_tool, False)

    duration = time.monotonic() - started
    store.finish_sync_log(
        log_id,
        status=status,
        campaigns_synced=campaigns_synced,
        events_inserted=events_inserted,
        duration_s=duration,
        error_message=error_message,
    )
    logger.info(
        "Pulse %s sync %s — %d campaigns, %d new events, %.1fs",
        source_tool, status, campaigns_synced, events_inserted, duration,
    )
    return {
        "source_tool": source_tool,
        "status":      status,
        "campaigns":   campaigns_synced,
        "events":      events_inserted,
        "duration_s":  round(duration, 1),
        "error":       error_message,
    }


def sync_all(triggered_by: str = "schedule") -> list[dict]:
    """Run every configured connector, sequentially."""
    results = []
    for source_tool in CONNECTORS:
        results.append(sync_source(source_tool, triggered_by))
    return results


def run_scheduled_sync() -> None:
    try:
        store.backfill_client_agency()
    except Exception as exc:
        logger.warning("Pulse: agency backfill skipped: %s", exc)
    sync_all("schedule")


# ── In-process hourly scheduler ───────────────────────────────────────────────

_scheduler_thread: threading.Thread | None = None
_scheduler_stop = threading.Event()


def _seconds_to_next_hour() -> float:
    now = datetime.now(timezone.utc)
    nxt = now.replace(minute=0, second=0, microsecond=0) + timedelta(hours=_SYNC_INTERVAL_HOURS)
    return max(1.0, (nxt - now).total_seconds())


def _scheduler_loop() -> None:
    if _scheduler_stop.wait(_STARTUP_DELAY_S):
        return
    while not _scheduler_stop.is_set():
        try:
            run_scheduled_sync()
        except Exception as exc:   # never let the loop die on one bad run
            logger.error("Pulse scheduler loop error: %s", exc)
        if _scheduler_stop.wait(_seconds_to_next_hour()):
            break


def start_scheduler() -> None:
    """Start the hourly Pulse sync thread (idempotent).

    A no-op when neither connector has an API key: without one there is nothing
    to fetch, and starting anyway would write a failed sync log every hour and
    bury real breakage in noise.
    """
    global _scheduler_thread
    if not any(c.is_configured() for c in CONNECTORS.values()):
        logger.info("Pulse: no connector configured — scheduler not started.")
        return
    if _scheduler_thread and _scheduler_thread.is_alive():
        return
    _scheduler_stop.clear()
    _scheduler_thread = threading.Thread(
        target=_scheduler_loop, name="pulse-sync-scheduler", daemon=True,
    )
    _scheduler_thread.start()


def stop_scheduler() -> None:
    _scheduler_stop.set()
