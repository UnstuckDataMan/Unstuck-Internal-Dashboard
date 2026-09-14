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

# ── How much a single run is allowed to do ────────────────────────────────────
# A real Smartlead account accumulates campaigns indefinitely — the live one has
# 1082, of which 34 are ACTIVE. Paging every campaign's statistics every hour
# would cost thousands of API calls, blow through the rate limit, and take
# longer than the hour it has to finish in, so runs would pile up on each other.
#
# Each run therefore does bounded work:
#   * DRAFTED / ARCHIVED — never synced. Drafts have no sends; archived
#     campaigns were deliberately shelved.
#   * ACTIVE             — synced every run. This is what clients are watching.
#   * everything else    — a rolling slice per run, least-recently-synced first
#     (never-synced first of all), so history backfills steadily and stays
#     reasonably fresh without a thundering herd.
#
# With the live numbers that is 34 + 25 = 59 campaigns per run instead of 1082.
# Raise PULSE_BACKFILL_PER_RUN temporarily to pull historical data in faster.
SKIP_STATUSES = frozenset({"DRAFTED", "ARCHIVED"})
LIVE_STATUSES = frozenset({"ACTIVE"})

try:
    _BACKFILL_PER_RUN = max(0, int(os.environ.get("PULSE_BACKFILL_PER_RUN", "25")))
except ValueError:
    _BACKFILL_PER_RUN = 25


def select_campaigns_for_run(stored: list[dict]) -> list[dict]:
    """Pick the bounded set of campaigns this run will sync.

    `stored` must already be ordered least-recently-synced first — see
    store.campaigns_for_sync(). Returns live campaigns plus the rolling slice.
    """
    live, rolling = [], []
    for row in stored:
        status = str(row.get("status") or "").upper()
        if status in SKIP_STATUSES:
            continue
        (live if status in LIVE_STATUSES else rolling).append(row)
    return live + rolling[:_BACKFILL_PER_RUN]

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

        # Refresh the campaign list first (names and statuses change), then
        # decide from the STORED rows which subset to actually pull events for
        # — the stored rows carry last_synced_at, which the rolling order needs.
        remote = connector.fetch_campaigns()
        store.upsert_campaigns(
            remote, source_tool=source_tool, channel=connector.CHANNEL,
        )

        stored = store.campaigns_for_sync(source_tool)
        selected = select_campaigns_for_run(stored)
        logger.info(
            "Pulse %s: %d campaigns known, %d selected for this run.",
            source_tool, len(stored), len(selected),
        )

        synced_ids: list[str] = []
        for row in selected:
            external_id = str(row.get("external_campaign_id") or "")
            label = row.get("name") or external_id
            if not external_id:
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
                failures.append(f"{label}: {exc}")
                continue

            events_inserted += store.insert_events(events)
            campaigns_synced += 1
            # Stamped only after a successful sync, so a failing campaign stays
            # at the head of the rolling order and is retried next run rather
            # than being rotated to the back as though it had succeeded.
            synced_ids.append(str(row["id"]))
            if _PACING_SECONDS:
                time.sleep(_PACING_SECONDS)

        store.mark_campaigns_synced(synced_ids)

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
