"""Offline tests for campaign stats tracking.

Covers the send-count queries (source filter, date buckets, chaser counts),
the parallel time breakdown, and the full sync_campaign_core counter pipeline
with a mocked sheet + mocked Supabase.
"""
from __future__ import annotations

from datetime import date, timedelta

from conftest import FakeResponse, in_list_values, param_values, params_list

import app.routers.campaigns as campaigns
import app.utils.auto_sync as auto_sync
import app.utils.google_sheets as gs
from app.utils.dates import today_utc   # same clock the app uses — never date.today()


def count_response(n: int):
    return FakeResponse(200, [], headers={"Content-Range": f"0-0/{n}"})


# ── _count_contacted / _count_chaser_contacted ────────────────────────────────

def test_count_contacted_filters_to_auto_sync_source(fake_sb):
    fake_sb.route("GET", "contacted_prospects", lambda c: count_response(42))
    n = campaigns._count_contacted("client-1", date_eq="2026-06-12")
    assert n == 42
    call = fake_sb.calls_to("GET", "contacted_prospects")[0]
    assert "eq.auto_sync" in param_values(call, "source"), \
        "send stats MUST only count auto_sync rows (scrub uploads would inflate them)"
    assert "eq.2026-06-12" in param_values(call, "contacted_at")
    assert call["headers"].get("Prefer") == "count=exact"


def test_count_contacted_campaign_name_filter_quotes_punctuation(fake_sb):
    fake_sb.route("GET", "contacted_prospects", lambda c: count_response(1))
    campaigns._count_contacted("client-1", campaign_names=['June, "Q2" Push'])
    call = fake_sb.calls_to("GET", "contacted_prospects")[0]
    raw = param_values(call, "campaign_name")[0]
    # Commas/quotes inside the name must be double-quoted and escaped, or
    # PostgREST would split the name into two bogus filter values.
    assert raw == 'in.("June, \\"Q2\\" Push")'


def test_count_chaser_uses_chaser_date_and_survives_missing_column(fake_sb):
    fake_sb.route("GET", "contacted_prospects", lambda c: count_response(7))
    n = campaigns._count_chaser_contacted("client-1", date_gte="2026-06-01",
                                          date_lte="2026-06-12")
    assert n == 7
    call = fake_sb.calls_to("GET", "contacted_prospects")[0]
    chaser_filters = param_values(call, "chaser_contacted_at")
    assert "not.is.null" in chaser_filters
    assert "gte.2026-06-01" in chaser_filters and "lte.2026-06-12" in chaser_filters

    # Column missing → Supabase 400 → graceful zero, never an exception
    fake_sb.route("GET", "contacted_prospects",
                  lambda c: FakeResponse(400, [], text="column does not exist"))
    assert campaigns._count_chaser_contacted("client-1") == 0


# ── _time_breakdown (parallelised buckets) ────────────────────────────────────

def test_time_breakdown_sums_initial_plus_chasers_per_bucket(fake_sb):
    def handler(call):
        is_chaser = bool(param_values(call, "chaser_contacted_at"))
        return count_response(2 if is_chaser else 5)

    fake_sb.route("GET", "contacted_prospects", handler)
    out = campaigns._time_breakdown("client-1")

    assert out["today"] == 7 and out["week"] == 7
    assert out["month"] == 7 and out["all_time"] == 7
    assert len(fake_sb.calls_to("GET", "contacted_prospects")) == 8  # 4 buckets × 2

    today  = today_utc()
    monday = today - timedelta(days=today.weekday())
    friday = monday + timedelta(days=4)
    today_calls = [c for c in fake_sb.calls_to("GET", "contacted_prospects")
                   if f"eq.{today.isoformat()}" in param_values(c, "contacted_at")]
    assert today_calls, "today bucket must filter contacted_at = today"
    assert monday.strftime("%b") in out["week_label"]

    # The working week ends FRIDAY — sends never go out at weekends, so a
    # Sunday end-boundary would widen the window and mislabel the range.
    week_calls = [c for c in fake_sb.calls_to("GET", "contacted_prospects")
                  if f"gte.{monday.isoformat()}" in param_values(c, "contacted_at")]
    assert week_calls, "week bucket must start on Monday"
    assert all(f"lte.{friday.isoformat()}" in param_values(c, "contacted_at")
               for c in week_calls), "week bucket must end on Friday, not Sunday"
    assert str(friday.day) in out["week_label"]


# ── sync_campaign_core: counters, date back-fill, stale cleanup ───────────────

def _patch_sheet_reads(monkeypatch, leads, sent):
    monkeypatch.setattr(gs, "read_leads", lambda sid: list(leads))
    monkeypatch.setattr(gs, "read_sent_with_dates", lambda sid: [dict(e) for e in sent])
    written = {"sent": [], "chaser": []}
    monkeypatch.setattr(gs, "write_sent_dates",
                        lambda sid, pairs: written["sent"].extend(pairs))
    monkeypatch.setattr(gs, "write_chaser_dates",
                        lambda sid, pairs: written["chaser"].extend(pairs))
    return written


LEADS = [
    {"email": "lead1@acme.com",  "status": "Lead"},
    {"email": "lead2@acme.com",  "status": "Lead"},        # same domain → dedupe
    {"email": "int@corp.com",    "status": "Interested"},
    {"email": "unsub@corp.com",  "status": "Unsubscribe"},
    {"email": "reply@corp.com",  "status": "Reply"},       # counted, never blocked
]
SENT = [
    {"email": "lead1@acme.com", "sent_date": "2026-06-01", "row_num": 2,
     "chaser_sent": True,  "chaser_date": ""},
    {"email": "fresh@corp.com", "sent_date": "",           "row_num": 3,
     "chaser_sent": False, "chaser_date": ""},
]


def test_sync_core_stamps_today_not_a_future_date(fake_sb, monkeypatch):
    """Regression: a row ticked sent with a blank Sent Date must be stamped with
    TODAY, never a future date.

    The sync used to stamp the *next working day*, so a send ticked today landed
    on tomorrow (or, on a Friday, on next Monday). That pushed every same-day
    send out of the Today bucket and out of the month-to-date window, and moved
    Friday's sends into the following week.
    """
    today = today_utc()
    written = _patch_sheet_reads(monkeypatch, [], SENT)
    fake_sb.route("POST", "contacted_prospects", lambda c: FakeResponse(201))
    fake_sb.route("GET",  "contacted_prospects", lambda c: FakeResponse(200, []))
    fake_sb.route("PATCH", "campaigns", lambda c: FakeResponse(204))

    auto_sync.sync_campaign_core("camp-1", "sheet-1", "client-1", "June Campaign")

    # Written to the sheet: exactly today, for both Sent Date and Chaser Date.
    stamped = [d for _row, d in written["sent"] + written["chaser"]]
    assert stamped, "the blank Sent Date row should have been back-filled"
    assert set(stamped) == {today.isoformat()}
    assert all(date.fromisoformat(d) <= today for d in stamped), \
        "a send must never be stamped with a future date"

    # And the same date reaches contacted_at, which is what the stats count.
    ins = fake_sb.calls_to("POST", "contacted_prospects")[0]["json"]
    fresh = next(r for r in ins if r["email"] == "fresh@corp.com")
    assert fresh["contacted_at"] == today.isoformat()


def test_sync_core_full_pipeline(fake_sb, monkeypatch):
    # A ticked row with no Sent Date is stamped TODAY — the day it was ticked.
    send_day = today_utc().isoformat()
    written = _patch_sheet_reads(monkeypatch, LEADS, SENT)

    fake_sb.route("POST", "dnc_entries", lambda c: FakeResponse(201))
    fake_sb.route("POST", "contacted_prospects", lambda c: FakeResponse(201))
    # Stale-cleanup pre-read: one current + one no-longer-ticked email
    fake_sb.route("GET", "contacted_prospects", lambda c: FakeResponse(200, [
        {"email": "lead1@acme.com"}, {"email": "gone@corp.com"},
    ]))
    fake_sb.route("PATCH", "campaigns", lambda c: FakeResponse(204))

    result = auto_sync.sync_campaign_core(
        "camp-1", "sheet-1", "client-1", "June Campaign",
        total_prospects=2, completed_at=None,
    )

    # Counters
    assert result == {"leads_added": 2, "interested_added": 1,
                      "unsubscribes_added": 1, "reply_count": 1,
                      "contacted_added": 2, "error": ""}

    # Date back-fill: missing Sent Date / Chaser Date stamped today
    assert written["sent"]   == [(3, send_day)]
    assert written["chaser"] == [(2, send_day)]

    # DNC rows: Lead→domain (deduped), Interested/Unsub→email, Reply absent
    dnc_post = fake_sb.calls_to("POST", "dnc_entries")[0]
    assert param_values(dnc_post, "on_conflict") == ["client_id,email"]
    assert "ignore-duplicates" in dnc_post["headers"]["Prefer"]
    rows = {r["email"]: r["reason"] for r in dnc_post["json"]}
    assert rows == {"acme.com": "lead", "int@corp.com": "interested",
                    "unsub@corp.com": "opt_out"}

    # Contacted: DELETE-then-INSERT, chaser_contacted_at on EVERY row
    deletes = fake_sb.calls_to("DELETE", "contacted_prospects")
    refresh = [d for d in deletes if param_values(d, "email")]
    assert in_list_values(param_values(refresh[0], "email")[0]) == \
        ["lead1@acme.com", "fresh@corp.com"]
    ins = fake_sb.calls_to("POST", "contacted_prospects")[0]["json"]
    by_email = {r["email"]: r for r in ins}
    assert by_email["lead1@acme.com"]["contacted_at"] == "2026-06-01"
    assert by_email["lead1@acme.com"]["chaser_contacted_at"] == send_day  # back-filled
    assert by_email["fresh@corp.com"]["contacted_at"] == send_day
    assert by_email["fresh@corp.com"]["chaser_contacted_at"] is None  # key present!
    assert all(r["campaign_name"] == "June Campaign" for r in ins)
    assert all(r["source"] == "auto_sync" for r in ins)

    # Stale cleanup deletes only the no-longer-ticked email
    stale = [d for d in deletes
             if param_values(d, "source") == ["eq.auto_sync"]
             and param_values(d, "email")
             and in_list_values(param_values(d, "email")[0]) == ["gone@corp.com"]]
    assert stale, "stale auto_sync row should be deleted"

    # Campaign counters PATCH (incl. chaser_count) + completed_at auto-stamp
    patch = fake_sb.calls_to("PATCH", "campaigns")[0]["json"]
    assert patch["sent_count"] == 2 and patch["chaser_count"] == 1
    assert patch["lead_count"] == 2 and patch["reply_count"] == 1
    assert patch["interested_count"] == 1 and patch["unsubscribe_count"] == 1
    assert "completed_at" in patch, "100% sent must auto-stamp completed_at"


def test_sync_core_skips_unchanged_rows(fake_sb, monkeypatch):
    """Steady state: every sent row already stored with identical dates →
    zero DELETE/INSERT churn on contacted_prospects (the diff optimisation)."""
    send_day = today_utc().isoformat()
    written = _patch_sheet_reads(monkeypatch, [], SENT)

    # Existing rows exactly match what the sync would write:
    # lead1: sheet date + chaser back-filled to send_day; fresh: back-filled date, no chaser.
    fake_sb.route("GET", "contacted_prospects", lambda c: FakeResponse(200, [
        {"email": "lead1@acme.com", "contacted_at": "2026-06-01",
         "chaser_contacted_at": send_day},
        {"email": "fresh@corp.com", "contacted_at": send_day,
         "chaser_contacted_at": None},
    ]))
    fake_sb.route("PATCH", "campaigns", lambda c: FakeResponse(204))

    result = auto_sync.sync_campaign_core("camp-1", "sheet-1", "client-1", "June Campaign")

    assert result["error"] == ""
    assert result["contacted_added"] == 0, "unchanged rows must not be rewritten"
    assert fake_sb.calls_to("DELETE", "contacted_prospects") == []
    assert fake_sb.calls_to("POST",   "contacted_prospects") == []
    # Counters still refresh every run
    assert fake_sb.calls_to("PATCH", "campaigns")


def test_sync_core_rewrites_when_existing_fetch_fails(fake_sb, monkeypatch):
    """If the diff fetch errors, fall back to rewriting everything (the old,
    always-correct behaviour) rather than skipping writes on unknown state."""
    _patch_sheet_reads(monkeypatch, [], SENT)
    fake_sb.route("GET", "contacted_prospects",
                  lambda c: FakeResponse(500, [], text="boom"))
    fake_sb.route("POST", "contacted_prospects", lambda c: FakeResponse(201))
    fake_sb.route("PATCH", "campaigns", lambda c: FakeResponse(204))

    result = auto_sync.sync_campaign_core("camp-1", "sheet-1", "client-1", "June Campaign")

    assert result["contacted_added"] == 2, "unknown existing state → write all"
    assert fake_sb.calls_to("POST", "contacted_prospects")
    # Stale cleanup must NOT run on unknown state (could wipe valid history)
    deletes = fake_sb.calls_to("DELETE", "contacted_prospects")
    stale_deletes = [d for d in deletes if param_values(d, "source") == ["eq.auto_sync"]]
    assert stale_deletes == []


def test_sync_core_failed_sheet_read_never_wipes_history(fake_sb, monkeypatch):
    def boom(sid):
        raise RuntimeError("quota exceeded")
    monkeypatch.setattr(gs, "read_leads", lambda sid: [])
    monkeypatch.setattr(gs, "read_sent_with_dates", boom)
    fake_sb.route("PATCH", "campaigns", lambda c: FakeResponse(204))

    result = auto_sync.sync_campaign_core("camp-1", "sheet-1", "client-1", "June")

    assert result["contacted_added"] == 0
    assert fake_sb.calls_to("DELETE", "contacted_prospects") == [], \
        "a transient read failure must never trigger the stale-row cleanup"


def test_sync_core_chaser_column_missing_falls_back(fake_sb, monkeypatch):
    """PATCH with chaser_count rejected (column absent) → base PATCH retried."""
    _patch_sheet_reads(monkeypatch, [], SENT[:1])
    fake_sb.route("POST", "contacted_prospects", lambda c: FakeResponse(201))
    fake_sb.route("GET", "contacted_prospects", lambda c: FakeResponse(200, []))

    patches = []
    def patch_handler(call):
        patches.append(call["json"])
        if "chaser_count" in call["json"]:
            return FakeResponse(400, [], text="chaser_count does not exist")
        return FakeResponse(204)
    fake_sb.route("PATCH", "campaigns", patch_handler)

    auto_sync.sync_campaign_core("camp-1", "sheet-1", "client-1", "June")
    assert len(patches) == 2
    assert "chaser_count" in patches[0] and "chaser_count" not in patches[1]
    assert patches[1]["sent_count"] == 1, "sent_count must survive the fallback"


# ── Reset client stats ────────────────────────────────────────────────────────

CAMPAIGN_ROW = {
    "id": "1", "campaign_name": "Camp", "client_id": "c1", "client_name": "X",
    "sheet_id": "s", "sheet_url": "u", "total_prospects": 10, "sent_count": 0,
    "completed": None, "completed_at": None, "tags": [], "lead_count": 0,
    "reply_count": 0, "interested_count": 0, "unsubscribe_count": 0,
    "paused": False, "chaser_count": 0, "created_at": "2026-06-01T00:00:00",
    "sender_profile_name": "P",
}


def test_reset_stats_deletes_only_auto_sync_rows_and_zeroes_counters(fake_sb, client):
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, [CAMPAIGN_ROW]))
    fake_sb.route("GET", "contacted_prospects", lambda c: count_response(0))
    fake_sb.route("DELETE", "contacted_prospects", lambda c: FakeResponse(204))
    fake_sb.route("PATCH", "campaigns", lambda c: FakeResponse(204))

    resp = client.post("/api/campaigns/reset-stats", params={"client_id": "c1"})
    assert resp.status_code == 200

    # The delete must be scoped to EXACTLY this client's auto_sync rows —
    # scrub/manual/CSV history (and everything else) must survive.
    deletes = fake_sb.calls_to("DELETE", "contacted_prospects")
    assert len(deletes) == 1
    assert dict(params_list(deletes[0])) == {
        "client_id": "eq.c1", "source": "eq.auto_sync",
    }
    assert fake_sb.calls_to("DELETE", "dnc_entries") == [], "DNC list untouched"

    patch = fake_sb.calls_to("PATCH", "campaigns")[0]
    assert dict(params_list(patch)) == {"client_id": "eq.c1"}
    assert patch["json"] == {"sent_count": 0, "lead_count": 0, "reply_count": 0,
                             "interested_count": 0, "unsubscribe_count": 0,
                             "chaser_count": 0}

    # Refreshed panel comes back with the reset button + confirmation modal
    assert "reset-stats-modal" in resp.text
    assert "Reset Stats" in resp.text


def test_reset_stats_retries_without_chaser_count_column(fake_sb, client):
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, [CAMPAIGN_ROW]))
    fake_sb.route("GET", "contacted_prospects", lambda c: count_response(0))
    fake_sb.route("DELETE", "contacted_prospects", lambda c: FakeResponse(204))

    patches = []
    def patch_handler(call):
        patches.append(call["json"])
        if "chaser_count" in call["json"]:
            return FakeResponse(400, [], text="chaser_count does not exist")
        return FakeResponse(204)
    fake_sb.route("PATCH", "campaigns", patch_handler)

    resp = client.post("/api/campaigns/reset-stats", params={"client_id": "c1"})
    assert resp.status_code == 200
    assert len(patches) == 2
    assert "chaser_count" not in patches[1]
    assert patches[1]["sent_count"] == 0


def test_reset_stats_surfaces_delete_failure_without_touching_counters(fake_sb, client):
    fake_sb.route("DELETE", "contacted_prospects",
                  lambda c: FakeResponse(500, [], text="boom"))
    resp = client.post("/api/campaigns/reset-stats", params={"client_id": "c1"})
    assert resp.status_code == 200          # HTMX swaps the visible error
    assert "Reset failed" in resp.text
    assert fake_sb.calls_to("PATCH", "campaigns") == [], \
        "counters must not be zeroed if the history delete failed"


def test_reset_stats_requires_client_id(fake_sb, client):
    resp = client.post("/api/campaigns/reset-stats")
    assert resp.status_code == 400
    assert "client_id" in resp.json()["error"]


# ── /api/campaigns dashboard stats ────────────────────────────────────────────

def test_campaigns_panel_buckets_and_pipeline(fake_sb, client):
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, [
        {"id": "1", "campaign_name": "Active",  "client_id": "c1", "client_name": "X",
         "sheet_id": "s", "sheet_url": "u", "total_prospects": 100, "sent_count": 40,
         "completed": None, "completed_at": None, "tags": [], "lead_count": 3,
         "reply_count": 2, "interested_count": 1, "unsubscribe_count": 0,
         "paused": False, "chaser_count": 5, "created_at": "2026-06-01T00:00:00",
         "sender_profile_name": "P"},
        {"id": "2", "campaign_name": "Paused",  "client_id": "c1", "client_name": "X",
         "sheet_id": "s", "sheet_url": "u", "total_prospects": 50, "sent_count": 10,
         "completed": None, "completed_at": None, "tags": [], "lead_count": 1,
         "reply_count": 0, "interested_count": 0, "unsubscribe_count": 0,
         "paused": True, "chaser_count": 0, "created_at": "2026-06-02T00:00:00",
         "sender_profile_name": "P"},
        {"id": "3", "campaign_name": "Done",    "client_id": "c1", "client_name": "X",
         "sheet_id": "s", "sheet_url": "u", "total_prospects": 20, "sent_count": 20,
         "completed": True, "completed_at": "2026-06-05T00:00:00", "tags": [],
         "lead_count": 4, "reply_count": 1, "interested_count": 0,
         "unsubscribe_count": 1, "paused": False, "chaser_count": 2,
         "created_at": "2026-06-03T00:00:00", "sender_profile_name": "P"},
    ]))
    fake_sb.route("GET", "contacted_prospects", lambda c: count_response(3))

    resp = client.get("/api/campaigns", params={"client_id": "c1"})
    assert resp.status_code == 200
    html = resp.text
    for name in ("Active", "Paused", "Done"):
        assert name in html
    # Pipeline = (100-40) from the active campaign only; paused/past excluded
    assert "60" in html
    # The auto-sync status pill renders in the panel regardless of run state.
    assert "sync-status" in html and "Auto-sync" in html


# ── Auto-sync run status ──────────────────────────────────────────────────────

def _reset_sync_status():
    with auto_sync._last_run_lock:
        auto_sync._last_run.update(started_at=None, finished_at=None, running=False,
                                   ok=0, failed=0, total=0, duration_s=None)


def test_run_auto_sync_is_sequential_and_records_status(fake_sb, monkeypatch):
    _reset_sync_status()
    rows = [
        {"id": "1", "sheet_id": "s1", "client_id": "c1", "client_name": "X",
         "campaign_name": "A", "total_prospects": 0, "completed_at": None},
        {"id": "2", "sheet_id": "s2", "client_id": "c1", "client_name": "X",
         "campaign_name": "B", "total_prospects": 0, "completed_at": None},
    ]
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, rows))

    order = []
    def fake_core(**kw):
        order.append(kw["campaign_id"])
        # second campaign reports an error → counts as a failure
        return {"error": "" if kw["campaign_id"] == "1" else "sheet boom",
                "leads_added": 0, "interested_added": 0,
                "unsubscribes_added": 0, "reply_count": 0, "contacted_added": 0}
    monkeypatch.setattr(auto_sync, "sync_campaign_core", fake_core)

    auto_sync.run_auto_sync()

    assert order == ["1", "2"], "campaigns must sync one at a time, in order"
    st = auto_sync.get_sync_status()
    assert st["running"] is False
    assert st["total"] == 2 and st["ok"] == 1 and st["failed"] == 1
    assert st["started_at"] and st["finished_at"]
    assert st["duration_s"] is not None
    # The failing campaign is recorded with its name, client, status and reason.
    assert len(st["failures"]) == 1
    fail = st["failures"][0]
    assert fail["name"] == "B" and fail["client"] == "X"
    assert "sheet boom" in fail["reason"]
    assert fail["status"] in ("active", "paused", "past")


def test_run_auto_sync_records_sheet_reads(fake_sb, monkeypatch):
    _reset_sync_status()
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, [
        {"id": "1", "sheet_id": "s1", "client_id": "c1", "client_name": "X",
         "campaign_name": "A", "total_prospects": 0, "completed_at": None},
    ]))
    # Each campaign sync consumes some real Sheets reads — the run must total them.
    def core_that_reads(**kw):
        gs._bump_sheet_reads(3)
        return {"error": "", "leads_added": 0, "interested_added": 0,
                "unsubscribes_added": 0, "reply_count": 0, "contacted_added": 0}
    monkeypatch.setattr(auto_sync, "sync_campaign_core", core_that_reads)

    auto_sync.run_auto_sync()
    assert auto_sync.get_sync_status()["sheet_reads"] == 3


def test_run_auto_sync_classifies_and_reports_failed_campaign_status(fake_sb, monkeypatch):
    _reset_sync_status()
    rows = [
        # active (no completion, not fully sent) — will fail
        {"id": "a", "sheet_id": "s", "client_id": "c1", "client_name": "Acme",
         "campaign_name": "Live", "total_prospects": 100, "sent_count": 10,
         "completed": False, "completed_at": None, "paused": False},
        # past (fully sent) — will fail
        {"id": "b", "sheet_id": "s", "client_id": "c1", "client_name": "Acme",
         "campaign_name": "Done", "total_prospects": 50, "sent_count": 50,
         "completed": False, "completed_at": None, "paused": False},
        # paused — will fail
        {"id": "c", "sheet_id": "s", "client_id": "c1", "client_name": "Acme",
         "campaign_name": "Halted", "total_prospects": 100, "sent_count": 5,
         "completed": False, "completed_at": None, "paused": True},
    ]
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, rows))
    monkeypatch.setattr(auto_sync, "sync_campaign_core",
                        lambda **kw: {"error": "Sheet read error: not found",
                                      "leads_added": 0, "interested_added": 0,
                                      "unsubscribes_added": 0, "reply_count": 0,
                                      "contacted_added": 0})
    auto_sync.run_auto_sync()

    st = auto_sync.get_sync_status()
    by_name = {f["name"]: f["status"] for f in st["failures"]}
    assert by_name == {"Live": "active", "Done": "past", "Halted": "paused"}


def test_sync_status_endpoint_reports_fresh_run(client):
    now = _now_utc()
    _seed_sync_status(now - timedelta(minutes=6), now - timedelta(minutes=5),
                      ok=3, failed=0, total=3)
    data = client.get("/api/campaigns/sync-status").json()
    assert data["state"] == "ok"
    assert data["ok"] == 3 and data["total"] == 3
    assert "ago" in data["ago_text"]


def test_sync_status_endpoint_flags_stale_run(client):
    now = _now_utc()
    _seed_sync_status(now - timedelta(hours=3), now - timedelta(hours=3),
                      ok=2, failed=0, total=2)
    data = client.get("/api/campaigns/sync-status").json()
    assert data["state"] == "stale", "a run older than the hourly cadence is stale"


def test_sync_status_endpoint_idle_before_first_run(client):
    _reset_sync_status()
    data = client.get("/api/campaigns/sync-status").json()
    assert data["state"] == "idle" and data["ran"] is False


def test_sync_now_starts_background_run(client, monkeypatch):
    _reset_sync_status()
    started = {"n": 0}
    # Replace the threaded target so the test doesn't spin a real sync; assert
    # the endpoint kicks it off and returns immediately.
    def fake_thread_target():
        started["n"] += 1
    class _FakeThread:
        def __init__(self, target=None, **kw): self._t = target
        def start(self): self._t()
    monkeypatch.setattr(campaigns.threading, "Thread", _FakeThread)
    monkeypatch.setattr("app.utils.auto_sync.run_auto_sync", fake_thread_target)

    resp = client.post("/api/campaigns/sync-now")
    assert resp.status_code == 200
    assert resp.json()["started"] is True
    assert started["n"] == 1


def test_sync_now_skips_when_already_running(client):
    now = _now_utc()
    with auto_sync._last_run_lock:
        auto_sync._last_run.update(started_at=now.isoformat(), finished_at=None,
                                   running=True, ok=0, failed=0, total=0, duration_s=None)
    resp = client.post("/api/campaigns/sync-now")
    assert resp.status_code == 409
    assert resp.json()["started"] is False
    _reset_sync_status()


def test_run_auto_sync_skips_when_run_lock_held():
    """A second trigger while a run holds the run-lock is a no-op (no status
    change), so the hourly tick can't clobber a manual run and vice versa."""
    _reset_sync_status()
    assert auto_sync._run_lock.acquire(blocking=False)
    try:
        auto_sync.run_auto_sync()   # should return immediately, record nothing
        assert auto_sync.get_sync_status()["started_at"] is None
    finally:
        auto_sync._run_lock.release()


def test_seconds_to_next_hour_within_bounds():
    s = auto_sync._seconds_to_next_hour()
    assert 0 < s <= 3600


def _now_utc():
    from datetime import datetime, timezone
    return datetime.now(timezone.utc)


def _seed_sync_status(started, finished, ok, failed, total):
    with auto_sync._last_run_lock:
        auto_sync._last_run.update(
            started_at=started.isoformat(), finished_at=finished.isoformat(),
            running=False, ok=ok, failed=failed, total=total, duration_s=12.0)
