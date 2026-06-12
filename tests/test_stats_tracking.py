"""Offline tests for campaign stats tracking.

Covers the send-count queries (source filter, date buckets, chaser counts),
the parallel time breakdown, and the full sync_campaign_core counter pipeline
with a mocked sheet + mocked Supabase.
"""
from __future__ import annotations

from datetime import date, timedelta

from conftest import FakeResponse, in_list_values, param_values

import app.routers.campaigns as campaigns
import app.utils.auto_sync as auto_sync
import app.utils.google_sheets as gs


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

    today  = date.today()
    monday = today - timedelta(days=today.weekday())
    today_calls = [c for c in fake_sb.calls_to("GET", "contacted_prospects")
                   if f"eq.{today.isoformat()}" in param_values(c, "contacted_at")]
    assert today_calls, "today bucket must filter contacted_at = today"
    assert monday.strftime("%b") in out["week_label"]


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


def test_sync_core_full_pipeline(fake_sb, monkeypatch):
    today = date.today().isoformat()
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
    assert written["sent"]   == [(3, today)]
    assert written["chaser"] == [(2, today)]

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
    assert by_email["lead1@acme.com"]["chaser_contacted_at"] == today  # back-filled
    assert by_email["fresh@corp.com"]["contacted_at"] == today
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
