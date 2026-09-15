"""
Offline tests for Outbound Pulse.

The connectors cannot be tested against live APIs from here, so the coverage
concentrates on the parts that are ours and that are easy to get subtly wrong:

  * the counting contract (sends count per send, qualifying stages per lead) —
    the property the whole funnel rests on;
  * idempotency, i.e. that re-syncing produces byte-identical dedupe keys;
  * agency scoping, i.e. that no query escapes the tenant filter;
  * the portal's access rules.

Uses the existing FakeSupabase harness from conftest.py — no network.
"""
from __future__ import annotations

from datetime import datetime, timezone

import pytest

from tests.conftest import FakeResponse

from app.utils.pulse import meet_alfred, normalize, smartlead, store

AGENCY = "11111111-1111-1111-1111-111111111111"
CAMPAIGN = "22222222-2222-2222-2222-222222222222"
CLIENT = "33333333-3333-3333-3333-333333333333"


@pytest.fixture(autouse=True)
def _pin_agency(monkeypatch):
    """Pin the agency id so tests don't depend on the lookup/create path."""
    store.reset_agency_cache()
    monkeypatch.setenv("PULSE_AGENCY_ID", AGENCY)
    yield
    store.reset_agency_cache()


# ── normalize: timestamps ─────────────────────────────────────────────────────

@pytest.mark.parametrize("raw", [
    "2025-03-04T10:30:00Z",
    "2025-03-04T10:30:00+00:00",
    "2025-03-04 10:30:00",
    "2025-03-04T10:30:00",
])
def test_parse_ts_normalizes_to_utc(raw):
    parsed = normalize.parse_ts(raw)
    assert parsed == datetime(2025, 3, 4, 10, 30, tzinfo=timezone.utc)


def test_parse_ts_shifts_offset_timestamps_to_utc():
    assert normalize.parse_ts("2025-03-04T12:30:00+02:00") == \
        datetime(2025, 3, 4, 10, 30, tzinfo=timezone.utc)


def test_parse_ts_handles_epoch_seconds_and_millis():
    seconds = normalize.parse_ts(1741084200)
    millis  = normalize.parse_ts(1741084200000)
    assert seconds == millis


@pytest.mark.parametrize("raw", ["", None, "not a date", "0000-00-00"])
def test_parse_ts_returns_none_rather_than_raising(raw):
    # A connector must not lose a whole campaign because one row has a bad date.
    assert normalize.parse_ts(raw) is None


# ── normalize: the counting contract ──────────────────────────────────────────

def _event(event_type, when, lead="a@b.com", step=""):
    return normalize.make_event(
        agency_id=AGENCY, campaign_id=CAMPAIGN, client_id=CLIENT,
        event_type=event_type, occurred_at=normalize.parse_ts(when),
        lead=lead, source_tool=normalize.SOURCE_SMARTLEAD, sequence_ref=step,
    )


def test_opens_collapse_to_one_event_per_lead():
    """Six opens at six times are one `opened` row — the funnel counts people."""
    keys = {
        _event("opened", f"2025-03-0{day}T10:00:00Z")["dedupe_key"]
        for day in range(1, 7)
    }
    assert len(keys) == 1


def test_sends_stay_distinct_per_step():
    """A 4-step sequence to one lead is 4 sends — that is send volume."""
    keys = {
        _event("sent", f"2025-03-0{step}T10:00:00Z", step=str(step))["dedupe_key"]
        for step in range(1, 5)
    }
    assert len(keys) == 4


def test_different_leads_never_share_a_dedupe_key():
    a = _event("replied", "2025-03-04T10:00:00Z", lead="a@b.com")
    b = _event("replied", "2025-03-04T10:00:00Z", lead="c@d.com")
    assert a["dedupe_key"] != b["dedupe_key"]


def test_dedupe_key_is_stable_across_runs():
    """Re-syncing the same window must produce identical keys, or the unique
    index cannot make the sync idempotent."""
    assert _event("replied", "2025-03-04T10:00:00Z")["dedupe_key"] == \
           _event("replied", "2025-03-04T10:00:00Z")["dedupe_key"]


def test_make_event_rejects_unknown_stage():
    with pytest.raises(ValueError):
        _event("clicked", "2025-03-04T10:00:00Z")


def test_lead_key_normalizes_case_and_linkedin_tracking_params():
    assert normalize.lead_key("A@B.com") == "a@b.com"
    assert normalize.lead_key("https://linkedin.com/in/jo/?utm=x") == \
        "https://linkedin.com/in/jo"
    assert normalize.lead_key(None, "", "fallback") == "fallback"


# ── normalize: reply classification ───────────────────────────────────────────

@pytest.mark.parametrize("category,expected", [
    ("Interested",            normalize.EVENT_POSITIVE_REPLY),
    ("interested - pricing",  normalize.EVENT_POSITIVE_REPLY),
    ("Meeting Booked",        normalize.EVENT_MEETING_BOOKED),
    ("Not Interested",        None),
    ("Out Of Office",         None),
    ("Some New Category",     None),   # unknown is never treated as positive
    ("",                      None),
    (None,                    None),
])
def test_classify_reply(category, expected):
    assert normalize.classify_reply(category) == expected


def test_funnel_rates():
    counts = {"sent": 1000, "opened": 400, "replied": 50,
              "positive_reply": 10, "meeting_booked": 4}
    rows = {r["key"]: r for r in normalize.funnel_with_rates(counts)}
    assert rows["sent"]["step_rate"] is None      # nothing above it
    assert rows["opened"]["step_rate"] == 40.0
    assert rows["replied"]["step_rate"] == 12.5   # of opened
    assert rows["replied"]["overall"] == 5.0      # of sent
    assert rows["meeting_booked"]["overall"] == 0.4


def test_opened_does_not_print_the_same_rate_twice():
    """The stage under `sent` has step_rate == overall by definition; showing
    both reads as a rendering bug."""
    rows = {r["key"]: r for r in normalize.funnel_with_rates(
        {"sent": 1000, "opened": 400, "replied": 50,
         "positive_reply": 10, "meeting_booked": 4})}
    assert rows["opened"]["step_rate"] == 40.0
    assert rows["opened"]["overall"] is None


def test_only_the_first_stage_is_flagged_as_top_of_funnel():
    """A later stage whose parent is zero also has no step rate, but it is not
    the top of the funnel and must not be labelled as one."""
    rows = normalize.funnel_with_rates(
        {"sent": 100, "opened": 20, "replied": 0,
         "positive_reply": 0, "meeting_booked": 0})
    by_key = {r["key"]: r for r in rows}
    assert by_key["sent"]["is_top"] is True
    assert by_key["meeting_booked"]["is_top"] is False
    assert by_key["meeting_booked"]["step_rate"] is None   # parent is zero


def test_a_zero_parent_stage_is_not_labelled_top_of_funnel(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "event_type": "sent", "events": 500},
    ]))
    r = client.get("/api/outbound-pulse/overview")
    # `sent` is the only stage with a value, so exactly one "top of funnel"
    # label should appear per funnel — not one for every empty stage below it.
    assert r.text.count("top of funnel") == 2   # agency total + the one client


def test_funnel_rates_survive_an_empty_funnel():
    rows = normalize.funnel_with_rates(normalize.empty_funnel())
    assert all(r["value"] == 0 for r in rows)
    assert all(r["step_rate"] is None for r in rows[1:])   # no divide-by-zero


# ── Smartlead mapping ─────────────────────────────────────────────────────────

def _smartlead_events(rows):
    return smartlead.events_from_statistics(
        rows, agency_id=AGENCY, campaign_id=CAMPAIGN, client_id=CLIENT,
    )


def test_smartlead_maps_a_full_lead_journey():
    events = _smartlead_events([{
        "lead_email":     "jo@acme.com",
        "sequence_number": 1,
        "sent_time":      "2025-03-01T09:00:00Z",
        "open_time":      "2025-03-01T11:00:00Z",
        "reply_time":     "2025-03-02T08:00:00Z",
        "lead_category":  "Interested",
    }])
    types = [e["event_type"] for e in events]
    assert types == ["sent", "opened", "replied", "positive_reply"]
    assert all(e["agency_id"] == AGENCY for e in events)
    assert all(e["client_id"] == CLIENT for e in events)


def test_smartlead_meeting_implies_positive_reply():
    """The funnel must never show more meetings than positive replies."""
    events = _smartlead_events([{
        "lead_email":    "jo@acme.com",
        "sent_time":     "2025-03-01T09:00:00Z",
        "reply_time":    "2025-03-02T08:00:00Z",
        "lead_category": "Meeting Booked",
    }])
    types = [e["event_type"] for e in events]
    assert "positive_reply" in types
    assert "meeting_booked" in types


def test_smartlead_tolerates_renamed_fields():
    """Field-name drift across API versions must not empty the funnel."""
    events = _smartlead_events([{
        "email":     "jo@acme.com",
        "sent_at":   "2025-03-01T09:00:00Z",
        "opened_at": "2025-03-01T11:00:00Z",
    }])
    assert [e["event_type"] for e in events] == ["sent", "opened"]


def test_smartlead_skips_rows_with_no_lead_identity():
    """Without a lead key the once-per-lead stages can't dedupe, so counting
    the row would inflate the funnel on every re-sync."""
    assert _smartlead_events([{"sent_time": "2025-03-01T09:00:00Z"}]) == []


def test_smartlead_unclassified_reply_is_not_positive():
    events = _smartlead_events([{
        "lead_email":    "jo@acme.com",
        "reply_time":    "2025-03-02T08:00:00Z",
        "lead_category": "Not Interested",
    }])
    assert [e["event_type"] for e in events] == ["replied"]


def test_smartlead_rows_unwraps_envelope_shapes():
    assert smartlead._rows({"data": [{"id": 1}]}) == [{"id": 1}]
    assert smartlead._rows([{"id": 1}]) == [{"id": 1}]
    assert smartlead._rows({"unexpected": "shape"}) == []


# ── Meet Alfred mapping ───────────────────────────────────────────────────────

def _alfred_events(rows):
    return meet_alfred.events_from_activities(
        rows, agency_id=AGENCY, campaign_id=CAMPAIGN, client_id=CLIENT,
    )


def test_alfred_maps_acceptance_to_opened():
    """The load-bearing stage mapping: a LinkedIn acceptance is the `opened`
    equivalent, which is what lets both channels share one funnel."""
    events = _alfred_events([{
        "profile url": "https://linkedin.com/in/jo",
        "sent at":     "2025-03-01T09:00:00Z",
        "accepted at": "2025-03-02T09:00:00Z",
    }])
    assert [e["event_type"] for e in events] == ["sent", "opened"]


def test_alfred_meeting_backfills_the_stages_above_it():
    """A tracked meeting with no reply timestamp must not produce a funnel that
    narrows and then widens."""
    events = _alfred_events([{
        "profile_url":    "https://linkedin.com/in/jo",
        "sent at":        "2025-03-01T09:00:00Z",
        "meeting booked": "2025-03-05T09:00:00Z",
    }])
    types = [e["event_type"] for e in events]
    assert "replied" in types
    assert "positive_reply" in types
    assert "meeting_booked" in types


def test_alfred_csv_import_matches_the_api_path():
    """The CSV and API paths share one normalization, so they cannot drift."""
    csv_bytes = (
        b"Profile URL,Sent At,Accepted At,Replied At,Category\n"
        b"https://linkedin.com/in/jo,2025-03-01T09:00:00Z,"
        b"2025-03-02T09:00:00Z,2025-03-03T09:00:00Z,Interested\n"
    )
    rows = meet_alfred.parse_csv(csv_bytes)
    assert rows[0]["profile url"] == "https://linkedin.com/in/jo"

    events = _alfred_events(rows)
    assert [e["event_type"] for e in events] == \
        ["sent", "opened", "replied", "positive_reply"]


def test_alfred_csv_handles_a_utf8_bom():
    rows = meet_alfred.parse_csv(b"\xef\xbb\xbfProfile URL,Sent At\nx,2025-03-01\n")
    assert "profile url" in rows[0]


# ── Store: agency scoping ─────────────────────────────────────────────────────

def test_every_funnel_query_is_agency_scoped(fake_sb):
    from tests.conftest import param_values

    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))
    store.funnel(client_id=CLIENT)
    store.funnel_by_client()
    store.funnel_by_channel()
    store.funnel_timeseries()

    calls = fake_sb.calls_to("GET", "pulse_funnel_daily")
    assert calls, "expected funnel queries"
    for call in calls:
        assert param_values(call, "agency_id") == [f"eq.{AGENCY}"], \
            "a funnel query escaped the tenant filter"


def test_date_range_sends_both_bounds(fake_sb):
    """Two filters on the same column can't live in a dict — the store must send
    them as repeated params or the upper bound is silently dropped."""
    from datetime import date
    from tests.conftest import param_values

    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))
    store.funnel(date_from=date(2025, 3, 1), date_to=date(2025, 3, 31))

    call = fake_sb.calls_to("GET", "pulse_funnel_daily")[0]
    assert sorted(param_values(call, "day")) == ["gte.2025-03-01", "lte.2025-03-31"]


def test_funnel_sums_daily_rows(fake_sb):
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"event_type": "sent",   "events": 100},
        {"event_type": "sent",   "events": 50},
        {"event_type": "opened", "events": 30},
    ]))
    counts = store.funnel()
    assert counts["sent"] == 150
    assert counts["opened"] == 30
    assert counts["meeting_booked"] == 0


def test_insert_events_reports_only_new_rows(fake_sb):
    """PostgREST returns just the inserted rows under ignore-duplicates, which
    is how a re-sync correctly reports zero new events."""
    fake_sb.route("POST", "pulse_campaign_events",
                  lambda call: FakeResponse(201, [{"id": "x"}]))
    assert store.insert_events([_event("sent", "2025-03-01T09:00:00Z")]) == 1

    fake_sb.route("POST", "pulse_campaign_events", lambda call: FakeResponse(201, []))
    assert store.insert_events([_event("sent", "2025-03-01T09:00:00Z")]) == 1 - 1


def test_insert_events_uses_the_dedupe_conflict_target(fake_sb):
    from tests.conftest import param_values

    fake_sb.route("POST", "pulse_campaign_events", lambda call: FakeResponse(201, []))
    store.insert_events([_event("sent", "2025-03-01T09:00:00Z")])

    call = fake_sb.calls_to("POST", "pulse_campaign_events")[0]
    assert param_values(call, "on_conflict") == ["agency_id,dedupe_key"]
    assert "ignore-duplicates" in call["headers"].get("Prefer", "")


def test_insert_events_survives_a_failing_chunk(fake_sb, monkeypatch):
    """One bad chunk must not lose the rest of a run."""
    monkeypatch.setattr(store, "INSERT_CHUNK", 1)
    calls = {"n": 0}

    def handler(call):
        calls["n"] += 1
        if calls["n"] == 1:
            return FakeResponse(500, [], text="boom")
        return FakeResponse(201, [{"id": "x"}])

    fake_sb.route("POST", "pulse_campaign_events", handler)
    inserted = store.insert_events([
        _event("sent", "2025-03-01T09:00:00Z", step="1"),
        _event("sent", "2025-03-02T09:00:00Z", step="2"),
    ])
    assert inserted == 1


# ── Endpoints ─────────────────────────────────────────────────────────────────

def test_pulse_page_renders(client, fake_sb):
    r = client.get("/outbound-pulse")
    assert r.status_code == 200
    assert "Outbound Pulse" in r.text


def test_overview_reports_a_missing_migration_instead_of_500ing(client, fake_sb):
    fake_sb.route("GET", "clients",
                  lambda call: FakeResponse(404, [], text="relation does not exist"))
    r = client.get("/api/outbound-pulse/overview")
    assert r.status_code == 200
    assert "outbound_pulse_schema.sql" in r.text


def test_overview_renders_a_funnel(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "event_type": "sent",           "events": 1000},
        {"client_id": CLIENT, "event_type": "opened",         "events": 400},
        {"client_id": CLIENT, "event_type": "replied",        "events": 50},
        {"client_id": CLIENT, "event_type": "meeting_booked", "events": 4},
    ]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))

    r = client.get("/api/outbound-pulse/overview")
    assert r.status_code == 200
    assert "Acme" in r.text
    assert "1,000" in r.text


def test_overview_flags_unmapped_campaigns(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(
        200, [{"id": CAMPAIGN, "client_id": None, "name": "Stray", "channel": "email",
               "source_tool": "smartlead", "external_campaign_id": "1",
               "status": "active", "last_synced_at": None, "created_at": None}]))

    r = client.get("/api/outbound-pulse/overview")
    assert r.status_code == 200
    # No client rows, so the empty state renders — but the warning must still
    # fire, because unmapped campaigns are exactly why a funnel looks empty.
    assert "not mapped to a client" in r.text


def test_sync_status_marks_an_unconfigured_connector(client, fake_sb, monkeypatch):
    monkeypatch.delenv("SMARTLEAD_API_KEY", raising=False)
    monkeypatch.delenv("MEET_ALFRED_API_KEY", raising=False)
    fake_sb.route("GET", "pulse_sync_logs", lambda call: FakeResponse(200, []))

    r = client.get("/api/outbound-pulse/sync-status")
    assert r.status_code == 200
    assert "not configured" in r.text


# ── Portal access ─────────────────────────────────────────────────────────────

def test_portal_rejects_an_unknown_token(client, fake_sb):
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))
    r = client.get("/portal/nope")
    assert r.status_code == 404
    assert "not valid" in r.text
    assert "noindex" in r.headers.get("x-robots-tag", "")
    assert "no-store" in r.headers.get("cache-control", "")


def test_portal_rejects_a_revoked_token(client, fake_sb):
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, [{
        "id": "a1", "client_id": CLIENT, "label": "",
        "expires_at": None, "revoked_at": "2025-01-01T00:00:00+00:00", "view_count": 3,
    }]))
    r = client.get("/portal/whatever")
    assert r.status_code == 404
    assert "no longer active" in r.text


def test_portal_rejects_an_expired_token(client, fake_sb):
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, [{
        "id": "a1", "client_id": CLIENT, "label": "",
        "expires_at": "2020-01-01T00:00:00+00:00", "revoked_at": None, "view_count": 0,
    }]))
    r = client.get("/portal/whatever")
    assert r.status_code == 404
    assert "expired" in r.text


def test_portal_looks_up_by_hash_not_plaintext(client, fake_sb):
    """The plaintext token must never reach the database."""
    from tests.conftest import param_values

    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))
    client.get("/portal/super-secret-token")

    call = fake_sb.calls_to("GET", "pulse_client_access")[0]
    sent = param_values(call, "token_hash")[0]
    assert "super-secret-token" not in sent
    assert sent.startswith("eq.") and len(sent) == len("eq.") + 64


def test_portal_renders_for_a_valid_token(client, fake_sb):
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, [{
        "id": "a1", "client_id": CLIENT, "label": "Sarah",
        "expires_at": None, "revoked_at": None, "view_count": 0,
    }]))
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"event_type": "sent", "events": 900, "channel": "email", "day": "2025-03-01"},
        {"event_type": "replied", "events": 45, "channel": "email", "day": "2025-03-01"},
    ]))

    r = client.get("/portal/valid-token")
    assert r.status_code == 200
    assert "Acme" in r.text
    assert "900" in r.text
    # Internal chrome must not leak into a client-facing page.
    assert "Dashboard" not in r.text
    assert "/outbound-pulse" not in r.text


def test_portal_ignores_a_client_id_in_the_query_string(client, fake_sb):
    """Scope comes from the token record only — a client cannot widen it."""
    from tests.conftest import param_values

    other = "99999999-9999-9999-9999-999999999999"
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, [{
        "id": "a1", "client_id": CLIENT, "label": "",
        "expires_at": None, "revoked_at": None, "view_count": 0,
    }]))
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    client.get(f"/portal/valid-token?client_id={other}&range=all")

    for call in fake_sb.calls_to("GET", "pulse_funnel_daily"):
        assert param_values(call, "client_id") == [f"eq.{CLIENT}"]


def test_portal_records_a_visit(client, fake_sb):
    """Engagement logging is how success-criteria question #2 gets answered."""
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, [{
        "id": "a1", "client_id": CLIENT, "label": "",
        "expires_at": None, "revoked_at": None, "view_count": 2,
    }]))
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    client.get("/portal/valid-token")

    visits = fake_sb.calls_to("POST", "pulse_portal_visits")
    assert len(visits) == 1
    assert visits[0]["json"]["client_id"] == CLIENT
    assert visits[0]["json"]["agency_id"] == AGENCY


def test_portal_exposes_no_mutating_routes():
    """Read-only by construction, not by convention."""
    from app.routers import pulse_portal

    for route in pulse_portal.router.routes:
        assert set(route.methods) <= {"GET", "HEAD"}, f"{route.path} is not read-only"


# ── Auth wiring ───────────────────────────────────────────────────────────────

def test_pulse_routes_map_to_the_pulse_tool():
    from app import auth

    assert auth.tool_for_path("/outbound-pulse") == "outbound_pulse"
    assert auth.tool_for_path("/outbound-pulse/clients/abc") == "outbound_pulse"
    assert auth.tool_for_path("/api/outbound-pulse/overview") == "outbound_pulse"


def test_portal_is_public_but_the_internal_module_is_not():
    from app import auth

    assert auth.is_public_path("/portal/abc")
    assert not auth.is_public_path("/outbound-pulse")
    assert not auth.is_public_path("/api/outbound-pulse/overview")


def test_staff_roles_get_pulse_and_viewers_do_not():
    from app import auth

    assert "outbound_pulse" in auth.ROLE_TOOLS["admin"]
    assert "outbound_pulse" in auth.ROLE_TOOLS["sdr"]
    assert "outbound_pulse" not in auth.ROLE_TOOLS["viewer"]


def test_pulse_module_is_gated_when_auth_is_on(client, monkeypatch):
    monkeypatch.setenv("AUTH_DISABLED", "0")
    r = client.get("/outbound-pulse", follow_redirects=False)
    assert r.status_code == 303
    assert r.headers["location"] == "/login"


def test_portal_is_reachable_when_auth_is_on(client, fake_sb, monkeypatch):
    """A client has no dashboard login — the gate must not bounce them."""
    monkeypatch.setenv("AUTH_DISABLED", "0")
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))
    r = client.get("/portal/some-token", follow_redirects=False)
    assert r.status_code == 404          # rejected by the token check, not the gate
    assert "location" not in r.headers


# ── Client detail page ────────────────────────────────────────────────────────

def _detail_routes(fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": "#7c3aed",
               "emoji": "\U0001F3E2", "active": True}]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(
        200, [{"id": CAMPAIGN, "client_id": CLIENT, "name": "Acme UK PR",
               "channel": "email", "source_tool": "smartlead",
               "external_campaign_id": "1", "status": "active",
               "last_synced_at": "2025-03-05T10:00:00+00:00", "created_at": None}]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": "email",
         "event_type": "sent", "events": 800, "day": "2025-03-01"},
        {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": "email",
         "event_type": "replied", "events": 40, "day": "2025-03-02"},
        {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": "linkedin",
         "event_type": "sent", "events": 200, "day": "2025-03-02"},
    ]))
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))


def test_client_detail_page_renders(client, fake_sb):
    _detail_routes(fake_sb)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert r.status_code == 200
    assert "Acme" in r.text
    assert "Acme UK PR" in r.text
    assert "1,000" in r.text                  # 800 email + 200 LinkedIn


def test_client_detail_page_404s_for_an_unknown_client(client, fake_sb):
    _detail_routes(fake_sb)
    r = client.get("/outbound-pulse/clients/44444444-4444-4444-4444-444444444444")
    assert r.status_code == 404


def test_client_detail_range_filter_narrows_the_query(client, fake_sb):
    from tests.conftest import param_values

    _detail_routes(fake_sb)
    client.get(f"/outbound-pulse/clients/{CLIENT}?range=all")
    for call in fake_sb.calls_to("GET", "pulse_funnel_daily"):
        assert param_values(call, "day") == []      # "all" sets no date bounds

    fake_sb.calls.clear()
    client.get(f"/outbound-pulse/clients/{CLIENT}?range=7d")
    assert any(param_values(call, "day")
               for call in fake_sb.calls_to("GET", "pulse_funnel_daily"))


def test_custom_dates_override_the_preset():
    from app.routers.outbound_pulse import _resolve_range

    rng = _resolve_range("30d", "2025-01-01", "2025-01-31")
    assert rng["preset"] == "custom"
    assert rng["from"].isoformat() == "2025-01-01"
    assert rng["to"].isoformat() == "2025-01-31"


def test_reversed_custom_dates_are_swapped_not_rejected():
    from app.routers.outbound_pulse import _resolve_range

    rng = _resolve_range("30d", "2025-01-31", "2025-01-01")
    assert rng["from"].isoformat() == "2025-01-01"
    assert rng["to"].isoformat() == "2025-01-31"


def test_unknown_preset_falls_back_to_30d():
    from app.routers.outbound_pulse import _resolve_range

    assert _resolve_range("bogus", "", "")["preset"] == "30d"


# ── Portal access management ──────────────────────────────────────────────────

def test_creating_a_portal_link_stores_only_a_hash(client, fake_sb):
    """The plaintext token is shown once and never persisted."""
    import hashlib
    import re

    captured = {}

    def create(call):
        captured.update(call["json"])
        return FakeResponse(201, [{"id": "a1"}])

    fake_sb.route("POST", "pulse_client_access", create)
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/access",
                    data={"label": "Sarah"})
    assert r.status_code == 200

    match = re.search(r"/portal/([A-Za-z0-9_\-]+)", r.text)
    assert match, "the plaintext link should be shown once"
    token = match.group(1)

    assert "token" not in captured
    assert captured["token_hash"] == hashlib.sha256(token.encode()).hexdigest()
    assert captured["agency_id"] == AGENCY
    assert captured["client_id"] == CLIENT
    assert captured["expires_at"]          # links are not permanent


def test_revoking_a_link_sets_revoked_at(client, fake_sb):
    patched = {}

    def patch(call):
        patched.update(call["json"])
        return FakeResponse(204, [])

    fake_sb.route("PATCH", "pulse_client_access", patch)
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))

    r = client.delete(f"/api/outbound-pulse/clients/{CLIENT}/access/a1")
    assert r.status_code == 200
    assert patched.get("revoked_at")


def test_engagement_panel_ranks_clients_by_views(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_portal_visits", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "viewed_at": "2025-03-05T10:00:00+00:00"},
        {"client_id": CLIENT, "viewed_at": "2025-03-04T10:00:00+00:00"},
    ]))
    r = client.get("/api/outbound-pulse/engagement")
    assert r.status_code == 200
    assert "Acme" in r.text
    assert ">2<" in r.text


# ── Campaign mapping ──────────────────────────────────────────────────────────

def test_mapping_a_campaign_repoints_its_events(client, fake_sb):
    """A remap must move the campaign's history with it, or the client's funnel
    silently loses everything synced before the mapping."""
    patched = []
    fake_sb.route("PATCH", "pulse_campaigns",
                  lambda call: (patched.append(("campaign", call)), FakeResponse(204, []))[1])
    fake_sb.route("PATCH", "pulse_campaign_events",
                  lambda call: (patched.append(("events", call)), FakeResponse(204, []))[1])
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, []))

    r = client.post(f"/api/outbound-pulse/campaigns/{CAMPAIGN}/client",
                    data={"client_id": CLIENT})
    assert r.status_code == 200

    tables = [t for t, _ in patched]
    assert tables == ["campaign", "events"]
    assert patched[1][1]["json"]["client_id"] == CLIENT


def test_unmapping_a_campaign_is_allowed(client, fake_sb):
    patched = {}
    fake_sb.route("PATCH", "pulse_campaigns",
                  lambda call: (patched.update(call["json"]), FakeResponse(204, []))[1])
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, []))

    client.post(f"/api/outbound-pulse/campaigns/{CAMPAIGN}/client", data={"client_id": ""})
    assert patched["client_id"] is None


# ── Meet Alfred CSV import endpoint ───────────────────────────────────────────

def test_csv_import_creates_a_campaign_and_events(client, fake_sb):
    fake_sb.route("POST", "pulse_campaigns",
                  lambda call: FakeResponse(201, [{"id": CAMPAIGN, "client_id": CLIENT}]))
    fake_sb.route("POST", "pulse_campaign_events",
                  lambda call: FakeResponse(201, [{"id": "e1"}, {"id": "e2"}]))
    fake_sb.route("POST", "pulse_sync_logs", lambda call: FakeResponse(201, [{"id": "log1"}]))

    csv_bytes = (
        b"Profile URL,Sent At,Accepted At\n"
        b"https://linkedin.com/in/jo,2025-03-01T09:00:00Z,2025-03-02T09:00:00Z\n"
    )
    r = client.post(
        "/api/outbound-pulse/import/meet-alfred",
        data={"campaign_name": "Acme LinkedIn", "client_id": CLIENT},
        files={"file": ("report.csv", csv_bytes, "text/csv")},
    )
    assert r.status_code == 200
    assert "Imported" in r.text

    created = fake_sb.calls_to("POST", "pulse_campaigns")[0]["json"]
    assert created["channel"] == "linkedin"
    assert created["source_tool"] == "meet_alfred"
    # Deterministic external id, so re-importing updates rather than duplicates.
    assert created["external_campaign_id"] == "csv:acme linkedin"


def test_csv_import_requires_a_campaign_name(client, fake_sb):
    r = client.post(
        "/api/outbound-pulse/import/meet-alfred",
        data={"campaign_name": "  "},
        files={"file": ("r.csv", b"Profile URL\nx\n", "text/csv")},
    )
    assert r.status_code == 200
    assert "err-box" in r.text
    assert fake_sb.calls_to("POST", "pulse_campaigns") == []


def test_csv_import_logs_a_failure_to_sync_logs(client, fake_sb):
    """A failed import must be visible in connector history, not just on screen."""
    fake_sb.route("POST", "pulse_sync_logs", lambda call: FakeResponse(201, [{"id": "log1"}]))
    fake_sb.route("POST", "pulse_campaigns", lambda call: FakeResponse(500, [], text="nope"))

    patched = {}
    fake_sb.route("PATCH", "pulse_sync_logs",
                  lambda call: (patched.update(call["json"]), FakeResponse(204, []))[1])

    r = client.post(
        "/api/outbound-pulse/import/meet-alfred",
        data={"campaign_name": "Acme"},
        files={"file": ("r.csv", b"Profile URL,Sent At\nx,2025-03-01\n", "text/csv")},
    )
    assert r.status_code == 200
    assert patched.get("status") == "error"


# ── Sync orchestration ────────────────────────────────────────────────────────

def test_sync_skips_an_unconfigured_connector(monkeypatch):
    from app.utils.pulse import sync as pulse_sync

    monkeypatch.delenv("SMARTLEAD_API_KEY", raising=False)
    result = pulse_sync.sync_source("smartlead", "test")
    assert result["status"] == "skipped"
    assert "not configured" in result["error"]


def _stored(cid, external, status, name="c", last_synced=None, client=None):
    return {"id": cid, "client_id": client, "external_campaign_id": external,
            "name": name, "status": status, "last_synced_at": last_synced}


def test_sync_reports_partial_when_one_campaign_fails(monkeypatch, fake_sb):
    """A single broken campaign must not be reported as a clean run."""
    from app.utils.pulse import sync as pulse_sync

    monkeypatch.setenv("SMARTLEAD_API_KEY", "k")
    monkeypatch.setattr(smartlead, "fetch_campaigns", lambda: [
        {"external_id": "1", "name": "Good", "status": "ACTIVE", "raw": {}},
        {"external_id": "2", "name": "Bad",  "status": "ACTIVE", "raw": {}},
    ])

    def fake_sync_campaign(*, external_campaign_id, **kw):
        if external_campaign_id == "2":
            raise smartlead.SmartleadError("upstream 500")
        return []

    monkeypatch.setattr(smartlead, "sync_campaign", fake_sync_campaign)
    monkeypatch.setattr(pulse_sync, "_PACING_SECONDS", 0)

    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, [
        _stored("id1", "1", "ACTIVE", "Good"),
        _stored("id2", "2", "ACTIVE", "Bad"),
    ]))
    fake_sb.route("POST", "pulse_campaigns", lambda call: FakeResponse(201, []))
    fake_sb.route("POST", "pulse_sync_logs", lambda call: FakeResponse(201, [{"id": "log1"}]))

    result = pulse_sync.sync_source("smartlead", "test")
    assert result["status"] == "partial"
    assert result["campaigns"] == 1
    assert "upstream 500" in result["error"]


def test_only_the_succeeding_campaign_is_stamped_as_synced(monkeypatch, fake_sb):
    """A failed campaign must stay at the head of the rolling order so the next
    run retries it, rather than being rotated to the back as if it succeeded."""
    from tests.conftest import param_values
    from app.utils.pulse import sync as pulse_sync

    monkeypatch.setenv("SMARTLEAD_API_KEY", "k")
    monkeypatch.setattr(smartlead, "fetch_campaigns", lambda: [])
    monkeypatch.setattr(pulse_sync, "_PACING_SECONDS", 0)

    def fake_sync_campaign(*, external_campaign_id, **kw):
        if external_campaign_id == "2":
            raise smartlead.SmartleadError("boom")
        return []

    monkeypatch.setattr(smartlead, "sync_campaign", fake_sync_campaign)
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, [
        _stored("id1", "1", "ACTIVE"), _stored("id2", "2", "ACTIVE"),
    ]))
    fake_sb.route("POST", "pulse_sync_logs", lambda call: FakeResponse(201, [{"id": "log1"}]))

    pulse_sync.sync_source("smartlead", "test")

    patches = fake_sb.calls_to("PATCH", "pulse_campaigns")
    assert len(patches) == 1
    ids = param_values(patches[0], "id")[0]
    assert "id1" in ids and "id2" not in ids


def test_bulk_campaign_upsert_omits_client_id(monkeypatch, fake_sb):
    """Under merge-duplicates PostgREST only updates columns present in the
    payload, so omitting client_id is what preserves a human's mapping. Sending
    it — even as null — would blank every mapping on the next sync."""
    from app.utils.pulse import sync as pulse_sync

    monkeypatch.setenv("SMARTLEAD_API_KEY", "k")
    monkeypatch.setattr(smartlead, "fetch_campaigns", lambda: [
        {"external_id": "1", "name": "Mapped", "status": "ACTIVE", "raw": {}}])
    monkeypatch.setattr(smartlead, "sync_campaign", lambda **kw: [])
    monkeypatch.setattr(pulse_sync, "_PACING_SECONDS", 0)

    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(
        200, [_stored(CAMPAIGN, "1", "ACTIVE", "Mapped", client=CLIENT)]))
    fake_sb.route("POST", "pulse_campaigns", lambda call: FakeResponse(201, []))
    fake_sb.route("POST", "pulse_sync_logs", lambda call: FakeResponse(201, [{"id": "log1"}]))

    pulse_sync.sync_source("smartlead", "test")

    body = fake_sb.calls_to("POST", "pulse_campaigns")[0]["json"]
    assert isinstance(body, list)                    # bulk, not one call each
    assert "client_id" not in body[0]
    # Nor last_synced_at — that would make every campaign look freshly synced
    # and destroy the rolling backfill order.
    assert "last_synced_at" not in body[0]


def test_campaign_selection_is_bounded_and_prioritised(monkeypatch):
    """The live account has 1082 campaigns, 34 active. A run must not walk them
    all, must never skip an active one, and must ignore drafts and archives."""
    from app.utils.pulse import sync as pulse_sync

    monkeypatch.setattr(pulse_sync, "_BACKFILL_PER_RUN", 25)
    stored = (
        [_stored(f"a{i}", str(i), "ACTIVE") for i in range(34)]
        + [_stored(f"c{i}", str(i), "COMPLETED") for i in range(577)]
        + [_stored(f"p{i}", str(i), "PAUSED") for i in range(390)]
        + [_stored(f"d{i}", str(i), "DRAFTED") for i in range(47)]
        + [_stored(f"r{i}", str(i), "ARCHIVED") for i in range(7)]
    )
    selected = pulse_sync.select_campaigns_for_run(stored)

    statuses = [s["status"] for s in selected]
    assert statuses.count("ACTIVE") == 34          # every live campaign, always
    assert "DRAFTED" not in statuses               # no sends to report
    assert "ARCHIVED" not in statuses              # deliberately shelved
    assert len(selected) == 34 + 25                # bounded, not 1082


def test_backfill_slice_follows_the_stored_order(monkeypatch):
    """store.campaigns_for_sync orders least-recently-synced first; the slice
    must respect that or the same campaigns get picked every run."""
    from app.utils.pulse import sync as pulse_sync

    monkeypatch.setattr(pulse_sync, "_BACKFILL_PER_RUN", 2)
    stored = [
        _stored("never", "1", "COMPLETED", last_synced=None),
        _stored("old",   "2", "COMPLETED", last_synced="2020-01-01T00:00:00+00:00"),
        _stored("fresh", "3", "COMPLETED", last_synced="2026-09-11T00:00:00+00:00"),
    ]
    assert [c["id"] for c in pulse_sync.select_campaigns_for_run(stored)] == ["never", "old"]


def test_zero_backfill_still_syncs_active_campaigns(monkeypatch):
    from app.utils.pulse import sync as pulse_sync

    monkeypatch.setattr(pulse_sync, "_BACKFILL_PER_RUN", 0)
    stored = [_stored("a", "1", "ACTIVE"), _stored("c", "2", "COMPLETED")]
    assert [c["id"] for c in pulse_sync.select_campaigns_for_run(stored)] == ["a"]


# ── raw_payload hygiene ───────────────────────────────────────────────────────

def test_smartlead_raw_payload_drops_message_bodies_and_names():
    """Live rows are ~2.4KB because they embed the email body. Storing that on
    every event is gigabytes of duplicated prospect content in a reporting
    table that has no use for it."""
    events = _smartlead_events([{
        "lead_email":    "jo@acme.com",
        "lead_name":     "Jo Bloggs",
        "email_subject": "Quick question about your Q4 plans",
        "email_message": "<html>" + ("x" * 4000) + "</html>",
        "sent_time":     "2025-03-01T09:00:00Z",
        "sequence_number": 1,
        "stats_id":      "s1",
        "open_count":    0,
    }])
    assert events
    stored = events[0]["raw_payload"]["row"]
    assert "email_message" not in stored
    assert "email_subject" not in stored
    assert "lead_name" not in stored
    # Diagnostics that make a mapping bug traceable are kept.
    assert stored["stats_id"] == "s1"
    assert stored["sequence_number"] == 1
    # The lead is still identifiable via the event's own lead_key.
    assert events[0]["lead_key"] == "jo@acme.com"


def test_smartlead_raw_payload_stays_small():
    import json
    events = _smartlead_events([{
        "lead_email":  "jo@acme.com",
        "email_message": "y" * 20000,
        "sent_time":   "2025-03-01T09:00:00Z",
    }])
    assert len(json.dumps(events[0]["raw_payload"])) < 400


def test_alfred_raw_payload_drops_names_and_truncates_wide_columns():
    events = _alfred_events([{
        "profile url": "https://linkedin.com/in/jo",
        "full name":   "Jo Bloggs",
        "message":     "hello there",
        "notes":       "z" * 5000,
        "sent at":     "2025-03-01T09:00:00Z",
    }])
    stored = events[0]["raw_payload"]["row"]
    assert "full name" not in stored
    assert "message" not in stored
    assert "notes" not in stored
    assert stored["profile url"] == "https://linkedin.com/in/jo"


def test_past_runs_do_not_make_an_unconfigured_connector_look_healthy(
    client, fake_sb, monkeypatch,
):
    """Pulling an API key breaks the sync. History from before that says nothing
    about now, so a successful old run must not render as a green "ok"."""
    monkeypatch.delenv("SMARTLEAD_API_KEY", raising=False)
    monkeypatch.delenv("MEET_ALFRED_API_KEY", raising=False)
    fake_sb.route("GET", "pulse_sync_logs", lambda call: FakeResponse(200, [{
        "id": "l1", "source_tool": "smartlead",
        "run_at": "2026-09-11T09:00:00+00:00", "finished_at": "2026-09-11T09:00:41+00:00",
        "status": "ok", "campaigns_synced": 4, "events_inserted": 318,
        "duration_s": 41.2, "triggered_by": "schedule", "error_message": None,
    }]))

    r = client.get("/api/outbound-pulse/sync-status")
    assert r.status_code == 200
    assert "state-ok" not in r.text
    assert "not configured" in r.text


def test_a_configured_connector_with_an_old_run_reads_as_stale(
    client, fake_sb, monkeypatch,
):
    monkeypatch.setenv("SMARTLEAD_API_KEY", "k")
    monkeypatch.setenv("MEET_ALFRED_API_KEY", "k")
    fake_sb.route("GET", "pulse_sync_logs", lambda call: FakeResponse(200, [{
        "id": "l1", "source_tool": "smartlead",
        "run_at": "2020-01-01T00:00:00+00:00", "finished_at": "2020-01-01T00:00:41+00:00",
        "status": "ok", "campaigns_synced": 4, "events_inserted": 318,
        "duration_s": 41.2, "triggered_by": "schedule", "error_message": None,
    }]))

    r = client.get("/api/outbound-pulse/sync-status")
    assert "state-stale" in r.text


# ── Campaign mapping filters ──────────────────────────────────────────────────
# The live account has 1082 campaigns, mostly DRAFTED. Without filtering, the
# mapping table is unusable for its actual job: finding the 34 ACTIVE ones.

def _filter_routes(fake_sb, campaigns=None):
    rows = campaigns if campaigns is not None else [
        {"id": "a1", "client_id": None, "channel": "email", "source_tool": "smartlead",
         "external_campaign_id": "1", "name": "Live one", "status": "ACTIVE",
         "last_synced_at": "2026-09-14T09:00:00+00:00", "created_at": None},
        {"id": "d1", "client_id": None, "channel": "email", "source_tool": "smartlead",
         "external_campaign_id": "2", "name": "Draft one", "status": "DRAFTED",
         "last_synced_at": None, "created_at": None},
    ]
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, rows))
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))


def test_status_filter_is_pushed_to_the_query(client, fake_sb):
    """Filtering must happen in the database, not by hiding rows in the browser
    — 1082 campaigns of HTML per keystroke is not a filter."""
    from tests.conftest import param_values

    _filter_routes(fake_sb)
    client.get("/api/outbound-pulse/campaigns?status=ACTIVE")

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "status")]
    assert listing, "status filter never reached PostgREST"
    assert param_values(listing[0], "status") == ["eq.ACTIVE"]


def test_synced_filter_maps_to_null_checks(client, fake_sb):
    from tests.conftest import param_values

    _filter_routes(fake_sb)
    client.get("/api/outbound-pulse/campaigns?synced=never")
    never = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
             if param_values(c, "last_synced_at")]
    assert param_values(never[0], "last_synced_at") == ["is.null"]

    fake_sb.calls.clear()
    client.get("/api/outbound-pulse/campaigns?synced=synced")
    synced = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
              if param_values(c, "last_synced_at")]
    assert param_values(synced[0], "last_synced_at") == ["not.is.null"]


def test_filters_combine(client, fake_sb):
    from tests.conftest import param_values

    _filter_routes(fake_sb)
    client.get("/api/outbound-pulse/campaigns"
               "?status=ACTIVE&source_tool=smartlead&channel=email&synced=synced")

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "status")]
    call = listing[0]
    assert param_values(call, "status") == ["eq.ACTIVE"]
    assert param_values(call, "source_tool") == ["eq.smartlead"]
    assert param_values(call, "channel") == ["eq.email"]
    assert param_values(call, "last_synced_at") == ["not.is.null"]
    # The tenant filter must survive the extra filters.
    assert param_values(call, "agency_id") == [f"eq.{AGENCY}"]


def test_unknown_filter_values_are_ignored_not_passed_through(client, fake_sb):
    """Filter values reach a query builder, so anything not recognised is
    dropped rather than forwarded."""
    from tests.conftest import param_values

    _filter_routes(fake_sb)
    client.get("/api/outbound-pulse/campaigns?channel=bogus&synced=bogus")

    for call in fake_sb.calls_to("GET", "pulse_campaigns"):
        assert param_values(call, "channel") == []
        assert param_values(call, "last_synced_at") == []


def test_filter_dropdowns_show_every_status_with_counts(client, fake_sb):
    """Options describe the whole set, not the filtered view — otherwise
    selecting ACTIVE would leave ACTIVE as the only option left."""
    _filter_routes(fake_sb)
    r = client.get("/api/outbound-pulse/campaigns?status=ACTIVE")
    assert r.status_code == 200
    assert "ACTIVE (1)" in r.text
    assert "DRAFTED (1)" in r.text          # still offered while ACTIVE is applied
    assert 'value="ACTIVE" selected' in r.text


def test_active_filter_shows_a_clear_control(client, fake_sb):
    _filter_routes(fake_sb)
    assert "Clear filters" in client.get(
        "/api/outbound-pulse/campaigns?status=ACTIVE").text
    assert "Clear filters" not in client.get("/api/outbound-pulse/campaigns").text


def test_empty_filter_result_says_so(client, fake_sb):
    _filter_routes(fake_sb, campaigns=[])
    r = client.get("/api/outbound-pulse/campaigns?status=ACTIVE")
    assert "No campaigns match these filters" in r.text
    # The never-synced explanation must not masquerade as an empty-filter result.
    assert "No campaigns synced yet" not in r.text


def test_draft_filter_warns_that_drafts_are_never_synced(client, fake_sb):
    _filter_routes(fake_sb)
    r = client.get("/api/outbound-pulse/campaigns?status=DRAFTED")
    assert "never synced" in r.text


def test_mapping_a_campaign_keeps_the_active_filters(client, fake_sb):
    """Mapping is done in batches inside a filtered view. Resetting to all 1082
    campaigns after each one would make the job unworkable."""
    from tests.conftest import param_values

    _filter_routes(fake_sb)
    fake_sb.route("PATCH", "pulse_campaigns", lambda call: FakeResponse(204, []))
    fake_sb.route("PATCH", "pulse_campaign_events", lambda call: FakeResponse(204, []))

    r = client.post("/api/outbound-pulse/campaigns/a1/client",
                    data={"client_id": CLIENT, "status": "ACTIVE",
                          "source_tool": "smartlead", "channel": "", "synced": ""})
    assert r.status_code == 200
    assert 'value="ACTIVE" selected' in r.text

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "status")]
    assert listing, "the re-render dropped the filter"
    assert param_values(listing[0], "status") == ["eq.ACTIVE"]


def test_mapping_does_not_make_the_table_refetch_itself(client, fake_sb):
    """The POST already returns the table. Firing pulseSynced would make the
    list re-request itself — a second full render of the filtered set."""
    _filter_routes(fake_sb)
    fake_sb.route("PATCH", "pulse_campaigns", lambda call: FakeResponse(204, []))
    fake_sb.route("PATCH", "pulse_campaign_events", lambda call: FakeResponse(204, []))

    r = client.post("/api/outbound-pulse/campaigns/a1/client",
                    data={"client_id": CLIENT})
    assert r.headers.get("hx-trigger") == "pulseMappingChanged"


# ── Correcting a wrong mapping ────────────────────────────────────────────────
# Remapping already worked; what was missing was finding the campaign again
# among ~1000 rows once it had been mapped to the wrong client.

def _remap_routes(fake_sb):
    rows = [
        {"id": "a1", "client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "external_campaign_id": "1", "name": "Wrongly mapped", "status": "ACTIVE",
         "last_synced_at": "2026-09-14T09:00:00+00:00", "created_at": None},
        {"id": "a2", "client_id": None, "channel": "email", "source_tool": "smartlead",
         "external_campaign_id": "2", "name": "Still unmapped", "status": "ACTIVE",
         "last_synced_at": None, "created_at": None},
    ]
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, rows))
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, [
        {"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True},
        {"id": "other", "name": "Globex", "color": None, "emoji": None, "active": True},
    ]))
    fake_sb.route("PATCH", "pulse_campaigns", lambda call: FakeResponse(204, []))
    fake_sb.route("PATCH", "pulse_campaign_events", lambda call: FakeResponse(204, []))


def test_filtering_by_client_finds_what_is_mapped_to_them(client, fake_sb):
    from tests.conftest import param_values

    _remap_routes(fake_sb)
    r = client.get(f"/api/outbound-pulse/campaigns?client_id={CLIENT}")
    assert r.status_code == 200

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "client_id")]
    assert param_values(listing[0], "client_id") == [f"eq.{CLIENT}"]


def test_unmapped_filter_is_distinct_from_no_filter(client, fake_sb):
    """"" means no filter; "none" means campaigns with no client. Conflating
    them would make the unmapped view silently show everything."""
    from tests.conftest import param_values

    _remap_routes(fake_sb)
    client.get("/api/outbound-pulse/campaigns?client_id=none")
    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "client_id")]
    assert param_values(listing[0], "client_id") == ["is.null"]

    fake_sb.calls.clear()
    client.get("/api/outbound-pulse/campaigns")
    for call in fake_sb.calls_to("GET", "pulse_campaigns"):
        assert param_values(call, "client_id") == []


def _filter_client_select(html: str) -> str:
    """Just the Client filter dropdown, not the per-row remap dropdowns."""
    import re
    match = re.search(
        r'<select name="client_id" class="filter-select">(.*?)</select>',
        html, re.S,
    )
    assert match, "client filter select not rendered"
    return match.group(1)


def test_client_filter_offers_only_clients_that_have_campaigns(client, fake_sb):
    """Offering every client when three have campaigns makes the filter a
    haystack of dead ends. The per-row remap dropdowns still list them all —
    you must be able to remap TO a client that has nothing yet."""
    _remap_routes(fake_sb)
    r = client.get("/api/outbound-pulse/campaigns")

    options = _filter_client_select(r.text)
    assert "Acme (1)" in options          # has a mapped campaign
    assert "Globex" not in options        # has none — not worth filtering by
    assert "— unmapped — (1)" in options

    # ...but Globex is still a remap target on the rows themselves.
    assert "Globex" in r.text


def test_remap_moves_the_campaign_and_its_history(client, fake_sb):
    """The fix for a wrong mapping: pick a different client. The campaign's
    events must follow, or the funnel keeps crediting the wrong client."""
    _remap_routes(fake_sb)

    r = client.post("/api/outbound-pulse/campaigns/a1/client",
                    data={"client_id": "other"})
    assert r.status_code == 200

    campaign_patch = fake_sb.calls_to("PATCH", "pulse_campaigns")[0]
    assert campaign_patch["json"]["client_id"] == "other"

    event_patch = fake_sb.calls_to("PATCH", "pulse_campaign_events")[0]
    assert event_patch["json"]["client_id"] == "other"


def test_remap_to_unmapped_clears_both(client, fake_sb):
    """Setting a campaign back to unmapped must also detach its events, or they
    stay credited to the old client while the campaign shows unmapped."""
    _remap_routes(fake_sb)

    client.post("/api/outbound-pulse/campaigns/a1/client", data={"client_id": ""})
    assert fake_sb.calls_to("PATCH", "pulse_campaigns")[0]["json"]["client_id"] is None
    assert fake_sb.calls_to("PATCH", "pulse_campaign_events")[0]["json"]["client_id"] is None


def test_remapping_keeps_the_client_filter(client, fake_sb):
    """The filter field is posted under its own name so it cannot be confused
    with the client_id being assigned."""
    from tests.conftest import param_values

    _remap_routes(fake_sb)
    r = client.post("/api/outbound-pulse/campaigns/a1/client",
                    data={"client_id": "other", "filter_client_id": CLIENT,
                          "status": "ACTIVE"})
    assert r.status_code == 200

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "client_id")]
    assert listing, "the re-render dropped the client filter"
    assert param_values(listing[0], "client_id") == [f"eq.{CLIENT}"]


def test_the_assigned_client_is_not_taken_from_the_filter(client, fake_sb):
    """Regression guard: two fields named client_id on one form would let the
    active filter overwrite the mapping the user actually chose."""
    _remap_routes(fake_sb)

    client.post("/api/outbound-pulse/campaigns/a1/client",
                data={"client_id": "other", "filter_client_id": CLIENT})

    assert fake_sb.calls_to("PATCH", "pulse_campaigns")[0]["json"]["client_id"] == "other"


def test_client_filter_view_explains_that_remapping_removes_the_row(client, fake_sb):
    _remap_routes(fake_sb)
    r = client.get(f"/api/outbound-pulse/campaigns?client_id={CLIENT}")
    assert "removes it from this filtered view" in r.text


# ── Name search ───────────────────────────────────────────────────────────────

def test_search_becomes_an_ilike_filter(client, fake_sb):
    from tests.conftest import param_values

    _remap_routes(fake_sb)
    client.get("/api/outbound-pulse/campaigns?q=acme")

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "name")]
    assert listing, "search never reached PostgREST"
    assert param_values(listing[0], "name") == ['ilike."*acme*"']


def test_search_quotes_names_containing_commas_and_parens(client, fake_sb):
    """Real campaign names look like "(CURATED) - 11-480 - B2C Health". An
    unquoted PostgREST value would have its commas and parens parsed as syntax
    and return the wrong rows, or none."""
    from tests.conftest import param_values

    _remap_routes(fake_sb)
    client.get("/api/outbound-pulse/campaigns?q=(CURATED) - 11-480, UK")

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "name")]
    value = param_values(listing[0], "name")[0]
    assert value.startswith('ilike."*') and value.endswith('*"')
    assert "(CURATED)" in value


def test_search_escapes_sql_wildcards():
    """A literal underscore must not match any character — "50_off" should not
    find "5000ff"."""
    assert store._ilike_value("50_off") == r'"*50\_off*"'
    assert store._ilike_value("100%") == r'"*100\%*"'


def test_search_term_is_length_capped(client, fake_sb):
    from tests.conftest import param_values

    _remap_routes(fake_sb)
    client.get("/api/outbound-pulse/campaigns?q=" + "a" * 500)

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "name")]
    assert len(param_values(listing[0], "name")[0]) < 130


def test_search_survives_a_mapping(client, fake_sb):
    from tests.conftest import param_values

    _remap_routes(fake_sb)
    client.post("/api/outbound-pulse/campaigns/a1/client",
                data={"client_id": "other", "q": "acme"})

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "name")]
    assert listing, "the re-render dropped the search term"


# ── Bulk mapping ──────────────────────────────────────────────────────────────

def test_bulk_map_assigns_every_selected_campaign(client, fake_sb):
    from tests.conftest import in_list_values, param_values

    _remap_routes(fake_sb)
    r = client.post("/api/outbound-pulse/campaigns/bulk-map",
                    data={"campaign_ids": ["a1", "a2"],
                          "target_client_id": "other"})
    assert r.status_code == 200

    patch = fake_sb.calls_to("PATCH", "pulse_campaigns")[0]
    assert patch["json"]["client_id"] == "other"
    assert sorted(in_list_values(param_values(patch, "id")[0])) == ["a1", "a2"]


def test_bulk_map_repoints_the_events_too(client, fake_sb):
    """Without this the funnel keeps crediting the old client for everything
    synced before the remap."""
    from tests.conftest import in_list_values, param_values

    _remap_routes(fake_sb)
    client.post("/api/outbound-pulse/campaigns/bulk-map",
                data={"campaign_ids": ["a1", "a2"], "target_client_id": "other"})

    patch = fake_sb.calls_to("PATCH", "pulse_campaign_events")[0]
    assert patch["json"]["client_id"] == "other"
    assert sorted(in_list_values(param_values(patch, "campaign_id")[0])) == ["a1", "a2"]


def test_bulk_map_is_agency_scoped(client, fake_sb):
    from tests.conftest import param_values

    _remap_routes(fake_sb)
    client.post("/api/outbound-pulse/campaigns/bulk-map",
                data={"campaign_ids": ["a1"], "target_client_id": "other"})

    for table in ("pulse_campaigns", "pulse_campaign_events"):
        for call in fake_sb.calls_to("PATCH", table):
            assert param_values(call, "agency_id") == [f"eq.{AGENCY}"]


def test_bulk_map_to_unmapped_clears_the_client(client, fake_sb):
    _remap_routes(fake_sb)
    client.post("/api/outbound-pulse/campaigns/bulk-map",
                data={"campaign_ids": ["a1"], "target_client_id": ""})

    assert fake_sb.calls_to("PATCH", "pulse_campaigns")[0]["json"]["client_id"] is None


def test_bulk_map_with_nothing_selected_is_refused(client, fake_sb):
    _remap_routes(fake_sb)
    r = client.post("/api/outbound-pulse/campaigns/bulk-map",
                    data={"target_client_id": "other"})
    assert "Select at least one campaign" in r.text
    assert fake_sb.calls_to("PATCH", "pulse_campaigns") == []


def test_bulk_map_does_not_reassign_to_the_active_client_filter(client, fake_sb):
    """The filter and the assignment target are separate fields. If they shared
    a name, bulk-mapping inside a filtered view would silently reassign every
    selected campaign back to the client being filtered on."""
    _remap_routes(fake_sb)
    client.post("/api/outbound-pulse/campaigns/bulk-map",
                data={"campaign_ids": ["a1"], "target_client_id": "other",
                      "client_id": CLIENT})         # the active filter

    assert fake_sb.calls_to("PATCH", "pulse_campaigns")[0]["json"]["client_id"] == "other"


def test_bulk_map_keeps_the_filtered_view(client, fake_sb):
    from tests.conftest import param_values

    _remap_routes(fake_sb)
    r = client.post("/api/outbound-pulse/campaigns/bulk-map",
                    data={"campaign_ids": ["a1"], "target_client_id": "other",
                          "client_id": "none", "status": "ACTIVE"})
    assert r.status_code == 200
    assert 'value="ACTIVE" selected' in r.text

    listing = [c for c in fake_sb.calls_to("GET", "pulse_campaigns")
               if param_values(c, "status")]
    assert param_values(listing[0], "client_id") == ["is.null"]


def test_bulk_map_chunks_large_selections(client, fake_sb):
    """The ids go into a PostgREST in.(...) filter on the URL; a few hundred
    UUIDs in one request would exceed the URL length limit and lose the lot."""
    _remap_routes(fake_sb)
    client.post("/api/outbound-pulse/campaigns/bulk-map",
                data={"campaign_ids": [f"id{i}" for i in range(120)],
                      "target_client_id": "other"})

    patches = fake_sb.calls_to("PATCH", "pulse_campaigns")
    assert len(patches) == 3                      # 120 ids at 50 per chunk
    for call in patches:
        assert len(call["url"]) < 4000


def test_bulk_map_fires_the_mapping_event_not_a_sync(client, fake_sb):
    _remap_routes(fake_sb)
    r = client.post("/api/outbound-pulse/campaigns/bulk-map",
                    data={"campaign_ids": ["a1"], "target_client_id": "other"})
    assert r.headers.get("hx-trigger") == "pulseMappingChanged"


def test_rows_carry_a_selection_checkbox(client, fake_sb):
    _remap_routes(fake_sb)
    r = client.get("/api/outbound-pulse/campaigns")
    assert 'name="campaign_ids"' in r.text
    assert 'id="bulk-select-all"' in r.text
    # The bulk bar starts hidden — it is an action on a selection, not furniture.
    assert 'id="pulse-bulk-bar" hidden' in r.text


# ── Failure messages name the real cause ──────────────────────────────────────
# A production statement timeout on pulse_funnel_daily surfaced as a bare
# "500 Server Error" plus advice to re-run a migration that had already run.

def _timeout_response(call):
    return FakeResponse(500, {
        "code": "57014",
        "message": "canceling statement due to statement timeout",
        "details": None, "hint": None,
    })


def test_timeout_surfaces_the_postgres_error_not_a_bare_500(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_funnel_daily", _timeout_response)

    r = client.get("/api/outbound-pulse/overview")
    assert r.status_code == 200
    assert "statement timeout" in r.text
    assert "57014" in r.text


def test_timeout_does_not_tell_you_to_rerun_the_schema_migration(client, fake_sb):
    """The migration had already run. Pointing at it wasted the investigation."""
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_funnel_daily", _timeout_response)

    r = client.get("/api/outbound-pulse/overview")
    assert "outbound_pulse_schema.sql" not in r.text
    assert "outbound_pulse_rollup.sql" in r.text


def test_missing_relation_still_points_at_the_schema_migration(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(404, {
        "code": "42P01", "message": 'relation "public.pulse_campaigns" does not exist',
    }))
    r = client.get("/api/outbound-pulse/overview")
    assert "outbound_pulse_schema.sql" in r.text


def test_unclassified_errors_get_no_misleading_advice(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(
        500, {"code": "XX000", "message": "something unexpected"}))

    r = client.get("/api/outbound-pulse/overview")
    assert "something unexpected" in r.text
    assert ".sql" not in r.text


def test_error_classification():
    assert store._describe_postgrest_error("t", FakeResponse(
        500, {"code": "57014", "message": "canceling statement due to statement timeout"})
    ).kind == "timeout"
    assert store._describe_postgrest_error("t", FakeResponse(
        404, {"code": "42P01", "message": "relation does not exist"})).kind == "missing"
    assert store._describe_postgrest_error("t", FakeResponse(
        404, {"code": "PGRST205", "message": "Could not find the table"})).kind == "missing"
    assert store._describe_postgrest_error("t", FakeResponse(
        500, {"code": "XX000", "message": "boom"})).kind == "error"


def test_non_json_error_body_does_not_mask_the_failure():
    exc = store._describe_postgrest_error("t", FakeResponse(
        502, None, text="<html>Bad Gateway</html>"))
    assert "HTTP 502" in str(exc)
    assert "Bad Gateway" in str(exc)


# ── Rollup migration file ─────────────────────────────────────────────────────
# The SQL itself is validated against a real Postgres (see the harness used when
# it was written). These guard the properties that make it safe to run on prod,
# so an edit cannot quietly remove them.

def _rollup_sql():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1]
            / "migrations" / "outbound_pulse_rollup.sql").read_text(encoding="utf-8")


def test_rollup_migration_is_one_transaction():
    sql = _rollup_sql()
    assert sql.index("BEGIN;") < sql.index("CREATE TABLE")
    assert sql.rstrip().endswith("COMMIT;")


def test_rollup_migration_locks_events_before_rebuilding():
    """Without the lock, a sync inserting mid-rebuild is double counted or lost."""
    sql = _rollup_sql()
    assert "LOCK TABLE pulse_campaign_events IN SHARE ROW EXCLUSIVE MODE" in sql
    assert sql.index("LOCK TABLE") < sql.index("TRUNCATE pulse_funnel_rollup")


def test_rollup_trigger_and_rebuild_bucket_days_identically():
    """If these drift, live and rebuilt rows land on different days near midnight."""
    sql = _rollup_sql()
    assert sql.count("(occurred_at AT TIME ZONE 'UTC')::date") == 2


def test_rollup_trigger_uses_a_transition_table():
    """Statement-level with a transition table is what excludes ON CONFLICT
    DO NOTHING duplicates from the counts."""
    sql = _rollup_sql()
    assert "REFERENCING NEW TABLE AS new_rows" in sql
    assert "FOR EACH STATEMENT" in sql


def test_rollup_view_keeps_the_column_contract():
    sql = _rollup_sql()
    view = sql[sql.index("CREATE VIEW pulse_funnel_daily"):]
    cols = ["agency_id", "client_id", "campaign_id", "channel",
            "source_tool", "event_type", "day", "events"]
    positions = [view.index(f".{c}") for c in cols]
    assert positions == sorted(positions), "view column order changed"
    assert "events      BIGINT" in sql


# ── The team's real Smartlead categories ──────────────────────────────────────
# Interested / Information Request = low-level interest; Meeting Request = the
# highest intent. Before this, Meeting Request counted only as a positive reply,
# so no live category ever reached the top stage and it read zero everywhere.

@pytest.mark.parametrize("category,expected", [
    ("Interested",          normalize.EVENT_POSITIVE_REPLY),
    ("Information Request", normalize.EVENT_POSITIVE_REPLY),
    ("Meeting Request",     normalize.EVENT_MEETING_BOOKED),
    ("Not Interested",      None),
    ("Do Not Contact",      None),
    ("Out Of Office",       None),
    ("Wrong Person",        None),
])
def test_live_smartlead_categories(category, expected):
    assert normalize.classify_reply(category) == expected


def test_meeting_request_is_no_longer_just_a_positive_reply():
    assert normalize.classify_reply("Meeting Request") != normalize.EVENT_POSITIVE_REPLY


def test_renamed_meeting_request_still_reaches_the_top_stage():
    assert normalize.classify_reply("Meeting Requested") == normalize.EVENT_MEETING_BOOKED
    assert normalize.classify_reply("Meeting Request ✅") == normalize.EVENT_MEETING_BOOKED


def test_negative_wins_when_a_category_contains_both_phrases():
    """When in doubt the funnel under-reports. A category naming both a refusal
    and a meeting must not count as a meeting."""
    assert normalize.classify_reply("Not interested in a meeting request") is None


def test_a_meeting_request_lead_counts_as_interested_too():
    """Every stage is a subset of the one above it: a lead asking for a meeting
    was also interested, so the funnel never shows more meetings than interest."""
    events = _smartlead_events([{
        "lead_email":    "jo@acme.com",
        "sent_time":     "2026-09-01T09:00:00Z",
        "reply_time":    "2026-09-02T09:00:00Z",
        "lead_category": "Meeting Request",
    }])
    types = [e["event_type"] for e in events]
    assert "positive_reply" in types
    assert "meeting_booked" in types


def test_resync_adds_the_meeting_stage_without_duplicating_interest():
    """History already synced under the old mapping has a positive_reply event
    but no meeting event. Re-syncing must add the meeting and reuse the existing
    interested row's dedupe key, so the backfill is automatic and exact."""
    row = {"lead_email": "jo@acme.com", "reply_time": "2026-09-02T09:00:00Z",
           "lead_category": "Meeting Request"}
    first = {e["event_type"]: e["dedupe_key"] for e in _smartlead_events([row])}
    second = {e["event_type"]: e["dedupe_key"] for e in _smartlead_events([row])}
    assert first["positive_reply"] == second["positive_reply"]
    assert first["meeting_booked"] == second["meeting_booked"]


def test_stage_labels_use_the_teams_vocabulary():
    """The top stage records meeting REQUESTS. Calling it "booked" would tell a
    client they have calls on the calendar that the data doesn't show."""
    assert normalize.STAGE_LABELS[normalize.EVENT_POSITIVE_REPLY] == "Interested"
    assert normalize.STAGE_LABELS[normalize.EVENT_MEETING_BOOKED] == "Meeting requested"


# ── Opens are shown only when they are tracked ────────────────────────────────

@pytest.mark.parametrize("opened,replied,tracked", [
    (0,    0,   False),   # nothing tracked, nothing replied
    (0,    40,  False),   # tracking off — the common case
    (12,   40,  False),   # tracking on for some campaigns only: incomplete
    (40,   40,  True),    # boundary: a lead opens before replying
    (900,  40,  True),    # tracking on
])
def test_opens_are_tracked(opened, replied, tracked):
    assert normalize.opens_are_tracked({"opened": opened, "replied": replied}) is tracked


def test_untracked_opens_are_dropped_from_the_funnel():
    rows = normalize.funnel_with_rates(
        {"sent": 1000, "opened": 0, "replied": 50,
         "positive_reply": 10, "meeting_booked": 4})
    assert [r["key"] for r in rows] == \
        ["sent", "replied", "positive_reply", "meeting_booked"]


def test_reply_rate_is_against_sent_when_opens_are_dropped():
    """With the Opened row gone, the reply rate must not be computed against a
    zero it no longer shows."""
    rows = {r["key"]: r for r in normalize.funnel_with_rates(
        {"sent": 1000, "opened": 0, "replied": 50,
         "positive_reply": 10, "meeting_booked": 4})}
    assert rows["replied"]["step_rate"] == 5.0            # 50 of 1000 sent
    assert rows["replied"]["overall"] is None             # would just repeat 5.0
    assert rows["positive_reply"]["step_rate"] == 20.0    # 10 of 50 replied
    assert rows["meeting_booked"]["step_rate"] == 40.0    # 4 of 10 interested


def test_tracked_opens_keep_their_stage():
    """A LinkedIn funnel (opened = connection accepted) or an email campaign
    with tracking on must still show the stage."""
    rows = normalize.funnel_with_rates(
        {"sent": 1000, "opened": 400, "replied": 50,
         "positive_reply": 10, "meeting_booked": 4})
    assert "opened" in [r["key"] for r in rows]


def _detail_routes_with(fake_sb, rows):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": "🏢", "active": True}]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(
        200, [{"id": CAMPAIGN, "client_id": CLIENT, "name": "Acme UK PR",
               "channel": "email", "source_tool": "smartlead",
               "external_campaign_id": "1", "status": "ACTIVE",
               "last_synced_at": None, "created_at": None}]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, rows))
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))


def _day(event_type, events, channel="email"):
    return {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": channel,
            "event_type": event_type, "events": events, "day": "2026-09-10"}


def test_client_page_hides_the_opened_column_when_untracked(client, fake_sb):
    _detail_routes_with(fake_sb, [_day("sent", 1000), _day("replied", 50)])
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert r.status_code == 200
    assert "<th>Opened</th>" not in r.text


def test_client_page_shows_the_opened_column_when_tracked(client, fake_sb):
    _detail_routes_with(fake_sb, [_day("sent", 1000), _day("opened", 400),
                                  _day("replied", 50)])
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert "<th>Opened</th>" in r.text


def test_client_page_reply_rate_is_replied_over_sent(client, fake_sb):
    """This column used to read funnel[2] by position. Once Opened can drop out,
    position 2 is the Interested stage — the column would have shown the wrong
    metric under the right heading."""
    _detail_routes_with(fake_sb, [_day("sent", 1000), _day("replied", 50),
                                  _day("positive_reply", 10)])
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert "<td>5.0%</td>" in r.text


def test_linkedin_opened_note_only_appears_with_the_opened_stage(client, fake_sb):
    untracked = [_day("sent", 800), _day("replied", 40),
                 _day("sent", 200, "linkedin"), _day("replied", 7, "linkedin")]
    _detail_routes_with(fake_sb, untracked)
    assert "connection request was accepted" not in \
        client.get(f"/outbound-pulse/clients/{CLIENT}").text

    tracked = untracked + [_day("opened", 300), _day("opened", 90, "linkedin")]
    _detail_routes_with(fake_sb, tracked)
    assert "connection request was accepted" in \
        client.get(f"/outbound-pulse/clients/{CLIENT}").text


# ── Client names: no icon ─────────────────────────────────────────────────────

def test_overview_client_card_has_no_icon(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": "🏢", "active": True}]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "event_type": "sent", "events": 100},
    ]))
    r = client.get("/api/outbound-pulse/overview")
    assert "pulse-client-name" in r.text
    assert "🏢" not in r.text
    assert "📈" not in r.text
    assert "pulse-client-emoji" not in r.text


def test_engagement_table_has_no_icon(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": "🏢", "active": True}]))
    fake_sb.route("GET", "pulse_portal_visits", lambda call: FakeResponse(
        200, [{"client_id": CLIENT, "viewed_at": "2026-09-10T10:00:00+00:00"}]))
    r = client.get("/api/outbound-pulse/engagement")
    assert "Acme" in r.text
    assert "🏢" not in r.text and "📈" not in r.text


# ── Client portal wording ─────────────────────────────────────────────────────

def test_portal_never_calls_meeting_requests_booked(client, fake_sb):
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, [{
        "id": "a1", "client_id": CLIENT, "label": "",
        "expires_at": None, "revoked_at": None, "view_count": 0,
    }]))
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"event_type": "sent", "events": 900, "channel": "email", "day": "2026-09-10"},
        {"event_type": "replied", "events": 45, "channel": "email", "day": "2026-09-10"},
        {"event_type": "meeting_booked", "events": 3, "channel": "email", "day": "2026-09-10"},
        {"event_type": "sent", "events": 100, "channel": "linkedin", "day": "2026-09-10"},
    ]))
    r = client.get("/portal/valid-token")
    assert r.status_code == 200
    assert "Meeting requests" in r.text
    assert "booked" not in r.text.lower()
    # Opens aren't tracked here, so the portal must not explain a stage it hides.
    assert "connection request was accepted" not in r.text
