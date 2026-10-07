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
    counts = {"sent": 1000, "opened": 400, "replied": 50, "positive_reply": 10,
              "information_request": 3, "meeting_booked": 1}
    rows = {r["key"]: r for r in normalize.funnel_with_rates(counts)}
    assert rows["sent"]["step_rate"] is None      # nothing above it
    assert rows["opened"]["step_rate"] == 40.0
    assert rows["replied"]["step_rate"] == 12.5   # of opened
    assert rows["replied"]["overall"] == 5.0      # of sent
    assert rows["leads"]["value"] == 4            # 3 info + 1 meeting; not interested
    assert rows["leads"]["step_rate"] == 8.0      # of replied
    assert rows["leads"]["overall"] == 0.4        # of sent = the lead rate


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
         "positive_reply": 0, "information_request": 0, "meeting_booked": 0})
    by_key = {r["key"]: r for r in rows}
    assert by_key["sent"]["is_top"] is True
    assert by_key["leads"]["is_top"] is False
    assert by_key["leads"]["step_rate"] is None   # parent is zero


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
    # Categories are outcomes now, not events — see the outcome tests.
    assert types == ["sent", "opened", "replied"]
    assert all(e["agency_id"] == AGENCY for e in events)
    assert all(e["client_id"] == CLIENT for e in events)


def test_smartlead_no_longer_emits_outcome_events():
    """Outcome events were append-only, so a lead could never leave a bucket.
    The connector must not write them any more."""
    events = _smartlead_events([{
        "lead_email":    "jo@acme.com",
        "sent_time":     "2025-03-01T09:00:00Z",
        "reply_time":    "2025-03-02T08:00:00Z",
        "lead_category": "Meeting Request",
    }])
    assert {e["event_type"] for e in events}.isdisjoint(normalize.OUTCOME_STAGES)


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
    # Replied is backfilled at the meeting time so it never reads below Leads;
    # the meeting itself becomes an outcome, not an event.
    assert types.count("replied") == 1
    assert set(types).isdisjoint(normalize.OUTCOME_STAGES)

    outcomes = meet_alfred.outcomes_from_activities(
        [{"profile_url": "https://linkedin.com/in/jo",
          "meeting booked": "2025-03-05T09:00:00Z"}],
        agency_id=AGENCY, campaign_id=CAMPAIGN)
    assert [o["stage"] for o in outcomes] == ["meeting_booked"]


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
    assert [e["event_type"] for e in events] == ["sent", "opened", "replied"]
    outcomes = meet_alfred.outcomes_from_activities(
        rows, agency_id=AGENCY, campaign_id=CAMPAIGN)
    assert [o["stage"] for o in outcomes] == ["positive_reply"]


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
    assert "Client Performance &amp; Reports" in r.text


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

def test_a_portal_link_is_stored_so_it_can_be_reopened(client, fake_sb):
    """Deliberately changed: the token used to be hashed and discarded, so a
    link was shown once and could only be re-issued. The team needs to open a
    client's report themselves, so it is kept.

    The hash remains the lookup key — the stored token is display only and is
    never part of resolving a link."""
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

    match = re.search(r"/r/([A-Za-z0-9_\-]+)", r.text)
    assert match, "the plaintext link should be shown once"
    token = match.group(1)

    assert captured["token"] == token
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
        return [], []

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
        return [], []

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
    monkeypatch.setattr(smartlead, "sync_campaign", lambda **kw: ([], []))
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
    ("Information Request", normalize.EVENT_INFORMATION_REQUEST),
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


def test_a_meeting_request_lead_is_not_also_interested():
    """Interested counts only prospects marked Interested. The categories are
    mutually exclusive, so a meeting request is a meeting request and nothing
    else."""
    assert normalize.classify_reply("Meeting Request") == normalize.EVENT_MEETING_BOOKED
    assert normalize.classify_reply("Information Request") == normalize.EVENT_INFORMATION_REQUEST
    assert normalize.classify_reply("Interested") == normalize.EVENT_POSITIVE_REPLY


def test_resyncing_the_same_lead_produces_the_same_outcome_key():
    """Outcomes upsert on (agency, campaign, lead). The same lead must map to the
    same row every sync, or a re-sync would add a lead instead of updating it."""
    row = {"lead_email": "Jo@Acme.com", "reply_time": "2026-09-02T09:00:00Z",
           "lead_category": "Meeting Request"}
    a = smartlead.outcomes_from_statistics([row], agency_id=AGENCY, campaign_id=CAMPAIGN)
    b = smartlead.outcomes_from_statistics([row], agency_id=AGENCY, campaign_id=CAMPAIGN)
    key = lambda o: (o["agency_id"], o["campaign_id"], o["lead_key"])
    assert [key(o) for o in a] == [key(o) for o in b]
    assert a[0]["lead_key"] == "jo@acme.com"


def test_stage_labels_use_the_teams_vocabulary():
    """The top stage records meeting REQUESTS. Calling it "booked" would tell a
    client they have calls on the calendar that the data doesn't show."""
    assert normalize.STAGE_LABELS[normalize.EVENT_POSITIVE_REPLY] == "Interested"
    assert normalize.STAGE_LABELS[normalize.EVENT_INFORMATION_REQUEST] == "Information requests"
    assert normalize.STAGE_LABELS[normalize.EVENT_MEETING_BOOKED] == "Meeting requests"
    assert "booked" not in " ".join(normalize.STAGE_LABELS.values()).lower()


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
         "positive_reply": 10, "information_request": 3, "meeting_booked": 1})
    assert [r["key"] for r in rows] == ["sent", "replied", "leads"]


def test_reply_rate_is_against_sent_when_opens_are_dropped():
    """With the Opened row gone, the reply rate must not be computed against a
    zero it no longer shows."""
    rows = {r["key"]: r for r in normalize.funnel_with_rates(
        {"sent": 1000, "opened": 0, "replied": 50,
         "positive_reply": 10, "information_request": 3, "meeting_booked": 1})}
    assert rows["replied"]["step_rate"] == 5.0    # 50 of 1000 sent
    assert rows["replied"]["overall"] is None     # would just repeat 5.0
    assert rows["leads"]["step_rate"] == 8.0      # 4 leads of 50 replied
    assert rows["leads"]["overall"] == 0.4        # the lead rate


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

# ── Leads and lead rate ───────────────────────────────────────────────────────
# Leads = Information requests + Meeting requests. Interested is tracked but is
# not a lead. Lead rate = leads / sent, the same base as the reply rate.

def test_lead_count_excludes_interested():
    counts = {"sent": 1000, "replied": 60, "positive_reply": 25,
              "information_request": 7, "meeting_booked": 3}
    assert normalize.lead_count(counts) == 10


def test_lead_rate_is_leads_over_sent():
    counts = {"sent": 2000, "positive_reply": 40,
              "information_request": 12, "meeting_booked": 8}
    assert normalize.lead_rate(counts) == 1.0


def test_lead_rate_has_no_value_without_sends():
    assert normalize.lead_rate({"information_request": 3}) is None


def test_outcome_breakdown_lists_all_three_with_share_of_replies():
    rows = normalize.outcome_breakdown(
        {"replied": 50, "positive_reply": 20, "information_request": 5, "meeting_booked": 5})
    by_key = {r["key"]: r for r in rows}
    assert [r["label"] for r in rows] == \
        ["Interested", "Information requests", "Meeting requests", "Unsubscribes"]
    assert by_key["positive_reply"]["of_replies"] == 40.0
    assert by_key["positive_reply"]["is_lead"] is False
    assert by_key["information_request"]["is_lead"] is True
    assert by_key["meeting_booked"]["is_lead"] is True


# ── Outcomes: one current row per replying lead ──────────────────────────────

def _outcomes(rows):
    return smartlead.outcomes_from_statistics(rows, agency_id=AGENCY, campaign_id=CAMPAIGN)


def test_outcomes_collapse_sequence_steps_to_one_row_per_lead():
    """A single upsert containing the same key twice fails in Postgres ("cannot
    affect row a second time") and would lose the whole batch."""
    out = _outcomes([
        {"lead_email": "jo@acme.com", "sequence_number": 1,
         "sent_time": "2026-09-01T09:00:00Z", "lead_category": "Interested"},
        {"lead_email": "jo@acme.com", "sequence_number": 2,
         "reply_time": "2026-09-03T09:00:00Z", "lead_category": "Interested"},
        {"lead_email": "jo@acme.com", "sequence_number": 3,
         "sent_time": "2026-09-05T09:00:00Z", "lead_category": "Interested"},
    ])
    assert len(out) == 1
    assert out[0]["stage"] == "positive_reply"
    assert out[0]["day"] == "2026-09-03"


def test_outcome_day_is_the_earliest_reply():
    out = _outcomes([
        {"lead_email": "jo@acme.com", "reply_time": "2026-09-09T09:00:00Z",
         "lead_category": "Meeting Request"},
        {"lead_email": "jo@acme.com", "reply_time": "2026-09-04T09:00:00Z",
         "lead_category": "Meeting Request"},
    ])
    assert out[0]["day"] == "2026-09-04"


def test_leads_that_never_replied_have_no_outcome():
    assert _outcomes([{"lead_email": "jo@acme.com", "sent_time": "2026-09-01T09:00:00Z",
                       "lead_category": "Interested"}]) == []


def test_a_non_positive_reply_is_still_written_with_no_stage():
    """Writing a NULL stage is what removes a lead from the counts when it is
    re-marked Not Interested — the upsert overwrites its old stage."""
    out = _outcomes([{"lead_email": "jo@acme.com", "reply_time": "2026-09-02T09:00:00Z",
                      "lead_category": "Not Interested"}])
    assert len(out) == 1
    assert out[0]["stage"] is None
    assert out[0]["category"] == "Not Interested"


def test_recategorised_lead_moves_bucket_on_the_next_sync():
    """The core reason outcomes are current state: the same lead, re-marked,
    produces the same row key with a different stage."""
    before = _outcomes([{"lead_email": "jo@acme.com", "reply_time": "2026-09-02T09:00:00Z",
                         "lead_category": "Interested"}])
    after = _outcomes([{"lead_email": "jo@acme.com", "reply_time": "2026-09-02T09:00:00Z",
                        "lead_category": "Meeting Request"}])
    assert before[0]["lead_key"] == after[0]["lead_key"]
    assert (before[0]["stage"], after[0]["stage"]) == ("positive_reply", "meeting_booked")


def test_make_outcome_rejects_unknown_stages():
    with pytest.raises(ValueError):
        normalize.make_outcome(agency_id=AGENCY, campaign_id=CAMPAIGN, lead="x",
                               category="", stage="booked_call",
                               replied_at=normalize.parse_ts("2026-09-02T09:00:00Z"))


def test_alfred_negative_category_beats_a_tracked_meeting_date():
    """The category is the team's latest judgement; an old meeting date in the
    export must not override an explicit "Not Interested"."""
    out = meet_alfred.outcomes_from_activities(
        [{"profile url": "https://linkedin.com/in/jo",
          "replied at": "2026-09-02T09:00:00Z",
          "meeting booked": "2026-09-05T09:00:00Z",
          "category": "Not Interested"}],
        agency_id=AGENCY, campaign_id=CAMPAIGN)
    assert out[0]["stage"] is None


# ── Writing outcomes ──────────────────────────────────────────────────────────

def test_upsert_outcomes_merges_on_the_lead_key(fake_sb):
    from tests.conftest import param_values

    fake_sb.route("POST", "pulse_lead_outcomes", lambda call: FakeResponse(201, []))
    rows = _outcomes([{"lead_email": "jo@acme.com", "reply_time": "2026-09-02T09:00:00Z",
                       "lead_category": "Interested"}])
    assert store.upsert_outcomes(rows) == 1

    call = fake_sb.calls_to("POST", "pulse_lead_outcomes")[0]
    assert param_values(call, "on_conflict") == ["agency_id,campaign_id,lead_key"]
    assert "merge-duplicates" in call["headers"].get("Prefer", "")


def test_upsert_outcomes_reports_a_shortfall(fake_sb, monkeypatch):
    monkeypatch.setattr(store, "INSERT_CHUNK", 1)
    calls = {"n": 0}

    def handler(call):
        calls["n"] += 1
        return FakeResponse(404 if calls["n"] == 2 else 201,
                            {"code": "42P01", "message": "relation does not exist"}
                            if calls["n"] == 2 else [])

    fake_sb.route("POST", "pulse_lead_outcomes", handler)
    rows = _outcomes([
        {"lead_email": f"l{i}@acme.com", "reply_time": "2026-09-02T09:00:00Z",
         "lead_category": "Interested"} for i in range(3)])
    assert store.upsert_outcomes(rows) == 2


def test_sync_flags_campaigns_whose_outcomes_did_not_save(monkeypatch, fake_sb):
    """Lost outcomes must show on the connector panel, and the campaign must stay
    at the head of the queue — not be stamped synced and rotated away."""
    from app.utils.pulse import sync as pulse_sync

    monkeypatch.setenv("SMARTLEAD_API_KEY", "k")
    monkeypatch.setattr(smartlead, "fetch_campaigns", lambda: [])
    monkeypatch.setattr(pulse_sync, "_PACING_SECONDS", 0)
    outcome = _outcomes([{"lead_email": "jo@acme.com",
                          "reply_time": "2026-09-02T09:00:00Z",
                          "lead_category": "Interested"}])
    monkeypatch.setattr(smartlead, "sync_campaign", lambda **kw: ([], outcome))

    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, [
        {"id": "id1", "client_id": None, "external_campaign_id": "1",
         "name": "Acme", "status": "ACTIVE", "last_synced_at": None}]))
    fake_sb.route("POST", "pulse_lead_outcomes", lambda call: FakeResponse(
        404, {"code": "42P01", "message": "relation does not exist"}))
    fake_sb.route("POST", "pulse_sync_logs", lambda call: FakeResponse(201, [{"id": "log1"}]))

    result = pulse_sync.sync_source("smartlead", "test")
    assert result["status"] == "error"
    assert "lead outcomes" in result["error"]
    assert fake_sb.calls_to("PATCH", "pulse_campaigns") == []


# ── Views ─────────────────────────────────────────────────────────────────────

_LEAD_ROWS = [
    {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": "email", "day": "2026-09-10",
     "event_type": "sent", "events": 2000},
    {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": "email", "day": "2026-09-10",
     "event_type": "replied", "events": 80},
    {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": "email", "day": "2026-09-10",
     "event_type": "positive_reply", "events": 30},
    {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": "email", "day": "2026-09-10",
     "event_type": "information_request", "events": 12},
    {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": "email", "day": "2026-09-10",
     "event_type": "meeting_booked", "events": 8},
]


def test_overview_shows_lead_rate_and_all_three_outcomes(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, _LEAD_ROWS))

    r = client.get("/api/outbound-pulse/overview")
    assert r.status_code == 200
    assert "Lead rate" in r.text
    assert "1.0%" in r.text                      # 20 leads / 2000 sent
    assert "Information requests" in r.text
    assert "Meeting requests" in r.text
    assert ">30<" in r.text                      # Interested stays separate


def test_client_detail_table_has_lead_columns(client, fake_sb):
    _detail_routes_with(fake_sb, _LEAD_ROWS)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert r.status_code == 200
    for heading in ("<th>Interested</th>", "<th>Info req.</th>", "<th>Meeting req.</th>",
                    "<th>Leads</th>", "<th>Lead rate</th>"):
        assert heading in r.text
    assert "<td>20</td>" in r.text               # leads
    assert "<td>1.0%</td>" in r.text             # lead rate
    assert "<td>4.0%</td>" in r.text             # reply rate, 80/2000


# ── The SQL backfill must classify exactly like Python ───────────────────────

def _sql_array(sql: str, name: str) -> tuple[str, ...]:
    import re
    # [^\]]* rather than .*? — a lazy match starts at the FIRST "ARRAY[" and
    # would swallow every list before the one named.
    match = re.search(r"ARRAY\[([^\]]*)\]\s+AS " + name, sql)
    assert match, f"ARRAY for {name} not found in migration"
    return tuple(re.findall(r"'([^']*)'", match.group(1)))


def _outcomes_sql():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1]
            / "migrations" / "outbound_pulse_outcomes.sql").read_text(encoding="utf-8")


@pytest.mark.parametrize("name,python", [
    ("meeting",     "_MEETING_CATEGORIES"),
    ("information", "_INFORMATION_CATEGORIES"),
    ("interested",  "_INTERESTED_CATEGORIES"),
    ("negative",    "_NEGATIVE_CATEGORIES"),
])
def test_backfill_term_lists_match_the_python_classifier(name, python):
    """The migration's backfill classifies history in SQL. If its term lists
    drift from normalize.py, the backfilled numbers disagree with what the next
    sync writes, and the dashboard shifts for no visible reason."""
    assert set(_sql_array(_outcomes_sql(), name)) == set(getattr(normalize, python))


def test_outcomes_migration_is_one_transaction_and_never_clobbers_live_rows():
    sql = _outcomes_sql()
    assert sql.index("BEGIN;") < sql.index("CREATE TABLE")
    assert sql.rstrip().endswith("COMMIT;")
    # A re-run must not overwrite a live sync's current category with history.
    assert "ON CONFLICT (agency_id, campaign_id, lead_key) DO NOTHING" in sql
    # The backfill classifier must not outlive the transaction.
    assert "DROP FUNCTION pulse_classify_backfill(text);" in sql


def test_view_reads_outcomes_from_the_table_not_legacy_events():
    """Legacy positive_reply / meeting_booked events stay in the rollup. If the
    view read them too, every outcome would count twice."""
    view = _outcomes_sql().split("CREATE VIEW pulse_funnel_daily AS", 1)[1]
    assert "WHERE r.event_type IN ('sent', 'opened', 'replied')" in view
    assert "FROM pulse_lead_outcomes o" in view
    assert "WHERE o.stage IS NOT NULL" in view


# ── Campaign mapping is its own tab ───────────────────────────────────────────

def test_page_has_a_tab_strip(client, fake_sb):
    r = client.get("/outbound-pulse")
    assert r.status_code == 200
    assert 'role="tablist"' in r.text
    assert 'data-tab="reporting"' in r.text
    assert 'data-tab="mapping"' in r.text
    assert "Campaign mapping" in r.text


def test_mapping_lives_in_its_own_panel_hidden_by_default(client, fake_sb):
    r = client.get("/outbound-pulse")
    assert '<section id="tab-mapping" role="tabpanel" aria-labelledby="tab-btn-mapping" hidden>' in r.text
    # The mapping table markup must sit inside that panel, not the reporting one.
    mapping_panel = r.text.split('<section id="tab-mapping"', 1)[1]
    assert 'id="pulse-campaigns-list"' in mapping_panel
    reporting_panel = r.text.split('<section id="tab-reporting"', 1)[1].split('<section id="tab-mapping"', 1)[0]
    assert 'id="pulse-campaigns-list"' not in reporting_panel
    assert 'id="pulse-overview"' in reporting_panel


def test_mapping_table_is_not_fetched_until_its_tab_is_opened(client, fake_sb):
    """One row per campaign, and this account has over a thousand — loading it
    with the page made every visit pay for a table most visits never open."""
    r = client.get("/outbound-pulse")
    block = r.text.split('id="pulse-campaigns-list"', 1)[1].split(">", 1)[0]
    assert 'hx-trigger="pulseMappingTab from:body"' in block
    assert "load" not in block.replace("pulseMappingTab", "")


def test_unmapped_warning_switches_tab_rather_than_scrolling(client, fake_sb):
    """The mapping panel is on another tab now, so scrolling to it would land
    on a hidden element."""
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, [
        {"id": CAMPAIGN, "client_id": None, "name": "Stray", "channel": "email",
         "source_tool": "smartlead", "external_campaign_id": "1", "status": "ACTIVE",
         "last_synced_at": None, "created_at": None}]))

    r = client.get("/api/outbound-pulse/overview")
    assert "not mapped to a client" in r.text
    assert "pulseShowMapping" in r.text
    assert "scrollIntoView" not in r.text


# ── Manual leads count as leads ───────────────────────────────────────────────

def test_manual_lead_is_a_lead_but_not_a_meeting_or_info_request():
    """The DNC & Merger tool records a lead without saying which kind, so it is
    its own stage rather than being folded into one that implies more."""
    assert normalize.EVENT_MANUAL_LEAD in normalize.LEAD_STAGES
    assert normalize.EVENT_MANUAL_LEAD not in (
        normalize.EVENT_INFORMATION_REQUEST, normalize.EVENT_MEETING_BOOKED)
    assert normalize.lead_count({"lead": 4}) == 4
    assert normalize.lead_count(
        {"lead": 4, "information_request": 2, "meeting_booked": 1}) == 7
    assert normalize.lead_rate({"sent": 1000, "lead": 5}) == 0.5


def test_interested_still_excludes_manual_leads():
    counts = {"replied": 50, "positive_reply": 20, "lead": 9}
    assert normalize.lead_count(counts) == 9
    by_key = {r["key"]: r for r in normalize.outcome_breakdown(counts)}
    assert by_key["positive_reply"]["value"] == 20
    assert by_key["positive_reply"]["is_lead"] is False
    assert by_key["lead"]["is_lead"] is True


def test_manual_lead_stage_is_hidden_when_there_are_none():
    """A Smartlead-only funnel should not show a permanent zero for a stage
    that source cannot produce."""
    keys = [r["key"] for r in normalize.outcome_breakdown(
        {"replied": 10, "positive_reply": 4, "meeting_booked": 1})]
    assert "lead" not in keys
    keys = [r["key"] for r in normalize.outcome_breakdown({"replied": 10, "lead": 3})]
    assert "lead" in keys


# ── Splitting by source ───────────────────────────────────────────────────────

def test_funnel_by_source_groups_the_three_tools(fake_sb):
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"source_tool": "smartlead", "event_type": "sent", "events": 900},
        {"source_tool": "smartlead", "event_type": "meeting_booked", "events": 4},
        {"source_tool": "meet_alfred", "event_type": "sent", "events": 200},
        {"source_tool": "manual", "event_type": "sent", "events": 300},
        {"source_tool": "manual", "event_type": "lead", "events": 6},
    ]))
    out = store.funnel_by_source()
    assert out["smartlead"]["sent"] == 900
    assert out["manual"]["lead"] == 6
    assert normalize.lead_count(out["manual"]) == 6
    assert normalize.lead_count(out["meet_alfred"]) == 0


def test_source_filter_reaches_the_query(fake_sb):
    from tests.conftest import param_values

    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))
    store.funnel(source_tool="manual")
    call = fake_sb.calls_to("GET", "pulse_funnel_daily")[0]
    assert param_values(call, "source_tool") == ["eq.manual"]
    assert param_values(call, "agency_id") == [f"eq.{AGENCY}"]


# ── Per-source tabs on the client view ────────────────────────────────────────

def _source_routes(fake_sb, rows):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, rows))
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))


_MIXED = [
    {"client_id": CLIENT, "source_tool": "smartlead", "channel": "email",
     "event_type": "sent", "events": 1000, "day": "2026-09-10"},
    {"client_id": CLIENT, "source_tool": "smartlead", "channel": "email",
     "event_type": "replied", "events": 40, "day": "2026-09-10"},
    {"client_id": CLIENT, "source_tool": "smartlead", "channel": "email",
     "event_type": "meeting_booked", "events": 5, "day": "2026-09-10"},
    {"client_id": CLIENT, "source_tool": "manual", "channel": "email",
     "event_type": "sent", "events": 500, "day": "2026-09-10"},
    {"client_id": CLIENT, "source_tool": "manual", "channel": "email",
     "event_type": "lead", "events": 7, "day": "2026-09-10"},
]


def test_client_page_has_a_tab_per_source_plus_combined(client, fake_sb):
    _source_routes(fake_sb, _MIXED)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert r.status_code == 200
    assert 'data-source="all"' in r.text
    for key, label in (("smartlead", "Smartlead"), ("meet_alfred", "Meet Alfred"),
                       ("manual", "Manual")):
        assert f'data-source="{key}"' in r.text
        assert label in r.text


def test_a_source_with_no_activity_is_disabled_not_hidden(client, fake_sb):
    """Meet Alfred has nothing in this range — the tab should still be visible,
    so its absence reads as "nothing yet", not "not connected"."""
    _source_routes(fake_sb, _MIXED)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    alfred = r.text.split('data-source="meet_alfred"', 1)[1].split(">", 1)[0]
    assert "disabled" in alfred
    smartlead = r.text.split('data-source="smartlead"', 1)[1].split(">", 1)[0]
    assert "disabled" not in smartlead


def test_each_source_panel_shows_only_that_source(client, fake_sb):
    _source_routes(fake_sb, _MIXED)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    manual = r.text.split('data-source-panel="manual"', 1)[1].split('data-source-panel=', 1)[0]
    assert "1,000" not in manual          # the Smartlead sends
    assert "500" in manual                # its own
    assert "Manual leads" in manual
    assert "DNC &amp; Merger tool" in manual


def test_manual_panel_states_its_caveats(client, fake_sb):
    """Manual leads are deduplicated by domain and replies carry no date; both
    make the numbers differ from the DNC & Merger tool, so both are stated."""
    _source_routes(fake_sb, _MIXED)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    manual = r.text.split('data-source-panel="manual"', 1)[1].split('data-source-panel=', 1)[0]
    assert "one per domain per client" in manual
    assert "manual replies are" in manual


def test_combined_notes_that_manual_has_no_replies(client, fake_sb):
    _source_routes(fake_sb, _MIXED)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    combined = r.text.split('data-source-panel="all"', 1)[1].split('data-source-panel=', 1)[0]
    assert "Smartlead and Meet Alfred only" in combined


def test_combined_totals_include_manual_leads(client, fake_sb):
    _source_routes(fake_sb, _MIXED)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    combined = r.text.split('data-source-panel="all"', 1)[1].split('data-source-panel=', 1)[0]
    assert "12 leads" in combined          # 5 meeting requests + 7 manual leads


# ── Date range selector ───────────────────────────────────────────────────────

def test_client_page_offers_a_custom_range(client, fake_sb):
    _source_routes(fake_sb, _MIXED)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert 'value="custom"' in r.text
    assert 'id="date-from"' in r.text and 'id="date-to"' in r.text


def test_custom_dates_filter_the_client_view(client, fake_sb):
    from tests.conftest import param_values

    _source_routes(fake_sb, _MIXED)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}"
                   "?date_from=2026-09-01&date_to=2026-09-15")
    assert r.status_code == 200
    # The inputs come back filled in, so the range survives a reload.
    assert 'value="2026-09-01"' in r.text and 'value="2026-09-15"' in r.text

    call = fake_sb.calls_to("GET", "pulse_funnel_daily")[0]
    assert sorted(param_values(call, "day")) == ["gte.2026-09-01", "lte.2026-09-15"]


def test_overview_accepts_a_custom_range(client, fake_sb):
    from tests.conftest import param_values

    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    r = client.get("/api/outbound-pulse/overview?date_from=2026-08-01&date_to=2026-08-31")
    assert r.status_code == 200
    call = fake_sb.calls_to("GET", "pulse_funnel_daily")[0]
    assert sorted(param_values(call, "day")) == ["gte.2026-08-01", "lte.2026-08-31"]


def test_overview_sends_the_date_inputs_with_every_refresh(client, fake_sb):
    """Without them in hx-include, picking a custom range then changing channel
    would silently drop back to the preset."""
    r = client.get("/outbound-pulse")
    overview = r.text.split('id="pulse-overview"', 1)[1].split(">", 1)[0]
    assert "#date-from" in overview and "#date-to" in overview


# ── Manual migration guards ───────────────────────────────────────────────────

def _manual_sql():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1]
            / "migrations" / "outbound_pulse_manual.sql").read_text(encoding="utf-8")


def test_manual_migration_reads_in_place_and_is_one_transaction():
    sql = _manual_sql()
    assert sql.index("BEGIN;") < sql.index("CREATE INDEX")
    assert sql.rstrip().endswith("COMMIT;")
    assert "pulse_manual_daily" in sql


def test_manual_view_counts_only_lead_and_interested_reasons():
    """opt_out and hand-added DNC entries are not campaign outcomes."""
    sql = _manual_sql()
    assert "d.reason IN ('lead', 'interested')" in sql
    assert "opt_out" not in sql.split("CREATE OR REPLACE VIEW", 1)[-1]


def test_manual_view_excludes_clients_without_an_agency():
    assert _manual_sql().count("c.agency_id IS NOT NULL") >= 2


def test_funnel_view_unions_manual_without_duplicating_events():
    sql = _manual_sql()
    view = sql.split("CREATE VIEW pulse_funnel_daily AS", 1)[1]
    assert "WHERE r.event_type IN ('sent', 'opened', 'replied')" in view
    assert "FROM pulse_lead_outcomes o" in view
    assert "FROM pulse_manual_daily" in view


# ── Trend buckets ─────────────────────────────────────────────────────────────

def _days(n, start="2024-04-11", sent=10):
    from datetime import date, timedelta
    d0 = date.fromisoformat(start)
    return [{"day": (d0 + timedelta(days=i)).isoformat(),
             "sent": sent, "replied": 1, "meeting_booked": 1} for i in range(n)]


def test_a_short_series_stays_one_column_per_day():
    from app.utils.pulse.normalize import bucket_timeseries

    out = bucket_timeseries(_days(30))
    assert out["unit"] == "day"
    assert len(out["buckets"]) == 30


def test_a_long_series_is_bucketed_instead_of_overflowing():
    """The all-time range is years wide. One column per day did not fit."""
    from app.utils.pulse.normalize import MAX_TREND_COLUMNS, bucket_timeseries

    for span in (61, 200, 413, 414, 900, 3000):
        out = bucket_timeseries(_days(span))
        assert len(out["buckets"]) <= MAX_TREND_COLUMNS, (span, out["unit"])
        assert out["unit"] != "day"


def test_bucketing_never_loses_or_invents_events():
    from app.utils.pulse.normalize import bucket_timeseries

    rows = _days(900)
    for unit_rows in (rows, _days(61), _days(30)):
        out = bucket_timeseries(unit_rows)
        assert sum(b["sent"] for b in out["buckets"]) == sum(r["sent"] for r in unit_rows)
        assert sum(b["replied"] for b in out["buckets"]) == len(unit_rows)


def test_buckets_are_whole_weeks_and_months_not_arbitrary_windows():
    """A reader must be able to name the period a bar covers."""
    from datetime import date

    from app.utils.pulse.normalize import bucket_timeseries

    weeks = bucket_timeseries(_days(200))
    assert weeks["unit"] == "week"
    for b in weeks["buckets"]:
        assert date.fromisoformat(b["day"]).weekday() == 0        # Monday

    months = bucket_timeseries(_days(900))
    assert months["unit"] == "month"
    for b in months["buckets"]:
        assert date.fromisoformat(b["day"]).day == 1


def test_buckets_come_back_oldest_first():
    from app.utils.pulse.normalize import bucket_timeseries

    for n in (30, 200, 900):
        days = [b["day"] for b in bucket_timeseries(_days(n))["buckets"]]
        assert days == sorted(days)


def test_an_empty_series_buckets_to_nothing_rather_than_raising():
    from app.utils.pulse.normalize import bucket_timeseries

    assert bucket_timeseries([])["buckets"] == []


def test_a_bucket_is_labelled_with_the_period_it_covers():
    from app.utils.pulse.normalize import bucket_timeseries

    assert bucket_timeseries(_days(5))["buckets"][0]["label"] == "11 Apr 2024"
    assert bucket_timeseries(_days(900))["buckets"][0]["label"] == "Apr 2024"
    assert "–" in bucket_timeseries(_days(200))["buckets"][0]["label"]


# ── Trend tooltip ─────────────────────────────────────────────────────────────

def test_trend_columns_carry_their_figures_for_the_tooltip(client, fake_sb):
    """A native `title` needed a second of stillness and showed nothing on
    hover, which is what made the strip look dead."""
    _detail_routes(fake_sb)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert 'class="trend-col"' in r.text
    for attr in ("data-label", "data-sent", "data-replied",
                 "data-leads", "data-lead-rate"):
        assert attr in r.text
    assert "trend-tip" in r.text


# ── Manual in the reporting scope picker ──────────────────────────────────────

def test_the_channel_picker_offers_manual(client, fake_sb):
    _detail_routes(fake_sb)
    for path in ("/outbound-pulse", f"/outbound-pulse/clients/{CLIENT}"):
        assert 'value="manual"' in client.get(path).text


def test_picking_manual_filters_by_source_not_channel(client, fake_sb):
    from tests.conftest import param_values

    _detail_routes(fake_sb)
    client.get(f"/outbound-pulse/clients/{CLIENT}?channel=manual")
    calls = fake_sb.calls_to("GET", "pulse_funnel_daily")
    assert calls
    for call in calls:
        assert param_values(call, "source_tool") == ["eq.manual"]
        assert param_values(call, "channel") == []


def test_each_scope_is_one_tool_so_they_do_not_double_count(client, fake_sb):
    """Manual outreach is email. Filtering Email by channel would have counted
    it under Email and under Manual, so Both stopped being their sum."""
    from tests.conftest import param_values

    from app.routers.outbound_pulse import REPORT_SCOPES

    assert len(set(REPORT_SCOPES.values())) == len(REPORT_SCOPES)

    for scope, source in REPORT_SCOPES.items():
        fake_sb.calls.clear()
        _detail_routes(fake_sb)
        client.get(f"/outbound-pulse/clients/{CLIENT}?channel={scope}")
        for call in fake_sb.calls_to("GET", "pulse_funnel_daily"):
            assert param_values(call, "source_tool") == [f"eq.{source}"]


def test_an_unknown_scope_falls_back_to_everything(client, fake_sb):
    from tests.conftest import param_values

    _detail_routes(fake_sb)
    client.get(f"/outbound-pulse/clients/{CLIENT}?channel=carrier-pigeon")
    for call in fake_sb.calls_to("GET", "pulse_funnel_daily"):
        assert param_values(call, "source_tool") == []


def test_the_mapping_filters_still_filter_by_a_real_channel(client, fake_sb):
    """Campaigns have a channel column; only the reporting toolbar changed."""
    from app.routers.outbound_pulse import _channel_filter

    assert _channel_filter("email") == "email"
    assert _channel_filter("manual") == ""


# ── Portal chrome ─────────────────────────────────────────────────────────────

def test_the_report_logo_file_is_in_the_repo():
    """The template would render a broken image if this were only on a laptop."""
    import pathlib

    logo = (pathlib.Path(__file__).resolve().parents[1]
            / "app" / "static" / "img" / "logo-report.png")
    assert logo.is_file() and logo.stat().st_size > 1000


# ── Account manager notes ─────────────────────────────────────────────────────

# ── The notes migration ───────────────────────────────────────────────────────

# ── The client link ───────────────────────────────────────────────────────────

def test_a_new_link_uses_the_short_path(client, fake_sb):
    fake_sb.route("POST", "pulse_client_access", lambda call: FakeResponse(
        201, [{"id": "a1", "label": "", "created_at": None, "expires_at": None,
               "revoked_at": None, "view_count": 0}]))
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/access", data={"label": ""})
    assert "/r/" in r.text
    assert "/portal/" not in r.text


def test_the_token_is_shorter_but_still_unguessable(client, fake_sb):
    """22 URL-safe characters is 128 bits. The old 43 were twice the length for
    security nobody was going to exhaust either way."""
    import re

    seen = {}

    def capture(call):
        seen.update(call.get("json") or {})
        return FakeResponse(201, [{"id": "a1", "label": "", "created_at": None,
                                   "expires_at": None, "revoked_at": None,
                                   "view_count": 0}])

    fake_sb.route("POST", "pulse_client_access", capture)
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/access", data={"label": ""})
    token = re.search(r"/r/([A-Za-z0-9_-]+)", r.text).group(1)
    assert len(token) == 22
    # Only the hash is stored, whatever the token's length.
    assert token not in seen["token_hash"]
    assert len(seen["token_hash"]) == 64


def test_links_already_handed_out_still_work(client, fake_sb):
    """A report link that stops working is a client emailing about it."""
    _portal_routes(fake_sb)
    assert client.get("/portal/valid-token").status_code == 200
    assert client.get("/r/valid-token").status_code == 200


def test_both_paths_look_up_the_same_hash(client, fake_sb):
    from tests.conftest import param_values

    _portal_routes(fake_sb)
    client.get("/portal/same-token")
    client.get("/r/same-token")
    hashes = [param_values(c, "token_hash")[0]
              for c in fake_sb.calls_to("GET", "pulse_client_access")]
    assert len(set(hashes)) == 1


def test_the_short_path_is_public_like_the_long_one():
    from app import auth

    assert auth.is_public_path("/r/sometoken")
    assert auth.is_public_path("/portal/sometoken")
    # Not a blanket pass for anything starting with r.
    assert not auth.is_public_path("/reply-bank")


def test_a_custom_domain_replaces_the_render_host(client, fake_sb, monkeypatch):
    """PORTAL_BASE_URL is what makes the link professional; the rest is length."""
    fake_sb.route("POST", "pulse_client_access", lambda call: FakeResponse(
        201, [{"id": "a1", "label": "", "created_at": None, "expires_at": None,
               "revoked_at": None, "view_count": 0}]))
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))

    monkeypatch.setenv("PORTAL_BASE_URL", "reports.unstuck.agency")
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/access", data={"label": ""})
    assert "https://reports.unstuck.agency/r/" in r.text
    assert "testserver" not in r.text


def test_a_custom_domain_keeps_a_scheme_it_was_given(client, fake_sb, monkeypatch):
    fake_sb.route("POST", "pulse_client_access", lambda call: FakeResponse(
        201, [{"id": "a1", "label": "", "created_at": None, "expires_at": None,
               "revoked_at": None, "view_count": 0}]))
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, []))

    monkeypatch.setenv("PORTAL_BASE_URL", "https://reports.unstuck.agency/")
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/access", data={"label": ""})
    assert "https://reports.unstuck.agency/r/" in r.text
    assert "agency//r/" not in r.text       # the trailing slash is stripped


# ── Headline rates ────────────────────────────────────────────────────────────

def _counts(**kw):
    c = normalize.empty_funnel()
    c.update(kw)
    return c


def test_reply_rate_counts_unsubscribes_as_responses():
    """An unsubscribe is a response to the mail. Leaving it out under-reports
    how many people the sequence actually reached."""
    c = _counts(sent=1000, replied=40, unsubscribed=10)
    assert normalize.response_count(c) == 50
    assert normalize.reply_rate(c) == 5.0


def test_reply_rate_is_unchanged_where_nothing_unsubscribed():
    c = _counts(sent=1000, replied=40)
    assert normalize.reply_rate(c) == 4.0


def test_interest_rate_is_the_interested_category_over_sends():
    """A sibling of the lead rate, not a total of positives: Interested is the
    soft category and Leads the hard one."""
    c = _counts(sent=1000, replied=40, positive_reply=12,
                information_request=3, meeting_booked=5)
    assert normalize.interest_rate(c) == 1.2
    assert normalize.lead_rate(c) == 0.8


def test_the_three_rates_share_one_base_and_narrow_left_to_right():
    c = _counts(sent=1000, replied=40, unsubscribed=10, positive_reply=12,
                information_request=3, meeting_booked=5)
    rates = normalize.headline_rates(c)
    assert [r["label"] for r in rates] == ["Reply rate", "Interest rate", "Lead rate"]
    assert [r["rate"] for r in rates] == sorted((r["rate"] for r in rates), reverse=True)


def test_a_rate_with_no_sends_is_a_dash_not_a_zero():
    rates = normalize.headline_rates(_counts(replied=3))
    assert all(r["rate"] is None for r in rates)


def test_unsubscribes_are_shown_as_a_share_of_sends_not_of_replies():
    """They are not one of the mutually exclusive reply categories, so a share
    of replies could read over 100%."""
    c = _counts(sent=1000, replied=10, unsubscribed=40)
    unsub = [o for o in normalize.outcome_breakdown(c) if o["key"] == "unsubscribed"][0]
    assert unsub["of_replies"] is None
    assert unsub["of_sent"] == 4.0
    assert unsub["is_negative"] is True


def test_unsubscribes_are_shown_even_at_zero():
    """Unlike a category a source cannot produce, "none unsubscribed" is a
    result worth seeing."""
    rows = normalize.outcome_breakdown(_counts(sent=500, replied=10))
    assert any(o["key"] == "unsubscribed" and o["value"] == 0 for o in rows)


# ── Unsubscribes from Smartlead ───────────────────────────────────────────────

def test_an_unsubscribe_is_recorded_alongside_the_reply_category():
    """Unsubscribing does not cancel a lead's category — both are true."""
    rows = _outcomes([{
        "lead_email": "a@x.com", "reply_time": "2025-03-04T10:00:00Z",
        "lead_category": "Interested", "is_unsubscribed": True,
    }])
    assert len(rows) == 1
    assert rows[0]["stage"] == normalize.EVENT_POSITIVE_REPLY
    assert rows[0]["unsub_day"] == "2025-03-04"


def test_a_lead_who_unsubscribed_without_replying_still_counts():
    rows = _outcomes([{
        "lead_email": "b@x.com", "sent_time": "2025-03-02T09:00:00Z",
        "is_unsubscribed": True,
    }])
    assert len(rows) == 1
    assert rows[0]["stage"] is None          # not a reply category
    assert rows[0]["unsub_day"] == "2025-03-02"


def test_an_unsubscribe_without_a_reply_is_dated_by_the_last_send():
    """The flag carries no timestamp, so the closest dated fact is used, and an
    unsubscribe cannot precede the mail that prompted it."""
    rows = _outcomes([
        {"lead_email": "c@x.com", "sent_time": "2025-03-01T09:00:00Z",
         "is_unsubscribed": True},
        {"lead_email": "c@x.com", "sent_time": "2025-03-06T09:00:00Z",
         "is_unsubscribed": True},
    ])
    assert rows[0]["unsub_day"] == "2025-03-06"


def test_a_reply_dates_the_unsubscribe_in_preference_to_a_send():
    rows = _outcomes([{
        "lead_email": "d@x.com", "sent_time": "2025-03-01T09:00:00Z",
        "reply_time": "2025-03-03T09:00:00Z", "is_unsubscribed": True,
    }])
    assert rows[0]["unsub_day"] == "2025-03-03"


def test_someone_merely_emailed_is_not_an_outcome():
    assert _outcomes([{"lead_email": "e@x.com", "sent_time": "2025-03-01T09:00:00Z"}]) == []


def test_the_unsubscribe_flag_is_read_in_every_shape_smartlead_sends_it():
    for flag in (True, "true", "t", 1, "1"):
        rows = _outcomes([{"lead_email": "f@x.com", "sent_time": "2025-03-01T09:00:00Z",
                           "is_unsubscribed": flag}])
        assert rows and rows[0]["unsub_day"] == "2025-03-01", flag
    for flag in (False, "false", "f", 0, "", None):
        rows = _outcomes([{"lead_email": "f@x.com", "sent_time": "2025-03-01T09:00:00Z",
                           "is_unsubscribed": flag}])
        assert rows == [], flag


def test_an_outcome_still_refuses_to_be_built_with_nothing_to_date_it():
    import pytest

    with pytest.raises(ValueError):
        normalize.make_outcome(agency_id=AGENCY, campaign_id=CAMPAIGN, lead="x",
                               category="Interested", replied_at=None)


# ── The blank panel ───────────────────────────────────────────────────────────

def test_the_channel_split_is_a_strip_below_the_numbers_not_a_sidebar(client, fake_sb):
    """A sidebar holding two lines of channel totals beside a column holding a
    funnel, three rate cards and five outcome boxes could not be balanced: it
    was either empty or mostly empty."""
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None,
               "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 500, "day": "2025-03-01"},
    ]))

    body = client.get("/api/outbound-pulse/overview").text
    assert "pulse-channel-split" not in body        # one channel, nothing to split


def test_the_total_card_keeps_both_columns_when_both_channels_have_data(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None,
               "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 500, "day": "2025-03-01"},
        {"client_id": CLIENT, "channel": "linkedin", "source_tool": "meet_alfred",
         "event_type": "sent", "events": 200, "day": "2025-03-01"},
    ]))

    body = client.get("/api/outbound-pulse/overview").text
    assert "pulse-channel-split" in body
    # Inside the card and after the numbers, rather than in a column beside them.
    card = body.split('class="pulse-total-card"', 1)[1]
    assert card.index("pulse-rates") < card.index("pulse-channel-split")


def test_the_rate_cards_render_on_the_overview(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None,
               "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 1000, "day": "2025-03-01"},
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "replied", "events": 40, "day": "2025-03-01"},
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "unsubscribed", "events": 10, "day": "2025-03-01"},
    ]))

    body = client.get("/api/outbound-pulse/overview").text
    for label in ("Reply rate", "Interest rate", "Lead rate", "Unsubscribes"):
        assert label in body
    assert "5.0%" in body               # (40 replies + 10 unsubs) / 1000


# ── The unsubscribes migration ────────────────────────────────────────────────

def _unsub_sql():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1]
            / "migrations" / "outbound_pulse_unsubscribes.sql").read_text(encoding="utf-8")


def test_the_unsubscribe_migration_is_one_transaction_and_re_runnable():
    sql = _unsub_sql()
    assert sql.index("BEGIN;") < sql.index("ALTER TABLE")
    assert "ADD COLUMN IF NOT EXISTS unsub_day" in sql
    assert sql.rstrip().endswith("COMMIT;")


def test_an_unsubscribe_is_counted_independently_of_the_reply_stage():
    """A lead marked Interested who then unsubscribes must count once in each,
    which a single `stage` column could not express."""
    view = _unsub_sql().split("CREATE VIEW pulse_funnel_daily AS", 1)[1]
    assert "WHERE o.stage IS NOT NULL" in view
    assert "WHERE o.unsub_day IS NOT NULL" in view


def test_the_manual_view_stops_excluding_opt_out():
    sql = _unsub_sql()
    assert "d.reason IN (\'lead\', \'interested\', \'opt_out\')" in sql
    assert "WHEN \'opt_out\'  THEN \'unsubscribed\'" in sql


def test_the_migration_backfills_from_stored_events():
    """Events already keep is_unsubscribed in raw_payload, so waiting for the
    rolling backfill to revisit every campaign would read zero for ~40 hours."""
    sql = _unsub_sql()
    assert "raw_payload->>\'is_unsubscribed\'" in sql
    assert "ON CONFLICT (agency_id, campaign_id, lead_key) DO NOTHING" in sql


# ── Published client reports ──────────────────────────────────────────────────

def _snapshot(sent=1000, replied=40, unsubscribed=10, leads=8):
    return {
        "version": 1,
        "counts": {"sent": sent, "opened": 0, "replied": replied,
                   "positive_reply": 12, "information_request": 3,
                   "meeting_booked": leads - 3, "lead": 0,
                   "unsubscribed": unsubscribed},
        "by_channel": {"email": {"sent": sent, "replied": replied,
                                 "information_request": 3, "meeting_booked": leads - 3}},
        "by_source": {},
        "trend": {"unit": "day", "buckets": [
            {"day": "2026-09-01", "label": "1 Sep 2026", "sent": 500, "replied": 20},
            {"day": "2026-09-02", "label": "2 Sep 2026", "sent": 500, "replied": 20},
        ]},
        "taken_at": "2026-10-01T09:00:00+00:00",
    }


def _report(rid="r1", start="2026-09-01", end="2026-09-30", status="published",
            body="<p>Volume dipped while we rewrote the opener.</p>", title=""):
    return {
        "id": rid, "client_id": CLIENT, "period_start": start, "period_end": end,
        "title": title, "body": body, "snapshot": _snapshot(),
        "status": status, "published_at": "2026-10-01T09:00:00+00:00",
        "created_by": "Dylan", "created_at": "2026-09-30T09:00:00+00:00",
        "updated_at": "2026-10-01T09:00:00+00:00",
    }


def _portal_routes(fake_sb, reports=None):
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, [{
        "id": "a1", "client_id": CLIENT, "label": "",
        "expires_at": None, "revoked_at": None, "view_count": 0,
    }]))
    fake_sb.route("GET", "clients", lambda call: FakeResponse(
        200, [{"id": CLIENT, "name": "Acme", "color": None,
               "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(
        200, [_report()] if reports is None else reports))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))


# ── The portal is a series of reports, not a live dashboard ──────────────────

def test_the_portal_renders_the_newest_published_report(client, fake_sb):
    _portal_routes(fake_sb)
    body = client.get("/r/valid-token").text
    assert body.count("<html") == 1
    assert "Acme" in body
    assert "September 2026" in body
    assert "1,000" in body                      # from the frozen snapshot


def test_the_portal_no_longer_offers_a_range_picker(client, fake_sb):
    """Letting a client re-slice the data is how they ended up reading numbers
    nobody had looked at before sending."""
    _portal_routes(fake_sb)
    body = client.get("/r/valid-token").text
    for gone in ("Last 30 days", "Last 90 days", "All time", "?range="):
        assert gone not in body


def test_a_published_report_reads_only_its_snapshot(client, fake_sb):
    """It must not re-query: a later sync can still add events inside a closed
    period, and the client would see a different number than they were sent."""
    _portal_routes(fake_sb)
    client.get("/r/valid-token")
    assert fake_sb.calls_to("GET", "pulse_funnel_daily") == []


def test_the_write_up_is_headed_with_the_period(client, fake_sb):
    _portal_routes(fake_sb)
    body = client.get("/r/valid-token").text
    assert "September 2026 Report &amp; Analytics" in body
    assert "From your account manager" not in body


def test_a_titled_report_uses_its_title_as_the_heading(client, fake_sb):
    _portal_routes(fake_sb, reports=[_report(title="Q3 review")])
    body = client.get("/r/valid-token").text
    assert "Q3 review Report &amp; Analytics" in body


def test_the_write_up_renders_as_markup_not_as_escaped_text(client, fake_sb):
    _portal_routes(fake_sb, reports=[_report(
        body="<p>Volume <strong>dipped</strong>.</p><ul><li>Rewrote the opener</li></ul>")])
    body = client.get("/r/valid-token").text
    assert "<strong>dipped</strong>" in body
    assert "<li>Rewrote the opener</li>" in body


# ── Paging through history ───────────────────────────────────────────────────

def _three_reports():
    return [
        _report("sep", "2026-09-01", "2026-09-30"),
        _report("aug", "2026-08-01", "2026-08-31"),
        _report("jul", "2026-07-01", "2026-07-31"),
    ]


def test_the_arrows_appear_once_there_is_a_history(client, fake_sb):
    _portal_routes(fake_sb, reports=_three_reports())
    body = client.get("/r/valid-token").text
    assert "report-nav" in body
    # The NEWEST is the highest number. "1 of 3" on the latest report read as
    # though the client were at the start of their history, not the end.
    assert "3 of 3" in body


def test_a_single_report_gets_no_arrows(client, fake_sb):
    _portal_routes(fake_sb)
    # The markup, not the stylesheet, which always carries the rule.
    assert 'class="report-nav"' not in client.get("/r/valid-token").text


def test_the_back_arrow_goes_to_the_previous_month(client, fake_sb):
    _portal_routes(fake_sb, reports=_three_reports())
    body = client.get("/r/valid-token").text
    assert "?report=aug" in body
    assert "?report=jul" not in body            # two away, not one


def test_the_middle_report_can_go_both_ways(client, fake_sb):
    _portal_routes(fake_sb, reports=_three_reports())
    body = client.get("/r/valid-token?report=aug").text
    assert "2 of 3" in body
    assert "?report=sep" in body                # newer
    assert "?report=jul" in body                # older
    assert "August 2026" in body


def test_the_oldest_report_has_nothing_older(client, fake_sb):
    _portal_routes(fake_sb, reports=_three_reports())
    body = client.get("/r/valid-token?report=jul").text
    assert "1 of 3" in body
    assert "?report=aug" in body
    assert "is-off" in body                     # the back arrow is disabled


def test_an_unknown_report_id_falls_back_to_the_newest(client, fake_sb):
    _portal_routes(fake_sb, reports=_three_reports())
    body = client.get("/r/valid-token?report=does-not-exist").text
    assert "3 of 3" in body
    assert "September 2026" in body


def test_a_report_id_cannot_widen_scope_to_another_client(client, fake_sb):
    """The id only ever selects from this token's own client history, which is
    the only list the route ever builds."""
    from tests.conftest import param_values

    _portal_routes(fake_sb, reports=_three_reports())
    client.get("/r/valid-token?report=someone-elses-report")
    call = fake_sb.calls_to("GET", "pulse_reports")[0]
    assert param_values(call, "client_id") == [f"eq.{CLIENT}"]
    assert param_values(call, "id") == []


def test_the_arrows_keep_the_reader_on_the_path_they_arrived_by(client, fake_sb):
    _portal_routes(fake_sb, reports=_three_reports())
    assert "/r/valid-token?report=aug" in client.get("/r/valid-token").text
    assert "/portal/valid-token?report=aug" in client.get("/portal/valid-token").text


# ── Drafts and empty states ──────────────────────────────────────────────────

def test_only_published_reports_are_served_to_a_client(client, fake_sb):
    from tests.conftest import param_values

    _portal_routes(fake_sb)
    client.get("/r/valid-token")
    call = fake_sb.calls_to("GET", "pulse_reports")[0]
    assert param_values(call, "status") == ["eq.published"]


def test_a_client_with_no_reports_yet_is_told_so(client, fake_sb):
    _portal_routes(fake_sb, reports=[])
    body = client.get("/r/valid-token").text
    assert "first report is on its way" in body
    assert "This link stays the same" in body


def test_a_report_read_failure_never_shows_a_client_a_stack_trace(client, fake_sb):
    _portal_routes(fake_sb)
    fake_sb.route("GET", "pulse_reports",
                  lambda call: FakeResponse(500, {"message": "relation does not exist"}))
    r = client.get("/r/valid-token")
    assert r.status_code == 200
    assert "first report is on its way" in r.text
    assert "relation" not in r.text


def test_the_portal_still_exposes_no_mutating_routes():
    from app.routers import pulse_portal

    for route in pulse_portal.router.routes:
        assert set(route.methods) <= {"GET", "HEAD"}, route.path


def test_a_visit_is_still_recorded(client, fake_sb):
    _portal_routes(fake_sb)
    client.get("/r/valid-token")
    assert fake_sb.calls_to("POST", "pulse_portal_visits")


# ── Building a report, internally ────────────────────────────────────────────

def test_publishing_freezes_the_figures_for_the_period(client, fake_sb):
    """Taken at publish, not at creation: a draft may have been opened weeks
    before it goes out."""
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 700, "day": "2026-09-04"},
    ]))

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert r.status_code == 200
    assert seen["status"] == "published"
    assert seen["published_at"]
    assert seen["snapshot"]["counts"]["sent"] == 700
    assert seen["snapshot"]["trend"]["buckets"]


def test_the_snapshot_covers_the_reports_period_not_today(client, fake_sb):
    from tests.conftest import param_values

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft", start="2026-07-01", end="2026-07-31")]))
    fake_sb.route("PATCH", "pulse_reports", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    days = []
    for call in fake_sb.calls_to("GET", "pulse_funnel_daily"):
        days.extend(param_values(call, "day"))
    assert "gte.2026-07-01" in days
    assert "lte.2026-07-31" in days


def test_a_write_up_is_sanitised_before_it_is_stored(client, fake_sb):
    """Cleaned once, where markup crosses from an author to a reader, rather
    than trusted on the way out to a client's page."""
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [_report()]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1",
                data={"title": "", "body": '<p onclick="x()">hi</p><script>bad()</script>'})
    assert seen["body"] == "<p>hi</p>"


def test_saving_a_draft_does_not_publish_it(client, fake_sb):
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1",
                data={"title": "", "body": "<p>draft</p>"})
    assert "status" not in seen
    assert "snapshot" not in seen


def test_unpublishing_keeps_the_snapshot(client, fake_sb):
    """So re-publishing without editing puts back exactly what was there."""
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [_report()]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/unpublish")
    assert seen["status"] == "draft"
    assert seen["published_at"] is None
    assert "snapshot" not in seen


def test_a_backwards_period_is_swapped_rather_than_rejected(client, fake_sb):
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, []))
    fake_sb.route("POST", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(201, [_report(status="draft")]))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports",
                data={"period_start": "2026-09-30", "period_end": "2026-09-01"})
    assert seen["period_start"] == "2026-09-01"
    assert seen["period_end"] == "2026-09-30"


def test_a_report_without_a_period_is_refused_with_a_reason(client, fake_sb):
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, []))
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports",
                    data={"period_start": "", "period_end": ""})
    assert "Pick a start and an end date" in r.text
    assert not fake_sb.calls_to("POST", "pulse_reports")


def test_a_new_draft_carries_no_snapshot(client, fake_sb):
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, []))
    fake_sb.route("POST", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(201, [_report(status="draft")]))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports",
                data={"period_start": "2026-09-01", "period_end": "2026-09-30"})
    assert seen["status"] == "draft"
    assert "snapshot" not in seen


def test_the_client_page_offers_a_reports_panel(client, fake_sb):
    _detail_routes(fake_sb)
    body = client.get(f"/outbound-pulse/clients/{CLIENT}").text
    assert "Reports" in body
    assert f"/api/outbound-pulse/clients/{CLIENT}/reports" in body
    assert "Notes for this client" not in body


def test_reports_are_agency_scoped_on_every_operation(client, fake_sb):
    from tests.conftest import param_values

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [_report()]))
    fake_sb.route("PATCH", "pulse_reports", lambda call: FakeResponse(200, []))
    fake_sb.route("DELETE", "pulse_reports", lambda call: FakeResponse(204, []))

    client.get(f"/api/outbound-pulse/clients/{CLIENT}/reports")
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1",
                data={"title": "", "body": "<p>x</p>"})
    client.delete(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1")

    for method in ("GET", "PATCH", "DELETE"):
        for call in fake_sb.calls_to(method, "pulse_reports"):
            assert param_values(call, "agency_id"), (method, call["params"])


# ── Report headings ──────────────────────────────────────────────────────────

def test_a_whole_calendar_month_is_named(pulse_reports_mod=None):
    from datetime import date

    from app.utils.pulse.reports import period_label

    assert period_label(date(2026, 9, 1), date(2026, 9, 30)) == "September 2026"
    assert period_label(date(2026, 2, 1), date(2026, 2, 28)) == "February 2026"


def test_a_partial_period_shows_its_dates():
    from datetime import date

    from app.utils.pulse.reports import period_label

    assert period_label(date(2026, 9, 5), date(2026, 9, 20)) == "5 Sep – 20 Sep 2026"
    assert "2025" in period_label(date(2025, 12, 20), date(2026, 1, 10))


def test_a_report_without_a_title_is_named_by_its_period():
    from app.utils.pulse.reports import report_heading

    assert report_heading(_report()) == "September 2026"
    assert report_heading(_report(title="Q3 review")) == "Q3 review"


# ── The write-up sanitiser ───────────────────────────────────────────────────

def test_the_sanitiser_keeps_the_formatting_a_write_up_needs():
    from app.utils.pulse.richtext import sanitize

    html = ("<p>Volume <strong>dipped</strong> and <em>recovered</em>.</p>"
            "<h2>Recommendations</h2><ul><li>Rewrite the opener</li></ul>")
    assert sanitize(html) == html


def test_the_sanitiser_drops_scripts_and_their_contents():
    from app.utils.pulse.richtext import sanitize

    assert sanitize("<script>alert(1)</script><p>after</p>") == "<p>after</p>"
    assert "alert" not in sanitize("<script>alert(1)</script>")


def test_the_sanitiser_drops_event_handlers():
    from app.utils.pulse.richtext import sanitize

    assert sanitize('<p onclick="steal()">click</p>') == "<p>click</p>"


def test_the_sanitiser_refuses_a_javascript_link():
    from app.utils.pulse.richtext import sanitize

    assert "javascript" not in sanitize('<a href="javascript:alert(1)">bad</a>')
    # Whitespace inside the scheme is a URL browsers still run.
    assert "script" not in sanitize('<a href="java\tscript:alert(1)">sneaky</a>').lower()


def test_the_sanitiser_keeps_a_real_link_and_makes_it_safe_to_follow():
    from app.utils.pulse.richtext import sanitize

    out = sanitize('<a href="https://unstuck.agency">us</a>')
    assert 'href="https://unstuck.agency"' in out
    assert 'rel="noopener noreferrer nofollow"' in out


def test_the_sanitiser_keeps_the_text_of_a_tag_it_drops():
    """A pasted <div> of prose should become prose, not disappear."""
    from app.utils.pulse.richtext import sanitize

    assert sanitize("<div><p>kept</p></div>") == "<p>kept</p>"


def test_the_sanitiser_closes_what_an_editor_left_open():
    from app.utils.pulse.richtext import sanitize

    assert sanitize("<p>unbalanced</div></p>") == "<p>unbalanced</p>"
    assert sanitize("<ul><li>one") == "<ul><li>one</li></ul>"


def test_an_emptied_editor_reads_as_empty():
    """A focused-then-cleared editor still emits <p><br></p>."""
    from app.utils.pulse.richtext import to_text

    assert to_text("<p><br></p>") == ""
    assert to_text("<p>real</p>") == "real"


# ── The reports migration ────────────────────────────────────────────────────

def _reports_sql():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1]
            / "migrations" / "outbound_pulse_reports.sql").read_text(encoding="utf-8")


def test_the_reports_migration_is_one_transaction_and_re_runnable():
    sql = _reports_sql()
    assert sql.index("BEGIN;") < sql.index("CREATE TABLE")
    assert "CREATE TABLE IF NOT EXISTS pulse_reports" in sql
    assert sql.rstrip().endswith("COMMIT;")


def test_one_report_per_client_per_period():
    """Publishing September twice corrects September rather than leaving the
    client two of them to choose between."""
    assert "pulse_reports_client_period_unique" in _reports_sql()


def test_a_report_cannot_end_before_it_starts():
    assert "CHECK (period_end >= period_start)" in _reports_sql()


def test_a_report_is_agency_scoped_and_dies_with_its_client():
    sql = _reports_sql()
    assert "agency_id    uuid NOT NULL REFERENCES agencies(id) ON DELETE CASCADE" in sql
    assert "client_id    uuid NOT NULL REFERENCES clients(id)  ON DELETE CASCADE" in sql


def test_the_sanitiser_unwraps_a_list_the_editor_put_inside_a_paragraph():
    """contenteditable emits <p><ul>...</ul></p> routinely. Browsers unwrap it
    on parse, but storing invalid markup that renders on a client's page is not
    something to leave to a parser to rescue."""
    from app.utils.pulse.richtext import sanitize

    assert sanitize("<p>intro</p><p><ul><li>one</li></ul></p>") ==         "<p>intro</p><ul><li>one</li></ul>"
    # And it leaves no empty shell behind.
    assert "<p></p>" not in sanitize("<p><ul><li>one</li></ul></p>")


def test_a_deliberate_blank_line_survives():
    from app.utils.pulse.richtext import sanitize

    assert sanitize("<p>a</p><p><br></p><p>b</p>") == "<p>a</p><p><br></p><p>b</p>"


def test_flattening_does_not_run_sentences_together():
    """The summary on the internal panel is this text; "sequence."
    + "Recommendations" read as one word."""
    from app.utils.pulse.richtext import to_text

    out = to_text("<p>the founder sequence.</p><h2>Recommendations</h2>"
                  "<ul><li>Hold the opener</li><li>Add a follow-up</li></ul>")
    assert out == "the founder sequence. Recommendations Hold the opener Add a follow-up"


def test_flattening_does_not_split_a_word_at_an_inline_tag():
    from app.utils.pulse.richtext import to_text

    assert to_text("<p>a<strong>b</strong>c</p>") == "abc"


# ── Copy A/B ──────────────────────────────────────────────────────────────────

def _sequence_payload(variants):
    return {"ok": True, "data": {"campaign_id": 1, "sequences": [{
        "seq_number": 1, "subject_line": "Quick question",
        "variants": [
            {"id": i, "variant_label": v[0], "is_baseline": i == 0,
             "stats": {"sent_count": v[1], "reply_count": v[2],
                       "positive_reply_count": v[3], "unsubscribed_count": v[4]}}
            for i, v in enumerate(variants)
        ],
    }]}}


def test_smartlead_variants_come_from_smartleads_own_endpoint(monkeypatch):
    """Smartlead computes this, so nothing re-derives it from our events, which
    would have needed a variant column on two tables and a rollup."""
    from datetime import date

    from app.utils.pulse import abtest, smartlead

    seen = {}

    def fake(path, params=None):
        seen["path"] = path
        seen["params"] = params or {}
        return _sequence_payload([("A", 600, 14, 6, 2), ("B", 620, 9, 2, 5)])

    monkeypatch.setattr(smartlead, "request_json", fake)
    steps = abtest.smartlead_variants("8801", date(2026, 9, 1), date(2026, 9, 30))

    assert seen["path"] == "/campaigns/8801/sequence-analytics"
    assert seen["params"]["start_date"].startswith("2026-09-01")
    assert seen["params"]["end_date"].startswith("2026-09-30")
    assert steps[0]["winner"] == "A"
    assert steps[0]["subject"] == "Quick question"


def test_a_variant_that_never_went_out_is_not_a_test(monkeypatch):
    from datetime import date

    from app.utils.pulse import abtest, smartlead

    monkeypatch.setattr(smartlead, "request_json",
                        lambda p, q=None: _sequence_payload(
                            [("A", 600, 14, 6, 2), ("B", 0, 0, 0, 0)]))
    steps = abtest.smartlead_variants("1", date(2026, 9, 1), date(2026, 9, 30))
    assert [v["variant"] for v in steps[0]["variants"]] == ["A"]
    assert steps[0]["winner"] is None
    assert "nothing to compare" in steps[0]["reason"]


def test_too_little_volume_is_not_called_a_winner():
    """A variant with barely any volume tops the table on noise alone."""
    from app.utils.pulse.abtest import MIN_SENDS_FOR_A_WINNER, _decide

    out = _decide([
        {"variant": "A", "is_baseline": True, "sent": 10, "reply": 2,
         "positive": 2, "unsubscribe": 0},
        {"variant": "B", "is_baseline": False, "sent": 9, "reply": 0,
         "positive": 0, "unsubscribe": 0},
    ])
    assert out["winner"] is None
    assert "Too little volume" in out["reason"]
    assert MIN_SENDS_FOR_A_WINNER > 10


def test_a_level_result_is_reported_as_level_not_as_a_coin_flip():
    from app.utils.pulse.abtest import _decide

    out = _decide([
        {"variant": "A", "is_baseline": True, "sent": 500, "reply": 10,
         "positive": 4, "unsubscribe": 1},
        {"variant": "B", "is_baseline": False, "sent": 500, "reply": 10,
         "positive": 4, "unsubscribe": 1},
    ])
    assert out["winner"] is None
    assert "level" in out["reason"]


def test_no_positive_responses_means_no_winner():
    from app.utils.pulse.abtest import _decide

    out = _decide([
        {"variant": "A", "is_baseline": True, "sent": 500, "reply": 6,
         "positive": 0, "unsubscribe": 2},
        {"variant": "B", "is_baseline": False, "sent": 500, "reply": 2,
         "positive": 0, "unsubscribe": 1},
    ])
    assert out["winner"] is None
    assert "No positive responses" in out["reason"]


def test_a_clear_result_names_a_winner_and_gives_rates():
    from app.utils.pulse.abtest import _decide

    out = _decide([
        {"variant": "A", "is_baseline": True, "sent": 1000, "reply": 40,
         "positive": 12, "unsubscribe": 4},
        {"variant": "B", "is_baseline": False, "sent": 1000, "reply": 20,
         "positive": 3, "unsubscribe": 9},
    ])
    assert out["winner"] == "A"
    assert out["reason"] == ""
    assert out["variants"][0]["positive_rate"] == 1.2
    assert out["variants"][0]["reply_rate"] == 4.0


def test_manual_variants_combine_across_a_clients_sheets(monkeypatch):
    from app.utils import google_sheets
    from app.utils.pulse import abtest

    sheets = {
        "s1": [{"variant": "S1/B1", "total": 300, "lead": 6, "interested": 2,
                "reply": 10, "unsubscribe": 3, "positive": 8},
               {"variant": "S2/B1", "total": 300, "lead": 1, "interested": 1,
                "reply": 4, "unsubscribe": 7, "positive": 2}],
        "s2": [{"variant": "S1/B1", "total": 200, "lead": 4, "interested": 1,
                "reply": 7, "unsubscribe": 1, "positive": 5}],
    }
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: sheets[sid])

    out = abtest.manual_variants(["s1", "s2"])
    by = {v["variant"]: v for v in out["variants"]}
    assert by["S1/B1"]["sent"] == 500          # combined across both sheets
    assert by["S1/B1"]["positive"] == 13
    assert out["winner"] == "S1/B1"
    assert out["sheets_read"] == 2


def test_an_unreadable_sheet_is_skipped_not_fatal(monkeypatch):
    """One revoked share should not hide every other campaign's result."""
    from app.utils import google_sheets
    from app.utils.pulse import abtest

    def reader(sid):
        if sid == "bad":
            raise RuntimeError("permission denied")
        return [{"variant": "S1/B1", "total": 400, "lead": 8, "interested": 2,
                 "reply": 12, "unsubscribe": 2, "positive": 10}]

    monkeypatch.setattr(google_sheets, "read_ab_stats", reader)
    out = abtest.manual_variants(["bad", "good"])
    assert out["sheets_read"] == 1
    assert out["sheets_skipped"] == 1
    assert out["variants"][0]["sent"] == 400


def test_the_two_sources_state_what_they_decided_on(monkeypatch):
    """They measure different things, so a combined winner would mean nothing."""
    from datetime import date

    from app.utils import google_sheets
    from app.utils.pulse import abtest, smartlead

    monkeypatch.setattr(smartlead, "request_json",
                        lambda p, q=None: _sequence_payload(
                            [("A", 600, 14, 6, 2), ("B", 620, 9, 2, 5)]))
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: [])

    steps = abtest.smartlead_variants("1", date(2026, 9, 1), date(2026, 9, 30))
    assert "Smartlead" in steps[0]["basis"]
    assert "Lead and Interested" in abtest.manual_variants([])["basis"]


def test_the_ab_panel_says_manual_figures_are_not_windowed(client, fake_sb, monkeypatch):
    """Comparing a month of Smartlead against the lifetime of a manual campaign
    would invite a false conclusion."""
    from app.utils import google_sheets
    from app.utils.pulse import smartlead

    _detail_routes(fake_sb)
    monkeypatch.setattr(smartlead, "is_configured", lambda: False)
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: [])

    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "Manual \u00b7 all time" in body or "Manual · all time" in body
    assert "whole life" in body


def test_a_missing_smartlead_key_is_explained_not_crashed(client, fake_sb, monkeypatch):
    from app.utils import google_sheets
    from app.utils.pulse import smartlead

    _detail_routes(fake_sb)
    monkeypatch.setattr(smartlead, "is_configured", lambda: False)
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: [])

    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab")
    assert r.status_code == 200
    assert "No Smartlead API key" in r.text


def test_one_failing_campaign_does_not_hide_the_others(client, fake_sb, monkeypatch):
    from app.utils import google_sheets
    from app.utils.pulse import abtest, smartlead
    from app.routers import outbound_pulse as router_mod

    _detail_routes(fake_sb)
    monkeypatch.setattr(smartlead, "is_configured", lambda: True)
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: [])

    def boom(external, start, end):
        raise RuntimeError("429")

    monkeypatch.setattr(router_mod.abtest, "smartlead_variants", boom)
    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab")
    assert r.status_code == 200
    assert "could not be read from Smartlead" in r.text


def test_the_client_page_offers_the_ab_panel(client, fake_sb):
    _detail_routes(fake_sb)
    body = client.get(f"/outbound-pulse/clients/{CLIENT}").text
    assert "Winning copy" in body
    assert f"/api/outbound-pulse/clients/{CLIENT}/ab" in body


# ── The write-up actually reaches the server ─────────────────────────────────

def test_publishing_also_saves_the_write_up(client, fake_sb):
    """Publish used to only snapshot, so whether the write-up survived depended
    on a separate save request with no ordering against it."""
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish",
                data={"title": "Q3 review", "body": "<p>Saved on publish.</p>",
                      "from_editor": "1"})
    assert seen["body"] == "<p>Saved on publish.</p>"
    assert seen["title"] == "Q3 review"
    assert seen["status"] == "published"
    assert "snapshot" in seen


def test_a_published_write_up_is_sanitised_too(client, fake_sb):
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish",
                data={"body": "<p onclick='x()'>hi</p><script>bad()</script>",
                      "from_editor": "1"})
    assert seen["body"] == "<p>hi</p>"


def test_the_publish_button_does_not_also_submit_the_form(client, fake_sb):
    """As type=submit it fired two competing requests — one saving, one
    snapshotting, with no ordering between them."""
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [_report()]))

    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/reports").text
    start = body.rindex("<button", 0, body.index("/publish"))
    publish = body[start:body.index(">", body.index("/publish"))]
    assert 'type="button"' in publish
    assert 'type="submit"' not in publish


def test_the_editor_sends_its_current_content_whatever_htmx_collected():
    """HTMX reads a form's values BEFORE htmx:configRequest fires, so syncing
    the hidden field in that handler was one step too late and posted an empty
    body. The handler now sets the parameter directly."""
    import pathlib

    editor = (pathlib.Path(__file__).resolve().parents[1] / "app" / "templates"
              / "partials" / "pulse_editor.html").read_text(encoding="utf-8")
    assert "e.detail.parameters['body'] = ed.area.innerHTML" in editor
    # And mirrored continuously, so a no-JS fallback textarea is never stale.
    assert "addEventListener('input', sync)" in editor


def test_the_portal_asks_for_reports_newest_period_first(client, fake_sb):
    """The arrows page in exactly this ordering, so it has to come from the
    query rather than from insertion order."""
    from tests.conftest import param_values

    _portal_routes(fake_sb)
    client.get("/r/valid-token")
    call = fake_sb.calls_to("GET", "pulse_reports")[0]
    # Newest period first, and among reports covering the same period the one
    # created last: that is a correction, not an older one being re-sent.
    assert param_values(call, "order") == ["period_end.desc,created_at.desc"]


# ── Publishing a draft from the row ──────────────────────────────────────────

def test_a_draft_can_be_published_without_opening_the_editor(client, fake_sb):
    """Publish used to live only inside the write-up form, so a draft you had
    already written could not be sent without clicking Edit first, and nothing
    on the row said so."""
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))

    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/reports").text
    row = body[:body.index("data-report-summary")]
    assert "/publish" in row


def test_publishing_without_a_body_does_not_wipe_the_saved_one(client, fake_sb):
    """The row's Publish button sends no editor content. Treating that as an
    empty write-up would publish the report with its text erased."""
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft", body="<p>Already written.</p>")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert seen["status"] == "published"
    assert "body" not in seen              # left exactly as it was
    assert "title" not in seen


def test_publishing_with_a_body_still_saves_it(client, fake_sb):
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish",
                data={"title": "Q3", "body": "<p>New text.</p>", "from_editor": "1"})
    assert seen["body"] == "<p>New text.</p>"
    assert seen["title"] == "Q3"


def test_a_write_up_can_still_be_deliberately_cleared(client, fake_sb):
    """Sending an empty body is different from sending none."""
    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft", body="<p>Old.</p>")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish",
                data={"body": "", "from_editor": "1"})
    assert seen["body"] == ""


def test_a_published_report_offers_unpublish_not_publish(client, fake_sb):
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [_report()]))

    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/reports").text
    row = body[:body.index("data-report-summary")]
    assert "/unpublish" in row
    assert "/publish\"" not in row.replace("/unpublish", "")


def test_a_draft_says_the_client_cannot_see_it(client, fake_sb):
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/reports").text
    assert "not on the client" in body


def test_a_published_report_carries_no_draft_notice(client, fake_sb):
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [_report()]))
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/reports").text
    assert "not on the client" not in body


def test_a_snapshot_can_actually_be_stored_as_json():
    """It is written to a jsonb column through requests, which serialises it.
    Trend buckets carried date objects, so publishing a period that had data
    raised inside the HTTP client, was swallowed, and looked like Publish
    simply doing nothing."""
    import json
    from datetime import date, timedelta

    from app.utils.pulse.normalize import bucket_timeseries

    d0 = date(2026, 1, 1)
    for span in (5, 200, 900, 3000):
        rows = [{"day": (d0 + timedelta(days=i)).isoformat(), "sent": 10}
                for i in range(span)]
        json.dumps(bucket_timeseries(rows))      # must not raise


def test_trend_bucket_dates_are_strings():
    from datetime import date, timedelta

    from app.utils.pulse.normalize import bucket_timeseries

    rows = [{"day": (date(2026, 1, 1) + timedelta(days=i)).isoformat(), "sent": 1}
            for i in range(200)]
    bucket = bucket_timeseries(rows)["buckets"][0]
    assert isinstance(bucket["start"], str)
    assert isinstance(bucket["end"], str)
    assert bucket["start"] == bucket["day"]


def test_publishing_a_period_with_data_stores_a_json_snapshot(client, fake_sb):
    """The case that broke: an empty period serialised fine, a period with
    activity did not."""
    import json

    seen = {}

    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 700, "day": "2026-09-04"},
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "replied", "events": 20, "day": "2026-09-11"},
    ]))

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert seen["snapshot"]["trend"]["buckets"]
    json.dumps(seen["snapshot"])                 # must not raise


def test_a_failed_publish_says_so_instead_of_looking_like_success(client, fake_sb):
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: FakeResponse(500, {"message": "nope"}))

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert "Could not publish" in r.text


def test_a_failed_save_says_so_too(client, fake_sb):
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [_report()]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: FakeResponse(500, {"message": "nope"}))

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1",
                    data={"title": "", "body": "<p>x</p>"})
    assert "Could not save" in r.text


# ── The date range controls ──────────────────────────────────────────────────

def _styles():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1] / "app" / "templates"
            / "pulse_styles.html").read_text(encoding="utf-8")


def test_the_custom_range_can_actually_be_hidden():
    """An author-origin `display` beats the browser's [hidden] rule, so the
    custom-range inputs were never hidden. People typed dates into boxes that
    looked live while the dropdown still said "Last 30 days"."""
    css = _styles()
    assert ".date-range[hidden] { display: none; }" in css
    assert css.index(".date-range {") < css.index(".date-range[hidden]")


def test_the_overview_date_inputs_are_disabled_outside_custom_mode(client, fake_sb):
    """HTMX omits disabled inputs but sends hidden ones, and any date beats the
    preset server-side — so a leftover date used to override every later
    preset choice and the Range control went dead."""
    _detail_routes(fake_sb)
    body = client.get("/outbound-pulse").text
    span = body[body.index('id="custom-range"'):body.index("apply-range")]
    assert span.count("disabled") == 2


def test_the_client_page_date_inputs_are_disabled_for_a_preset(client, fake_sb):
    _detail_routes(fake_sb)
    body = client.get(f"/outbound-pulse/clients/{CLIENT}?range=30d").text
    span = body[body.index('id="custom-range"'):body.index("apply-range")]
    assert span.count("disabled") == 2


def test_the_client_page_date_inputs_are_live_for_a_custom_range(client, fake_sb):
    _detail_routes(fake_sb)
    body = client.get(
        f"/outbound-pulse/clients/{CLIENT}?date_from=2026-09-01&date_to=2026-09-30").text
    span = body[body.index('id="custom-range"'):body.index("apply-range")]
    assert "disabled" not in span
    assert "2026-09-01" in span and "2026-09-30" in span


def test_the_overview_shows_the_range_it_actually_applied(client, fake_sb):
    """The template hardcoded 30d as selected, so after a custom range the
    dropdown and the figures on screen disagreed."""
    _detail_routes(fake_sb)

    body = client.get("/outbound-pulse?range=90d").text
    select = body[body.index('id="range-select"'):body.index("</select>")]
    assert '<option value="90d" selected>' in select.replace(" >", ">")

    body = client.get("/outbound-pulse?date_from=2026-09-01&date_to=2026-09-30").text
    select = body[body.index('id="range-select"'):body.index("</select>")]
    assert 'value="custom" selected' in select


def test_a_preset_still_resolves_to_that_preset_server_side(client, fake_sb):
    """The regression in one line: an empty date must never beat the preset."""
    from app.routers.outbound_pulse import _resolve_range

    assert _resolve_range("90d", "", "")["preset"] == "90d"
    assert _resolve_range("90d", "2026-09-01", "")["preset"] == "custom"


def test_the_ab_panel_asks_for_the_same_range_as_the_page(client, fake_sb):
    """It shares the page's date inputs through hx-include. While those were
    enabled for presets too, A/B always resolved to a custom range and its
    header printed raw dates instead of "last 30 days"."""
    _detail_routes(fake_sb)
    body = client.get(f"/outbound-pulse/clients/{CLIENT}?range=30d").text
    assert 'hx-include="#range-select, #date-from, #date-to"' in body
    # Disabled for a preset, so only `range` reaches the A/B endpoint.
    span = body[body.index('id="custom-range"'):body.index("apply-range")]
    assert span.count("disabled") == 2


# ── Naming and the home screen ───────────────────────────────────────────────

def _index_html():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1] / "app" / "templates"
            / "index.html").read_text(encoding="utf-8")


def test_the_tool_is_named_client_performance_and_reports(client, fake_sb):
    _detail_routes(fake_sb)
    for path in ("/outbound-pulse", f"/outbound-pulse/clients/{CLIENT}"):
        body = client.get(path).text
        assert "Client Performance &amp; Reports" in body
        assert "Outbound Pulse" not in body


def test_the_urls_and_the_tool_key_are_deliberately_unchanged():
    """Renaming the tool key would silently drop every per-person tools_add /
    tools_remove grant in app_users, and renaming the SOP key orphans the SOP
    content stored against it. Only the labels moved."""
    from app import auth

    assert "outbound_pulse" in auth.TOOLS
    assert auth.tool_for_path("/outbound-pulse") == "outbound_pulse"
    assert auth.tool_for_path("/api/outbound-pulse/overview") == "outbound_pulse"


def test_the_client_links_are_untouched_by_the_rename(client, fake_sb):
    _portal_routes(fake_sb)
    assert client.get("/r/valid-token").status_code == 200
    assert client.get("/portal/valid-token").status_code == 200


def test_the_home_screen_has_a_reporting_and_analytics_category():
    html = _index_html()
    assert "screen-reporting" in html
    assert "Reporting &amp; Analytics" in html
    assert "navigateTo('screen-reporting')" in html


def test_the_tool_moved_out_of_campaign_management():
    """It lives in the new category, and in only one place."""
    html = _index_html()
    assert html.count('href="/outbound-pulse"') == 1
    reporting = html[html.index('id="screen-reporting"'):]
    reporting = reporting[:reporting.index('id="screen-operations"')]
    assert 'href="/outbound-pulse"' in reporting


def test_the_new_category_counts_its_tools():
    """The count is written by JS from a map of screen ids. Leave the new one
    out and its card shows a stale hardcoded number forever."""
    html = _index_html()
    assert "'count-reporting': 'screen-reporting'" in html
    assert 'id="count-reporting"' in html


def test_the_new_card_is_still_permission_gated():
    html = _index_html()
    reporting = html[html.index('id="screen-reporting"'):]
    reporting = reporting[:reporting.index('id="screen-operations"')]
    assert "{% if 'outbound_pulse' in tools %}" in reporting


# ── Hiding clients from the overview ─────────────────────────────────────────

BD = "55555555-5555-5555-5555-555555555555"


def _two_clients(fake_sb, prefs=None):
    """Acme plus the business-development client, both with activity."""
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, [
        {"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True},
        {"id": BD, "name": "Unstuck - Business Development", "color": None,
         "emoji": None, "active": True},
    ]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 1000, "day": "2026-09-01"},
        {"client_id": BD, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 400, "day": "2026-09-01"},
    ]))
    # prefs: None = no row (never chosen), else a list for the stored row.
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(
        200, [] if prefs is None else [{"mode": "exclude", "client_ids": prefs}]))


# ── The three states, which is the whole point of the schema ────────────────

def test_a_user_who_has_never_chosen_gets_the_default_exclusion(client, fake_sb):
    _two_clients(fake_sb, prefs=None)
    body = client.get("/api/outbound-pulse/overview").text
    assert "Acme" in body
    assert "Unstuck - Business Development" not in body.split("Hiding")[0]


def test_a_user_who_cleared_their_exclusions_sees_every_client(client, fake_sb):
    """The test that fails if "never chosen" and "chose nothing" are ever
    collapsed into one state. Saving an empty list has to stick."""
    _two_clients(fake_sb, prefs=[])
    body = client.get("/api/outbound-pulse/overview").text
    assert "Acme" in body
    assert "Unstuck - Business Development" in body
    assert "Hiding" not in body


def test_a_saved_exclusion_hides_only_what_was_chosen(client, fake_sb):
    _two_clients(fake_sb, prefs=[CLIENT])
    body = client.get("/api/outbound-pulse/overview").text
    assert "Unstuck - Business Development" in body
    assert "Hiding Acme" in body


# ── Totals and the channel strip have to agree with the grid ────────────────

def test_a_hidden_client_leaves_the_totals_not_just_the_grid(client, fake_sb):
    _two_clients(fake_sb, prefs=None)        # BD hidden by default
    body = client.get("/api/outbound-pulse/overview").text
    assert "1,000" in body                   # Acme alone
    assert "1,400" not in body               # not Acme + BD


def test_the_channel_strip_matches_the_totals_above_it(client, fake_sb):
    """The strip used to come from a separate, unfiltered query, so with
    anything hidden it stopped adding up to the totals printed above it."""
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, [
        {"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True},
        {"id": BD, "name": "Unstuck - Business Development", "color": None,
         "emoji": None, "active": True},
    ]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 1000, "day": "2026-09-01"},
        {"client_id": CLIENT, "channel": "linkedin", "source_tool": "meet_alfred",
         "event_type": "sent", "events": 200, "day": "2026-09-01"},
        {"client_id": BD, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 400, "day": "2026-09-01"},
    ]))

    body = client.get("/api/outbound-pulse/overview").text
    assert "1,200" in body        # the total: Acme's email + LinkedIn
    assert "1,600" not in body    # BD's 400 is in neither the total nor the strip
    strip = body[body.index("By channel"):]
    assert "400" not in strip


def test_the_overview_names_the_clients_it_is_hiding(client, fake_sb):
    """A total with something taken out of it is otherwise indistinguishable
    from the real figure."""
    _two_clients(fake_sb, prefs=None)
    body = client.get("/api/outbound-pulse/overview").text
    assert "Hiding Unstuck - Business Development" in body


# ── Scope: the things an exclusion must NOT touch ───────────────────────────

def test_a_hidden_client_still_has_a_working_page(client, fake_sb):
    """Hiding is about your list, not about access."""
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(
        200, [{"mode": "exclude", "client_ids": [CLIENT]}]))
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert r.status_code == 200
    assert "Acme" in r.text


def test_the_client_portal_ignores_overview_exclusions(client, fake_sb):
    _portal_routes(fake_sb)
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(
        200, [{"mode": "exclude", "client_ids": [CLIENT]}]))
    r = client.get("/r/valid-token")
    assert r.status_code == 200
    assert "September 2026" in r.text


def test_publishing_a_report_never_reads_view_preferences(client, fake_sb):
    """The strongest statement that a personal view filter cannot reach a
    client's frozen figures."""
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert fake_sb.calls_to("GET", "pulse_user_prefs") == []


def test_exclusions_do_not_narrow_the_campaign_mapping_picker(client, fake_sb):
    """You must still be able to map a campaign to a hidden client."""
    _two_clients(fake_sb, prefs=[BD])
    body = client.get("/api/outbound-pulse/campaigns").text
    assert "Unstuck - Business Development" in body


# ── Saving, resetting, scoping ──────────────────────────────────────────────

def test_saving_upserts_on_the_user_and_the_agency(client, fake_sb):
    seen = {}
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, [
        {"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(200, []))
    fake_sb.route("POST", "pulse_user_prefs",
                  lambda call: (seen.update(call), FakeResponse(201, []))[1])

    r = client.post("/api/outbound-pulse/exclusions", data={"excluded": [CLIENT]})
    assert r.status_code == 200
    assert seen["params"]["on_conflict"] == "agency_id,user_email"
    assert "merge-duplicates" in seen["headers"].get("Prefer", "")
    assert seen["json"]["client_ids"] == [CLIENT]
    assert seen["json"]["user_email"] == seen["json"]["user_email"].lower()


def test_saving_refreshes_the_overview(client, fake_sb):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(200, []))
    fake_sb.route("POST", "pulse_user_prefs", lambda call: FakeResponse(201, []))
    r = client.post("/api/outbound-pulse/exclusions", data={})
    assert r.headers.get("HX-Trigger") == "pulseRefresh"


def test_saving_nothing_writes_an_empty_list_rather_than_doing_nothing(client, fake_sb):
    """"Show me everything" is a real answer and has to be storable."""
    seen = {}
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(200, []))
    fake_sb.route("POST", "pulse_user_prefs",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(201, []))[1])
    client.post("/api/outbound-pulse/exclusions", data={})
    assert seen["client_ids"] == []


def test_resetting_deletes_the_row_rather_than_saving_an_empty_list(client, fake_sb):
    """The only way back to the house default."""
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(200, []))
    fake_sb.route("DELETE", "pulse_user_prefs", lambda call: FakeResponse(204, []))

    r = client.delete("/api/outbound-pulse/exclusions")
    assert r.headers.get("HX-Trigger") == "pulseRefresh"
    assert fake_sb.calls_to("DELETE", "pulse_user_prefs")
    assert not fake_sb.calls_to("POST", "pulse_user_prefs")


def test_an_unknown_client_id_is_dropped_rather_than_written(client, fake_sb):
    """A uuid[] column rejects a non-uuid, which would fail the whole save."""
    seen = {}
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, [
        {"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True}]))
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(200, []))
    fake_sb.route("POST", "pulse_user_prefs",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(201, []))[1])

    client.post("/api/outbound-pulse/exclusions",
                data={"excluded": [CLIENT, "not-a-uuid", "../../etc"]})
    assert seen["client_ids"] == [CLIENT]


def test_preferences_are_agency_scoped(client, fake_sb):
    from tests.conftest import param_values

    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(200, []))
    fake_sb.route("DELETE", "pulse_user_prefs", lambda call: FakeResponse(204, []))

    client.get("/api/outbound-pulse/exclusions")
    client.delete("/api/outbound-pulse/exclusions")
    for method in ("GET", "DELETE"):
        for call in fake_sb.calls_to(method, "pulse_user_prefs"):
            assert param_values(call, "agency_id")


# ── Resilience and the default ──────────────────────────────────────────────

def test_a_missing_preferences_table_does_not_break_the_overview(client, fake_sb):
    """Losing the overview because a view preference could not be read would
    be a far worse failure than losing the preference."""
    _two_clients(fake_sb, prefs=None)
    fake_sb.route("GET", "pulse_user_prefs",
                  lambda call: FakeResponse(500, {"message": "does not exist"}))
    r = client.get("/api/outbound-pulse/overview")
    assert r.status_code == 200
    assert "Acme" in r.text


def test_the_default_client_is_matched_by_name_not_by_id():
    """The id differs between Supabase projects; this has to work on a fresh
    database with nothing seeded."""
    from app.routers.outbound_pulse import _default_excluded_ids

    found = _default_excluded_ids([
        {"id": "x", "name": "  unstuck - business development  "},
        {"id": "y", "name": "Acme"},
    ])
    assert found == {"x"}


def test_the_default_exclusion_is_env_overridable(monkeypatch):
    import importlib

    from app.routers import outbound_pulse as mod

    monkeypatch.setenv("PULSE_DEFAULT_EXCLUDED_CLIENTS", "Acme, Other Co")
    reloaded = importlib.reload(mod)
    try:
        assert reloaded.DEFAULT_EXCLUDED_CLIENT_NAMES == ("acme", "other co")
    finally:
        monkeypatch.delenv("PULSE_DEFAULT_EXCLUDED_CLIENTS", raising=False)
        importlib.reload(mod)


# ── The migration ───────────────────────────────────────────────────────────

def _prefs_sql():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1]
            / "migrations" / "outbound_pulse_user_prefs.sql").read_text(encoding="utf-8")


def test_the_prefs_migration_is_one_transaction_and_re_runnable():
    sql = _prefs_sql()
    assert sql.index("BEGIN;") < sql.index("CREATE TABLE")
    assert "CREATE TABLE IF NOT EXISTS pulse_user_prefs" in sql
    assert sql.rstrip().endswith("COMMIT;")


def test_row_presence_is_what_distinguishes_never_chosen_from_chose_nothing():
    sql = _prefs_sql()
    assert "PRIMARY KEY (agency_id, user_email)" in sql
    assert "excluded_client_ids uuid[] NOT NULL DEFAULT '{}'" in sql


def test_the_prefs_table_has_no_foreign_key_to_app_users():
    """With AUTH_DISABLED the user is dev@local, who has no app_users row — a
    key would break saving in exactly the mode that exists as a rollback.

    Checks the DDL with comments stripped; the comment block explains the
    decision and naturally mentions the table.
    """
    ddl = chr(10).join(line for line in _prefs_sql().splitlines()
                       if not line.lstrip().startswith("--"))
    assert "app_users" not in ddl
    assert "REFERENCES agencies(id)" in ddl      # the one key it does have


def test_the_email_key_is_forced_lowercase_by_the_database():
    assert "CHECK (user_email = lower(user_email)" in _prefs_sql()


# ── Hand-entered figures ─────────────────────────────────────────────────────

# ── The naming trap ─────────────────────────────────────────────────────────

def test_hand_entry_is_a_different_source_from_the_dnc_tool():
    """One is automatic and already in this database; the other is typed in."""
    assert normalize.SOURCE_MANUAL != normalize.SOURCE_HAND_ENTRY
    labels = normalize.SOURCE_LABELS
    assert labels[normalize.SOURCE_MANUAL] != labels[normalize.SOURCE_HAND_ENTRY]


def test_no_source_is_labelled_just_manual():
    """That word is what made the two confusable. Each is named by provenance."""
    assert "Manual" not in normalize.SOURCE_LABELS.values()


def test_the_dnc_source_key_is_unchanged():
    """It is written into the by_source block of every snapshot on disk."""
    assert normalize.SOURCE_MANUAL == "manual"


def test_the_channel_picker_offers_hand_entered_figures(client, fake_sb):
    _detail_routes(fake_sb)
    for path in ("/outbound-pulse", f"/outbound-pulse/clients/{CLIENT}"):
        assert 'value="hand_entry"' in client.get(path).text


def test_picking_it_filters_by_source_not_channel(client, fake_sb):
    from tests.conftest import param_values

    _detail_routes(fake_sb)
    client.get(f"/outbound-pulse/clients/{CLIENT}?channel=hand_entry")
    calls = fake_sb.calls_to("GET", "pulse_funnel_daily")
    assert calls
    for call in calls:
        assert param_values(call, "source_tool") == ["eq.hand_entry"]


def test_the_four_scopes_still_partition_the_data():
    from app.routers.outbound_pulse import REPORT_SCOPES

    assert len(set(REPORT_SCOPES.values())) == len(REPORT_SCOPES) == 4


def test_the_source_tab_says_a_human_typed_it(client, fake_sb):
    """A synced number and a typed one must never look identical."""
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "linkedin", "source_tool": "hand_entry",
         "event_type": "sent", "events": 900, "day": "2026-09-01"},
    ]))
    body = client.get(f"/outbound-pulse/clients/{CLIENT}").text
    assert "Typed in by an account manager" in body
    assert "recorded against the first of" in body


# ── Bounces ─────────────────────────────────────────────────────────────────

def test_bounces_are_a_counted_stage():
    assert normalize.EVENT_BOUNCED in normalize.FUNNEL_STAGES
    assert normalize.EVENT_BOUNCED in normalize.empty_funnel()


def test_a_bounce_is_not_a_response():
    """Putting it in the reply rate would inflate every one in the tool,
    including the ones already frozen into published reports."""
    assert normalize.EVENT_BOUNCED not in normalize.RESPONSE_STAGES
    base = _counts(sent=1000, replied=40)
    withb = _counts(sent=1000, replied=40, bounced=90)
    assert normalize.reply_rate(base) == normalize.reply_rate(withb)


def test_a_bounce_is_not_a_lead_and_not_an_exclusive_outcome():
    assert normalize.EVENT_BOUNCED not in normalize.LEAD_STAGES
    assert normalize.EVENT_BOUNCED not in normalize.OUTCOME_STAGES
    c = _counts(sent=1000, bounced=50, meeting_booked=3)
    assert normalize.lead_count(c) == 3


def test_bounces_are_hidden_at_zero_but_unsubscribes_are_not():
    """Only one source can report a bounce, so a permanent zero would read as
    broken rather than as good news."""
    keys = [o["key"] for o in normalize.outcome_breakdown(_counts(sent=500, replied=9))]
    assert "unsubscribed" in keys
    assert "bounced" not in keys
    keys = [o["key"] for o in normalize.outcome_breakdown(
        _counts(sent=500, replied=9, bounced=4))]
    assert "bounced" in keys


def test_bounce_rate_is_a_share_of_sends():
    assert normalize.bounce_rate(_counts(sent=1000, bounced=25)) == 2.5


def test_a_snapshot_written_before_bounces_existed_still_renders(client, fake_sb):
    _portal_routes(fake_sb)       # its snapshot has no "bounced" key
    assert client.get("/r/valid-token").status_code == 200


# ── Dating ──────────────────────────────────────────────────────────────────

def test_any_day_in_a_month_resolves_to_the_first():
    from datetime import date

    from app.utils.pulse import hand_entry

    for value in ("2026-09", "2026-09-01", "2026-09-17", "2026-09-30"):
        assert hand_entry.parse_month(value) == date(2026, 9, 1)
    assert hand_entry.parse_month("") is None
    assert hand_entry.parse_month("nonsense") is None


def test_month_bounds_cover_the_whole_month():
    from datetime import date

    from app.utils.pulse import hand_entry

    assert hand_entry.month_bounds(date(2026, 9, 1)) == (date(2026, 9, 1), date(2026, 9, 30))
    assert hand_entry.month_bounds(date(2026, 12, 1)) == (date(2026, 12, 1), date(2026, 12, 31))
    assert hand_entry.month_bounds(date(2024, 2, 1)) == (date(2024, 2, 1), date(2024, 2, 29))


def test_a_monthly_figure_is_never_spread_across_days():
    """Dividing a month by 30 would invent a daily series nobody measured."""
    sql = _hand_sql()
    view = sql.split("CREATE OR REPLACE VIEW pulse_hand_entry_daily", 1)[1]
    assert "generate_series" not in view
    assert "/" not in view.split("CROSS JOIN LATERAL")[1].split(")")[0]
    assert "e.period_month" in view


# ── Entry and storage ───────────────────────────────────────────────────────

def _hand_routes(fake_sb, entries=None, seen=None):
    fake_sb.route("GET", "pulse_hand_entries",
                  lambda call: FakeResponse(200, entries or []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))

    def capture(call):
        if seen is not None:
            seen.update(call)
        return FakeResponse(201, [{"id": "h1"}])

    fake_sb.route("POST", "pulse_hand_entries", capture)


def test_saving_upserts_on_client_channel_and_month(client, fake_sb):
    seen = {}
    _hand_routes(fake_sb, seen=seen)

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                    data={"channel": "linkedin", "period_month": "2026-09",
                          "sent": "900", "opened": "310", "replied": "44",
                          "confirm": "1"})
    assert r.status_code == 200
    assert seen["params"]["on_conflict"] == \
        "agency_id,client_id,channel,period_month"
    body = seen["json"]
    assert body["period_month"] == "2026-09-01"
    assert body["channel"] == "linkedin"
    assert body["sent"] == 900 and body["opened"] == 310
    assert body["entered_by"]


def test_an_entry_payload_is_json_serialisable(client, fake_sb):
    """The month crosses this boundary as a string; a date object would fail
    only in production."""
    import json

    seen = {}
    _hand_routes(fake_sb, seen=seen)
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                data={"channel": "email", "period_month": "2026-09",
                      "sent": "10", "confirm": "1"})
    json.dumps(seen["json"])
    assert isinstance(seen["json"]["period_month"], str)


def test_a_blank_figure_reads_as_zero(client, fake_sb):
    seen = {}
    _hand_routes(fake_sb, seen=seen)
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                data={"channel": "email", "period_month": "2026-09",
                      "sent": "400", "confirm": "1"})
    assert seen["json"]["bounced"] == 0


def test_a_negative_figure_is_refused_with_a_reason(client, fake_sb):
    _hand_routes(fake_sb)
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                    data={"channel": "email", "period_month": "2026-09",
                          "sent": "-5", "confirm": "1"})
    assert "cannot be negative" in r.text
    assert not fake_sb.calls_to("POST", "pulse_hand_entries")


def test_a_non_numeric_figure_is_refused_with_a_reason(client, fake_sb):
    _hand_routes(fake_sb)
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                    data={"channel": "email", "period_month": "2026-09",
                          "sent": "lots", "confirm": "1"})
    assert "must be a number" in r.text


def test_a_missing_month_is_refused(client, fake_sb):
    _hand_routes(fake_sb)
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                    data={"channel": "email", "period_month": "", "sent": "10"})
    assert "Pick a month" in r.text


def test_saving_warns_once_when_a_connector_already_covers_that_month(client, fake_sb):
    """A partly-synced month is real, so this warns rather than refusing."""
    fake_sb.route("GET", "pulse_hand_entries", lambda call: FakeResponse(200, []))
    fake_sb.route("POST", "pulse_hand_entries", lambda call: FakeResponse(201, [{"id": "h1"}]))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 500, "day": "2026-09-04"},
    ]))

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                    data={"channel": "email", "period_month": "2026-09", "sent": "10"})
    assert "already reported sends" in r.text
    assert "Smartlead" in r.text
    assert not fake_sb.calls_to("POST", "pulse_hand_entries")

    # Confirming goes through.
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                    data={"channel": "email", "period_month": "2026-09",
                          "sent": "10", "confirm": "1"})
    assert fake_sb.calls_to("POST", "pulse_hand_entries")


def test_saving_reloads_the_page_so_the_funnel_cannot_be_stale(client, fake_sb):
    """This page is server-rendered; the funnel and tabs above the panel are
    stale the moment a figure lands."""
    _hand_routes(fake_sb)
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                    data={"channel": "email", "period_month": "2026-09",
                          "sent": "10", "confirm": "1"})
    assert r.headers.get("HX-Refresh") == "true"


def test_entries_are_agency_scoped_on_read_and_delete(client, fake_sb):
    from tests.conftest import param_values

    _hand_routes(fake_sb)
    fake_sb.route("DELETE", "pulse_hand_entries", lambda call: FakeResponse(204, []))
    client.get(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries")
    client.delete(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries/h1")
    for method in ("GET", "DELETE"):
        for call in fake_sb.calls_to(method, "pulse_hand_entries"):
            assert param_values(call, "agency_id")


def test_the_form_names_people_not_events():
    """Every other source counts each person once per stage and the database
    enforces it. This one cannot be enforced, so the wording carries it."""
    from app.utils.pulse import hand_entry

    assert hand_entry.METRIC_LABELS["email"]["opened"] == "Leads who opened"
    assert hand_entry.METRIC_LABELS["linkedin"]["opened"] == \
        "Connection requests accepted"


def test_acceptances_and_opens_share_one_stage():
    """Two columns feeding one stage would let a single entry count both."""
    from app.utils.pulse import hand_entry

    assert hand_entry.STAGE_FOR_METRIC["opened"] == normalize.EVENT_OPENED
    assert "accepted" not in hand_entry.METRICS
    # No separate column in the table either, so an entry cannot report both.
    ddl = _hand_sql().split("CREATE TABLE", 1)[1].split(");", 1)[0]
    assert "accepted" not in ddl
    stages = [v for v in hand_entry.STAGE_FOR_METRIC.values()]
    assert len(stages) == len(set(stages))      # one metric per stage


# ── Into the funnel, with no special-casing ─────────────────────────────────

def test_a_hand_entry_reaches_the_funnel_like_any_other_source(client, fake_sb):
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "linkedin", "source_tool": "hand_entry",
         "event_type": "sent", "events": 900, "day": "2026-09-01"},
        {"client_id": CLIENT, "channel": "linkedin", "source_tool": "hand_entry",
         "event_type": "meeting_booked", "events": 6, "day": "2026-09-01"},
    ]))
    body = client.get(f"/outbound-pulse/clients/{CLIENT}").text
    assert "900" in body
    assert "Entered by hand" in body


def test_a_published_snapshot_includes_hand_entered_figures(client, fake_sb):
    """It joins the same view, so report_snapshot needs no special case —
    there is no mention of the source in the router's snapshot code."""
    import inspect

    from app.routers import outbound_pulse as mod

    seen = {}
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "linkedin", "source_tool": "hand_entry",
         "event_type": "sent", "events": 900, "day": "2026-09-01"},
    ]))

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert seen["snapshot"]["by_source"]["hand_entry"]["sent"] == 900
    assert "hand_entry" not in inspect.getsource(mod.report_snapshot)


# ── The migration ───────────────────────────────────────────────────────────

def _hand_sql():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1]
            / "migrations" / "outbound_pulse_hand_entry.sql").read_text(encoding="utf-8")


def test_the_hand_entry_migration_is_one_transaction_and_re_runnable():
    sql = _hand_sql()
    assert sql.index("BEGIN;") < sql.index("CREATE TABLE")
    assert "CREATE TABLE IF NOT EXISTS pulse_hand_entries" in sql
    assert sql.rstrip().endswith("COMMIT;")


def test_the_funnel_view_unions_all_five_branches():
    view = _hand_sql().split("CREATE VIEW pulse_funnel_daily AS", 1)[1]
    for branch in ("pulse_funnel_rollup", "pulse_lead_outcomes",
                   "o.unsub_day IS NOT NULL", "pulse_manual_daily",
                   "pulse_hand_entry_daily"):
        assert branch in view, branch


def test_the_migration_does_not_touch_the_dnc_manual_view():
    """It adds a source; it does not alter the existing one."""
    assert "CREATE OR REPLACE VIEW pulse_manual_daily" not in _hand_sql()


def test_zero_figures_produce_no_funnel_rows():
    assert "WHERE v.events > 0" in _hand_sql()


def test_one_entry_per_client_channel_and_month():
    assert "pulse_hand_entries_unique" in _hand_sql()
    assert "(agency_id, client_id, channel, period_month)" in _hand_sql()


def test_a_month_must_be_stored_as_its_first_day():
    assert "date_trunc('month', period_month)" in _hand_sql()


def test_figures_cannot_be_negative_in_the_database():
    sql = _hand_sql()
    for col in ("sent", "opened", "replied", "meetings", "bounced", "unsubscribed"):
        assert f"CHECK ({col}" in sql.replace("  ", " ") or f"{col}         >= 0" in sql


# ── Precision below 0.1% ─────────────────────────────────────────────────────

def _pct_of(leads, sent):
    from app.utils.pulse.template_filters import _pct

    c = _counts(sent=sent, meeting_booked=leads)
    return _pct(normalize.lead_rate(c))


def test_a_small_rate_no_longer_collapses_to_zero():
    """At one decimal place a client with four leads from forty thousand sends
    read exactly the same as one with none."""
    assert _pct_of(4, 40000) == "0.01%"
    assert _pct_of(0, 40000) == "0.0%"
    assert _pct_of(4, 40000) != _pct_of(0, 40000)


def test_rates_at_a_tenth_of_a_percent_and_above_are_unchanged():
    """Every existing figure in the tool keeps the shape it had."""
    assert _pct_of(12, 1000) == "1.2%"
    assert _pct_of(50, 1000) == "5.0%"
    assert _pct_of(1, 1000) == "0.1%"
    assert _pct_of(3, 1000) == "0.3%"


def test_a_very_small_rate_still_shows_a_figure():
    """Two significant figures, so something that happened never prints as 0.0%."""
    assert _pct_of(1, 1000000) == "0.0001%"
    assert _pct_of(1, 5000) == "0.02%"


def test_a_rate_is_never_rendered_in_scientific_notation():
    """str(0.00001) is "1e-05", which is not a percentage anyone wants on a
    client report."""
    from app.utils.pulse.template_filters import _pct

    for sent in (10 ** n for n in range(3, 8)):
        rendered = _pct_of(1, sent)
        assert "e" not in rendered.lower(), (sent, rendered)
        assert rendered.endswith("%")


def test_trailing_zeros_are_stripped_below_a_tenth():
    """0.04%, not 0.040% — the number shows the precision it actually has."""
    assert _pct_of(4, 10000) == "0.04%"


def test_no_base_is_still_an_em_dash():
    from app.utils.pulse.template_filters import _pct

    assert _pct(None) == "\u2014"
    assert normalize.lead_rate(_counts(sent=0, meeting_booked=3)) is None


def test_a_nonsense_value_does_not_crash_a_client_report():
    from app.utils.pulse.template_filters import _pct

    assert _pct("not a number") == "\u2014"


def test_the_funnel_step_rates_go_through_the_same_filter(client, fake_sb):
    """Those four sites interpolated the float directly, so they would have
    printed 0.0% while the boxes beside them printed 0.01%."""
    import pathlib

    root = pathlib.Path(__file__).resolve().parents[1] / "app" / "templates"
    for name in ("partials/pulse_funnel.html", "portal.html"):
        html = (root / name).read_text(encoding="utf-8")
        assert "step_rate }}%" not in html, name
        assert "overall }}%" not in html, name
        assert "step_rate | pulse_pct" in html, name


def test_a_low_rate_reads_the_same_everywhere_on_a_report(client, fake_sb):
    """The headline box, the funnel step and the outcome share would otherwise
    disagree about the same number."""
    _portal_routes(fake_sb, reports=[_report()])
    body = client.get("/r/valid-token").text
    assert "0.0%" not in body or "0.4%" in body


# ── The winner on the client's report ────────────────────────────────────────

def _step(winner="A", subject="Quick question", campaign="Acme UK",
          copy=None, reason=""):
    return {
        "campaign": campaign, "step": 1, "subject": subject,
        "basis": "Smartlead's own positive replies",
        "winner": winner, "reason": reason, "total_sent": 2000,
        "variants": [
            {"variant": "A", "is_baseline": True, "sent": 1000, "reply": 40,
             "positive": 12, "unsubscribe": 3, "reply_rate": 4.0,
             "positive_rate": 1.2, "variant_copy": copy},
            {"variant": "B", "is_baseline": False, "sent": 1000, "reply": 18,
             "positive": 3, "unsubscribe": 9, "reply_rate": 1.8,
             "positive_rate": 0.3, "variant_copy": None},
        ],
    }


def test_only_decided_winners_reach_the_report():
    """"Too early to tell" is useful on the internal panel, where it is
    actionable. On a client's report it reads as an excuse."""
    from app.utils.pulse.abtest import winners_for_report

    assert len(winners_for_report([_step()])) == 1
    assert winners_for_report([_step(winner=None, reason="2 variations are level")]) == []


def test_the_report_gets_the_result_not_the_method():
    """A client should not be handed every version we tried, including the ones
    that did badly — that says more about our process than their campaign."""
    from app.utils.pulse.abtest import winners_for_report

    win = winners_for_report([_step()])[0]
    assert win["label"] == "Version A"
    assert win["reply_rate"] == 4.0
    assert "variants" not in win
    assert "B" not in str(win)


def test_written_down_copy_beats_the_step_subject():
    """Smartlead reports one subject for the whole step, so it is the same
    string for every variant and cannot distinguish the winner."""
    from app.utils.pulse.abtest import winners_for_report

    win = winners_for_report([_step(copy={"subject": "The one that won",
                                          "body": "<p>Hello</p>"})])[0]
    assert win["subject"] == "The one that won"
    assert win["body"] == "<p>Hello</p>"


def test_the_manual_winner_joins_the_report_too():
    from app.utils.pulse.abtest import winners_for_report

    manual = {
        "winner": "S1/B1", "reason": "", "total_sent": 700,
        "basis": "your Lead and Interested statuses",
        "variants": [{"variant": "S1/B1", "is_baseline": False, "sent": 700,
                      "reply": 21, "positive": 13, "unsubscribe": 6,
                      "reply_rate": 3.0, "positive_rate": 1.9, "variant_copy": None}],
    }
    out = winners_for_report([], manual)
    assert len(out) == 1
    assert out[0]["campaign"] == "Manual campaigns"


def test_the_winner_is_frozen_into_the_snapshot(client, fake_sb, monkeypatch):
    """Read live on the portal it would be a figure that moves, printed next to
    a column of figures that cannot."""
    from app.utils.pulse import abtest, smartlead

    seen = {}
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, []))
    monkeypatch.setattr(smartlead, "is_configured", lambda: True)
    monkeypatch.setattr(abtest, "smartlead_variants",
                        lambda e, s, en: [_step()])
    monkeypatch.setattr(abtest, "manual_variants",
                        lambda ids: {"variants": [], "winner": None,
                                     "reason": "", "total_sent": 0,
                                     "sheets_read": 0, "sheets_skipped": 0,
                                     "basis": ""})

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert seen["snapshot"]["ab"][0]["label"] == "Version A"


def test_a_failing_ab_lookup_never_blocks_a_publish(client, fake_sb, monkeypatch):
    """The report's own numbers come from our database. Losing a write-up
    because Smartlead was slow would be a far worse trade."""
    from app.utils.pulse import abtest, smartlead

    seen = {}
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        _report(status="draft")]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(200, []))[1])
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, []))
    monkeypatch.setattr(smartlead, "is_configured", lambda: True)

    def boom(*a, **k):
        raise RuntimeError("429 from Smartlead")

    monkeypatch.setattr(abtest, "smartlead_variants", boom)
    monkeypatch.setattr(abtest, "manual_variants", boom)

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert r.status_code == 200
    assert seen["status"] == "published"
    assert seen["snapshot"]["ab"] == []


def test_the_report_shows_the_winner_below_reply_outcomes(client, fake_sb):
    snap = _snapshot()
    snap["ab"] = [{"label": "Version A", "campaign": "Acme UK",
                   "subject": "Quick question about {{company}}",
                   "body": "<p>Short and direct.</p>",
                   "reply_rate": 4.0, "positive_rate": 1.2, "sent": 1000,
                   "basis": "Smartlead's own positive replies"}]
    report = _report()
    report["snapshot"] = snap
    _portal_routes(fake_sb, reports=[report])

    body = client.get("/r/valid-token").text
    assert "What worked best" in body
    assert "Quick question about" in body
    assert "Short and direct." in body
    assert body.index("Reply outcomes") < body.index("What worked best")


def test_a_report_published_before_ab_existed_still_renders(client, fake_sb):
    _portal_routes(fake_sb)        # its snapshot has no "ab" key at all
    r = client.get("/r/valid-token")
    assert r.status_code == 200
    assert "What worked best" not in r.text


# ── Recording what a variant said ───────────────────────────────────────────

def _copy_routes(fake_sb, rows=None, seen=None):
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_copy_variants",
                  lambda call: FakeResponse(200, rows or []))

    def capture(call):
        if seen is not None:
            seen.update(call)
        return FakeResponse(201, [{"id": "cv1"}])

    fake_sb.route("POST", "pulse_copy_variants", capture)


def test_saving_copy_upserts_on_the_variant(client, fake_sb, monkeypatch):
    from app.utils.pulse import smartlead

    seen = {}
    _copy_routes(fake_sb, seen=seen)
    monkeypatch.setattr(smartlead, "is_configured", lambda: False)

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/copy",
                    data={"source_tool": "smartlead", "variant_key": "A",
                          "campaign_ref": "Acme UK",
                          "subject": "Quick question",
                          "body": "<p>Hello</p>"})
    assert r.status_code == 200
    assert seen["params"]["on_conflict"] == \
        "agency_id,client_id,source_tool,campaign_ref,variant_key"
    assert seen["json"]["variant_key"] == "A"
    assert seen["json"]["subject"] == "Quick question"
    assert seen["json"]["entered_by"]


def test_recorded_copy_is_sanitised_before_it_is_stored(client, fake_sb, monkeypatch):
    """It renders on a client-facing report."""
    from app.utils.pulse import smartlead

    seen = {}
    _copy_routes(fake_sb, seen=seen)
    monkeypatch.setattr(smartlead, "is_configured", lambda: False)

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/copy",
                data={"source_tool": "smartlead", "variant_key": "A",
                      "body": "<p onclick='x()'>hi</p><script>bad()</script>"})
    assert seen["json"]["body"] == "<p>hi</p>"


def test_empty_copy_is_refused_with_a_reason(client, fake_sb, monkeypatch):
    from app.utils.pulse import smartlead

    _copy_routes(fake_sb)
    monkeypatch.setattr(smartlead, "is_configured", lambda: False)

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/copy",
                    data={"source_tool": "smartlead", "variant_key": "A",
                          "subject": "  ", "body": "<p><br></p>"})
    assert "Add a subject line or some copy" in r.text
    assert not fake_sb.calls_to("POST", "pulse_copy_variants")


def test_copy_with_no_campaign_matches_any_campaign():
    """A Copy Bank combination is reused wherever it ran, so it should not have
    to be typed once per campaign."""
    from app.routers.outbound_pulse import _copy_for

    index = {("manual", "", "S1/B1"): {"subject": "shared"}}
    assert _copy_for(index, "manual", "Any Campaign", "S1/B1")["subject"] == "shared"


def test_campaign_specific_copy_wins_over_the_shared_one():
    from app.routers.outbound_pulse import _copy_for

    index = {("smartlead", "", "A"): {"subject": "shared"},
             ("smartlead", "Acme UK", "A"): {"subject": "specific"}}
    assert _copy_for(index, "smartlead", "Acme UK", "A")["subject"] == "specific"


def test_a_copy_read_failure_does_not_take_the_ab_panel_down(client, fake_sb, monkeypatch):
    from app.utils import google_sheets
    from app.utils.pulse import smartlead

    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_copy_variants",
                  lambda call: FakeResponse(500, {"message": "does not exist"}))
    monkeypatch.setattr(smartlead, "is_configured", lambda: False)
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: [])

    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab")
    assert r.status_code == 200


# ── The migration ───────────────────────────────────────────────────────────

def _copy_sql():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1] / "migrations"
            / "outbound_pulse_copy_variants.sql").read_text(encoding="utf-8")


def test_the_copy_migration_is_one_transaction_and_re_runnable():
    sql = _copy_sql()
    assert sql.index("BEGIN;") < sql.index("CREATE TABLE")
    assert "CREATE TABLE IF NOT EXISTS pulse_copy_variants" in sql
    assert sql.rstrip().endswith("COMMIT;")


def test_one_copy_record_per_variant():
    assert "pulse_copy_variants_unique" in _copy_sql()


def test_a_copy_record_needs_a_subject_or_a_body():
    assert "length(btrim(subject)) > 0 OR length(btrim(body)) > 0" in _copy_sql()


def test_the_recorded_copy_key_cannot_collide_with_a_dict_method():
    """`v.copy` in Jinja resolves to dict.copy — the built-in method, always
    truthy — so every winner rendered as "copy recorded" with nothing in it.
    Jinja's v['copy'] falls back to the attribute too, so the key has to differ
    from anything on dict."""
    import pathlib

    assert not hasattr({}, "variant_copy")
    ab = (pathlib.Path(__file__).resolve().parents[1] / "app" / "templates"
          / "partials" / "pulse_ab.html").read_text(encoding="utf-8")
    assert "v.copy" not in ab
    assert "v.variant_copy" in ab


def test_a_variant_with_no_recorded_copy_offers_to_add_it(client, fake_sb, monkeypatch):
    from app.utils import google_sheets
    from app.utils.pulse import abtest, smartlead

    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, []))
    monkeypatch.setattr(smartlead, "is_configured", lambda: True)
    monkeypatch.setattr(abtest, "smartlead_variants", lambda e, s, en: [_step()])
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: [])

    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "Add the copy for A" in body
    assert "Recorded by" not in body


def test_a_variant_with_recorded_copy_shows_it(client, fake_sb, monkeypatch):
    from app.utils import google_sheets
    from app.utils.pulse import abtest, smartlead

    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, [
        {"id": "cv1", "client_id": CLIENT, "source_tool": "smartlead",
         "campaign_ref": "Acme UK PR", "variant_key": "A",
         "subject": "The winning line", "body": "<p>Body</p>",
         "entered_by": "Dylan", "updated_at": None}]))
    monkeypatch.setattr(smartlead, "is_configured", lambda: True)
    monkeypatch.setattr(abtest, "smartlead_variants",
                        lambda e, s, en: [_step(campaign="Acme UK PR")])
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: [])

    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "The winning line" in body
    assert "Recorded by Dylan" in body
    assert "Add the copy for A" not in body


# ── Show only, as well as hide ───────────────────────────────────────────────

def _three_clients(fake_sb, prefs=None):
    """Acme, Northfield and the business-development client, all with activity."""
    OTHER = "66666666-6666-6666-6666-666666666666"
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, [
        {"id": CLIENT, "name": "Acme", "color": None, "emoji": None, "active": True},
        {"id": OTHER, "name": "Northfield", "color": None, "emoji": None, "active": True},
        {"id": BD, "name": "Unstuck - Business Development", "color": None,
         "emoji": None, "active": True},
    ]))
    fake_sb.route("GET", "pulse_campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 1000, "day": "2026-09-01"},
        {"client_id": OTHER, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 2000, "day": "2026-09-01"},
        {"client_id": BD, "channel": "email", "source_tool": "smartlead",
         "event_type": "sent", "events": 400, "day": "2026-09-01"},
    ]))
    fake_sb.route("GET", "pulse_user_prefs",
                  lambda call: FakeResponse(200, [] if prefs is None else [prefs]))
    return OTHER


def test_show_only_keeps_just_the_chosen_clients(client, fake_sb):
    other = _three_clients(fake_sb, {"mode": "only", "client_ids": [CLIENT]})
    body = client.get("/api/outbound-pulse/overview").text
    assert "Acme" in body
    assert "Northfield" not in body
    assert "1,000" in body               # Acme alone
    assert "3,400" not in body           # not the whole agency


def test_hide_and_show_only_are_the_same_selection_applied_two_ways(client, fake_sb):
    """One client ticked: hidden under exclude, the only one left under only."""
    _three_clients(fake_sb, {"mode": "exclude", "client_ids": [CLIENT]})
    hidden = client.get("/api/outbound-pulse/overview").text
    assert "Acme" not in hidden.split("Hiding")[0]

    _three_clients(fake_sb, {"mode": "only", "client_ids": [CLIENT]})
    only = client.get("/api/outbound-pulse/overview").text
    assert "Acme" in only
    assert "Northfield" not in only


def test_show_only_names_what_it_is_showing_not_what_it_hid(client, fake_sb):
    """Listing everything left out would be most of the agency."""
    _three_clients(fake_sb, {"mode": "only", "client_ids": [CLIENT]})
    body = client.get("/api/outbound-pulse/overview").text
    assert "Showing only Acme" in body
    assert "Hiding" not in body


def test_hide_mode_still_names_what_it_hid(client, fake_sb):
    _three_clients(fake_sb, {"mode": "exclude", "client_ids": [BD]})
    body = client.get("/api/outbound-pulse/overview").text
    assert "Hiding Unstuck - Business Development" in body
    assert "Showing only" not in body


def test_the_panel_offers_both_modes(client, fake_sb):
    _three_clients(fake_sb)
    body = client.get("/api/outbound-pulse/exclusions").text
    assert 'value="exclude"' in body
    assert 'value="only"' in body
    assert "Hide these" in body
    assert "Show only these" in body


def test_the_panel_summary_counts_the_right_way_round(client, fake_sb):
    _three_clients(fake_sb, {"mode": "only", "client_ids": [CLIENT]})
    assert "Showing 1 client" in client.get("/api/outbound-pulse/exclusions").text

    _three_clients(fake_sb, {"mode": "exclude", "client_ids": [CLIENT]})
    assert "1 client hidden" in client.get("/api/outbound-pulse/exclusions").text


def test_the_mode_is_saved_with_the_selection(client, fake_sb):
    seen = {}
    _three_clients(fake_sb)
    fake_sb.route("POST", "pulse_user_prefs",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(201, []))[1])

    client.post("/api/outbound-pulse/exclusions",
                data={"mode": "only", "excluded": [CLIENT]})
    assert seen["mode"] == "only"
    assert seen["client_ids"] == [CLIENT]


def test_show_only_nothing_is_refused_rather_than_saved(client, fake_sb):
    """An overview with nothing in it reads as broken, not as a filter."""
    _three_clients(fake_sb)
    fake_sb.route("POST", "pulse_user_prefs", lambda call: FakeResponse(201, []))

    r = client.post("/api/outbound-pulse/exclusions", data={"mode": "only"})
    assert "Pick at least one client to show" in r.text
    assert not fake_sb.calls_to("POST", "pulse_user_prefs")


def test_a_rejected_save_keeps_the_mode_the_user_picked(client, fake_sb):
    """Snapping the radio back to what is stored would hide what went wrong."""
    _three_clients(fake_sb)
    r = client.post("/api/outbound-pulse/exclusions", data={"mode": "only"})
    body = r.text
    only_radio = body[body.index('value="only"'):body.index('value="only"') + 60]
    assert "checked" in only_radio


def test_hiding_nothing_is_still_storable(client, fake_sb):
    """Exclude mode with an empty list is "show me everything" and must stay
    saveable — it is the state that distinguishes chosen from never-chosen."""
    seen = {}
    _three_clients(fake_sb)
    fake_sb.route("POST", "pulse_user_prefs",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(201, []))[1])

    r = client.post("/api/outbound-pulse/exclusions", data={"mode": "exclude"})
    assert r.headers.get("HX-Trigger") == "pulseRefresh"
    assert seen["client_ids"] == []
    assert seen["mode"] == "exclude"


def test_an_unknown_mode_falls_back_to_hiding(client, fake_sb):
    seen = {}
    _three_clients(fake_sb)
    fake_sb.route("POST", "pulse_user_prefs",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(201, []))[1])

    client.post("/api/outbound-pulse/exclusions",
                data={"mode": "sideways", "excluded": [CLIENT]})
    assert seen["mode"] == "exclude"


def test_a_stored_mode_that_makes_no_sense_reads_as_hiding(fake_sb):
    """A row written by an older or a broken client must not blank the view."""
    from app.utils.pulse import store

    fake_sb.route("GET", "pulse_user_prefs", lambda call: FakeResponse(
        200, [{"mode": "nonsense", "client_ids": [CLIENT]}]))
    assert store.get_user_filter("dev@local")["mode"] == store.FILTER_EXCLUDE


def test_resetting_still_returns_to_the_default_whichever_mode(client, fake_sb):
    _three_clients(fake_sb, {"mode": "only", "client_ids": [CLIENT]})
    fake_sb.route("DELETE", "pulse_user_prefs", lambda call: FakeResponse(204, []))

    r = client.delete("/api/outbound-pulse/exclusions")
    assert r.headers.get("HX-Trigger") == "pulseRefresh"
    assert fake_sb.calls_to("DELETE", "pulse_user_prefs")


# ── The migration ───────────────────────────────────────────────────────────

def _mode_sql():
    import pathlib
    return (pathlib.Path(__file__).resolve().parents[1] / "migrations"
            / "outbound_pulse_user_prefs_mode.sql").read_text(encoding="utf-8")


def test_the_mode_migration_is_one_transaction_and_re_runnable():
    sql = _mode_sql()
    assert sql.index("BEGIN;") < sql.index("ALTER TABLE")
    assert "ADD COLUMN IF NOT EXISTS mode" in sql
    # The rename is guarded, so a second run does not fail on a missing column.
    assert "column_name = 'excluded_client_ids'" in sql
    assert sql.rstrip().endswith("COMMIT;")


def test_only_mode_cannot_be_stored_empty():
    assert "mode <> 'only' OR cardinality(client_ids) > 0" in _mode_sql()


def test_the_mode_is_constrained_to_the_two_it_supports():
    assert "CHECK (mode IN ('exclude', 'only'))" in _mode_sql()


# ── The portal link is readable from the client profile ──────────────────────

def _access_rows(fake_sb, rows):
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_client_access", lambda call: FakeResponse(200, rows))


def test_a_stored_link_is_shown_and_openable(client, fake_sb):
    _access_rows(fake_sb, [{
        "id": "a1", "label": "Sarah", "created_by": "dylan@unstuck-agency.com",
        "created_at": "2026-09-01T09:00:00+00:00", "expires_at": None,
        "revoked_at": None, "last_used_at": None, "view_count": 3,
        "token": "abc123token",
    }])
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/access").text
    assert "/r/abc123token" in body
    assert ">Open<" in body
    assert "not recoverable" not in body


def test_a_link_issued_before_tokens_were_stored_says_so(client, fake_sb):
    """A hash cannot be reversed. Saying "not recoverable" beats showing a
    broken link or pretending the row is fine."""
    _access_rows(fake_sb, [{
        "id": "a1", "label": "Old one", "created_by": "", "created_at": None,
        "expires_at": None, "revoked_at": None, "last_used_at": None,
        "view_count": 0, "token": None,
    }])
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/access").text
    assert "not recoverable" in body
    assert ">Open<" not in body


def test_a_revoked_link_is_never_offered_to_open(client, fake_sb):
    _access_rows(fake_sb, [{
        "id": "a1", "label": "Gone", "created_by": "", "created_at": None,
        "expires_at": None, "revoked_at": "2026-09-05T09:00:00+00:00",
        "last_used_at": None, "view_count": 2, "token": "stillhere",
    }])
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/access").text
    assert "/r/stillhere" not in body
    assert "revoked" in body


def test_the_link_is_read_back_with_the_access_list(client, fake_sb):
    from tests.conftest import param_values

    _access_rows(fake_sb, [])
    client.get(f"/api/outbound-pulse/clients/{CLIENT}/access")
    call = fake_sb.calls_to("GET", "pulse_client_access")[0]
    assert "token" in param_values(call, "select")[0]


def test_resolving_a_link_still_matches_on_the_hash_only(client, fake_sb):
    """The stored token is display only. If it ever became the lookup key, a
    read of the table would be enough to authenticate."""
    from tests.conftest import param_values

    _portal_routes(fake_sb)
    client.get("/r/valid-token")
    call = fake_sb.calls_to("GET", "pulse_client_access")[0]
    assert param_values(call, "token_hash")
    assert param_values(call, "token") == []


def test_the_visible_links_migration_keeps_the_hash_as_the_key():
    import pathlib

    sql = (pathlib.Path(__file__).resolve().parents[1] / "migrations"
           / "outbound_pulse_visible_links.sql").read_text(encoding="utf-8")
    assert "ADD COLUMN IF NOT EXISTS token TEXT" in sql
    assert "DROP COLUMN" not in sql           # token_hash stays
    assert sql.rstrip().endswith("COMMIT;")


# ── Manual campaigns on the client profile ──────────────────────────────────

def test_manual_campaigns_are_listed_on_the_client_page(client, fake_sb):
    """They cannot appear in the funnel's campaign table: their outcomes are
    recorded against the client, so pulse_manual_daily carries a NULL
    campaign_id by design."""
    _detail_routes(fake_sb)
    fake_sb.route("GET", "campaigns", lambda call: FakeResponse(200, [
        {"id": "mc1", "campaign_name": "Acme - PR Directors",
         "sender_profile_name": "Sarah", "sheet_url": "https://sheets/x",
         "total_prospects": 420, "sent_count": 410,
         "created_at": "2026-09-02T09:00:00+00:00"},
    ]))
    body = client.get(f"/outbound-pulse/clients/{CLIENT}").text
    assert "Manual campaigns" in body
    assert "Acme - PR Directors" in body
    assert "420" in body
    assert "https://sheets/x" in body


def test_the_manual_section_is_absent_when_there_are_none(client, fake_sb):
    _detail_routes(fake_sb)
    fake_sb.route("GET", "campaigns", lambda call: FakeResponse(200, []))
    body = client.get(f"/outbound-pulse/clients/{CLIENT}").text
    assert "Manual campaigns" not in body


def test_manual_campaigns_are_scoped_to_the_client(client, fake_sb):
    from tests.conftest import param_values

    _detail_routes(fake_sb)
    fake_sb.route("GET", "campaigns", lambda call: FakeResponse(200, []))
    client.get(f"/outbound-pulse/clients/{CLIENT}")
    for call in fake_sb.calls_to("GET", "campaigns"):
        assert param_values(call, "client_id") == [f"eq.{CLIENT}"]


def test_a_failing_campaigns_read_does_not_break_the_client_page(client, fake_sb):
    """A supporting list, not the page."""
    _detail_routes(fake_sb)
    fake_sb.route("GET", "campaigns",
                  lambda call: FakeResponse(500, {"message": "nope"}))
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert r.status_code == 200
    assert "Acme" in r.text


# ── Hand entry is LinkedIn only ─────────────────────────────────────────────

def test_hand_entry_is_linkedin_only():
    """Meet Alfred is the LinkedIn channel, and it is the one source nothing
    here can read for itself. Email is covered by Smartlead and the DNC tool."""
    from app.utils.pulse import hand_entry

    assert hand_entry.CHANNEL == "linkedin"
    assert hand_entry.CHANNELS == ("linkedin",)


def test_the_entry_form_offers_no_channel_choice(client, fake_sb):
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_hand_entries", lambda call: FakeResponse(200, []))
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries").text
    assert "<select name=\"channel\"" not in body
    assert 'name="channel" value="linkedin"' in body
    # And the labels are the LinkedIn ones.
    assert "Connection requests sent" in body
    assert "Emails sent" not in body


def test_a_posted_email_channel_is_ignored(client, fake_sb):
    """Nothing writes an email row any more, whatever the form says."""
    seen = {}
    fake_sb.route("GET", "pulse_hand_entries", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, []))
    fake_sb.route("POST", "pulse_hand_entries",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(201, [{"id": "h1"}]))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                data={"channel": "email", "period_month": "2026-09",
                      "sent": "100", "confirm": "1"})
    assert seen["channel"] == "linkedin"


# ── A manual variant label resolves against the Copy Bank ────────────────────

def test_a_variant_label_is_an_index_into_the_copy_bank():
    from app.utils import copy_bank

    assert copy_bank.parse_variant("S2/B1") == (1, 0)
    assert copy_bank.parse_variant(" S3 / B2 ") == (2, 1)
    assert copy_bank.parse_variant("A") == (None, None)
    assert copy_bank.parse_variant("") == (None, None)


def test_blank_copy_cards_are_dropped_before_indexing():
    """The index counts what a person can see in Copy Bank, and Copy Bank does
    not render empty cards. Counting them here would offset every label."""
    from app.utils import copy_bank

    out = copy_bank.extract({"email": {
        "subjects":   ["First", "   ", "Third"],
        "variations": [{"body": "Body one"}, {"body": ""}, {"body": "Body three"}],
    }})
    assert out["subjects"] == ["First", "Third"]
    assert out["bodies"] == ["Body one", "Body three"]


def test_a_label_past_the_end_of_the_arrays_resolves_to_nothing():
    """Copy edited after a campaign went out renumbers everything below it."""
    from app.utils import copy_bank

    copy = {"subjects": ["One"], "bodies": ["Body"]}
    assert copy_bank.variant_in(copy, "S1/B1") == {"subject": "One", "body": "Body"}
    assert copy_bank.variant_in(copy, "S4/B1") is None
    assert copy_bank.variant_in(copy, "S1/B9") is None


def _manual_ab(fake_sb, monkeypatch, variants=None, campaigns=None, templates=None):
    """A client page with one manual campaign and no Smartlead."""
    from app.utils import google_sheets
    from app.utils.pulse import smartlead

    _detail_routes(fake_sb)
    monkeypatch.setattr(smartlead, "is_configured", lambda: False)
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: variants
                        if variants is not None else [
        {"variant": "S1/B1", "total": 300, "lead": 9, "interested": 3,
         "reply": 14, "unsubscribe": 2, "positive": 12},
        {"variant": "S1/B2", "total": 300, "lead": 1, "interested": 0,
         "reply": 5, "unsubscribe": 6, "positive": 1},
    ])
    fake_sb.route("GET", "campaigns", lambda call: FakeResponse(200, campaigns
                  if campaigns is not None else
                  [{"id": "mc1", "sheet_id": "s1", "campaign_name": "Acme MSP",
                    "completed_at": None, "copy_territory": "uk",
                    "copy_industry": "msp"}]))
    fake_sb.route("GET", "copy_bank_templates", lambda call: FakeResponse(200,
                  templates if templates is not None else
                  [{"content": {"email": {
                      "subjects":   ["Quick one about your MSP", "Second subject"],
                      "variations": [{"body": "Hi {{first_name}},\n\nSaw you run MSP."},
                                     {"body": "Different body."}],
                  }}}]))
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, []))


def test_the_winning_variants_copy_is_read_out_of_the_copy_bank(client, fake_sb, monkeypatch):
    """S1/B1 means subject 1 with body 1 of the entry the merge recorded, so
    nobody should have to type copy that is already written down."""
    _manual_ab(fake_sb, monkeypatch)
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "Quick one about your MSP" in body
    assert "Saw you run MSP." in body
    assert "From the Copy Bank" in body
    assert "Add the copy for S1/B1" not in body


def test_the_copy_bank_entry_is_looked_up_by_what_the_merge_recorded(client, fake_sb, monkeypatch):
    from tests.conftest import param_values

    _manual_ab(fake_sb, monkeypatch)
    client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab")
    keys = [v for call in fake_sb.calls_to("GET", "copy_bank_templates")
            for v in param_values(call, "key")]
    assert f"eq.__c__{CLIENT}__uk_msp" in keys


def test_copy_typed_by_hand_beats_the_copy_bank(client, fake_sb, monkeypatch):
    """Writing it down is somebody correcting what the index resolves to."""
    _manual_ab(fake_sb, monkeypatch)
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, [
        {"source_tool": "manual", "campaign_ref": "", "variant_key": "S1/B1",
         "subject": "What we actually sent", "body": "<p>Corrected.</p>",
         "entered_by": "Sarah"},
    ]))
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "What we actually sent" in body
    assert "Recorded by Sarah" in body
    assert "Quick one about your MSP" not in body


def test_each_campaign_resolves_against_its_own_copy_bank_entry(client, fake_sb, monkeypatch):
    """Per campaign there is no ambiguity: a campaign is one send of one Copy
    Bank entry, so its labels resolve against that entry and nothing else."""
    from tests.conftest import param_values
    from app.utils import google_sheets
    from app.utils.pulse import smartlead

    _detail_routes(fake_sb)
    monkeypatch.setattr(smartlead, "is_configured", lambda: False)
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: [
        {"variant": "S1/B1", "total": 300, "lead": 9, "interested": 3,
         "reply": 14, "unsubscribe": 2, "positive": 12},
        {"variant": "S1/B2", "total": 300, "lead": 1, "interested": 0,
         "reply": 5, "unsubscribe": 6, "positive": 1},
    ])
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "campaigns", lambda call: FakeResponse(200, [
        {"id": "mc1", "sheet_id": "s1", "campaign_name": "MSP push",
         "completed_at": None, "copy_territory": "uk", "copy_industry": "msp"},
        {"id": "mc2", "sheet_id": "s2", "campaign_name": "Media push",
         "completed_at": None, "copy_territory": "uk", "copy_industry": "media"},
    ]))

    def templates(call):
        key = param_values(call, "key")[0]
        subject = "MSP subject" if "msp" in key else "Media subject"
        return FakeResponse(200, [{"content": {"email": {
            "subjects": [subject], "variations": [{"body": subject + " body"}]}}}])

    fake_sb.route("GET", "copy_bank_templates", templates)

    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "MSP push" in body and "Media push" in body
    assert "MSP subject" in body
    assert "Media subject" in body


def test_the_combined_block_refuses_copy_when_campaigns_differ(client, fake_sb, monkeypatch):
    """Added up, "S1/B1" can be two different messages and the sheet cannot
    say which one a prospect got."""
    _manual_ab(fake_sb, monkeypatch, campaigns=[
        {"id": "mc1", "sheet_id": "s1", "campaign_name": "A", "completed_at": None,
         "copy_territory": "uk", "copy_industry": "msp"},
        {"id": "mc2", "sheet_id": "s2", "campaign_name": "B", "completed_at": None,
         "copy_territory": "uk", "copy_industry": "media"},
    ])
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "All 2 campaigns together" in body
    assert "different copy per campaign" in body
    assert "2 different Copy Bank entries" in body


def test_the_combined_block_is_absent_for_a_single_campaign(client, fake_sb, monkeypatch):
    """There is nothing to add up, and repeating the same table under a
    different heading would read as a second result."""
    _manual_ab(fake_sb, monkeypatch)
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "campaigns together" not in body


def test_copy_is_saved_against_one_campaign(client, fake_sb, monkeypatch):
    """The campaign's id, not its name — renaming a campaign must not lose the
    copy somebody typed for it."""
    _manual_ab(fake_sb, monkeypatch, campaigns=[
        {"id": "mc1", "sheet_id": "s1", "campaign_name": "MSP push",
         "completed_at": None, "copy_territory": "uk", "copy_industry": "msp"},
    ])
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert 'name="campaign_ref" value="mc1"' in body


def test_copy_typed_for_one_campaign_does_not_leak_to_another(client, fake_sb, monkeypatch):
    _manual_ab(fake_sb, monkeypatch, templates=[], campaigns=[
        {"id": "mc1", "sheet_id": "s1", "campaign_name": "First",
         "completed_at": None, "copy_territory": "", "copy_industry": ""},
        {"id": "mc2", "sheet_id": "s2", "campaign_name": "Second",
         "completed_at": None, "copy_territory": "", "copy_industry": ""},
    ])
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, [
        {"source_tool": "manual", "campaign_ref": "mc1", "variant_key": "S1/B1",
         "subject": "Only for the first campaign", "body": "<p>x</p>",
         "entered_by": "Sarah"},
    ]))
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert body.count("Only for the first campaign") == 2   # shown once, prefilled once
    assert "Add the copy for S1/B1" in body                  # the other campaign


def test_copy_recorded_for_every_campaign_still_shows_under_each(client, fake_sb, monkeypatch):
    """Copy saved before the breakdown existed carries no campaign, and must
    not disappear now that one is asked for."""
    _manual_ab(fake_sb, monkeypatch, templates=[], campaigns=[
        {"id": "mc1", "sheet_id": "s1", "campaign_name": "First",
         "completed_at": None, "copy_territory": "", "copy_industry": ""},
        {"id": "mc2", "sheet_id": "s2", "campaign_name": "Second",
         "completed_at": None, "copy_territory": "", "copy_industry": ""},
    ])
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, [
        {"source_tool": "manual", "campaign_ref": "", "variant_key": "S1/B1",
         "subject": "Recorded once for all", "body": "<p>x</p>",
         "entered_by": "Sarah"},
    ]))
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "Add the copy for S1/B1" not in body


def test_the_breakdown_reads_each_sheet_once(monkeypatch):
    """Per-campaign and combined come out of the same reads. A second pass
    would double every Google Sheets call for a figure already computed."""
    from app.utils import google_sheets
    from app.utils.pulse import abtest

    reads = []

    def reader(sid):
        reads.append(sid)
        return [{"variant": "S1/B1", "total": 300, "lead": 9, "interested": 3,
                 "reply": 14, "unsubscribe": 2, "positive": 12},
                {"variant": "S1/B2", "total": 300, "lead": 1, "interested": 0,
                 "reply": 5, "unsubscribe": 6, "positive": 1}]

    monkeypatch.setattr(google_sheets, "read_ab_stats", reader)
    out = abtest.manual_breakdown([
        {"ref": "mc1", "name": "A", "sheet_id": "s1"},
        {"ref": "mc2", "name": "B", "sheet_id": "s2"},
    ])
    assert reads == ["s1", "s2"]
    assert len(out["campaigns"]) == 2
    assert out["combined"]["variants"][0]["sent"] == 600      # 300 + 300


def test_campaigns_are_ordered_with_the_biggest_first(monkeypatch):
    """The campaign with the most behind it is the one to weigh most, and the
    one that most often has a verdict at all."""
    from app.utils import google_sheets
    from app.utils.pulse import abtest

    sheets = {
        "small": [{"variant": "S1/B1", "total": 50, "lead": 1, "interested": 0,
                   "reply": 2, "unsubscribe": 0, "positive": 1},
                  {"variant": "S1/B2", "total": 50, "lead": 0, "interested": 0,
                   "reply": 1, "unsubscribe": 1, "positive": 0}],
        "big":   [{"variant": "S1/B1", "total": 900, "lead": 20, "interested": 5,
                   "reply": 40, "unsubscribe": 3, "positive": 25},
                  {"variant": "S1/B2", "total": 880, "lead": 2, "interested": 1,
                   "reply": 12, "unsubscribe": 9, "positive": 3}],
    }
    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: sheets[sid])
    out = abtest.manual_breakdown([
        {"ref": "a", "name": "Small", "sheet_id": "small"},
        {"ref": "b", "name": "Big", "sheet_id": "big"},
    ])
    assert [b["name"] for b in out["campaigns"]] == ["Big", "Small"]


def test_a_campaign_too_small_to_call_says_so_rather_than_naming_a_winner(monkeypatch):
    """Splitting by campaign means less volume each. That is the truth the
    combined figure was hiding, not a regression."""
    from app.utils import google_sheets
    from app.utils.pulse import abtest

    monkeypatch.setattr(google_sheets, "read_ab_stats", lambda sid: [
        {"variant": "S1/B1", "total": 12, "lead": 1, "interested": 0,
         "reply": 1, "unsubscribe": 0, "positive": 1},
        {"variant": "S1/B2", "total": 11, "lead": 0, "interested": 0,
         "reply": 0, "unsubscribe": 0, "positive": 0},
    ])
    out = abtest.manual_breakdown([{"ref": "a", "name": "Tiny", "sheet_id": "s"}])
    assert out["campaigns"][0]["winner"] is None
    assert "Too little volume" in out["campaigns"][0]["reason"]


def test_a_report_names_the_campaign_each_winner_came_from(monkeypatch):
    """"Manual campaigns" was accurate for one combined verdict and would be a
    lie against a figure covering one campaign."""
    from app.utils.pulse.abtest import winners_for_report

    blocks = [
        {"winner": "S1/B1", "name": "MSP push", "basis": "your Lead statuses",
         "variants": [{"variant": "S1/B1", "sent": 700, "reply_rate": 3.0,
                       "positive_rate": 1.9, "variant_copy": {"subject": "Won",
                                                              "body": "<p>b</p>"}}]},
        {"winner": "S2/B1", "name": "Media push", "basis": "your Lead statuses",
         "variants": [{"variant": "S2/B1", "sent": 400, "reply_rate": 2.0,
                       "positive_rate": 1.0, "variant_copy": None}]},
    ]
    out = winners_for_report([], blocks)
    assert [w["campaign"] for w in out] == ["MSP push", "Media push"]
    assert out[0]["subject"] == "Won"


def test_a_report_published_before_the_breakdown_still_replays(monkeypatch):
    """Re-publishing an old report runs its combined verdict through here."""
    from app.utils.pulse.abtest import winners_for_report

    combined = {
        "winner": "S1/B1", "basis": "your Lead and Interested statuses",
        "variants": [{"variant": "S1/B1", "sent": 700, "reply_rate": 3.0,
                      "positive_rate": 1.9, "variant_copy": None}],
    }
    out = winners_for_report([], combined)
    assert out[0]["campaign"] == "Manual campaigns"


def test_a_campaign_with_no_copy_source_resolves_nothing(client, fake_sb, monkeypatch):
    """Choosing a copy source is optional in the merge, so plenty of campaigns
    have none — that is a missing lookup, not an error."""
    _manual_ab(fake_sb, monkeypatch, campaigns=[
        {"sheet_id": "s1", "campaign_name": "A", "completed_at": None,
         "copy_territory": None, "copy_industry": None},
    ])
    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab")
    assert r.status_code == 200
    assert "Add the copy for S1/B1" in r.text
    assert "different Copy Bank entries" not in r.text


def test_copy_bank_text_is_escaped_before_it_reaches_a_page(client, fake_sb, monkeypatch):
    """Copy Bank stores plain text and renders it with textContent, so it has
    never been escaped — and this body goes on to a client-facing report."""
    _manual_ab(fake_sb, monkeypatch, templates=[{"content": {"email": {
        "subjects":   ["Subject"],
        "variations": [{"body": "<script>alert(1)</script> 5 > 3"}],
    }}}])
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab").text
    assert "<script>alert(1)</script>" not in body
    assert "&lt;script&gt;" in body


def test_line_breaks_in_copy_bank_text_survive_as_structure():
    from app.utils.pulse import richtext

    out = richtext.from_text("Line one\nLine two\n\nNew paragraph")
    assert out == "<p>Line one<br>Line two</p><p>New paragraph</p>"
    assert richtext.from_text("   ") == ""


def test_the_copy_bank_entry_is_fetched_once_for_the_whole_table(client, fake_sb, monkeypatch):
    """Once per variant row would be a request per row for text that does not
    change between them."""
    _manual_ab(fake_sb, monkeypatch)
    client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab")
    assert len(fake_sb.calls_to("GET", "copy_bank_templates")) == 1


def test_a_published_report_carries_the_resolved_copy(client, fake_sb, monkeypatch):
    """The point of resolving it: the client reads the message, not a letter."""
    saved = {}
    _manual_ab(fake_sb, monkeypatch)
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        {"id": "r1", "client_id": CLIENT, "period_start": "2026-09-01",
         "period_end": "2026-09-30", "title": "", "body": "<p>Hi</p>",
         "snapshot": None, "status": "draft", "published_at": None,
         "created_by": "", "created_at": None, "updated_at": None},
    ]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (saved.update(call.get("json") or {}),
                                FakeResponse(204, []))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    winners = (saved.get("snapshot") or {}).get("ab") or []
    assert any("Quick one about your MSP" in (w.get("subject") or "")
               for w in winners), winners


def test_a_copy_bank_read_failure_does_not_take_the_panel_down(client, fake_sb, monkeypatch):
    _manual_ab(fake_sb, monkeypatch)
    fake_sb.route("GET", "copy_bank_templates",
                  lambda call: FakeResponse(500, {"message": "nope"}))
    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/ab")
    assert r.status_code == 200
    assert "Add the copy for S1/B1" in r.text


def test_copy_banks_own_lookup_still_falls_back_to_the_bizdev_key(client, fake_sb):
    """The key format and the fallback moved into a shared helper. Copy Bank's
    own endpoint has to resolve exactly what it did before."""
    from tests.conftest import param_values

    seen = []

    def handler(call):
        key = param_values(call, "key")[0]
        seen.append(key)
        if key == "eq.uk_msp":
            return FakeResponse(200, [{"content": {"email": {
                "subjects": ["Migrated"], "variations": [{"body": "Body"}]}}}])
        return FakeResponse(200, [])

    fake_sb.route("GET", "copy_bank_templates", handler)
    out = client.get("/api/copy-bank/templates/acme-id/uk/msp").json()
    assert out["subjects"] == ["Migrated"]
    assert seen == ["eq.__c__acme-id__uk_msp", "eq.uk_msp"]


# ── Clearing a client's report history ───────────────────────────────────────

def _report_rows(fake_sb, rows=None, on_delete=None):
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200,
        rows if rows is not None else [
            {"id": "r1", "client_id": CLIENT, "period_start": "2026-09-01",
             "period_end": "2026-09-30", "title": "", "body": "<p>x</p>",
             "snapshot": None, "status": "published",
             "published_at": "2026-10-01T09:00:00+00:00", "created_by": "",
             "created_at": None, "updated_at": None},
            {"id": "r2", "client_id": CLIENT, "period_start": "2026-08-01",
             "period_end": "2026-08-31", "title": "", "body": "",
             "snapshot": None, "status": "draft", "published_at": None,
             "created_by": "", "created_at": None, "updated_at": None},
        ]))
    if on_delete is not None:
        fake_sb.route("DELETE", "pulse_reports", on_delete)


def test_the_panel_offers_to_clear_the_history(client, fake_sb):
    _report_rows(fake_sb)
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/reports").text
    assert "Clear this client" in body
    assert "Type DELETE to confirm" in body


def test_there_is_nothing_to_clear_when_there_are_no_reports(client, fake_sb):
    _report_rows(fake_sb, rows=[])
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/reports").text
    assert "Clear this client" not in body


def test_clearing_deletes_every_report_for_that_client(client, fake_sb):
    _report_rows(fake_sb, on_delete=lambda call: FakeResponse(
        200, [{"id": "r1"}, {"id": "r2"}]))
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/clear",
                    data={"confirm": "DELETE"})
    assert r.status_code == 200
    assert "Cleared 2 reports" in r.text
    assert len(fake_sb.calls_to("DELETE", "pulse_reports")) == 1


def test_the_wipe_is_scoped_to_one_client_and_one_agency(client, fake_sb):
    """A DELETE missing a filter is the one typo this must not have."""
    from tests.conftest import param_values

    _report_rows(fake_sb, on_delete=lambda call: FakeResponse(200, []))
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/clear",
                data={"confirm": "DELETE"})
    call = fake_sb.calls_to("DELETE", "pulse_reports")[0]
    assert param_values(call, "client_id") == [f"eq.{CLIENT}"]
    assert param_values(call, "agency_id")


def test_the_wrong_word_deletes_nothing(client, fake_sb):
    _report_rows(fake_sb, on_delete=lambda call: FakeResponse(200, []))
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/clear",
                    data={"confirm": "delete all"})
    assert "Type DELETE to confirm" in r.text
    assert "Nothing has been deleted" in r.text
    assert not fake_sb.calls_to("DELETE", "pulse_reports")


def test_an_empty_confirmation_deletes_nothing(client, fake_sb):
    _report_rows(fake_sb, on_delete=lambda call: FakeResponse(200, []))
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/clear",
                data={"confirm": ""})
    assert not fake_sb.calls_to("DELETE", "pulse_reports")


def test_a_refused_confirmation_leaves_the_form_open(client, fake_sb):
    """Collapsing it would make a typo feel like the button had stopped working."""
    _report_rows(fake_sb)
    body = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/clear",
                       data={"confirm": "nope"}).text
    assert "data-wipe-form hidden" not in body
    assert "data-wipe-form " in body


def test_a_wipe_that_removes_nothing_does_not_claim_the_link_is_cleared(client, fake_sb):
    """Somebody else got there first. Claiming credit for a delete that did
    nothing reads identically to the real thing otherwise."""
    _report_rows(fake_sb, on_delete=lambda call: FakeResponse(200, []))
    body = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/clear",
                       data={"confirm": "DELETE"}).text
    assert "nothing left to clear" in body
    assert "no history on it" not in body


def test_a_failed_wipe_says_so_rather_than_claiming_success(client, fake_sb):
    _report_rows(fake_sb, on_delete=lambda call: FakeResponse(500, {"message": "no"}))
    body = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/clear",
                       data={"confirm": "DELETE"}).text
    assert "Could not clear the reports" in body
    assert "Cleared" not in body


def test_clearing_nothing_is_not_reported_as_a_failure():
    """Nothing was there and the delete did not happen must not read the same
    to the caller."""
    from app.utils.pulse import store

    assert store.delete_all_reports("") is None


def test_the_confirm_form_is_hidden_until_it_is_opened(client, fake_sb):
    """An author-origin `display` beats the browser's [hidden], which is how
    the date range controls were broken for a week."""
    _detail_routes(fake_sb)
    body = client.get(f"/outbound-pulse/clients/{CLIENT}").text
    assert ".report-wipe-form[hidden]" in body


# ── Typed figures survive the round trip to the server ───────────────────────

def _hand_save_routes(fake_sb, covered=None):
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_hand_entries", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_funnel_daily",
                  lambda call: FakeResponse(200, covered or []))


def test_the_overlap_warning_gives_the_typed_figures_back(client, fake_sb):
    """The warning is a round trip. Coming back to an empty form meant "save
    again to confirm" confirmed a month of zeros, and the figures were gone."""
    _hand_save_routes(fake_sb, covered=[
        {"client_id": CLIENT, "campaign_id": CAMPAIGN, "channel": "linkedin",
         "event_type": "sent", "events": 500, "day": "2026-09-04"},
    ])
    body = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                       data={"period_month": "2026-09", "sent": "820",
                             "replied": "38", "meetings": "4",
                             "source_note": "Meet Alfred, 2 Oct"}).text
    assert "already reported sends" in body
    assert 'value="820"' in body
    assert 'value="38"' in body
    assert 'value="4"' in body
    assert 'value="Meet Alfred, 2 Oct"' in body


def test_confirming_after_the_warning_saves_the_real_figures(client, fake_sb):
    saved = {}
    _hand_save_routes(fake_sb)
    fake_sb.route("POST", "pulse_hand_entries",
                  lambda call: (saved.update(call.get("json") or {}),
                                FakeResponse(201, [{"id": "h1"}]))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                data={"period_month": "2026-09", "sent": "820", "replied": "38",
                      "meetings": "4", "confirm": "1"})
    assert saved["sent"] == 820
    assert saved["replied"] == 38
    assert saved["meetings"] == 4
    assert saved["period_month"] == "2026-09-01"


def test_a_bad_figure_does_not_wipe_the_good_ones(client, fake_sb):
    _hand_save_routes(fake_sb)
    body = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                       data={"period_month": "2026-09", "sent": "820",
                             "replied": "lots", "confirm": "1"}).text
    assert "must be a number" in body
    assert 'value="820"' in body


def test_a_month_of_nothing_is_refused(client, fake_sb):
    """The funnel view skips zero rows, so this would save a record that shows
    up nowhere and looks exactly like the save having failed."""
    _hand_save_routes(fake_sb)
    fake_sb.route("POST", "pulse_hand_entries",
                  lambda call: FakeResponse(201, [{"id": "h1"}]))
    body = client.post(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries",
                       data={"period_month": "2026-09", "confirm": "1"}).text
    assert "Enter at least one figure" in body
    assert not fake_sb.calls_to("POST", "pulse_hand_entries")


def test_an_empty_form_renders_empty_boxes_not_zeros(client, fake_sb):
    """A field pre-filled with 0 reads as a figure somebody entered."""
    _hand_save_routes(fake_sb)
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/hand-entries").text
    assert 'value="0"' not in body


def test_hand_entered_figures_reach_a_published_report(client, fake_sb):
    """The whole point of typing them in. They land on the first of the month,
    so a report covering that month includes them."""
    seen = {}
    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_funnel_daily", lambda call: FakeResponse(200, [
        {"client_id": CLIENT, "campaign_id": None, "channel": "linkedin",
         "event_type": "sent", "events": 820, "day": "2026-09-01",
         "source_tool": "hand_entry"},
        {"client_id": CLIENT, "campaign_id": None, "channel": "linkedin",
         "event_type": "replied", "events": 38, "day": "2026-09-01",
         "source_tool": "hand_entry"},
    ]))
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        {"id": "r1", "client_id": CLIENT, "period_start": "2026-09-01",
         "period_end": "2026-09-30", "title": "", "body": "", "snapshot": None,
         "status": "draft", "published_at": None, "created_by": "",
         "created_at": None, "updated_at": None},
    ]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(204, []))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert seen["snapshot"]["counts"]["sent"] == 820
    assert seen["snapshot"]["counts"]["replied"] == 38


# ── The client report ────────────────────────────────────────────────────────

def _published(fake_sb, snapshot):
    _portal_routes(fake_sb, reports=[{
        "id": "sep", "client_id": CLIENT, "period_start": "2026-09-01",
        "period_end": "2026-09-30", "title": "", "body": "<p>Note</p>",
        "snapshot": snapshot, "status": "published",
        "published_at": "2026-10-01T09:00:00+00:00", "created_by": "",
        "created_at": "2026-10-01T09:00:00+00:00", "updated_at": None,
    }])


def test_the_winning_copy_names_its_parts(client, fake_sb):
    """A subject line and an opening paragraph look alike once they are both
    just quoted text on a page."""
    _published(fake_sb, {
        "version": 1, "counts": {"sent": 900, "replied": 30, "lead": 5},
        "funnel": [], "by_channel": {}, "trend": {"unit": "day", "buckets": []},
        "ab": [{"label": "Version A", "campaign": "UK PR", "sent": 900,
                "subject": "Quick question", "body": "<p>Hi there,</p>",
                "reply_rate": 3.3, "positive_rate": 1.1, "basis": ""}],
    })
    body = client.get("/r/valid-token").text
    assert "Subject line" in body
    assert "Main email body" in body
    assert body.index("Subject line") < body.index("Main email body")


def test_the_outcome_cards_drop_the_share_line(client, fake_sb):
    _published(fake_sb, {
        "version": 1,
        "counts": {"sent": 1000, "replied": 100, "interested": 31,
                   "info_request": 4, "meeting_booked": 9, "unsubscribed": 22},
        "funnel": [], "by_channel": {}, "trend": {"unit": "day", "buckets": []},
        "ab": [],
    })
    body = client.get("/r/valid-token").text
    assert "Interested" in body                  # the cards are still there
    assert "outcome-share" not in body
    assert "of replies</div>" not in body


def test_the_client_report_has_no_daily_trend(client, fake_sb):
    """A send is dated when the campaign is uploaded, not when each message
    goes out, so the strip showed a cliff on upload day and little after it."""
    _published(fake_sb, {
        "version": 1, "counts": {"sent": 900, "replied": 30},
        "funnel": [], "by_channel": {}, "ab": [],
        "trend": {"unit": "day", "buckets": [
            {"start": "2026-09-01", "end": "2026-09-01", "label": "1 Sep", "sent": 800},
            {"start": "2026-09-02", "end": "2026-09-02", "label": "2 Sep", "sent": 12},
            {"start": "2026-09-03", "end": "2026-09-03", "label": "3 Sep", "sent": 9},
        ]},
    })
    body = client.get("/r/valid-token").text
    assert "Activity over time" not in body
    assert "trend-strip" not in body


def test_the_internal_page_keeps_its_trend(client, fake_sb):
    """It is the same data; the difference is that the shape is understood
    there, and account managers need it to spot a stalled send."""
    _detail_routes(fake_sb)
    assert "trend-strip" in client.get(f"/outbound-pulse/clients/{CLIENT}").text


# ── LinkedIn is counted on its own, not folded into the email funnel ─────────

def test_a_funnel_can_have_another_taken_out_of_it():
    from app.utils.pulse.normalize import without

    total = {"sent": 1000, "replied": 60, "meeting_booked": 5}
    part = {"sent": 300, "replied": 20, "meeting_booked": 2}
    assert without(total, part) == {"sent": 700, "replied": 40, "meeting_booked": 3}


def test_subtracting_never_goes_below_zero():
    """The two sides are counted independently. A late sync landing between
    them must not put a negative send count on a client's report."""
    from app.utils.pulse.normalize import without

    assert without({"sent": 10}, {"sent": 40})["sent"] == 0


def test_subtracting_nothing_leaves_the_figures_alone():
    from app.utils.pulse.normalize import without

    assert without({"sent": 10, "replied": 2}, {}) == {"sent": 10, "replied": 2}


def _report_with(fake_sb, counts, by_channel):
    _portal_routes(fake_sb, reports=[{
        "id": "sep", "client_id": CLIENT, "period_start": "2026-09-01",
        "period_end": "2026-09-30", "title": "", "body": "", "status": "published",
        "published_at": "2026-10-01T09:00:00+00:00", "created_by": "",
        "created_at": "2026-10-01T09:00:00+00:00", "updated_at": None,
        "snapshot": {"version": 1, "counts": counts, "by_channel": by_channel,
                     "trend": {"unit": "day", "buckets": []}, "ab": []},
    }])


_MIXED_TOTAL = {"sent": 2571, "replied": 98, "meeting_booked": 9,
                "info_request": 4, "positive_reply": 25, "unsubscribed": 15}
_MIXED_LINKEDIN = {"sent": 820, "replied": 38, "meeting_booked": 4,
                   "info_request": 0, "positive_reply": 6, "unsubscribed": 2,
                   "opened": 265}


def test_linkedin_is_taken_out_of_the_headline_figures(client, fake_sb):
    """A connection request is not an email. Added together they produced a
    "sent" the client could not reconcile with anything."""
    _report_with(fake_sb, _MIXED_TOTAL, {"linkedin": _MIXED_LINKEDIN})
    body = client.get("/r/valid-token").text
    assert "1,751" in body              # 2571 emails - 820 connection requests
    assert "2,571" not in body
    assert "Emails sent" in body


def test_linkedin_gets_its_own_card(client, fake_sb):
    _report_with(fake_sb, _MIXED_TOTAL, {"linkedin": _MIXED_LINKEDIN})
    body = client.get("/r/valid-token").text
    assert "Connection requests sent" in body
    assert "820" in body
    assert "Requests accepted" in body
    assert "265" in body


def test_the_old_combined_channel_card_is_gone(client, fake_sb):
    """It said both channels roll into the funnel above, which is no longer
    true and was the thing being corrected."""
    _report_with(fake_sb, _MIXED_TOTAL, {"linkedin": _MIXED_LINKEDIN})
    body = client.get("/r/valid-token").text
    assert "Email and LinkedIn" not in body
    assert "roll into the funnel above" not in body


def test_the_email_rates_are_not_diluted_by_linkedin(client, fake_sb):
    """60 email replies against 1,751 emails is 3.4%. The combined figure, 98
    against 2,571, reads 3.8% — LinkedIn replies lifting an email rate, which
    is the average of two unlike platforms rather than either one of them."""
    _report_with(fake_sb, _MIXED_TOTAL, {"linkedin": _MIXED_LINKEDIN})
    body = client.get("/r/valid-token").text
    assert "3.4%" in body
    # And LinkedIn's own rate is quoted against connection requests, not
    # emails: (38 replies + 2 opt-outs) / 820 requests. Reply rate counts
    # unsubscribes as responses everywhere in this tool.
    assert "4.9%" in body


def test_a_client_with_no_linkedin_gets_no_card(client, fake_sb):
    """A card of zeros implies we tried and nothing happened."""
    _report_with(fake_sb, {"sent": 900, "replied": 30, "meeting_booked": 3}, {})
    body = client.get("/r/valid-token").text
    assert "Connection requests sent" not in body
    assert "900" in body


def test_a_linkedin_only_client_is_not_told_there_was_no_activity(client, fake_sb):
    """The "nothing happened" gate counts both platforms. Only the email half
    of the page is skipped."""
    _report_with(fake_sb, dict(_MIXED_LINKEDIN), {"linkedin": dict(_MIXED_LINKEDIN)})
    body = client.get("/r/valid-token").text
    assert "No campaign activity" not in body
    assert "Connection requests sent" in body
    assert "Your email funnel" not in body
    assert "Emails sent" not in body


def test_a_row_with_no_channel_still_reaches_the_client(client, fake_sb):
    """Split by subtraction, not by reading the email channel directly, so a
    channel we do not recognise cannot drop out of both halves of the page."""
    _report_with(fake_sb,
                 {"sent": 1000, "replied": 50, "meeting_booked": 4},
                 {"linkedin": {"sent": 200, "replied": 10, "meeting_booked": 1}})
    body = client.get("/r/valid-token").text
    assert "800" in body               # 1000 - 200, including any unchannelled


def test_a_report_published_before_the_split_still_separates_linkedin(client, fake_sb):
    """by_channel has been in every snapshot from the start, so this applies
    to reports already sent rather than only to new ones."""
    _report_with(fake_sb, _MIXED_TOTAL, {"linkedin": _MIXED_LINKEDIN,
                                         "email": {"sent": 1751}})
    body = client.get("/r/valid-token").text
    assert "1,751" in body
    assert "Connection requests sent" in body


def test_a_snapshot_with_no_channels_at_all_renders(client, fake_sb):
    """Nothing to subtract: the totals stand as the email figures."""
    _report_with(fake_sb, {"sent": 900, "replied": 30}, {})
    r = client.get("/r/valid-token")
    assert r.status_code == 200
    assert "900" in r.text


# ── Linking a client to their Performance Tracker sheet ──────────────────────

# A real tracker opens with a banner naming the client and puts the column
# names on row 2, so the fixture is a grid with that shape rather than rows
# already keyed by a header we assumed.
TRACKER_HEADERS = ["Date/Time", "Lead", "Channel", "Headcount", "Job role",
                   "Industry", "Location", "ICP rating", "Category"]

TRACKER_GRID = [
    ["Acme"] + [""] * (len(TRACKER_HEADERS) - 1),
    list(TRACKER_HEADERS),
    ["9/23/2025", "a.com", "Smartlead", "51-200", "Head of Growth",
     "Marketing", "United Kingdom", "8", "Lead"],
    ["2/18/2026", "b.com", "LinkedIn", "11-50", "Founder", "SaaS",
     "Ireland", "-", "Lead"],
    ["3/2/2026", "c.com", "Manual", "1-10", "Managing Director", "PR",
     "United Kingdom", "10", "Interested"],
]

SHEET_LINK = ("https://docs.google.com/spreadsheets/d/"
              "1OUnow7VQXJSTgzuJk_WxgkDJDRWE_uuQiU3TXPv5I9E/edit?gid=7#gid=7")
SHEET_KEY = "1OUnow7VQXJSTgzuJk_WxgkDJDRWE_uuQiU3TXPv5I9E"


def _tracker_row(**over):
    row = {"id": "t1", "client_id": CLIENT, "sheet_id": SHEET_KEY,
           "sheet_url": SHEET_LINK, "tab_title": "All leads",
           "linked_by": "Dylan", "last_error": "",
           "last_checked_at": "2026-10-07T09:00:00+00:00",
           "created_at": None, "updated_at": None}
    row.update(over)
    return row


def _tracker_routes(fake_sb, monkeypatch, linked=None, grid=None, headers=None,
                    reader=None):
    from app.utils import google_sheets

    _detail_routes(fake_sb)
    fake_sb.route("GET", "pulse_client_trackers",
                  lambda call: FakeResponse(200, [linked] if linked else []))
    fake_sb.route("POST", "pulse_client_trackers",
                  lambda call: FakeResponse(201, [_tracker_row()]))
    fake_sb.route("PATCH", "pulse_client_trackers", lambda call: FakeResponse(204))
    fake_sb.route("DELETE", "pulse_client_trackers", lambda call: FakeResponse(204))

    if grid is None:
        grid = (list(TRACKER_GRID) if headers is None
                else [["Acme"] + [""] * (len(headers) - 1), list(headers)])
    monkeypatch.setattr(google_sheets, "is_configured", lambda: True)
    monkeypatch.setattr(google_sheets, "read_tracker_values",
                        reader or (lambda sid, tab="All leads": grid))
    return fake_sb


def test_a_client_with_no_tracker_is_offered_the_form(client, fake_sb, monkeypatch):
    _tracker_routes(fake_sb, monkeypatch)
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/tracker").text
    assert 'name="sheet_url"' in body
    assert "Show lead quality" not in body, "nothing to show until a sheet is linked"


def test_the_client_page_does_not_touch_the_sheet(client, fake_sb, monkeypatch):
    """Reading a few thousand rows out of Google on every page view is the
    cost the A/B panel is already behind a button to avoid."""
    from app.utils import google_sheets

    def explode(*a, **k):
        raise AssertionError("the page must not read the sheet")

    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    monkeypatch.setattr(google_sheets, "read_tracker_values", explode)

    assert client.get(f"/outbound-pulse/clients/{CLIENT}").status_code == 200


def test_linking_stores_the_key_and_keeps_the_pasted_url(client, fake_sb, monkeypatch):
    saved = {}
    _tracker_routes(fake_sb, monkeypatch)
    fake_sb.route("POST", "pulse_client_trackers",
                  lambda call: (saved.update(call.get("json") or {}),
                                FakeResponse(201, [_tracker_row()]))[1])

    client.post(f"/api/outbound-pulse/clients/{CLIENT}/tracker",
                data={"sheet_url": SHEET_LINK})
    assert saved["sheet_id"] == SHEET_KEY, "the #gid must not be part of the key"
    assert saved["sheet_url"] == SHEET_LINK
    assert saved["tab_title"] == "All leads"


def test_a_link_that_is_not_a_sheet_is_refused_before_it_is_stored(client, fake_sb, monkeypatch):
    """extract_sheet_id hands back whatever it was given when the URL does not
    match, so without this a Notion page becomes a sheet id that 404s on every
    later read with no clue why."""
    _tracker_routes(fake_sb, monkeypatch)
    body = client.post(f"/api/outbound-pulse/clients/{CLIENT}/tracker",
                       data={"sheet_url": "https://notion.so/a-page"}).text
    assert "does not look like a Google Sheets link" in body
    assert not fake_sb.calls_to("POST", "pulse_client_trackers")


def test_the_locked_raw_leads_tab_is_refused_by_name(client, fake_sb, monkeypatch):
    _tracker_routes(fake_sb, monkeypatch)
    body = client.post(f"/api/outbound-pulse/clients/{CLIENT}/tracker",
                       data={"sheet_url": SHEET_LINK, "tab_title": "Raw Leads"}).text
    assert "locked" in body
    assert not fake_sb.calls_to("POST", "pulse_client_trackers")


def test_a_sheet_is_read_before_it_is_stored(client, fake_sb, monkeypatch):
    """A wrong link should say so while somebody is looking at it, not fail
    silently the first time a report is published."""
    from app.utils import google_sheets

    def missing(sid, tab="All leads"):
        raise google_sheets.TrackerUnavailable(
            f"That spreadsheet has no tab called “{tab}”.")

    _tracker_routes(fake_sb, monkeypatch)
    monkeypatch.setattr(google_sheets, "read_tracker_values", missing)

    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/tracker",
                    data={"sheet_url": SHEET_LINK, "tab_title": "Leeds"})
    assert r.status_code == 200, "a validation failure re-renders, it does not 4xx"
    assert "no tab called" in r.text
    assert not fake_sb.calls_to("POST", "pulse_client_trackers")


def test_a_sheet_with_no_date_column_is_refused(client, fake_sb, monkeypatch):
    """Without a date nothing can be tied to a report period, and a period
    figure built from undated rows would be fiction."""
    _tracker_routes(fake_sb, monkeypatch,
                    headers=["Lead", "Industry", "ICP rating"])
    body = client.post(f"/api/outbound-pulse/clients/{CLIENT}/tracker",
                       data={"sheet_url": SHEET_LINK}).text
    assert "no date column" in body
    assert not fake_sb.calls_to("POST", "pulse_client_trackers")


def test_unlinking_deletes_the_row_rather_than_blanking_it(client, fake_sb, monkeypatch):
    """No row means never linked, which is a different state from a linked
    sheet that cannot be read."""
    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    client.delete(f"/api/outbound-pulse/clients/{CLIENT}/tracker")
    assert fake_sb.calls_to("DELETE", "pulse_client_trackers")
    assert not fake_sb.calls_to("POST", "pulse_client_trackers")


def test_every_tracker_query_is_scoped_to_the_agency(client, fake_sb, monkeypatch):
    from tests.conftest import param_values

    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    client.get(f"/api/outbound-pulse/clients/{CLIENT}/tracker")
    client.delete(f"/api/outbound-pulse/clients/{CLIENT}/tracker")
    for method in ("GET", "DELETE"):
        for call in fake_sb.calls_to(method, "pulse_client_trackers"):
            assert param_values(call, "agency_id"), f"{method} is unscoped"


# ── The breakdown panel ──────────────────────────────────────────────────────

def test_the_panel_ships_its_rows_for_filtering_in_the_browser(client, fake_sb, monkeypatch):
    """One sheet read gives us the whole table; a round trip per filter click
    would be a Google request per click for information we already hold."""
    import json
    import re

    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    body = client.get(
        f"/api/outbound-pulse/clients/{CLIENT}/tracker/leads?range=all").text
    rows = json.loads(re.search(r"data-trk-rows>(.*?)</script>", body, re.S).group(1))
    assert len(rows) == 3
    assert {r["seniority"] for r in rows} == {"head", "founder", "director"}


def test_the_panel_states_how_many_were_rated(client, fake_sb, monkeypatch):
    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    body = client.get(
        f"/api/outbound-pulse/clients/{CLIENT}/tracker/leads?range=all").text
    assert "data-trk-rated" in body


def test_an_unreadable_sheet_degrades_the_panel_not_the_page(client, fake_sb, monkeypatch):
    from app.utils import google_sheets

    def denied(sid, tab="All leads"):
        raise google_sheets.TrackerUnavailable(
            "We do not have access to that sheet. Share it with bot@x.com.")

    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    monkeypatch.setattr(google_sheets, "read_tracker_values", denied)

    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/tracker/leads")
    assert r.status_code == 200
    assert "Share it with bot@x.com" in r.text


def test_a_failed_read_is_remembered_against_the_sheet(client, fake_sb, monkeypatch):
    """So the panel can say when the sheet stopped being readable instead of
    failing afresh every time somebody opens it."""
    from app.utils import google_sheets

    saved = {}
    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    fake_sb.route("PATCH", "pulse_client_trackers",
                  lambda call: (saved.update(call.get("json") or {}),
                                FakeResponse(204))[1])
    monkeypatch.setattr(google_sheets, "read_tracker_values",
                        lambda sid, tab="All leads": (_ for _ in ()).throw(
                            google_sheets.TrackerUnavailable("gone")))

    client.get(f"/api/outbound-pulse/clients/{CLIENT}/tracker/leads")
    assert saved.get("last_error") == "gone"


def test_a_good_read_clears_the_remembered_error(client, fake_sb, monkeypatch):
    saved = {}
    _tracker_routes(fake_sb, monkeypatch,
                    linked=_tracker_row(last_error="it was broken"))
    fake_sb.route("PATCH", "pulse_client_trackers",
                  lambda call: (saved.update(call.get("json") or {}),
                                FakeResponse(204))[1])

    client.get(f"/api/outbound-pulse/clients/{CLIENT}/tracker/leads")
    assert saved.get("last_error") == ""


def test_the_panel_defaults_to_the_range_the_page_is_showing(client, fake_sb, monkeypatch):
    """So the numbers agree with the funnel above them. Every row still ships,
    so widening to all time is a click rather than another sheet read."""
    import json
    import re

    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/tracker/leads"
                      "?range=custom&date_from=2026-02-01&date_to=2026-02-28").text
    rows = json.loads(re.search(r"data-trk-rows>(.*?)</script>", body, re.S).group(1))
    assert len(rows) == 3, "all three rows still ship"
    assert "2026-02-01" in body and "2026-02-28" in body


def test_the_ordering_tables_travel_with_the_rows(client, fake_sb, monkeypatch):
    """The browser recounts a filtered view, and must not hold its own copy of
    a business rule that could drift from Python's."""
    import json
    import re

    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    body = client.get(
        f"/api/outbound-pulse/clients/{CLIENT}/tracker/leads?range=all").text
    config = json.loads(re.search(r"data-trk-config>(.*?)</script>", body, re.S).group(1))
    assert [k for k, _ in config["seniority"]][0] == "founder"
    assert [k for k, _ in config["bands"]][-1] == "unrated"


def test_a_tab_with_no_leads_says_so_rather_than_drawing_empty_charts(client, fake_sb, monkeypatch):
    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row(),
                    grid=[["Acme"] + [""] * 8, list(TRACKER_HEADERS)])
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/tracker/leads").text
    assert "no leads on it yet" in body
    assert "data-trk-bars" not in body


def test_the_panel_refuses_to_guess_when_the_tracker_is_gone(client, fake_sb, monkeypatch):
    _tracker_routes(fake_sb, monkeypatch, linked=None)
    body = client.get(f"/api/outbound-pulse/clients/{CLIENT}/tracker/leads").text
    assert "No Performance Tracker sheet is linked" in body


def test_the_client_page_carries_the_lead_quality_panel(client, fake_sb, monkeypatch):
    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    body = client.get(f"/outbound-pulse/clients/{CLIENT}").text
    assert "Lead quality" in body
    assert f"/api/outbound-pulse/clients/{CLIENT}/tracker" in body


def test_filter_controls_that_hide_carry_their_own_hidden_rule(client, fake_sb, monkeypatch):
    """An author-origin `display` beats the browser's [hidden]. This module has
    shipped controls that could not be hidden twice already."""
    _tracker_routes(fake_sb, monkeypatch, linked=_tracker_row())
    body = client.get(f"/outbound-pulse/clients/{CLIENT}").text
    assert ".trk-link-form[hidden]" in body
    assert ".trk-linked-row[hidden]" in body


def test_the_migration_keeps_one_tracker_per_client():
    import pathlib

    sql = (pathlib.Path(__file__).resolve().parents[1] / "migrations"
           / "outbound_pulse_tracker.sql").read_text(encoding="utf-8")
    assert "pulse_client_trackers_unique" in sql
    assert "(agency_id, client_id)" in sql
    assert "ON DELETE CASCADE" in sql
    assert sql.rstrip().endswith("COMMIT;")


# ── Lead quality on the client's own report ──────────────────────────────────

def _publish_routes(fake_sb, monkeypatch, linked=True, grid=None, reader=None):
    """A draft report ready to publish, with a tracker behind it."""
    from app.utils import google_sheets
    from app.utils.pulse import smartlead

    seen = {}
    _detail_routes(fake_sb)
    monkeypatch.setattr(smartlead, "is_configured", lambda: False)
    fake_sb.route("GET", "pulse_copy_variants", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "campaigns", lambda call: FakeResponse(200, []))
    fake_sb.route("GET", "pulse_client_trackers",
                  lambda call: FakeResponse(200, [_tracker_row()] if linked else []))
    fake_sb.route("PATCH", "pulse_client_trackers", lambda call: FakeResponse(204))
    fake_sb.route("GET", "pulse_reports", lambda call: FakeResponse(200, [
        {"id": "r1", "client_id": CLIENT, "period_start": "2026-02-01",
         "period_end": "2026-03-31", "title": "", "body": "", "snapshot": None,
         "status": "draft", "published_at": None, "created_by": "",
         "created_at": None, "updated_at": None},
    ]))
    fake_sb.route("PATCH", "pulse_reports",
                  lambda call: (seen.update(call.get("json") or {}),
                                FakeResponse(204))[1])
    monkeypatch.setattr(google_sheets, "is_configured", lambda: True)
    monkeypatch.setattr(google_sheets, "read_tracker_values",
                        reader or (lambda sid, tab="All leads":
                                   TRACKER_GRID if grid is None else grid))
    return seen


def test_publishing_freezes_the_lead_quality_block(client, fake_sb, monkeypatch):
    seen = _publish_routes(fake_sb, monkeypatch)
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    icp = seen["snapshot"]["icp"]
    # Feb-Mar covers two of the three fixture leads: one ungraded, one at 10.
    assert icp["total"] == 2
    assert icp["rated"] == 1
    assert icp["average"] == 10.0


def test_the_frozen_block_carries_an_all_time_figure(client, fake_sb, monkeypatch):
    """A month of outbound is a handful of graded leads, so the period average
    on its own is noise."""
    seen = _publish_routes(fake_sb, monkeypatch)
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    all_time = seen["snapshot"]["icp"]["all_time"]
    assert all_time["total"] == 3
    assert all_time["rated"] == 2
    assert all_time["average"] == 9.0               # (8 + 10) / 2


def test_the_frozen_block_is_json_serialisable(client, fake_sb, monkeypatch):
    """A date object in a snapshot is what previously made Publish silently do
    nothing at all."""
    import json

    seen = _publish_routes(fake_sb, monkeypatch)
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    json.dumps(seen["snapshot"])


def test_a_report_keeps_job_titles_internal(client, fake_sb, monkeypatch):
    """The client's report wants the shape of who we reached, not a list of
    individual prospects."""
    seen = _publish_routes(fake_sb, monkeypatch)
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert "roles" not in seen["snapshot"]["icp"]


def test_publishing_survives_an_unreadable_tracker(client, fake_sb, monkeypatch):
    """A revoked share must not cost an account manager their write-up."""
    from app.utils import google_sheets

    def gone(sid, tab="All leads"):
        raise google_sheets.TrackerUnavailable("no access")

    seen = _publish_routes(fake_sb, monkeypatch, reader=gone)
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert r.status_code == 200
    assert seen["status"] == "published"
    assert seen["snapshot"]["icp"] == {}


def test_a_client_with_no_tracker_publishes_without_the_block(client, fake_sb, monkeypatch):
    seen = _publish_routes(fake_sb, monkeypatch, linked=False)
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert seen["snapshot"]["icp"] == {}


def test_a_period_with_no_leads_gets_no_block(client, fake_sb, monkeypatch):
    """Rather than a block of zeros, which reads as a campaign that failed."""
    seen = _publish_routes(fake_sb, monkeypatch, grid=[
        ["Acme"] + [""] * 8, list(TRACKER_HEADERS),
        ["9/23/2025", "a.com", "Smartlead", "51-200", "Head of Growth",
         "Marketing", "United Kingdom", "8", "Lead"],
    ])
    client.post(f"/api/outbound-pulse/clients/{CLIENT}/reports/r1/publish")
    assert seen["snapshot"]["icp"] == {}


def _icp_report(fake_sb, icp):
    _portal_routes(fake_sb, reports=[{
        "id": "sep", "client_id": CLIENT, "period_start": "2026-09-01",
        "period_end": "2026-09-30", "title": "", "body": "", "status": "published",
        "published_at": "2026-10-01T09:00:00+00:00", "created_by": "",
        "created_at": "2026-10-01T09:00:00+00:00", "updated_at": None,
        "snapshot": {"version": 1, "counts": {"sent": 900, "replied": 30},
                     "by_channel": {}, "trend": {"unit": "day", "buckets": []},
                     "ab": [], "icp": icp},
    }])


_ICP_BLOCK = {
    "version": 1, "total": 45, "rated": 30, "unrated": 15,
    "average": 8.0, "median": 8.0, "top_rated": 21, "top_rated_from": 8,
    "bands": [{"key": "9_10", "label": "ICP 9–10", "count": 12, "pct": 26.7},
              {"key": "7_8", "label": "ICP 7–8", "count": 13, "pct": 28.9},
              {"key": "5_6", "label": "ICP 5–6", "count": 3, "pct": 6.7},
              {"key": "1_4", "label": "ICP 1–4", "count": 2, "pct": 4.4},
              {"key": "unrated", "label": "Not yet rated", "count": 15, "pct": 33.3}],
    "headcount": [{"key": "1-10", "label": "1-10", "count": 20, "pct": 44.4},
                  {"key": "11-50", "label": "11-50", "count": 25, "pct": 55.6}],
    "industry":  [{"key": "Marketing", "label": "Marketing", "count": 45, "pct": 100.0}],
    "seniority": [{"key": "founder", "label": "Founder / Owner", "count": 18, "pct": 40.0},
                  {"key": "director", "label": "Director", "count": 27, "pct": 60.0}],
    "location":  [{"key": "United Kingdom", "label": "United Kingdom",
                   "count": 45, "pct": 100.0}],
    "all_time": {"total": 120, "rated": 95, "average": 7.4},
}


def test_the_report_shows_lead_quality(client, fake_sb):
    _icp_report(fake_sb, _ICP_BLOCK)
    body = client.get("/r/valid-token").text
    assert "Who we reached" in body
    assert "Leads assessed" in body
    assert "Seniority reached" in body


def test_the_report_always_states_the_grading_coverage(client, fake_sb):
    """An average over two thirds of the leads must not read as an average
    over all of them."""
    _icp_report(fake_sb, _ICP_BLOCK)
    body = client.get("/r/valid-token").text
    assert "30 of 45 graded" in body
    assert "30 of 45 have been graded so far" in body


def test_ungraded_leads_are_shown_as_their_own_band(client, fake_sb):
    _icp_report(fake_sb, _ICP_BLOCK)
    body = client.get("/r/valid-token").text
    assert "Not yet rated" in body
    assert "is-muted" in body, "absence must not be styled as a poor grade"


def test_the_all_time_average_sits_beside_the_period_one(client, fake_sb):
    _icp_report(fake_sb, _ICP_BLOCK)
    body = client.get("/r/valid-token").text
    assert "Average to date" in body
    assert "7.4" in body
    assert "across all 120 leads" in body


def test_a_single_value_breakdown_becomes_a_sentence(client, fake_sb):
    """One full-width bar at 100% says nothing."""
    _icp_report(fake_sb, _ICP_BLOCK)
    body = client.get("/r/valid-token").text
    assert "Every lead:" in body
    assert "United Kingdom" in body


def test_a_period_where_nothing_was_graded_shows_no_average(client, fake_sb):
    """0.0 average ICP is the single worst thing this section could print."""
    block = dict(_ICP_BLOCK, rated=0, average=None, top_rated=0,
                 all_time={"total": 120, "rated": 0, "average": None})
    _icp_report(fake_sb, block)
    body = client.get("/r/valid-token").text
    assert "Who we reached" in body           # the volume breakdowns still show
    assert "Average ICP rating" not in body
    assert "None have been graded yet" in body


def test_a_report_published_before_lead_quality_existed_still_renders(client, fake_sb):
    """The compatibility story for every report already sent."""
    _portal_routes(fake_sb)
    r = client.get("/r/valid-token")
    assert r.status_code == 200
    assert "Who we reached" not in r.text


def test_an_empty_block_renders_no_card(client, fake_sb):
    _icp_report(fake_sb, {})
    assert "Who we reached" not in client.get("/r/valid-token").text


def test_the_portal_with_no_report_at_all_still_renders(client, fake_sb):
    """Exercises the no-report fallback dict, which has to carry every key the
    template reads or it 500s."""
    _portal_routes(fake_sb, reports=[])
    assert client.get("/r/valid-token").status_code == 200


def test_the_lead_quality_card_reaches_a_linkedin_only_client(client, fake_sb):
    """It sits outside the email block: a client run only on LinkedIn still has
    leads and still wants to know how well they matched."""
    _portal_routes(fake_sb, reports=[{
        "id": "sep", "client_id": CLIENT, "period_start": "2026-09-01",
        "period_end": "2026-09-30", "title": "", "body": "", "status": "published",
        "published_at": "2026-10-01T09:00:00+00:00", "created_by": "",
        "created_at": "2026-10-01T09:00:00+00:00", "updated_at": None,
        "snapshot": {"version": 1,
                     "counts": {"sent": 500, "replied": 20},
                     "by_channel": {"linkedin": {"sent": 500, "replied": 20}},
                     "trend": {"unit": "day", "buckets": []},
                     "ab": [], "icp": _ICP_BLOCK},
    }])
    body = client.get("/r/valid-token").text
    assert "Your email funnel" not in body
    assert "Who we reached" in body


def test_the_portal_still_reads_only_its_snapshot(client, fake_sb, monkeypatch):
    """A published report must never reach out to Google when a client opens
    it — the figures were frozen at publish and must not move."""
    from app.utils import google_sheets

    def explode(*a, **k):
        raise AssertionError("the portal must not read the sheet")

    monkeypatch.setattr(google_sheets, "read_tracker_values", explode)
    _icp_report(fake_sb, _ICP_BLOCK)
    assert client.get("/r/valid-token").status_code == 200
