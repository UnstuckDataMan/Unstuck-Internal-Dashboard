"""Offline tests for the Copy Bank endpoints: default-chaser fallback and the
A/B-winner tracing that links campaigns back to the copy they used."""
from __future__ import annotations

from conftest import FakeResponse, param_values

import app.routers.copy_bank as cb
import app.utils.google_sheets as gs


# ── _parse_variant ─────────────────────────────────────────────────────────────

def test_parse_variant_maps_to_zero_based_indices():
    assert cb._parse_variant("S2/B1") == (1, 0)
    assert cb._parse_variant(" S1 / B3 ") == (0, 2)
    assert cb._parse_variant("bogus") == (None, None)
    assert cb._parse_variant("") == (None, None)


# ── Default-chaser fallback in copy_bank_template ──────────────────────────────

def _cb_templates_router(profiles_content, per_key_content):
    """Route GET copy_bank_templates to profiles row vs per-key row by `key`."""
    def handler(call):
        key = param_values(call, "key")[0]
        if key == "eq.__cb_profiles__":
            return FakeResponse(200, [{"content": profiles_content}])
        return FakeResponse(200, [{"content": per_key_content}])
    return handler


def test_chaser_falls_back_to_profile_default(fake_sb, client):
    fake_sb.route("GET", "copy_bank_templates", _cb_templates_router(
        profiles_content=[{"client_id": "c1", "name": "Acme",
                           "defaultChaser": {"steps": [{"body": "Just circling back!"}]}}],
        per_key_content={"chaser": {"steps": []}},   # industry has no chaser of its own
    ))
    resp = client.get("/api/copy-bank/templates/c1/US/SaaS", params={"channel": "chaser"})
    assert resp.status_code == 200
    assert resp.json()["bodies"] == ["Just circling back!"]


def test_chaser_prefers_its_own_over_default(fake_sb, client):
    fake_sb.route("GET", "copy_bank_templates", _cb_templates_router(
        profiles_content=[{"client_id": "c1",
                           "defaultChaser": {"steps": [{"body": "DEFAULT"}]}}],
        per_key_content={"chaser": {"steps": [{"body": "Industry-specific chaser"}]}},
    ))
    resp = client.get("/api/copy-bank/templates/c1/US/SaaS", params={"channel": "chaser"})
    assert resp.json()["bodies"] == ["Industry-specific chaser"]


def test_chaser_default_absent_returns_empty(fake_sb, client):
    fake_sb.route("GET", "copy_bank_templates", _cb_templates_router(
        profiles_content=[{"client_id": "c1"}],       # no defaultChaser
        per_key_content={"chaser": {"steps": []}},
    ))
    resp = client.get("/api/copy-bank/templates/c1/US/SaaS", params={"channel": "chaser"})
    assert resp.json()["bodies"] == []


# ── A/B winner tracing ─────────────────────────────────────────────────────────

def test_ab_winner_aggregates_and_maps_indices(fake_sb, client, monkeypatch):
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, [
        {"campaign_name": "June Push", "sheet_id": "s1"},
        {"campaign_name": "July Push", "sheet_id": "s2"},
    ]))
    ab = {
        "s1": [
            {"variant": "S1/B1", "total": 20, "lead": 1, "interested": 0, "reply": 1, "unsubscribe": 0},
            {"variant": "S2/B1", "total": 22, "lead": 4, "interested": 1, "reply": 0, "unsubscribe": 0},
        ],
        "s2": [
            {"variant": "S2/B1", "total": 20, "lead": 5, "interested": 2, "reply": 1, "unsubscribe": 0},
        ],
    }
    monkeypatch.setattr(gs, "read_ab_stats", lambda sid: ab[sid])

    resp = client.get("/api/copy-bank/ab-winner",
                      params={"client_id": "c1", "territory": "US", "industry": "SaaS"})
    assert resp.status_code == 200
    d = resp.json()
    assert d["has_data"] is True
    # S2/B1 aggregates to lead=9, interested=3 → positive 12 over 42 sends
    assert d["winner"]["variant"] == "S2/B1"
    assert d["winner"]["subject_idx"] == 1 and d["winner"]["body_idx"] == 0
    assert d["winner"]["positive"] == 12 and d["winner"]["total"] == 42
    assert d["winner"]["rate"] == round(12 / 42 * 100, 1)
    assert set(d["campaigns"]) == {"June Push", "July Push"}

    # The campaign lookup filtered by the exact copy source
    call = fake_sb.calls_to("GET", "campaigns")[0]
    assert "eq.US" in param_values(call, "copy_territory")
    assert "eq.SaaS" in param_values(call, "copy_industry")
    assert "eq.c1" in param_values(call, "client_id")


def test_ab_winner_tie_flagged(fake_sb, client, monkeypatch):
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, [
        {"campaign_name": "X", "sheet_id": "s1"},
    ]))
    monkeypatch.setattr(gs, "read_ab_stats", lambda sid: [
        {"variant": "S1/B1", "total": 10, "lead": 2, "interested": 0, "reply": 0, "unsubscribe": 0},
        {"variant": "S2/B1", "total": 10, "lead": 2, "interested": 0, "reply": 0, "unsubscribe": 0},
    ])
    d = client.get("/api/copy-bank/ab-winner",
                   params={"client_id": "c1", "territory": "US", "industry": "SaaS"}).json()
    assert d["has_data"] is True and d["is_tie"] is True and d["winner"] is None


def test_ab_winner_no_matching_campaigns(fake_sb, client):
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, []))
    d = client.get("/api/copy-bank/ab-winner",
                   params={"client_id": "c1", "territory": "US", "industry": "SaaS"}).json()
    assert d["has_data"] is False


def test_ab_winner_columns_not_migrated(fake_sb, client):
    fake_sb.route("GET", "campaigns",
                  lambda c: FakeResponse(400, [], text="column copy_territory does not exist"))
    d = client.get("/api/copy-bank/ab-winner",
                   params={"client_id": "c1", "territory": "US", "industry": "SaaS"}).json()
    assert d["has_data"] is False
