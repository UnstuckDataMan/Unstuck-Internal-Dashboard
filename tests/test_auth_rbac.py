"""Offline tests for the auth gate + RBAC layer (app/auth.py + middleware).

Pure-logic tests run without the app. Gate tests use the shared TestClient and
re-enable auth per-test (the suite default is AUTH_DISABLED=1, set in conftest)
by deleting the flag and monkeypatching the session-user resolver — no real
Google or Supabase calls.
"""
from __future__ import annotations

import pytest

from app import auth


# ── Pure logic: effective_tools ─────────────────────────────────────────────────

def test_effective_tools_role_union_add_minus_remove():
    user = {"role": "viewer", "tools_add": ["dnc"], "tools_remove": ["campaigns"]}
    tools = auth.effective_tools(user)
    assert "dnc" in tools                 # added on top of role
    assert "campaigns" not in tools       # revoked from role
    assert {"bd_targeting", "launch_checker"} <= tools   # remaining role defaults


def test_effective_tools_admin_gets_everything():
    assert auth.effective_tools({"role": "admin"}) == set(auth.TOOLS)


def test_effective_tools_clamps_unknown_keys():
    tools = auth.effective_tools({"role": "viewer", "tools_add": ["not_a_real_tool"]})
    assert "not_a_real_tool" not in tools


def test_effective_tools_unknown_role_is_empty():
    assert auth.effective_tools({"role": "nope"}) == set()


# ── Pure logic: tool_for_path (longest-prefix) ───────────────────────────────────

@pytest.mark.parametrize("path, tool", [
    ("/api/admin/users",                    "admin_users"),
    ("/admin/users",                        "admin_users"),
    ("/api/admin/merge-unstuck-profiles",   "copy_bank"),   # copy-bank maintenance, not user admin
    ("/api/dnc/scrub",                      "dnc"),
    ("/dnc-removal",                        "dnc"),
    ("/api/merge/generate",                 "mail_merge"),
    ("/api/copy-bank/approve",              "copy_bank"),
    ("/api/export/google/callback",         "copy_bank"),
    ("/api/campaigns/reset-stats",          "campaigns"),
    ("/client-profiles",                    "client_profiles"),
    ("/api/client-profiles",                "client_profiles"),
    ("/api/client-profiles/backfill",       "client_profiles"),
    ("/api/normalize/download/abc",         "city"),
    ("/city-state",                         "city"),
    ("/api/gender",                         "gender"),
])
def test_tool_for_path(path, tool):
    assert auth.tool_for_path(path) == tool


def test_tool_for_path_unmapped_is_none():
    assert auth.tool_for_path("/") is None
    assert auth.tool_for_path("/logout") is None
    assert auth.tool_for_path("/healthz") is None


# ── Pure logic: public allowlist ─────────────────────────────────────────────────

def test_is_public_path():
    for p in ("/login", "/logout", "/healthz", "/auth/login", "/auth/callback",
              "/static/img/logo.png", "/favicon.ico"):
        assert auth.is_public_path(p), p
    for p in ("/", "/api/dnc/clients", "/dnc-removal", "/admin/users"):
        assert not auth.is_public_path(p), p


# ── Gate behaviour via TestClient ────────────────────────────────────────────────

@pytest.fixture
def auth_on(monkeypatch):
    """Enable the auth gate for a test (suite default is disabled)."""
    monkeypatch.delenv("AUTH_DISABLED", raising=False)
    return monkeypatch


def _login_as(monkeypatch, role, tools=None, email="user@unstuck-agency.com"):
    user = {
        "email": email,
        "name":  "Test User",
        "role":  role,
        "tools": tools if tools is not None else sorted(auth.effective_tools({"role": role})),
    }
    monkeypatch.setattr(auth, "get_session_user", lambda request: user)
    return user


def test_unauthenticated_browser_redirects_to_login(client, auth_on):
    r = client.get("/", follow_redirects=False)
    assert r.status_code == 303
    assert r.headers["location"] == "/login"


def test_unauthenticated_api_returns_401_with_hx_redirect(client, auth_on):
    r = client.get("/api/campaigns", follow_redirects=False)
    assert r.status_code == 401
    assert r.headers.get("HX-Redirect") == "/login"


def test_healthz_is_public(client, auth_on):
    assert client.get("/healthz").status_code == 200


def test_login_page_is_public(client, auth_on):
    r = client.get("/login")
    assert r.status_code == 200
    assert "Sign in with Google" in r.text


def test_non_admin_forbidden_on_reset_stats(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "sdr")   # sdr HAS the campaigns tool, but reset-stats is admin-only
    r = client.post("/api/campaigns/reset-stats", follow_redirects=False)
    assert r.status_code == 403


def test_non_admin_forbidden_on_purge(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "sdr")
    r = client.post("/api/dnc/contacted/purge", follow_redirects=False)
    assert r.status_code == 403


def test_non_admin_forbidden_on_campaign_delete(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "sdr")
    r = client.delete("/api/campaigns/some-id", follow_redirects=False)
    assert r.status_code == 403


def test_viewer_blocked_from_dnc_tool_by_middleware(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "viewer")   # viewer lacks the dnc tool entirely
    r = client.get("/api/dnc/clients", follow_redirects=False)
    assert r.status_code == 403


def test_viewer_blocked_from_client_profiles(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "viewer")   # viewer lacks the client_profiles tool
    r = client.get("/client-profiles", follow_redirects=False)
    assert r.status_code == 403


def test_sdr_allowed_into_client_profiles(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "sdr")      # staff bundle includes client_profiles
    r = client.get("/client-profiles", follow_redirects=False)
    assert r.status_code == 200


def test_non_admin_blocked_from_profiles_backfill(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "sdr")      # has the tool, but backfill is admin-only
    r = client.get("/api/client-profiles/backfill", follow_redirects=False)
    assert r.status_code == 403


def test_admin_allowed_into_admin_users(client, auth_on, monkeypatch, fake_sb):
    _login_as(monkeypatch, "admin")
    r = client.get("/api/admin/users", follow_redirects=False)
    assert r.status_code == 200


def test_non_admin_blocked_from_admin_users(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "reviewer")   # reviewer lacks admin_users tool
    r = client.get("/api/admin/users", follow_redirects=False)
    assert r.status_code == 403


def test_index_hides_cards_for_viewer(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "viewer")
    html = client.get("/").text
    assert "Launch Checker" in html     # viewer has launch_checker
    assert "DNC & Merger" not in html   # viewer lacks dnc / mail_merge
    assert "Client Profiles" not in html  # viewer lacks client_profiles
    assert "/admin/users" not in html   # not an admin → no admin link


def test_index_shows_admin_link_for_admin(client, auth_on, monkeypatch):
    _login_as(monkeypatch, "admin")
    html = client.get("/").text
    assert "/admin/users" in html
    assert "DNC & Merger" in html
    assert "Client Profiles" in html
