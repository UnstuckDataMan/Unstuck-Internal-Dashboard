"""
Google Workspace SSO login for the dashboard.

Reuses the OAuth2 client already configured for the Copy Bank export
(GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET / APP_BASE_URL) — no new credentials.

Flow:
    /login         → branded page with a "Sign in with Google" button
    /auth/login    → redirect to Google's consent screen (state in session)
    /auth/callback → verify the ID token, check the email is on the @domain and
                     present + active in app_users, then store the user in the
                     session and bounce to "/"
    /logout        → clear the session
"""
from __future__ import annotations

import logging
import os
import secrets

import requests as http_req
from fastapi import APIRouter, Request
from fastapi.responses import RedirectResponse

from app.deps import templates
from app import auth
from app.utils.supabase import SUPABASE_URL, sb_headers as _sb_headers, sb_configured as _sb_configured

# Google sometimes returns a superset of the requested scopes (it always adds
# "openid"); relax oauthlib so fetch_token doesn't raise on the difference.
os.environ.setdefault("OAUTHLIB_RELAX_TOKEN_SCOPE", "1")

log = logging.getLogger(__name__)
router = APIRouter()

SCOPES = ["openid", "email", "profile"]


# ── Config helpers ───────────────────────────────────────────────────────────

def _google_configured() -> bool:
    return bool(os.environ.get("GOOGLE_CLIENT_ID") and os.environ.get("GOOGLE_CLIENT_SECRET"))


def _client_config() -> dict:
    return {
        "web": {
            "client_id":     os.environ.get("GOOGLE_CLIENT_ID", ""),
            "client_secret": os.environ.get("GOOGLE_CLIENT_SECRET", ""),
            "auth_uri":      "https://accounts.google.com/o/oauth2/auth",
            "token_uri":     "https://oauth2.googleapis.com/token",
        }
    }


def _redirect_uri(request: Request) -> str:
    base = os.environ.get("APP_BASE_URL", "").rstrip("/") or str(request.base_url).rstrip("/")
    return f"{base}/auth/callback"


# ── app_users lookup ─────────────────────────────────────────────────────────

def lookup_user(email: str) -> dict | None:
    """Return the active app_users row for `email`, or None if absent/inactive."""
    if not _sb_configured():
        return None
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/app_users",
            headers=_sb_headers(),
            params={
                "email":  f"eq.{email}",
                "active": "is.true",
                "select": "email,name,role,tools_add,tools_remove,active",
                "limit":  "1",
            },
            timeout=10,
        )
        if r.status_code == 200 and r.json():
            return r.json()[0]
    except Exception as exc:  # noqa: BLE001 — login must fail closed, never 500
        log.error("app_users lookup failed for %s: %s", email, exc)
    return None


def _stamp_last_login(email: str) -> None:
    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/app_users",
            headers=_sb_headers(),
            params={"email": f"eq.{email}"},
            json={"last_login": "now()"},
            timeout=10,
        )
    except Exception:  # non-fatal: a missing timestamp must never block login
        pass


# ── Routes ─────────────────────────────────────────────────────────────────────

@router.get("/login")
async def login(request: Request, error: str = ""):
    # Already signed in → straight to the dashboard.
    if auth.get_session_user(request):
        return RedirectResponse("/", status_code=303)
    return templates.TemplateResponse("login.html", {
        "request":         request,
        "error":           error,
        "google_ready":    _google_configured(),
        "allowed_domain":  auth.allowed_domain(),
    })


@router.get("/auth/login")
async def auth_login(request: Request):
    if not _google_configured():
        return RedirectResponse("/login?error=not_configured", status_code=303)
    try:
        from google_auth_oauthlib.flow import Flow
    except ImportError:
        return RedirectResponse("/login?error=not_configured", status_code=303)

    state = secrets.token_urlsafe(24)
    request.session["oauth_state"] = state

    flow = Flow.from_client_config(_client_config(), scopes=SCOPES,
                                   redirect_uri=_redirect_uri(request))
    auth_url, _ = flow.authorization_url(
        access_type="online",
        include_granted_scopes="true",
        state=state,
        prompt="select_account",
        hd=auth.allowed_domain(),          # hint Google to the Workspace domain
    )
    return RedirectResponse(auth_url, status_code=303)


@router.get("/auth/callback")
async def auth_callback(request: Request, code: str = "", state: str = "", error: str = ""):
    if error or not code:
        return RedirectResponse("/login?error=cancelled", status_code=303)

    if not state or state != request.session.pop("oauth_state", None):
        return RedirectResponse("/login?error=bad_state", status_code=303)

    try:
        from google_auth_oauthlib.flow import Flow
        from google.oauth2 import id_token as google_id_token
        from google.auth.transport import requests as google_requests
    except ImportError:
        return RedirectResponse("/login?error=not_configured", status_code=303)

    try:
        flow = Flow.from_client_config(_client_config(), scopes=SCOPES,
                                       redirect_uri=_redirect_uri(request))
        flow.fetch_token(code=code)
        claims = google_id_token.verify_oauth2_token(
            flow.credentials.id_token,
            google_requests.Request(),
            os.environ.get("GOOGLE_CLIENT_ID"),
            clock_skew_in_seconds=10,
        )
    except Exception as exc:  # noqa: BLE001
        log.error("Google OAuth callback failed: %s", exc)
        return RedirectResponse("/login?error=oauth_failed", status_code=303)

    email = (claims.get("email") or "").strip().lower()
    domain = auth.allowed_domain()

    # Defence in depth: verified email on the right Workspace domain.
    if not claims.get("email_verified") or not email.endswith("@" + domain):
        log.warning("Login rejected (domain/verify): %s hd=%s", email, claims.get("hd"))
        return RedirectResponse("/login?error=denied", status_code=303)

    # The real gate: the email must be provisioned and active in app_users.
    record = lookup_user(email)
    if not record:
        log.warning("Login rejected (not provisioned): %s", email)
        return RedirectResponse("/login?error=not_provisioned", status_code=303)

    request.session["user"] = {
        "email": email,
        "name":  record.get("name") or claims.get("name") or email.split("@")[0],
        "role":  record.get("role") or "sdr",
        "tools": sorted(auth.effective_tools(record)),
    }
    _stamp_last_login(email)
    return RedirectResponse("/", status_code=303)


@router.get("/logout")
async def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)
