"""
Authentication & role-based access control for the internal dashboard.

Two enforcement layers, both server-authoritative:

  1. A default-deny request gate (`access_response`, wired as HTTP middleware in
     app/main.py) that requires a logged-in session for every path except an
     explicit public allowlist, then checks the request's *tool* against the
     user's effective tool set.
  2. Per-endpoint dependencies (`require_admin`, `require_role`) layered on top
     of the gate for data-loss / privileged operations.

The logged-in user lives in the signed session cookie as `request.session["user"]`:

    {"email": ..., "name": ..., "role": ..., "tools": [<tool keys>]}

`tools` is precomputed at login (see effective_tools) so the hot path never
touches Supabase. Role/override changes therefore take effect on next login.

The whole layer is bypassed when AUTH_DISABLED is truthy — that env flag is both
the offline-test switch (set in tests/conftest.py) and the production rollback
lever (flip it in the Render dashboard without a code deploy).
"""
from __future__ import annotations

import os
from typing import Optional

from fastapi import HTTPException, Request
from starlette.responses import HTMLResponse, JSONResponse, RedirectResponse, Response


# ── Tool & role registry ───────────────────────────────────────────────────────
# Tool keys are the unit of access. Each gates one card on the home screen and
# the routes behind it (see ROUTE_TOOL_MAP). Keep these in sync with index.html.

TOOLS: tuple[str, ...] = (
    "dnc",
    "mail_merge",
    "copy_bank",
    "reply_bank",
    "bd_targeting",
    "campaigns",
    "launch_checker",
    "targeting_checker",
    "gender",
    "city",
    "admin_users",          # user-management screen — admins only
)

# Path-prefix → tool. Longest matching prefix wins, so more specific entries
# (e.g. /api/admin/users) override broader ones. A path with no match resolves
# to None and is allowed for any logged-in user (e.g. "/", "/logout").
ROUTE_TOOL_MAP: tuple[tuple[str, str], ...] = (
    ("/dnc-removal",                      "dnc"),
    ("/api/dnc",                          "dnc"),
    ("/api/merge",                        "mail_merge"),
    ("/copy-bank",                        "copy_bank"),
    ("/api/copy-bank",                    "copy_bank"),
    ("/api/export/google",                "copy_bank"),
    ("/api/admin/merge-unstuck-profiles", "copy_bank"),   # copy-bank maintenance, not user admin
    ("/reply-bank",                       "reply_bank"),
    ("/bd-targeting",                     "bd_targeting"),
    ("/api/bd-targeting",                 "bd_targeting"),
    ("/api/campaigns",                    "campaigns"),
    ("/launch-checker",                   "launch_checker"),
    ("/api/launch-checker",               "launch_checker"),
    ("/targeting-checker",                "targeting_checker"),
    ("/gender",                           "gender"),
    ("/api/gender",                       "gender"),
    ("/city-state",                       "city"),
    ("/api/normalize",                    "city"),
    ("/admin/users",                      "admin_users"),
    ("/api/admin/users",                  "admin_users"),
)

_ALL_TOOLS = set(TOOLS)
# Everything except the user-management screen — the default bundle for staff.
_STAFF_TOOLS = _ALL_TOOLS - {"admin_users"}

# Default tool bundle per role. Starting points — fine-tune per person with the
# tools_add / tools_remove overrides in the /admin/users screen.
ROLE_TOOLS: dict[str, set[str]] = {
    "admin":    set(_ALL_TOOLS),
    "reviewer": set(_STAFF_TOOLS),
    "sdr":      set(_STAFF_TOOLS),
    "viewer":   {"campaigns", "bd_targeting", "launch_checker"},
}

# Roles allowed to approve/publish Copy Bank drafts.
COPY_APPROVER_ROLES = {"admin", "reviewer"}

# Paths reachable without a session.
_PUBLIC_EXACT = {"/healthz", "/login", "/logout", "/favicon.ico"}
_PUBLIC_PREFIXES = ("/auth/", "/static/")


# ── Config helpers ───────────────────────────────────────────────────────────

def _truthy(val: str) -> bool:
    return str(val).strip().lower() in {"1", "true", "yes", "on"}


def auth_enabled() -> bool:
    """Read live each request so tests (and the prod kill switch) can toggle it."""
    return not _truthy(os.environ.get("AUTH_DISABLED", ""))


def allowed_domain() -> str:
    return os.environ.get("ALLOWED_EMAIL_DOMAIN", "unstuck-agency.com").strip().lower()


# ── Tool / role resolution ─────────────────────────────────────────────────────

def tool_for_path(path: str) -> Optional[str]:
    """Resolve a request path to its tool key via longest-prefix match."""
    best_tool: Optional[str] = None
    best_len = -1
    for prefix, tool in ROUTE_TOOL_MAP:
        if (path == prefix or path.startswith(prefix + "/")) and len(prefix) > best_len:
            best_tool, best_len = tool, len(prefix)
    return best_tool


def effective_tools(user: dict) -> set[str]:
    """ROLE_TOOLS[role] ∪ tools_add − tools_remove, clamped to known tools."""
    tools = set(ROLE_TOOLS.get((user.get("role") or "").lower(), set()))
    tools |= set(user.get("tools_add") or [])
    tools -= set(user.get("tools_remove") or [])
    return tools & _ALL_TOOLS


def get_session_user(request: Request) -> Optional[dict]:
    """The logged-in user dict from the session, or None."""
    try:
        return request.session.get("user")
    except (AssertionError, AttributeError):
        # SessionMiddleware not installed (shouldn't happen) — treat as anon.
        return None


# ── Request classification ──────────────────────────────────────────────────────

def is_public_path(path: str) -> bool:
    return path in _PUBLIC_EXACT or path.startswith(_PUBLIC_PREFIXES)


def _wants_json(request: Request) -> bool:
    """API/HTMX callers get status codes; plain browser navigations get redirects."""
    return request.url.path.startswith("/api/") or "HX-Request" in request.headers


# ── Gate decision (called by the HTTP middleware in main.py) ────────────────────

def access_response(request: Request) -> Optional[Response]:
    """
    Return a Response to short-circuit the request (deny), or None to allow it.
    """
    if not auth_enabled():
        return None

    path = request.url.path
    if is_public_path(path):
        return None

    user = get_session_user(request)
    if not user:
        if _wants_json(request):
            return JSONResponse(
                {"error": "not_authenticated"},
                status_code=401,
                headers={"HX-Redirect": "/login"},
            )
        return RedirectResponse("/login", status_code=303)

    tool = tool_for_path(path)
    if tool and tool not in set(user.get("tools") or []):
        if _wants_json(request):
            return JSONResponse({"error": "forbidden", "tool": tool}, status_code=403)
        return HTMLResponse(_forbidden_html(tool), status_code=403)

    return None


def _forbidden_html(tool: str) -> str:
    return (
        "<!doctype html><html><head><meta charset='utf-8'><title>No access</title>"
        "<style>body{font-family:-apple-system,Segoe UI,Arial,sans-serif;background:#07041a;"
        "color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}"
        ".box{text-align:center;max-width:420px;padding:32px}"
        "a{color:#c4b5fd;text-decoration:none}</style></head>"
        "<body><div class='box'><h1 style='font-size:22px'>No access to this tool</h1>"
        "<p style='color:rgba(255,255,255,.6);line-height:1.6'>Your account isn't permitted to "
        f"use the <b>{tool}</b> tool. Ask an admin if you think this is a mistake.</p>"
        "<p style='margin-top:24px'><a href='/'>← Back to dashboard</a></p></div></body></html>"
    )


# ── Endpoint dependencies (privileged operations) ───────────────────────────────

def require_login(request: Request) -> dict:
    """FastAPI dependency: ensure a session user exists. No-op when auth disabled."""
    if not auth_enabled():
        return {"email": "dev@local", "name": "Dev", "role": "admin",
                "tools": sorted(_ALL_TOOLS)}
    user = get_session_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    return user


def require_admin(request: Request) -> dict:
    user = require_login(request)
    if auth_enabled() and (user.get("role") or "").lower() != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    return user


def require_role(roles: set[str] | tuple[str, ...]):
    """Factory: dependency that requires the user's role be in `roles`."""
    allowed = {r.lower() for r in roles}

    def _dep(request: Request) -> dict:
        user = require_login(request)
        if auth_enabled() and (user.get("role") or "").lower() not in allowed:
            raise HTTPException(status_code=403, detail="Insufficient role")
        return user

    return _dep
