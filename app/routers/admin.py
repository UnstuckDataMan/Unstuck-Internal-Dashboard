"""
Admin-only user management for the dashboard (app_users CRUD).

Every route here is gated by `require_admin`, so non-admins get 403 even if the
tool-level middleware were bypassed. When auth is disabled (offline tests / the
prod kill switch) require_admin is a no-op, matching the rest of the app.
"""
from __future__ import annotations

import json
import logging

import requests as http_req
from fastapi import APIRouter, Depends, Request
from fastapi.responses import JSONResponse

from app.deps import templates
from app import auth
from app.utils.supabase import SUPABASE_URL, sb_headers as _sb_headers, sb_configured as _sb_configured

log = logging.getLogger(__name__)
router = APIRouter(dependencies=[Depends(auth.require_admin)])

_ROLES = ("admin", "reviewer", "sdr", "viewer")
_SELECT = "email,name,role,tools_add,tools_remove,active,created_at,last_login"


def _not_configured() -> JSONResponse:
    return JSONResponse({"ok": False, "error": "Supabase not configured"}, status_code=503)


def _clean_tools(values) -> list[str]:
    """Keep only known tool keys; drop everything else."""
    if not isinstance(values, list):
        return []
    return [t for t in values if t in auth.TOOLS]


# ── Page ─────────────────────────────────────────────────────────────────────

@router.get("/admin/users")
async def admin_users_page(request: Request):
    return templates.TemplateResponse("admin_users.html", {
        "request":    request,
        "tools":      list(auth.TOOLS),
        "roles":      list(_ROLES),
        "role_tools": {r: sorted(ts) for r, ts in auth.ROLE_TOOLS.items()},
    })


# ── API ──────────────────────────────────────────────────────────────────────

@router.get("/api/admin/users")
async def list_users():
    if not _sb_configured():
        return _not_configured()
    r = http_req.get(
        f"{SUPABASE_URL}/rest/v1/app_users",
        headers=_sb_headers(),
        params={"select": _SELECT, "order": "role.asc,email.asc"},
        timeout=15,
    )
    if r.status_code != 200:
        return JSONResponse({"ok": False, "error": r.text[:300]}, status_code=502)
    return JSONResponse({"ok": True, "users": r.json()})


@router.post("/api/admin/users")
async def upsert_user(request: Request):
    if not _sb_configured():
        return _not_configured()
    body = await request.json()
    email = (body.get("email") or "").strip().lower()
    if not email or "@" not in email:
        return JSONResponse({"ok": False, "error": "Valid email required"}, status_code=400)

    role = (body.get("role") or "sdr").lower()
    if role not in _ROLES:
        return JSONResponse({"ok": False, "error": f"Unknown role: {role}"}, status_code=400)

    row = {
        "email":        email,
        "name":         (body.get("name") or "").strip() or None,
        "role":         role,
        "tools_add":    _clean_tools(body.get("tools_add")),
        "tools_remove": _clean_tools(body.get("tools_remove")),
        "active":       bool(body.get("active", True)),
    }
    # Upsert keyed on the email primary key (on_conflict is required for a
    # non-PK-default merge — see the project's "merge-duplicates" rule).
    r = http_req.post(
        f"{SUPABASE_URL}/rest/v1/app_users",
        headers=_sb_headers("resolution=merge-duplicates,return=representation"),
        params={"on_conflict": "email"},
        json=row,
        timeout=15,
    )
    if r.status_code not in (200, 201):
        return JSONResponse({"ok": False, "error": r.text[:300]}, status_code=502)
    return JSONResponse({"ok": True})


@router.delete("/api/admin/users")
async def delete_user(request: Request, email: str = ""):
    if not _sb_configured():
        return _not_configured()
    email = (email or "").strip().lower()
    if not email:
        return JSONResponse({"ok": False, "error": "email required"}, status_code=400)

    # Guard against locking everyone out: never delete the last active admin.
    me = (auth.get_session_user(request) or {}).get("email", "").lower()
    if email == me:
        return JSONResponse({"ok": False, "error": "You can't remove your own account."},
                            status_code=400)

    r = http_req.delete(
        f"{SUPABASE_URL}/rest/v1/app_users",
        headers=_sb_headers(),
        params={"email": f"eq.{email}"},
        timeout=15,
    )
    if r.status_code not in (200, 204):
        return JSONResponse({"ok": False, "error": r.text[:300]}, status_code=502)
    return JSONResponse({"ok": True})
