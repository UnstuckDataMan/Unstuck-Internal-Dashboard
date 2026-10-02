"""
Credit Control — invoices, clients, commissions, insights and operational
settings for billing/retention tracking. Four dashboard entry points
(Credit Control, Commission, Operational Settings, Insights) share one
static front-end file and one data store.

The front-end (app/templates/credit_control/app.html) is the same file that
used to be Credit Control's entire standalone app, served locally by
serve.ps1 against a local data.json. It is served here byte-for-byte except
for four small, deliberate edits (see the diff in that file's history): the
data load/save URLs, and a small bit of JS that reads the request path to
decide which tabs to show. All of the actual business logic — churn
grace-period detection, due-date month assignment, commission math,
retention calculations — is completely untouched, on purpose: re-deriving
any of that against a new data shape is exactly the risk this design avoids.
Data itself lives in Supabase as a single JSONB blob
(credit_control_store, see migrations/credit_control_schema.sql) holding the
exact same {settings, clients, invoices} shape the app has always used, so
none of its ~3,000 lines of calculation code needed to change.

Access: all four tool keys (credit_control, insights, commission,
operational_settings) are admin-only by default — see ROLE_TOOLS in
app/auth.py. The two /api/credit-control/* data endpoints are shared by all
four pages, so unlike every other router here they can't be gated by a
single entry in ROUTE_TOOL_MAP (one prefix -> one tool); instead they check
the logged-in user's tool set directly against all four keys.
"""
from __future__ import annotations

from pathlib import Path

import requests
from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse

from app import auth
from app.utils.supabase import SUPABASE_URL, sb_service_headers as sb_headers, sb_service_configured as sb_configured

router = APIRouter()

_APP_HTML_PATH = Path(__file__).resolve().parent.parent / "templates" / "credit_control" / "app.html"

# Any one of these four grants access to the shared data endpoints — a user
# holding only 'insights', say, still needs to be able to load the data that
# powers the Insights tab.
_CC_TOOLS = {"credit_control", "insights", "commission", "operational_settings"}


def _require_cc_access(request: Request) -> dict:
    user = auth.require_login(request)
    if auth.auth_enabled() and not (_CC_TOOLS & set(user.get("tools") or [])):
        raise HTTPException(status_code=403, detail="No access to Credit Control data")
    return user


def _serve_app() -> HTMLResponse:
    return HTMLResponse(_APP_HTML_PATH.read_text(encoding="utf-8"))


@router.get("/credit-control")
async def credit_control_page(_user: dict = Depends(_require_cc_access)):
    return _serve_app()


@router.get("/commission")
async def commission_page(_user: dict = Depends(_require_cc_access)):
    return _serve_app()


@router.get("/operational-settings")
async def operational_settings_page(_user: dict = Depends(_require_cc_access)):
    return _serve_app()


@router.get("/insights")
async def insights_page(_user: dict = Depends(_require_cc_access)):
    return _serve_app()


@router.get("/api/credit-control/data")
async def get_data(_user: dict = Depends(_require_cc_access)):
    if not sb_configured():
        return JSONResponse({"error": "Supabase service key is not configured (SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY)."}, status_code=503)

    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/credit_control_store",
        headers=sb_headers(),
        params={"select": "data,updated_at", "id": "eq.1"},
        timeout=30,
    )
    r.raise_for_status()
    rows = r.json()
    if not rows:
        return JSONResponse(
            {"error": "No Credit Control data yet — run scripts/import_credit_control_data.py first."},
            status_code=404,
        )

    # _ccVersion rides along inside the same JSON the front-end already treats
    # as an opaque blob (DB = await r.json()) — it round-trips through
    # save/load for free without the front-end needing to know it exists,
    # and gives /save something to detect a concurrent edit against.
    payload = dict(rows[0]["data"])
    payload["_ccVersion"] = rows[0]["updated_at"]
    return JSONResponse(payload)


@router.post("/api/credit-control/save")
async def save_data(request: Request, user: dict = Depends(_require_cc_access)):
    if not sb_configured():
        return JSONResponse({"error": "Supabase service key is not configured (SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY)."}, status_code=503)

    try:
        body = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid JSON")
    if not isinstance(body, dict) or not all(k in body for k in ("settings", "clients", "invoices")):
        raise HTTPException(status_code=400, detail="Expected {settings, clients, invoices}")

    client_version = body.pop("_ccVersion", None)

    current = requests.get(
        f"{SUPABASE_URL}/rest/v1/credit_control_store",
        headers=sb_headers(),
        params={"select": "data,updated_at", "id": "eq.1"},
        timeout=30,
    )
    current.raise_for_status()
    current_rows = current.json()

    # Optimistic concurrency: if someone else saved since this client last
    # loaded, refuse the blind overwrite rather than silently clobbering
    # their edit (today's local-file save has no such check at all).
    if current_rows and client_version and current_rows[0]["updated_at"] != client_version:
        return JSONResponse(
            {"error": "conflict", "detail": "Data changed since you last loaded it — reload and reapply your change."},
            status_code=409,
        )

    # Backup before overwrite, mirroring serve.ps1's timestamped backups/*.json.
    if current_rows:
        requests.post(
            f"{SUPABASE_URL}/rest/v1/credit_control_backups",
            headers=sb_headers("return=minimal"),
            json={"data": current_rows[0]["data"], "saved_by": user.get("email", "unknown")},
            timeout=30,
        ).raise_for_status()

    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/credit_control_store",
        headers=sb_headers("resolution=merge-duplicates,return=minimal"),
        params={"on_conflict": "id"},
        json={"id": 1, "data": body, "updated_by": user.get("email", "unknown")},
        timeout=30,
    )
    r.raise_for_status()
    return JSONResponse({"ok": True})
