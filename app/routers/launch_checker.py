from datetime import datetime, timezone, timedelta
from typing import List, Optional

import os
import requests as http_req
from fastapi import APIRouter, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from app.deps import templates

router = APIRouter()


# ── Supabase helpers ─────────────────────────────────────────────────────────

def _sb_headers(prefer: str = "") -> dict:
    key = os.environ.get("SUPABASE_ANON_KEY", "")
    h = {
        "apikey":        key,
        "Authorization": f"Bearer {key}",
        "Content-Type":  "application/json",
    }
    if prefer:
        h["Prefer"] = prefer
    return h


def _sb_configured() -> bool:
    return bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_ANON_KEY"))


def _sb_url() -> str:
    return os.environ.get("SUPABASE_URL", "")


# ── Page ─────────────────────────────────────────────────────────────────────

@router.get("/launch-checker")
async def launch_checker(request: Request):
    return templates.TemplateResponse("launch_checker.html", {"request": request})


# ── Pydantic models ──────────────────────────────────────────────────────────

class RecordIn(BaseModel):
    id:                  str
    launch_name:         str
    user_name:           str
    role:                str
    client_name:         str
    client_type:         str
    channels:            list
    active_checks:       list
    completed_sub_steps: list
    status:              str
    timestamp_started:   str
    timestamp_updated:   str
    timestamp_completed: Optional[str] = None


class RecordPatch(BaseModel):
    completed_sub_steps: Optional[list] = None
    status:              Optional[str]  = None
    timestamp_updated:   Optional[str]  = None
    timestamp_completed: Optional[str]  = None


# ── API endpoints ─────────────────────────────────────────────────────────────

@router.get("/api/launch-checker/records")
def get_records():
    if not _sb_configured():
        return JSONResponse([], status_code=200)

    url = _sb_url()

    # Auto-delete records older than 2 months
    cutoff = (datetime.now(timezone.utc) - timedelta(days=60)).isoformat()
    http_req.delete(
        f"{url}/rest/v1/launch_checker_records",
        params={"timestamp_started": f"lt.{cutoff}"},
        headers=_sb_headers(),
        timeout=10,
    )

    # Return remaining records, newest first
    resp = http_req.get(
        f"{url}/rest/v1/launch_checker_records",
        params={"select": "*", "order": "timestamp_started.desc"},
        headers=_sb_headers(),
        timeout=10,
    )
    return resp.json() if resp.ok else []


@router.post("/api/launch-checker/records")
def create_record(body: RecordIn):
    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured"}, status_code=503)

    url  = _sb_url()
    resp = http_req.post(
        f"{url}/rest/v1/launch_checker_records",
        headers=_sb_headers("return=representation"),
        json=body.dict(),
        timeout=10,
    )
    if resp.ok:
        rows = resp.json()
        return rows[0] if rows else body.dict()
    return JSONResponse({"error": resp.text}, status_code=500)


@router.patch("/api/launch-checker/records/{record_id}")
def update_record(record_id: str, body: RecordPatch):
    if not _sb_configured():
        return JSONResponse({"error": "Supabase not configured"}, status_code=503)

    patch_data = {k: v for k, v in body.dict().items() if v is not None}
    if not patch_data:
        return JSONResponse({"ok": True})

    url  = _sb_url()
    resp = http_req.patch(
        f"{url}/rest/v1/launch_checker_records",
        params={"id": f"eq.{record_id}"},
        headers=_sb_headers("return=minimal"),
        json=patch_data,
        timeout=10,
    )
    return JSONResponse({"ok": resp.ok})
