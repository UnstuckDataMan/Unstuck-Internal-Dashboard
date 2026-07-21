import datetime
import os
import uuid

import requests as http_req
from fastapi import APIRouter, Depends, File, HTTPException, Request, UploadFile
from pydantic import BaseModel

from app import auth
from app.deps import templates

router = APIRouter()

_SUPABASE_URL = os.environ.get("SUPABASE_URL", "")
_SUPABASE_KEY = os.environ.get("SUPABASE_ANON_KEY", "")

SOP_GROUPS = [
    {"group": "Dashboard", "items": [
        {"key": "home",               "label": "Home"},
    ]},
    {"group": "Data Tools", "items": [
        {"key": "gender_classifier",  "label": "Gender Classifier"},
        {"key": "city_state",         "label": "City / State Normaliser"},
    ]},
    {"group": "DNC & Merger", "items": [
        {"key": "dnc_scrub",          "label": "Scrub Prospects"},
        {"key": "dnc_merge",          "label": "Mail Merge"},
        {"key": "dnc_campaigns",      "label": "Campaigns"},
        {"key": "dnc_manage",         "label": "Manage DNC Lists"},
        {"key": "dnc_contacted",      "label": "Manage Contacted"},
    ]},
    {"group": "Outreach", "items": [
        {"key": "reply_bank",         "label": "Reply Bank"},
        {"key": "copy_bank",          "label": "Copy Bank"},
    ]},
    {"group": "Targeting", "items": [
        {"key": "bd_targeting",       "label": "BD Targeting"},
        {"key": "launch_checker",     "label": "Launch Checker"},
        {"key": "targeting_checker",  "label": "Targeting Checker"},
    ]},
    {"group": "Admin", "items": [
        {"key": "client_profiles",    "label": "Client Profiles"},
    ]},
]


def _headers(prefer: str = "") -> dict:
    h = {
        "apikey":        _SUPABASE_KEY,
        "Authorization": f"Bearer {_SUPABASE_KEY}",
        "Content-Type":  "application/json",
    }
    if prefer:
        h["Prefer"] = prefer
    return h


@router.get("/sops")
async def sops_page(request: Request):
    user = auth.get_session_user(request)
    return templates.TemplateResponse("sops.html", {
        "request":    request,
        "active":     "sops",
        "user":       user,
        "sop_groups": SOP_GROUPS,
        "can_edit":   user and user.get("role") in {"admin", "reviewer"},
    })


@router.get("/api/sops/{key}")
async def get_sop(key: str):
    resp = http_req.get(
        f"{_SUPABASE_URL}/rest/v1/sops",
        params={"page_key": f"eq.{key}", "select": "page_key,title,content,updated_at,updated_by"},
        headers=_headers(),
        timeout=10,
    )
    rows = resp.json() if resp.ok else []
    if not rows:
        return {"page_key": key, "title": "", "content": "", "updated_at": None, "updated_by": ""}
    return rows[0]


class SopBody(BaseModel):
    title: str = ""
    content: str = ""


# NOTE: upload-image route must be before /{key} so FastAPI doesn't treat
# "upload-image" as a key value.
@router.post("/api/sops/upload-image")
async def upload_image(
    file: UploadFile = File(...),
    _: None = Depends(auth.require_role({"admin", "reviewer"})),
):
    ext = (file.filename or "img").rsplit(".", 1)[-1].lower() or "png"
    path = f"sop-media/{uuid.uuid4()}.{ext}"
    data = await file.read()
    resp = http_req.post(
        f"{_SUPABASE_URL}/storage/v1/object/{path}",
        data=data,
        headers={
            "apikey":        _SUPABASE_KEY,
            "Authorization": f"Bearer {_SUPABASE_KEY}",
            "Content-Type":  file.content_type or "image/png",
        },
        timeout=30,
    )
    if resp.status_code not in (200, 201):
        raise HTTPException(500, f"Upload failed: {resp.text}")
    return {"url": f"{_SUPABASE_URL}/storage/v1/object/public/{path}"}


@router.post("/api/sops/{key}")
async def save_sop(
    key: str,
    body: SopBody,
    request: Request,
    _: None = Depends(auth.require_role({"admin", "reviewer"})),
):
    user = auth.get_session_user(request)
    data = {
        "page_key":   key,
        "title":      body.title,
        "content":    body.content,
        "updated_at": datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "updated_by": (user or {}).get("email", ""),
    }
    resp = http_req.post(
        f"{_SUPABASE_URL}/rest/v1/sops",
        json=data,
        params={"on_conflict": "page_key"},
        headers=_headers("resolution=merge-duplicates,return=minimal"),
        timeout=10,
    )
    if resp.status_code not in (200, 201):
        raise HTTPException(500, f"Save failed: {resp.text}")
    return {"ok": True}
