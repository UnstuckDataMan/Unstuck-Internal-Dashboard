import os

import requests as http_req
from fastapi import APIRouter, Request

from app.deps import templates

router = APIRouter()


@router.get("/copy-bank")
async def copy_bank(request: Request):
    return templates.TemplateResponse("copy_bank.html", {
        "request":           request,
        "supabase_url":      os.environ.get("SUPABASE_URL", ""),
        "supabase_anon_key": os.environ.get("SUPABASE_ANON_KEY", ""),
    })


def _sb_headers():
    key = os.environ.get("SUPABASE_ANON_KEY", "")
    return {"apikey": key, "Authorization": f"Bearer {key}"}


@router.get("/api/copy-bank/profiles")
def copy_bank_profiles():
    url = os.environ.get("SUPABASE_URL", "")
    resp = http_req.get(
        f"{url}/rest/v1/copy_bank_templates",
        params={"key": "eq.__cb_profiles__", "select": "content"},
        headers=_sb_headers(),
        timeout=10,
    )
    rows = resp.json()
    if not rows or not isinstance(rows[0].get("content"), list):
        return []
    return [
        {
            "client_id":   p["client_id"],
            "name":        p["name"],
            "territories": p.get("territories", []),
            "industries":  p.get("industries", []),
        }
        for p in rows[0]["content"]
        if p.get("type") != "bizdev"
    ]


@router.get("/api/copy-bank/templates/{client_id}/{territory}/{industry}")
def copy_bank_template(client_id: str, territory: str, industry: str):
    url    = os.environ.get("SUPABASE_URL", "")
    cb_key = f"__c__{client_id}__{territory}_{industry}"
    resp   = http_req.get(
        f"{url}/rest/v1/copy_bank_templates",
        params={"key": f"eq.{cb_key}", "select": "content"},
        headers=_sb_headers(),
        timeout=10,
    )
    rows = resp.json()
    if not rows:
        return {"subjects": [], "bodies": []}
    c        = rows[0].get("content") or {}
    email    = c.get("email") or {}
    subjects = [s for s in (email.get("subjects") or []) if s and s.strip()]
    bodies   = [v["body"] for v in (email.get("variations") or []) if v.get("body", "").strip()]
    return {"subjects": subjects, "bodies": bodies}
