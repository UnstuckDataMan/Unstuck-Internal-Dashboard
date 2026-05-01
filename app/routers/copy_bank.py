import json
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
            "type":        p.get("type", "client"),
            "territories": p.get("territories", []),
            "industries":  p.get("industries", []),
        }
        for p in rows[0]["content"]
    ]


@router.get("/api/copy-bank/templates/{client_id}/{territory}/{industry}")
def copy_bank_template(client_id: str, territory: str, industry: str, channel: str = "email"):
    url = os.environ.get("SUPABASE_URL", "")

    # Bizdev content uses simple territory_industry keys; clients use __c__ prefix
    if client_id == "bizdev":
        cb_key = f"{territory}_{industry}"
    else:
        cb_key = f"__c__{client_id}__{territory}_{industry}"

    resp = http_req.get(
        f"{url}/rest/v1/copy_bank_templates",
        params={"key": f"eq.{cb_key}", "select": "content"},
        headers=_sb_headers(),
        timeout=10,
    )
    rows = resp.json()

    # Fallback: if no client key found, try bizdev key format (covers migrated profiles)
    if not rows and client_id != "bizdev":
        cb_key = f"{territory}_{industry}"
        resp = http_req.get(
            f"{url}/rest/v1/copy_bank_templates",
            params={"key": f"eq.{cb_key}", "select": "content"},
            headers=_sb_headers(),
            timeout=10,
        )
        rows = resp.json()

    if not rows:
        return {"subjects": [], "bodies": []}

    c = rows[0].get("content") or {}

    # Select channel data — flyout only available for Biz Dev
    ch_key = "flyout" if channel == "flyout" else "email"
    ch     = c.get(ch_key) or {}

    subjects = [s for s in (ch.get("subjects") or []) if s and s.strip()]
    bodies   = [v["body"] for v in (ch.get("variations") or []) if v.get("body", "").strip()]
    return {"subjects": subjects, "bodies": bodies}


@router.post("/api/admin/merge-unstuck-profiles")
def merge_unstuck_profiles():
    url = os.environ.get("SUPABASE_URL", "")
    write_headers = {
        **_sb_headers(),
        "Content-Type": "application/json",
        "Prefer": "return=representation",
    }

    resp = http_req.get(
        f"{url}/rest/v1/copy_bank_templates",
        params={"key": "eq.__cb_profiles__", "select": "content"},
        headers=_sb_headers(),
        timeout=10,
    )
    rows = resp.json()
    if not rows or not isinstance(rows[0].get("content"), list):
        return {"error": "No profiles found"}

    content = rows[0]["content"]

    bizdev = next((p for p in content if p.get("type") == "bizdev" and p.get("name") == "Unstuck Agency"), None)
    client = next((p for p in content if p.get("type") == "client" and p.get("name") == "Unstuck Agency"), None)

    if not bizdev or not client:
        return {
            "message":      "Nothing to merge",
            "bizdev_found": bool(bizdev),
            "client_found": bool(client),
        }

    # Union territories and industries — client order first, then any bizdev extras
    merged_territories = list(dict.fromkeys(
        (client.get("territories") or []) + (bizdev.get("territories") or [])
    ))
    merged_industries = list(dict.fromkeys(
        (client.get("industries") or []) + (bizdev.get("industries") or [])
    ))
    client["territories"] = merged_territories
    client["industries"]  = merged_industries

    # Remove bizdev entry, keep everything else (client entry already updated in-place)
    new_content = [p for p in content
                   if not (p.get("type") == "bizdev" and p.get("name") == "Unstuck Agency")]

    patch_resp = http_req.patch(
        f"{url}/rest/v1/copy_bank_templates",
        params={"key": "eq.__cb_profiles__"},
        data=json.dumps({"content": new_content}),
        headers=write_headers,
        timeout=10,
    )

    return {
        "merged":      True,
        "client_id":   client["client_id"],
        "territories": merged_territories,
        "industries":  merged_industries,
        "status":      patch_resp.status_code,
    }
