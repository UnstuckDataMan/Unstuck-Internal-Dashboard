import json
import logging
import os

import requests as http_req
from fastapi import APIRouter, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from app.deps import templates

log = logging.getLogger(__name__)

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


@router.get("/api/copy-bank/senders")
def copy_bank_senders():
    url = os.environ.get("SUPABASE_URL", "")
    resp = http_req.get(
        f"{url}/rest/v1/copy_bank_templates",
        params={"key": "eq.__cb_senders__", "select": "content"},
        headers=_sb_headers(),
        timeout=10,
    )
    rows = resp.json()
    if rows and isinstance(rows[0].get("content"), list):
        return rows[0]["content"]
    return ["Robyn", "Glen", "Leo", "Chris", "Devlin"]


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


# ── Approval workflow helpers ──────────────────────────────────────────────────

def _sb_write_headers(prefer: str = "return=minimal"):
    return {
        **_sb_headers(),
        "Content-Type": "application/json",
        "Prefer": prefer,
    }


def _send_slack_approval_notification(client_name: str, territory: str, industry: str, approved_by: str) -> dict:
    """Fire a Slack message confirming copy has been approved and published."""
    webhook_url = os.environ.get("SLACK_WEBHOOK_URL", "")
    if not webhook_url:
        return {"ok": False, "error": "SLACK_WEBHOOK_URL not set"}
    payload = {
        "text": (
            f"✅ *Copy Approved & Published*\n"
            f"*Client:* {client_name}  |  *Territory:* {territory}  |  *Industry:* {industry}\n"
            f"Approved by *{approved_by}* — copy is now live."
        )
    }
    try:
        r = http_req.post(webhook_url, json=payload, timeout=8)
        r.raise_for_status()
        return {"ok": True}
    except Exception as exc:
        log.error("Slack approval notification failed: %s", exc)
        return {"ok": False, "error": str(exc)}


def _send_slack_notification(client_name: str, territory: str, industry: str) -> dict:
    """Fire a Slack incoming-webhook message tagging the three reviewers.
    Returns {"ok": True} on success or {"ok": False, "error": "..."} on failure."""
    webhook_url = os.environ.get("SLACK_WEBHOOK_URL", "")
    if not webhook_url:
        msg = "SLACK_WEBHOOK_URL not set"
        log.warning("%s — skipping Slack notification", msg)
        return {"ok": False, "error": msg}

    ollie = os.environ.get("SLACK_OLLIE_ID", "Ollie")
    chris = os.environ.get("SLACK_CHRIS_ID", "Chris")
    leo   = os.environ.get("SLACK_LEO_ID", "Leo")

    # Format mentions: if the value looks like a Slack member ID use <@ID>, else plain name
    def mention(val: str) -> str:
        return f"<@{val}>" if val.startswith("U") and len(val) >= 9 else val

    payload = {
        "text": (
            f"📋 *Copy Approval Request*\n"
            f"*Client:* {client_name}  |  *Territory:* {territory}  |  *Industry:* {industry}\n"
            f"{mention(ollie)} {mention(chris)} {mention(leo)} — "
            f"please review in the Copy Bank admin panel."
        )
    }
    try:
        r = http_req.post(webhook_url, json=payload, timeout=8)
        r.raise_for_status()
        return {"ok": True}
    except Exception as exc:
        log.error("Slack notification failed: %s", exc)
        return {"ok": False, "error": str(exc)}


@router.get("/api/copy-bank/test-slack")
def test_slack():
    """Diagnostic endpoint — fires a test Slack notification and returns the result."""
    webhook_url = os.environ.get("SLACK_WEBHOOK_URL", "")
    ollie = os.environ.get("SLACK_OLLIE_ID", "")
    chris = os.environ.get("SLACK_CHRIS_ID", "")
    leo   = os.environ.get("SLACK_LEO_ID", "")
    env_status = {
        "SLACK_WEBHOOK_URL": f"{'set (' + webhook_url[:40] + '…)' if webhook_url else 'NOT SET'}",
        "SLACK_OLLIE_ID":    ollie or "NOT SET",
        "SLACK_CHRIS_ID":    chris or "NOT SET",
        "SLACK_LEO_ID":      leo   or "NOT SET",
    }
    slack_result = _send_slack_notification("Test Client", "US", "PR")
    return JSONResponse({"env": env_status, "slack": slack_result})


# ── Pydantic request bodies ────────────────────────────────────────────────────

class ApprovalRequestBody(BaseModel):
    key:              str
    client_name:      str
    territory:        str
    industry:         str
    content:          dict
    requested_by:     str = ""
    previous_content: dict = {}


class ApproveBody(BaseModel):
    key:         str
    approved_by: str   # 'Ollie' | 'Chris' | 'Leo'
    content:     dict  # final published content (reviewer may have edited before approving)


# ── Approval endpoints ─────────────────────────────────────────────────────────

@router.post("/api/copy-bank/request-approval")
def request_approval(body: ApprovalRequestBody):
    """Save a pending draft and fire a Slack notification to the reviewers."""
    url = os.environ.get("SUPABASE_URL", "")

    # Delete any existing pending rows for this key so there is always exactly one
    http_req.delete(
        f"{url}/rest/v1/copy_bank_pending",
        params={"key": f"eq.{body.key}", "status": "eq.pending"},
        headers=_sb_write_headers(),
        timeout=10,
    )

    # Insert the new pending request row
    resp = http_req.post(
        f"{url}/rest/v1/copy_bank_pending",
        headers={**_sb_write_headers("return=minimal")},
        json={
            "key":              body.key,
            "client_name":      body.client_name,
            "territory":        body.territory,
            "industry":         body.industry,
            "content":          body.content,
            "requested_by":     body.requested_by,
            "previous_content": body.previous_content,
            "status":           "pending",
        },
        timeout=10,
    )
    if not resp.ok:
        try:
            err_body = resp.json()
        except Exception:
            err_body = resp.text
        log.error("copy_bank_pending insert failed %s: %s", resp.status_code, err_body)
        return JSONResponse({"ok": False, "error": err_body}, status_code=500)

    slack = _send_slack_notification(body.client_name, body.territory, body.industry)
    return JSONResponse({"ok": True, "slack": slack})


@router.get("/api/copy-bank/pending-all")
def get_all_pending():
    """Return all pending draft rows (used to pre-load on page init)."""
    url = os.environ.get("SUPABASE_URL", "")
    resp = http_req.get(
        f"{url}/rest/v1/copy_bank_pending",
        params={"status": "eq.pending", "select": "*", "order": "created_at.desc", "limit": "500"},
        headers=_sb_headers(),
        timeout=10,
    )
    return resp.json() if resp.ok else []


@router.get("/api/copy-bank/pending/{key:path}")
def get_pending(key: str):
    """Return the current pending draft for a given key, or null if none."""
    url = os.environ.get("SUPABASE_URL", "")
    resp = http_req.get(
        f"{url}/rest/v1/copy_bank_pending",
        params={"key": f"eq.{key}", "status": "eq.pending", "select": "*", "order": "created_at.desc", "limit": "1"},
        headers=_sb_headers(),
        timeout=10,
    )
    rows = resp.json() if resp.ok else []
    return JSONResponse({"pending": rows[0] if rows else None})


@router.post("/api/copy-bank/approve")
def approve_copy(body: ApproveBody):
    """
    Approve a pending draft:
      1. Fetch the pending row
      2. Publish its content to copy_bank_templates
      3. Write an entry to copy_approval_logs
      4. Mark the pending row as approved
    """
    url = os.environ.get("SUPABASE_URL", "")

    # 1. Fetch most recent pending row
    resp = http_req.get(
        f"{url}/rest/v1/copy_bank_pending",
        params={"key": f"eq.{body.key}", "status": "eq.pending", "select": "*", "order": "created_at.desc", "limit": "1"},
        headers=_sb_headers(),
        timeout=10,
    )
    rows = resp.json() if resp.ok else []
    if not rows:
        return JSONResponse({"ok": False, "error": "No pending draft found"}, status_code=404)

    pending = rows[0]

    # 2. Publish to copy_bank_templates — use reviewer's final content (may differ from original)
    pub = http_req.post(
        f"{url}/rest/v1/copy_bank_templates",
        params={"on_conflict": "key"},
        headers=_sb_write_headers("resolution=merge-duplicates,return=minimal"),
        json={"key": body.key, "content": body.content},
        timeout=10,
    )
    if not pub.ok:
        try:
            err_body = pub.json()
        except Exception:
            err_body = pub.text
        log.error("copy_bank_templates publish failed %s: %s", pub.status_code, err_body)
        return JSONResponse({"ok": False, "error": f"Publish failed: {err_body}"}, status_code=500)

    # 3. Write approval log
    try:
        http_req.post(
            f"{url}/rest/v1/copy_approval_logs",
            headers=_sb_write_headers(),
            json={
                "key":                       body.key,
                "client_name":               pending["client_name"],
                "territory":                 pending["territory"],
                "industry":                  pending["industry"],
                "content_snapshot":          body.content,
                "previous_content_snapshot": pending.get("previous_content", {}),
                "approved_by":               body.approved_by,
                "requested_by":              pending.get("requested_by", ""),
                "requested_at":              pending["created_at"],
            },
            timeout=10,
        ).raise_for_status()
    except Exception as exc:
        log.error("Failed to write approval log: %s", exc)

    # 4. Mark pending row as approved
    try:
        http_req.patch(
            f"{url}/rest/v1/copy_bank_pending",
            params={"key": f"eq.{body.key}"},
            headers=_sb_write_headers(),
            json={"status": "approved"},
            timeout=10,
        ).raise_for_status()
    except Exception as exc:
        log.error("Failed to update pending status: %s", exc)

    # 5. Notify Slack that copy is live
    _send_slack_approval_notification(
        pending["client_name"], pending["territory"], pending["industry"], body.approved_by
    )

    return JSONResponse({"ok": True})


@router.get("/api/copy-bank/approval-logs")
def approval_logs():
    """Return pending requests and approval history for the logs panel."""
    url = os.environ.get("SUPABASE_URL", "")

    logs_resp = http_req.get(
        f"{url}/rest/v1/copy_approval_logs",
        params={"select": "*", "order": "approved_at.desc", "limit": "100"},
        headers=_sb_headers(),
        timeout=10,
    )
    pending_resp = http_req.get(
        f"{url}/rest/v1/copy_bank_pending",
        params={"select": "*", "status": "eq.pending", "order": "created_at.desc", "limit": "100"},
        headers=_sb_headers(),
        timeout=10,
    )

    return JSONResponse({
        "logs":    logs_resp.json()    if logs_resp.ok    else [],
        "pending": pending_resp.json() if pending_resp.ok else [],
    })


# ── Admin: merge Unstuck profiles ─────────────────────────────────────────────

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


# ── Apostrophe normalisation ───────────────────────────────────────────────────

_CURLY_APOS_TABLE = str.maketrans({
    "‘": "'",   # LEFT  SINGLE QUOTATION MARK  →  straight apostrophe
    "’": "'",   # RIGHT SINGLE QUOTATION MARK  →  straight apostrophe
    "‚": "'",   # SINGLE LOW-9 QUOTATION MARK  →  straight apostrophe
    "‛": "'",   # SINGLE HIGH-REVERSED-9       →  straight apostrophe
})


def _normalise_str(s: str) -> str:
    return s.translate(_CURLY_APOS_TABLE) if isinstance(s, str) else s


def _normalise_value(v):
    """Recursively normalise apostrophes in any JSON-compatible value."""
    if isinstance(v, str):
        return v.translate(_CURLY_APOS_TABLE)
    if isinstance(v, dict):
        return {k: _normalise_value(val) for k, val in v.items()}
    if isinstance(v, list):
        return [_normalise_value(item) for item in v]
    return v


@router.post("/api/copy-bank/normalize-apostrophes")
def normalize_apostrophes():
    """
    One-time (idempotent) migration: replace curly/smart apostrophes with
    straight apostrophes in every copy_bank_templates row that contains them.
    Returns counts of rows inspected, changed, and any errors.
    """
    url = os.environ.get("SUPABASE_URL", "")

    # Fetch all template rows (skip meta keys — profiles list, etc.)
    resp = http_req.get(
        f"{url}/rest/v1/copy_bank_templates",
        params={"select": "key,content"},
        headers=_sb_headers(),
        timeout=30,
    )
    if not resp.ok:
        return JSONResponse({"ok": False, "error": resp.text}, status_code=500)

    rows      = resp.json()
    inspected = 0
    changed   = 0
    errors    = []

    for row in rows:
        key     = row.get("key", "")
        content = row.get("content")
        if content is None:
            continue

        inspected += 1
        normalised = _normalise_value(content)

        # Only write back if something actually changed (cheap json round-trip comparison)
        if json.dumps(normalised, ensure_ascii=False) == json.dumps(content, ensure_ascii=False):
            continue

        patch = http_req.patch(
            f"{url}/rest/v1/copy_bank_templates",
            params={"key": f"eq.{key}"},
            headers=_sb_write_headers("return=minimal"),
            json={"content": normalised},
            timeout=10,
        )
        if patch.ok:
            changed += 1
        else:
            errors.append({"key": key, "status": patch.status_code, "body": patch.text[:200]})
            log.error("normalize_apostrophes: failed to patch %s — %s", key, patch.text[:200])

    return JSONResponse({
        "ok":        len(errors) == 0,
        "inspected": inspected,
        "changed":   changed,
        "errors":    errors,
    })
