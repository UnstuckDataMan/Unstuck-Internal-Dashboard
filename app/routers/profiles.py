"""
Central Client Profiles manager.

The `clients` table is the single source of truth for a client profile. This
router does CRUD on it and, after every write, mirrors the relevant slices out
to the legacy per-tool stores so the existing tools keep working unchanged:

  • sender_profiles        — Mail Merge sender dropdown (row named after client)
  • copy_bank_templates    — Copy Bank's __cb_profiles__ array entry
  • campaigns.client_name  — denormalised name, patched on rename

Mirroring runs in a daemon thread and is best-effort / non-fatal: the canonical
clients row is already saved, so a downstream hiccup never fails the user's save
(house style — see working-style memory).

The extended profile columns (territories, sender_emails, …) are added by
migrations/client_profiles_schema.sql. Until that runs, the router gracefully
falls back to the always-present id,name columns so nothing 500s.
"""
from __future__ import annotations

import threading
from html import escape

import requests as http_req
from fastapi import APIRouter, Depends, Form, Query, Request
from fastapi.responses import HTMLResponse, JSONResponse

from app import auth
from app.deps import templates
from app.utils.supabase import (
    SUPABASE_URL,
    sb_headers as _sb_headers,
    sb_configured as _sb_configured,
)

router = APIRouter()

# Canonical profile columns. Selected when present; falls back to id,name when
# the extended columns haven't been migrated yet.
_SELECT_FULL = (
    "id,name,type,territories,industries,industry_labels,"
    "senders,sender_emails,color,emoji,booking_link,notes,active,updated_at"
)
_SELECT_BASE = "id,name"


# ── helpers ─────────────────────────────────────────────────────────────────────

def _truthy(val: str) -> bool:
    return str(val).strip().lower() in {"1", "true", "yes", "on"}


def _split_multi(raw: str) -> list[str]:
    """Split a comma/newline-separated input into a clean, de-duplicated list."""
    if not raw:
        return []
    out: list[str] = []
    for chunk in raw.replace("\r", "\n").split("\n"):
        for part in chunk.split(","):
            p = part.strip()
            if p and p not in out:
                out.append(p)
    return out


def _build_payload(*, name, type, territories, industries, senders,
                   sender_emails, color, emoji, booking_link, notes, active) -> dict:
    """Assemble the canonical clients row from the submitted form fields."""
    inds = _split_multi(industries)
    return {
        "name":            name.strip(),
        "type":            (type or "client").strip() or "client",
        "territories":     _split_multi(territories),
        "industries":      inds,
        "industry_labels": {i: i for i in inds},
        "senders":         _split_multi(senders),
        "sender_emails":   _split_multi(sender_emails),
        "color":           (color or "").strip() or None,
        "emoji":           (emoji or "").strip() or None,
        "booking_link":    (booking_link or "").strip() or None,
        "notes":           (notes or "").strip() or None,
        "active":          _truthy(active),
    }


def _fetch_profiles() -> tuple[list[dict], str]:
    """All profiles, newest columns first with graceful fallback. Returns (rows, error)."""
    if not _sb_configured():
        return [], "Supabase is not configured."
    err = ""
    for sel in (_SELECT_FULL, _SELECT_BASE):
        try:
            r = http_req.get(
                f"{SUPABASE_URL}/rest/v1/clients",
                headers=_sb_headers(),
                params={"select": sel, "order": "name.asc"},
                timeout=10,
            )
            r.raise_for_status()
            return r.json(), ""
        except Exception as exc:
            err = str(exc)
    return [], err


def _fetch_profile(profile_id: str) -> dict | None:
    """A single profile row (for the edit form), with the same select fallback."""
    if not _sb_configured():
        return None
    for sel in (_SELECT_FULL, _SELECT_BASE):
        try:
            r = http_req.get(
                f"{SUPABASE_URL}/rest/v1/clients",
                headers=_sb_headers(),
                params={"select": sel, "id": f"eq.{profile_id}"},
                timeout=10,
            )
            r.raise_for_status()
            rows = r.json()
            return rows[0] if rows else None
        except Exception:
            continue
    return None


def _list_response(request: Request, trigger: bool = False):
    """Render the list partial. `trigger` fires an HX-Trigger so the page JS can
    reset/close the editor after a successful mutation."""
    profiles, error = _fetch_profiles()
    resp = templates.TemplateResponse(
        "partials/client_profiles_list.html",
        {"request": request, "profiles": profiles, "error": error},
    )
    if trigger:
        resp.headers["HX-Trigger"] = "profilesChanged"
    return resp


def _error(msg: str) -> HTMLResponse:
    """Inline error, retargeted to the editor's error slot (DNC add-entry pattern)."""
    return HTMLResponse(
        content=f'<div class="err-box" style="margin-top:0">{escape(msg)}</div>',
        headers={"HX-Retarget": "#profile-error", "HX-Reswap": "innerHTML"},
    )


# ── mirror / auto-sync ───────────────────────────────────────────────────────────

def _mirror_async(profile: dict) -> None:
    """Fire-and-forget the cross-tool sync so the HTTP response returns immediately."""
    threading.Thread(target=_sync_profile_to_tools, args=(profile,), daemon=True).start()


def _sync_profile_to_tools(profile: dict) -> None:
    if not _sb_configured():
        return
    _mirror_sender_profile(profile)
    _mirror_copy_bank_profile(profile)
    _mirror_campaign_client_name(profile)


def _mirror_sender_profile(profile: dict) -> None:
    """Upsert a sender_profiles row named after the client so the Mail Merge
    dropdown shows the profile's sending inboxes. Additive — per-team sender
    profiles are untouched. The old-named row is removed after a rename."""
    name   = (profile.get("name") or "").strip()
    emails = profile.get("sender_emails") or []
    if not name or not emails:
        return
    try:
        http_req.post(
            f"{SUPABASE_URL}/rest/v1/sender_profiles",
            headers=_sb_headers("resolution=merge-duplicates"),
            params={"on_conflict": "profile_name"},
            json={"profile_name": name, "emails": emails},
            timeout=10,
        )
    except Exception:
        pass  # non-fatal: canonical row already saved
    prev = (profile.get("_prev_name") or "").strip()
    if prev and prev != name:
        try:
            http_req.delete(
                f"{SUPABASE_URL}/rest/v1/sender_profiles",
                headers=_sb_headers(),
                params={"profile_name": f"eq.{prev}"},
                timeout=10,
            )
        except Exception:
            pass


def _mirror_copy_bank_profile(profile: dict) -> None:
    """Update (or append) this client's entry in the copy_bank_templates
    __cb_profiles__ array so Copy Bank reflects the central profile."""
    cid = profile.get("id")
    if not cid:
        return
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/copy_bank_templates",
            headers=_sb_headers(),
            params={"key": "eq.__cb_profiles__", "select": "content"},
            timeout=10,
        )
        r.raise_for_status()
        rows = r.json()
    except Exception:
        return
    content = rows[0]["content"] if rows and isinstance(rows[0].get("content"), list) else []

    new_entry = {
        "client_id":      cid,
        "name":           profile.get("name"),
        "type":           profile.get("type") or "client",
        "territories":    profile.get("territories") or [],
        "industries":     profile.get("industries") or [],
        "industryLabels": profile.get("industry_labels") or {},
        "senders":        profile.get("senders") or [],
        "color":          profile.get("color"),
        "emoji":          profile.get("emoji"),
    }
    entry = next((p for p in content if str(p.get("client_id")) == str(cid)), None)
    if entry is None:
        content.append(new_entry)
    else:
        entry.update(new_entry)

    try:
        http_req.post(
            f"{SUPABASE_URL}/rest/v1/copy_bank_templates",
            headers=_sb_headers("resolution=merge-duplicates,return=minimal"),
            params={"on_conflict": "key"},
            json={"key": "__cb_profiles__", "content": content},
            timeout=10,
        )
    except Exception:
        pass


def _mirror_campaign_client_name(profile: dict) -> None:
    """Propagate a rename to the denormalised campaigns.client_name."""
    cid  = profile.get("id")
    name = (profile.get("name") or "").strip()
    prev = (profile.get("_prev_name") or "").strip()
    if not cid or not name or not prev or name == prev:
        return
    try:
        http_req.patch(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers("return=minimal"),
            params={"client_id": f"eq.{cid}"},
            json={"client_name": name},
            timeout=15,
        )
    except Exception:
        pass


# ── routes ──────────────────────────────────────────────────────────────────────

@router.get("/client-profiles")
async def client_profiles_page(request: Request):
    return templates.TemplateResponse(
        "client_profiles.html", {"request": request, "active": "client_profiles", "sop_key": "client_profiles"}
    )


@router.get("/api/client-profiles")
async def list_profiles(request: Request):
    return _list_response(request)


@router.get("/api/client-profiles/backfill", dependencies=[Depends(auth.require_admin)])
def backfill_profiles(save: str = Query("false")):
    """Seed the new clients profile columns from the existing sender_profiles and
    __cb_profiles__ stores so the manager opens fully populated. Dry-run unless
    save=true. Idempotent. Declared before /{profile_id} so the literal wins."""
    if not _sb_configured():
        return JSONResponse({"error": "Supabase is not configured."}, status_code=503)

    try:
        rc = http_req.get(
            f"{SUPABASE_URL}/rest/v1/clients",
            headers=_sb_headers(),
            params={"select": "id,name", "order": "name.asc"},
            timeout=15,
        )
        rc.raise_for_status()
        clients = rc.json()
    except Exception as exc:
        return JSONResponse({"error": f"clients fetch failed: {exc}"}, status_code=500)

    sp_by_name: dict[str, list] = {}
    try:
        rs = http_req.get(
            f"{SUPABASE_URL}/rest/v1/sender_profiles",
            headers=_sb_headers(),
            params={"select": "profile_name,emails"},
            timeout=15,
        )
        if rs.ok:
            for sp in rs.json():
                sp_by_name[(sp.get("profile_name") or "").strip()] = sp.get("emails") or []
    except Exception:
        pass

    cb_by_id: dict[str, dict] = {}
    try:
        rb = http_req.get(
            f"{SUPABASE_URL}/rest/v1/copy_bank_templates",
            headers=_sb_headers(),
            params={"key": "eq.__cb_profiles__", "select": "content"},
            timeout=15,
        )
        rows = rb.json() if rb.ok else []
        content = rows[0]["content"] if rows and isinstance(rows[0].get("content"), list) else []
        for p in content:
            cb_by_id[str(p.get("client_id"))] = p
    except Exception:
        pass

    do_save = save.lower() == "true"
    report: list[dict] = []
    for c in clients:
        cid  = str(c.get("id"))
        name = (c.get("name") or "").strip()
        patch: dict = {}
        cb = cb_by_id.get(cid)
        if cb:
            patch["territories"]     = cb.get("territories") or []
            patch["industries"]      = cb.get("industries") or []
            patch["industry_labels"] = cb.get("industryLabels") or {}
            patch["senders"]         = cb.get("senders") or []
            if cb.get("color"):
                patch["color"] = cb["color"]
            if cb.get("emoji"):
                patch["emoji"] = cb["emoji"]
            if cb.get("type"):
                patch["type"] = cb["type"]
        if name in sp_by_name:
            patch["sender_emails"] = sp_by_name[name]
        if not patch:
            continue

        entry = {"id": cid, "name": name, "patch": patch}
        if do_save:
            try:
                pr = http_req.patch(
                    f"{SUPABASE_URL}/rest/v1/clients",
                    headers=_sb_headers("return=minimal"),
                    params={"id": f"eq.{cid}"},
                    json=patch,
                    timeout=15,
                )
                entry["status"] = pr.status_code
            except Exception as exc:
                entry["error"] = str(exc)
        report.append(entry)

    return JSONResponse({"ok": True, "would_save": do_save, "count": len(report), "profiles": report})


@router.get("/api/client-profiles/{profile_id}")
async def get_profile(profile_id: str):
    """Single profile as JSON — used to prefill the edit form."""
    profile = _fetch_profile(profile_id)
    if profile is None:
        return JSONResponse({"error": "Profile not found."}, status_code=404)
    return JSONResponse(profile)


@router.post("/api/client-profiles")
async def create_profile(
    request:       Request,
    name:          str = Form(...),
    type:          str = Form("client"),
    territories:   str = Form(""),
    industries:    str = Form(""),
    senders:       str = Form(""),
    sender_emails: str = Form(""),
    color:         str = Form(""),
    emoji:         str = Form(""),
    booking_link:  str = Form(""),
    notes:         str = Form(""),
    active:        str = Form("true"),
):
    name_norm = name.strip()
    if not name_norm:
        return _error("Please enter a client name.")
    if not _sb_configured():
        return _error("Supabase is not configured.")

    payload = _build_payload(
        name=name, type=type, territories=territories, industries=industries,
        senders=senders, sender_emails=sender_emails, color=color, emoji=emoji,
        booking_link=booking_link, notes=notes, active=active,
    )

    # Try the full row; fall back to name-only if the extended columns aren't
    # migrated yet, so the app still works pre-migration.
    new_row = None
    last = ""
    for body in (payload, {"name": name_norm}):
        try:
            r = http_req.post(
                f"{SUPABASE_URL}/rest/v1/clients",
                headers=_sb_headers("return=representation"),
                json=body,
                timeout=10,
            )
            if r.status_code == 409:
                return _error(f'A client named "{name_norm}" already exists.')
            r.raise_for_status()
            new_row = r.json()[0]
            break
        except Exception as exc:
            last = str(exc)
    if new_row is None:
        return _error(f"Could not create profile: {last}")

    _mirror_async({**payload, "id": new_row["id"]})
    return _list_response(request, trigger=True)


@router.patch("/api/client-profiles/{profile_id}")
async def update_profile(
    request:       Request,
    profile_id:    str,
    name:          str = Form(...),
    type:          str = Form("client"),
    territories:   str = Form(""),
    industries:    str = Form(""),
    senders:       str = Form(""),
    sender_emails: str = Form(""),
    color:         str = Form(""),
    emoji:         str = Form(""),
    booking_link:  str = Form(""),
    notes:         str = Form(""),
    active:        str = Form("true"),
):
    name_norm = name.strip()
    if not name_norm:
        return _error("Please enter a client name.")
    if not _sb_configured():
        return _error("Supabase is not configured.")

    prev_name = (_fetch_profile(profile_id) or {}).get("name") or ""

    payload = _build_payload(
        name=name, type=type, territories=territories, industries=industries,
        senders=senders, sender_emails=sender_emails, color=color, emoji=emoji,
        booking_link=booking_link, notes=notes, active=active,
    )

    ok = False
    last = ""
    for body in (payload, {"name": name_norm}):
        try:
            r = http_req.patch(
                f"{SUPABASE_URL}/rest/v1/clients",
                headers=_sb_headers("return=minimal"),
                params={"id": f"eq.{profile_id}"},
                json=body,
                timeout=10,
            )
            if r.status_code == 409:
                return _error(f'A client named "{name_norm}" already exists.')
            r.raise_for_status()
            ok = True
            break
        except Exception as exc:
            last = str(exc)
    if not ok:
        return _error(f"Could not update profile: {last}")

    _mirror_async({**payload, "id": profile_id, "_prev_name": prev_name})
    return _list_response(request, trigger=True)


@router.delete("/api/client-profiles/{profile_id}", dependencies=[Depends(auth.require_admin)])
async def delete_profile(request: Request, profile_id: str):
    if not _sb_configured():
        return _error("Supabase is not configured.")
    try:
        r = http_req.delete(
            f"{SUPABASE_URL}/rest/v1/clients",
            headers=_sb_headers(),
            params={"id": f"eq.{profile_id}"},
            timeout=10,
        )
        r.raise_for_status()
    except Exception as exc:
        return _error(f"Could not delete profile: {exc}")
    # No trigger: deleting from the list must not reset an in-progress edit.
    return _list_response(request)
