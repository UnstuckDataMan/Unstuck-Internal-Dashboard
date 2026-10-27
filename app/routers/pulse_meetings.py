"""
Read AI meetings on the reporting tool: the webhook that receives them, and
the two Overview panels built from them (Onboarding notes, Read AI summary),
plus the inbox for meetings that matched no client.

Internal only. Nothing here reaches the client's report link or a report
snapshot: the summary names pain points in the client's own words, which is
for the account manager, not the client.

The webhook is public to the auth gate (app/auth.py) and checks Read AI's
HMAC signature itself. Everything under /api/outbound-pulse is gated by the
outbound_pulse tool like the rest of the reporting tool.
"""
from __future__ import annotations

import json
import logging
import os
from datetime import datetime

from fastapi import APIRouter, Depends, Form, Query, Request
from fastapi.responses import HTMLResponse, JSONResponse

from app import auth
from app.deps import templates
from app.utils import llm
from app.utils.pulse import meeting_insights as mi
from app.utils.pulse import meetings_store as ms
from app.utils.pulse import readai, richtext, store
from app.utils.pulse.store import PulseNotReady

router = APIRouter()
logger = logging.getLogger(__name__)

# Read AI posts the whole transcript; a long meeting is a few MB at most.
_MAX_WEBHOOK_BYTES = 15_000_000
_MAX_NOTES_CHARS = richtext.MAX_BODY_CHARS


def _webhook_secret() -> str:
    return os.environ.get("READ_AI_WEBHOOK_SECRET", "").strip()


# ── Webhook ───────────────────────────────────────────────────────────────────

@router.post("/webhooks/read-ai")
async def read_ai_webhook(request: Request):
    """Receive one Read AI meeting report.

    Returns quickly: the AI work runs in a background thread. A 5xx makes Read
    AI retry (up to five times), so storage failures answer 503; a body that
    will never parse answers 400 so it is not retried pointlessly.
    """
    secret = _webhook_secret()
    if not secret:
        return JSONResponse({"error": "Read AI webhook is not configured."}, status_code=503)
    raw = await request.body()
    if len(raw) > _MAX_WEBHOOK_BYTES:
        return JSONResponse({"error": "Payload too large."}, status_code=413)
    if not readai.verify_signature(raw, request.headers.get("X-Read-Signature", ""), secret):
        return JSONResponse({"error": "Invalid signature."}, status_code=401)
    try:
        parsed = readai.parse_payload(json.loads(raw))
    except (ValueError, TypeError) as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)
    if parsed["trigger"] != "meeting_end":
        return JSONResponse({"result": "ignored", "reason": "only meeting_end is used"})
    try:
        outcome = mi.ingest(parsed)
    except PulseNotReady as exc:
        logger.error("Read AI webhook could not store session %s: %s",
                     parsed["session_id"], exc)
        return JSONResponse({"error": "Storage unavailable."}, status_code=503)
    return JSONResponse(outcome)


# ── Shared view helpers ───────────────────────────────────────────────────────

def _date(value, *, short: bool = False) -> str:
    """"12 Oct 2026" (or "12 Oct") from an ISO timestamp; "" when unknown."""
    if not value:
        return ""
    try:
        d = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return ""
    return f"{d.day} {d:%b}" if short else f"{d.day} {d:%b %Y}"


def _clients() -> list[dict]:
    return store.list_clients(active_only=True)


def _client(client_id: str) -> dict | None:
    return next((c for c in store.list_clients() if str(c["id"]) == client_id), None)


def _state(request: Request, kind: str, title: str, body: str, *, target: str = "",
           retry_url: str = "") -> HTMLResponse:
    return templates.TemplateResponse("partials/pulse_state.html", {
        "request": request, "kind": kind, "title": title, "body": body,
        "target": target, "retry_url": retry_url,
    })


def _not_ready(request: Request, exc: PulseNotReady, *, target: str, retry_url: str):
    kind = getattr(exc, "kind", "error")
    if kind == "missing":
        return _state(request, "setup", "Read AI storage isn't set up yet",
                      "Run migrations/outbound_pulse_readai.sql in the Supabase SQL editor, "
                      "then reload this page.")
    if kind == "setup":
        return _state(request, "setup", "Read AI storage isn't set up yet", str(exc))
    logger.warning("Read AI panel failed: %s", exc)
    return _state(request, "error", "Couldn't load this panel",
                  "The database didn't respond. Try again in a moment.",
                  target=target, retry_url=retry_url)


def _meeting_view(m: dict) -> dict:
    analysis = m.get("analysis") or {}
    domains = sorted(readai.external_domains(m.get("participants") or []))
    report_url = str(m.get("report_url") or "")
    return {
        **m,
        "id": str(m["id"]),
        # Rendered as a link, so only a web address gets through.
        "report_url": report_url if report_url.startswith("https://") else "",
        "date": _date(m.get("start_time") or m.get("received_at")),
        "short_date": _date(m.get("start_time") or m.get("received_at"), short=True),
        "positives": analysis.get("positives") or [],
        "negatives": analysis.get("negatives") or [],
        "summary": analysis.get("summary") or "",
        "truncated": bool(analysis.get("truncated")),
        "domains": domains,
    }


# ── Onboarding notes panel ────────────────────────────────────────────────────

def _notes_panel(request: Request, client_id: str, *, editing: bool = False,
                 draft_body: str | None = None, error: str = "", notice: str = ""):
    url = f"/api/outbound-pulse/clients/{client_id}/onboarding-notes"
    try:
        notes = ms.get_notes(client_id)
        onboarding = ms.onboarding_meeting(client_id)
        source = None
        if notes and notes.get("source_meeting_id"):
            source = (onboarding if onboarding and str(onboarding["id"]) == str(notes["source_meeting_id"])
                      else ms.get_meeting(str(notes["source_meeting_id"]), with_transcript=False))
        pending_source = None
        if notes and notes.get("pending_meeting_id"):
            pending_source = ms.get_meeting(str(notes["pending_meeting_id"]), with_transcript=False)
    except PulseNotReady as exc:
        return _not_ready(request, exc, target="#pulse-onboarding", retry_url=url)

    notes = notes or {}
    has_body = bool(richtext.to_text(notes.get("body") or "").strip())
    return templates.TemplateResponse("partials/pulse_onboarding_notes.html", {
        "request": request,
        "client_id": client_id,
        "notes": notes,
        "has_body": has_body,
        "body_for_editor": (notes.get("body") or "") if draft_body is None else draft_body,
        "version": notes.get("version") or 0,
        "editing": editing,
        "error": error,
        "notice": notice,
        "source": _meeting_view(source) if source else None,
        "pending_source": _meeting_view(pending_source) if pending_source else None,
        "edited_at": _date(notes.get("edited_at")),
        "generated_at": _date(notes.get("ai_generated_at")),
        "onboarding": _meeting_view(onboarding) if onboarding else None,
        "drafting": mi.is_drafting(client_id),
        "ai_ready": llm.is_configured(),
    })


@router.get("/api/outbound-pulse/clients/{client_id}/onboarding-notes")
async def onboarding_notes(request: Request, client_id: str):
    return _notes_panel(request, client_id)


def _expected_version(raw: str) -> int:
    try:
        return int(raw or 0)
    except ValueError:
        return -1


@router.post("/api/outbound-pulse/clients/{client_id}/onboarding-notes")
async def save_onboarding_notes(
    request:   Request,
    client_id: str,
    body:      str = Form(""),
    version:   str = Form("0"),
    user:      dict = Depends(auth.require_login),
):
    clean = richtext.sanitize(body[: _MAX_NOTES_CHARS * 2])
    try:
        previous = ms.get_notes(client_id)
        if _expected_version(version) != int((previous or {}).get("version") or 0):
            raise ms.Conflict("Someone else changed these notes. Your text is still here: "
                              "copy it, reload to see their version, then add it back.")
        ms.write_notes(client_id, previous, {
            "body": clean,
            "edited_by": user.get("name") or user.get("email") or "",
            "edited_at": ms.now_iso(),
        })
    except ms.Conflict as exc:
        return _notes_panel(request, client_id, editing=True, draft_body=clean, error=str(exc))
    except PulseNotReady:
        return _notes_panel(request, client_id, editing=True, draft_body=clean,
                            error="Couldn't save the notes. The database didn't respond. "
                                  "Your text is still here; try again.")
    return _notes_panel(request, client_id, notice="Notes saved.")


@router.post("/api/outbound-pulse/clients/{client_id}/onboarding-notes/draft/{action}")
async def onboarding_draft_action(
    request:   Request,
    client_id: str,
    action:    str,
    version:   str = Form("0"),
    user:      dict = Depends(auth.require_login),
):
    if action not in ("apply", "dismiss"):
        return _notes_panel(request, client_id, error="Unknown action.")
    try:
        previous = ms.get_notes(client_id)
        if not previous or not previous.get("pending_draft"):
            return _notes_panel(request, client_id)
        if _expected_version(version) != int(previous.get("version") or 0):
            raise ms.Conflict("Someone else changed these notes. Reload to see their version.")
        patch = {"pending_draft": "", "pending_meeting_id": None}
        if action == "apply":
            patch.update({
                "body": previous["pending_draft"],
                "source_meeting_id": previous.get("pending_meeting_id"),
                "ai_generated_at": ms.now_iso(),
                "edited_by": user.get("name") or user.get("email") or "",
                "edited_at": ms.now_iso(),
            })
        ms.write_notes(client_id, previous, patch)
    except ms.Conflict as exc:
        return _notes_panel(request, client_id, error=str(exc))
    except PulseNotReady as exc:
        return _not_ready(request, exc, target="#pulse-onboarding",
                          retry_url=f"/api/outbound-pulse/clients/{client_id}/onboarding-notes")
    return _notes_panel(request, client_id,
                        notice="Notes replaced with the new draft." if action == "apply"
                        else "Kept your notes. The new draft was discarded.")


@router.post("/api/outbound-pulse/clients/{client_id}/onboarding-notes/regenerate")
async def regenerate_onboarding_notes(request: Request, client_id: str):
    if not llm.is_configured():
        return _notes_panel(request, client_id,
                            error="OpenAI isn't connected, so notes can't be drafted. "
                                  "Ask an admin to add OPENAI_API_KEY in Render.")
    try:
        meeting = ms.onboarding_meeting(client_id)
    except PulseNotReady as exc:
        return _not_ready(request, exc, target="#pulse-onboarding",
                          retry_url=f"/api/outbound-pulse/clients/{client_id}/onboarding-notes")
    if not meeting:
        return _notes_panel(request, client_id,
                            error="No onboarding meeting yet. Mark one in the Read AI summary first.")
    mi._drafting.add(client_id)
    mi.run_async(mi.redraft_notes, client_id, str(meeting["id"]))
    return _notes_panel(request, client_id)


# ── Read AI summary panel ─────────────────────────────────────────────────────

def _readai_panel(request: Request, client_id: str, *, notice: str = "", error: str = ""):
    url = f"/api/outbound-pulse/clients/{client_id}/read-ai"
    try:
        meetings = [_meeting_view(m) for m in ms.list_client_meetings(client_id)]
        insights = ms.get_insights(client_id) or {}
    except PulseNotReady as exc:
        return _not_ready(request, exc, target="#pulse-readai", retry_url=url)

    try:
        domains = next((c.get("domains") or [] for c in ms.client_domains()
                        if str(c["id"]) == client_id), [])
    except PulseNotReady:
        domains = None  # column not migrated; the template says so

    labels = {m["id"]: f"{m['title']} · {m['short_date']}" if m["short_date"] else m["title"]
              for m in meetings}
    with_findings = {m["id"] for m in meetings
                     if m.get("analysis_status") == "done" and (m["positives"] or m["negatives"])}
    built_from = {str(i) for i in insights.get("meeting_ids") or []}
    analysing = any(m.get("analysis_status") in ("pending", "running") for m in meetings)
    building = insights.get("status") in ("pending", "running")
    overview = insights.get("overview") or {}
    return templates.TemplateResponse("partials/pulse_readai.html", {
        "request": request,
        "client_id": client_id,
        "meetings": meetings,
        "overview": overview,
        "has_overview": bool(overview.get("summary") or overview.get("successes")
                             or overview.get("pain_points")),
        "insights": insights,
        "generated_at": _date(insights.get("generated_at")),
        "stale": bool(with_findings) and with_findings != built_from and not building,
        "busy": analysing or building,
        "labels": labels,
        "domains": domains,
        "readai_ready": bool(_webhook_secret()),
        "ai_ready": llm.is_configured(),
        "notice": notice,
        "error": error,
    })


@router.get("/api/outbound-pulse/clients/{client_id}/read-ai")
async def read_ai_panel(request: Request, client_id: str):
    return _readai_panel(request, client_id)


@router.post("/api/outbound-pulse/clients/{client_id}/read-ai/refresh")
async def refresh_read_ai(request: Request, client_id: str):
    if not llm.is_configured():
        return _readai_panel(request, client_id,
                             error="OpenAI isn't connected, so the summary can't be refreshed. "
                                   "Ask an admin to add OPENAI_API_KEY in Render.")
    try:
        ms.save_insights(client_id, {"status": "pending", "error": ""})
    except PulseNotReady as exc:
        return _not_ready(request, exc, target="#pulse-readai",
                          retry_url=f"/api/outbound-pulse/clients/{client_id}/read-ai")
    mi.run_async(mi.refresh_client, client_id)
    return _readai_panel(request, client_id)


# ── One meeting ───────────────────────────────────────────────────────────────

@router.get("/api/outbound-pulse/read-ai/meetings/{meeting_id}")
async def meeting_transcript(request: Request, meeting_id: str, block: int = Query(-1)):
    try:
        meeting = ms.get_meeting(meeting_id)
    except PulseNotReady:
        meeting = None
        error = "Couldn't load this transcript. The database didn't respond. Close and try again."
    else:
        error = "" if meeting else "This meeting no longer exists."
    return templates.TemplateResponse("partials/pulse_readai_transcript.html", {
        "request": request,
        "meeting": _meeting_view(meeting) if meeting else None,
        "blocks": _with_clock((meeting or {}).get("transcript") or []),
        "highlight": block,
        "error": error,
    })


def _with_clock(blocks: list[dict]) -> list[dict]:
    """Each block with "12:05", its offset from the start of the meeting."""
    starts = [b.get("start_ms") for b in blocks if isinstance(b.get("start_ms"), int)]
    origin = min(starts) if starts else None
    out = []
    for b in blocks:
        clock = ""
        if origin is not None and isinstance(b.get("start_ms"), int):
            seconds = max(0, (b["start_ms"] - origin) // 1000)
            clock = f"{seconds // 60}:{seconds % 60:02d}"
        out.append({**b, "clock": clock})
    return out


def _after_meeting_change(request: Request, view: str, client_id: str, notice: str = ""):
    if view == "inbox":
        return _inbox_panel(request, notice=notice)
    return _readai_panel(request, client_id, notice=notice)


@router.post("/api/outbound-pulse/read-ai/meetings/{meeting_id}/assign")
async def assign_meeting(
    request:    Request,
    meeting_id: str,
    client_id:  str = Form(""),
    view:       str = Form("client"),
):
    """Assign a meeting to a client, unlink it (""), or ignore it ("ignore")."""
    try:
        meeting = ms.get_meeting(meeting_id, with_transcript=False)
        if not meeting:
            return _after_meeting_change(request, view, client_id, "That meeting no longer exists.")
        previous_client = str(meeting.get("client_id") or "")
        if client_id == "ignore":
            ms.update_meeting(meeting_id, {"client_id": None, "match_status": "ignored",
                                           "is_onboarding": False})
            return _after_meeting_change(request, view, previous_client,
                                         f"Ignored “{meeting.get('title')}”.")
        if not client_id:
            ms.update_meeting(meeting_id, {"client_id": None, "match_status": "unmatched",
                                           "is_onboarding": False})
            if previous_client:
                mi.run_async(mi.refresh_client, previous_client)
            return _after_meeting_change(request, view, previous_client,
                                         f"Unlinked “{meeting.get('title')}”. "
                                         "It's waiting in the Read AI inbox on the reporting overview.")
        client = _client(client_id)
        if not client:
            return _after_meeting_change(request, view, previous_client, "Choose a client first.")
        onboarding = mi.should_mark_onboarding(meeting, client_id)
        ms.update_meeting(meeting_id, {"client_id": client_id, "match_status": "manual",
                                       "is_onboarding": onboarding})
        mi.run_async(mi.process_meeting, meeting_id, draft_notes=onboarding)
        if previous_client and previous_client != client_id:
            mi.run_async(mi.refresh_client, previous_client)
        return _after_meeting_change(request, view, client_id,
                                     f"Assigned “{meeting.get('title')}” to {client['name']}.")
    except PulseNotReady as exc:
        if view == "inbox":
            return HTMLResponse('<div class="err-box">Couldn’t assign the meeting. '
                                'The database didn’t respond. Try again.</div>')
        return _not_ready(request, exc, target="#pulse-readai",
                          retry_url=f"/api/outbound-pulse/clients/{client_id}/read-ai")


@router.post("/api/outbound-pulse/read-ai/meetings/{meeting_id}/onboarding")
async def mark_onboarding(request: Request, meeting_id: str):
    try:
        meeting = ms.get_meeting(meeting_id, with_transcript=False)
        if not meeting or not meeting.get("client_id"):
            return HTMLResponse('<div class="err-box">That meeting isn’t assigned to a client.</div>')
        client_id = str(meeting["client_id"])
        ms.set_onboarding(client_id, meeting_id)
    except PulseNotReady as exc:
        return _not_ready(request, exc, target="#pulse-readai", retry_url="")
    if llm.is_configured():
        mi._drafting.add(client_id)
        mi.run_async(mi.redraft_notes, client_id, meeting_id)
    response = _readai_panel(request, client_id,
                             notice=f"“{meeting.get('title')}” is now the onboarding meeting.")
    # The notes panel listens for this and reloads to show the new draft.
    response.headers["HX-Trigger"] = "onboardingChanged"
    return response


@router.post("/api/outbound-pulse/read-ai/meetings/{meeting_id}/retry")
async def retry_meeting(request: Request, meeting_id: str):
    try:
        meeting = ms.get_meeting(meeting_id, with_transcript=False)
        if not meeting or not meeting.get("client_id"):
            return HTMLResponse('<div class="err-box">That meeting isn’t assigned to a client.</div>')
        ms.update_meeting(meeting_id, {"analysis_status": "pending", "analysis_error": ""})
    except PulseNotReady as exc:
        return _not_ready(request, exc, target="#pulse-readai", retry_url="")
    mi.run_async(mi.process_meeting, meeting_id)
    return _readai_panel(request, str(meeting["client_id"]))


# ── Inbox ─────────────────────────────────────────────────────────────────────

_STATUS_LABELS = {
    "matched": ("Filed", "success"),
    "manual": ("Filed", "success"),
    "ambiguous": ("Waiting to assign", "warning"),
    "unmatched": ("Waiting to assign", "warning"),
    "ignored": ("Ignored", ""),
}


def _attendee_note(m: dict) -> str:
    """Why a meeting was filed where it was, in the attendees' terms."""
    people = m.get("participants") or []
    if not people:
        return "Read AI sent no attendees"
    if m["domains"]:
        return ", ".join(m["domains"])
    if any(p.get("email") for p in people):
        return "only Unstuck or personal email addresses"
    return "no attendee emails from Read AI"


def _inbox_panel(request: Request, *, notice: str = ""):
    """Meetings waiting for a client, plus the latest arrivals and what happened.

    The recent list is what shows the webhook is working: without it, a
    meeting filed as internal, or one that never arrived, looked the same as
    an empty inbox.
    """
    try:
        waiting = [_meeting_view(m) for m in ms.list_inbox()]
        recent = [_meeting_view(m) for m in ms.list_recent()]
        clients = _clients()
    except PulseNotReady as exc:
        if getattr(exc, "kind", "") in ("missing", "setup"):
            # Not set up yet: the overview page gets no box for it.
            return HTMLResponse("")
        logger.warning("Read AI inbox failed: %s", exc)
        return templates.TemplateResponse("partials/pulse_readai_inbox.html", {
            "request": request, "error": True, "waiting": [], "recent": [],
            "clients": [], "notice": notice, "readai_ready": bool(_webhook_secret()),
        })
    if not waiting and not recent and not notice and not _webhook_secret():
        return HTMLResponse("")
    names = {str(c["id"]): c["name"] for c in store.list_clients()}
    for m in waiting + recent:
        m["candidates"] = [{"id": str(i), "name": names.get(str(i), "")}
                           for i in m.get("candidate_client_ids") or [] if str(i) in names]
        m["client_name"] = names.get(str(m.get("client_id") or ""), "")
        m["status_label"], m["status_tone"] = _STATUS_LABELS.get(
            m.get("match_status"), ("Waiting to assign", "warning"))
        m["attendees"] = _attendee_note(m)
        m["received"] = _date(m.get("received_at"))
    return templates.TemplateResponse("partials/pulse_readai_inbox.html", {
        "request": request, "waiting": waiting, "recent": recent, "clients": clients,
        "notice": notice, "error": False, "readai_ready": bool(_webhook_secret()),
    })


@router.get("/api/outbound-pulse/read-ai/inbox")
async def read_ai_inbox(request: Request):
    return _inbox_panel(request)
