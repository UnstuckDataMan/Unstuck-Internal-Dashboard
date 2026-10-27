"""
Storage for Read AI meetings, their AI overview, and onboarding notes.

Service-role only, like growth_store: these tables hold client conversations
and are closed to the anon key (migrations/outbound_pulse_readai.sql). Every
query is scoped to the agency, the same rule store.py follows.
"""
from __future__ import annotations

from datetime import datetime, timezone

import requests

from app.utils.supabase import SUPABASE_URL, sb_service_headers, sb_service_configured
from app.utils.pulse.store import PulseNotReady, current_agency_id, _describe_postgrest_error

MEETINGS = "pulse_meetings"
INSIGHTS = "pulse_meeting_insights"
NOTES = "pulse_onboarding_notes"

# Everything the panels show. The transcript is left out: it is the bulk of
# the row and only the transcript viewer and the AI read it.
LIST_COLUMNS = (
    "id,client_id,title,start_time,end_time,report_url,read_summary,participants,"
    "match_status,candidate_client_ids,is_onboarding,analysis,analysis_status,"
    "analysis_error,analysed_at,received_at"
)


class Conflict(ValueError):
    """Someone else saved the notes since this version was read."""


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _ready() -> None:
    if not sb_service_configured():
        raise PulseNotReady("Read AI storage needs SUPABASE_SERVICE_ROLE_KEY in Render.",
                            kind="setup")


def _request(method: str, table: str, *, params=None, json=None, prefer: str = ""):
    _ready()
    try:
        r = getattr(requests, method)(
            f"{SUPABASE_URL}/rest/v1/{table}",
            headers=sb_service_headers(prefer),
            params=params,
            json=json,
            timeout=20,
        )
    except requests.RequestException as exc:
        raise PulseNotReady(f"Query on {table} failed: {exc.__class__.__name__}") from exc
    if r.status_code >= 400 and r.status_code != 409:
        raise _describe_postgrest_error(table, r)
    return r


def _rows(r) -> list[dict]:
    try:
        body = r.json()
    except ValueError:
        return []
    return body if isinstance(body, list) else []


def _scope(**params) -> dict:
    return {"agency_id": f"eq.{current_agency_id()}", **params}


# ── Clients ───────────────────────────────────────────────────────────────────

def client_domains() -> list[dict]:
    """Every client with the domains it owns. Missing column -> PulseNotReady."""
    scope = f"(agency_id.eq.{current_agency_id()},agency_id.is.null)"
    r = _request("get", "clients", params={"select": "id,name,domains", "or": scope})
    return _rows(r)


# ── Meetings ──────────────────────────────────────────────────────────────────

def find_by_session(session_id: str) -> dict | None:
    r = _request("get", MEETINGS, params=_scope(
        session_id=f"eq.{session_id}",
        select="id,request_id,client_id,match_status,is_onboarding,transcript",
        limit="1"))
    rows = _rows(r)
    return rows[0] if rows else None


def insert_meeting(row: dict) -> dict | None:
    """Insert one meeting. None when the session already exists (a race)."""
    r = _request("post", MEETINGS, json={**row, "agency_id": current_agency_id()},
                 prefer="return=representation")
    if r.status_code == 409:
        return None
    rows = _rows(r)
    return rows[0] if rows else None


def update_meeting(meeting_id: str, patch: dict) -> None:
    _request("patch", MEETINGS, params=_scope(id=f"eq.{meeting_id}"),
             json={**patch, "updated_at": now_iso()}, prefer="return=minimal")


def get_meeting(meeting_id: str, *, with_transcript: bool = True) -> dict | None:
    columns = LIST_COLUMNS + (",transcript,action_items,owner_email" if with_transcript else "")
    r = _request("get", MEETINGS, params=_scope(id=f"eq.{meeting_id}", select=columns,
                                                limit="1"))
    rows = _rows(r)
    return rows[0] if rows else None


def list_client_meetings(client_id: str, *, with_transcript: bool = False) -> list[dict]:
    columns = LIST_COLUMNS + (",transcript" if with_transcript else "")
    r = _request("get", MEETINGS, params=_scope(
        client_id=f"eq.{client_id}", select=columns,
        order="start_time.desc.nullslast,received_at.desc", limit="200"))
    return _rows(r)


def list_inbox() -> list[dict]:
    r = _request("get", MEETINGS, params=_scope(
        client_id="is.null", match_status="in.(ambiguous,unmatched)",
        select=LIST_COLUMNS, order="received_at.desc", limit="50"))
    return _rows(r)


def onboarding_meeting(client_id: str) -> dict | None:
    r = _request("get", MEETINGS, params=_scope(
        client_id=f"eq.{client_id}", is_onboarding="is.true",
        select=LIST_COLUMNS, order="start_time.asc.nullslast", limit="1"))
    rows = _rows(r)
    return rows[0] if rows else None


def set_onboarding(client_id: str, meeting_id: str) -> None:
    """Make `meeting_id` the client's only onboarding meeting."""
    _request("patch", MEETINGS,
             params=_scope(client_id=f"eq.{client_id}", is_onboarding="is.true",
                           id=f"neq.{meeting_id}"),
             json={"is_onboarding": False, "updated_at": now_iso()}, prefer="return=minimal")
    update_meeting(meeting_id, {"is_onboarding": True})


# ── AI overview ───────────────────────────────────────────────────────────────

def get_insights(client_id: str) -> dict | None:
    r = _request("get", INSIGHTS, params=_scope(client_id=f"eq.{client_id}",
                                                select="*", limit="1"))
    rows = _rows(r)
    return rows[0] if rows else None


def save_insights(client_id: str, patch: dict) -> None:
    _request("post", INSIGHTS,
             params={"on_conflict": "agency_id,client_id"},
             json={**patch, "agency_id": current_agency_id(), "client_id": client_id,
                   "updated_at": now_iso()},
             prefer="resolution=merge-duplicates,return=minimal")


# ── Onboarding notes ──────────────────────────────────────────────────────────

def get_notes(client_id: str) -> dict | None:
    r = _request("get", NOTES, params=_scope(client_id=f"eq.{client_id}",
                                             select="*", limit="1"))
    rows = _rows(r)
    return rows[0] if rows else None


def write_notes(client_id: str, previous: dict | None, patch: dict) -> dict:
    """Optimistic write: fails with Conflict if the row moved since `previous`."""
    body = {**patch, "updated_at": now_iso(),
            "version": int((previous or {}).get("version") or 0) + 1}
    if previous:
        r = _request("patch", NOTES,
                     params=_scope(client_id=f"eq.{client_id}",
                                   version=f"eq.{previous.get('version')}"),
                     json=body, prefer="return=representation")
    else:
        r = _request("post", NOTES,
                     json={**body, "agency_id": current_agency_id(), "client_id": client_id},
                     prefer="return=representation")
    rows = _rows(r)
    if r.status_code == 409 or not rows:
        raise Conflict("Someone else changed these notes. Reload to see their version.")
    return rows[0]
