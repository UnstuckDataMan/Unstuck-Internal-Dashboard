import datetime
import logging
import os
import uuid

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from pydantic import BaseModel
from typing import List

log = logging.getLogger(__name__)
router = APIRouter()

# In-memory OAuth sessions: state_id → credential dict
# (fine for a single-server internal tool; use Redis for multi-instance)
_sessions: dict = {}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

# ── Helpers ────────────────────────────────────────────────────────────────────

def _google_configured() -> bool:
    return bool(os.environ.get("GOOGLE_CLIENT_ID") and os.environ.get("GOOGLE_CLIENT_SECRET"))


def _get_redirect_uri(request: Request) -> str:
    base = os.environ.get("APP_BASE_URL", "").rstrip("/")
    if not base:
        base = str(request.base_url).rstrip("/")
    return f"{base}/api/export/google/callback"


def _client_config() -> dict:
    return {
        "web": {
            "client_id":     os.environ.get("GOOGLE_CLIENT_ID", ""),
            "client_secret": os.environ.get("GOOGLE_CLIENT_SECRET", ""),
            "auth_uri":      "https://accounts.google.com/o/oauth2/auth",
            "token_uri":     "https://oauth2.googleapis.com/token",
        }
    }


# ── OAuth endpoints ────────────────────────────────────────────────────────────

@router.get("/api/export/google/status")
def google_status():
    """Check whether Google integration credentials are configured."""
    return JSONResponse({"configured": _google_configured()})


@router.get("/api/export/google/auth")
def google_auth(request: Request):
    """Redirect to Google OAuth consent screen (called inside a popup window)."""
    if not _google_configured():
        return HTMLResponse(
            "<script>window.opener&&window.opener.postMessage("
            "{type:'google_auth_error',error:'Google credentials not configured — set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET'},'*');"
            "window.close();</script>"
        )
    try:
        from google_auth_oauthlib.flow import Flow
    except ImportError:
        return HTMLResponse(
            "<script>window.opener&&window.opener.postMessage("
            "{type:'google_auth_error',error:'google-auth-oauthlib not installed'},'*');"
            "window.close();</script>"
        )

    state = str(uuid.uuid4())
    flow  = Flow.from_client_config(_client_config(), scopes=SCOPES,
                                    redirect_uri=_get_redirect_uri(request))
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        state=state,
        prompt="consent",
    )
    # Persist code_verifier so the callback can complete PKCE exchange
    _sessions[state] = {"code_verifier": getattr(flow, "code_verifier", None)}
    return RedirectResponse(auth_url)


@router.get("/api/export/google/callback", name="google_callback")
def google_callback(request: Request,
                    code: str  = None,
                    state: str = None,
                    error: str = None):
    """OAuth callback — exchanges code for token and closes the popup."""

    def _close(msg_type: str, **kw):
        import json as _json
        payload = _json.dumps({"type": msg_type, **kw})
        return HTMLResponse(
            f"<script>window.opener&&window.opener.postMessage({payload},'*');window.close();</script>"
        )

    if error or not code:
        return _close("google_auth_error", error=error or "cancelled")

    try:
        from google_auth_oauthlib.flow import Flow
    except ImportError:
        return _close("google_auth_error", error="google-auth-oauthlib not installed")

    try:
        flow = Flow.from_client_config(_client_config(), scopes=SCOPES,
                                       redirect_uri=_get_redirect_uri(request))
        # Restore code_verifier from the auth step to satisfy PKCE
        prior = _sessions.get(state, {})
        if prior.get("code_verifier"):
            flow.code_verifier = prior["code_verifier"]
        flow.fetch_token(code=code)
        creds = flow.credentials
        _sessions[state] = {
            "token":         creds.token,
            "refresh_token": creds.refresh_token,
            "token_uri":     creds.token_uri,
            "client_id":     creds.client_id,
            "client_secret": creds.client_secret,
            "scopes":        list(creds.scopes or SCOPES),
        }
        return _close("google_auth_success", session_id=state)
    except Exception as exc:
        log.error("Google OAuth callback error: %s", exc)
        return _close("google_auth_error", error=str(exc))


# ── Google Sheets creation ─────────────────────────────────────────────────────

class ExportSection(BaseModel):
    territory: str
    industry:  str
    channel:   str
    subjects:  List[str] = []
    bodies:    List[str] = []


class CreateSheetBody(BaseModel):
    session_id:   str
    profile_name: str
    sections:     List[ExportSection]


@router.post("/api/export/google/create-sheet")
def create_google_sheet(body: CreateSheetBody):
    session = _sessions.get(body.session_id)
    if not session:
        return JSONResponse(
            {"ok": False, "error": "Not authenticated with Google. Please connect first."},
            status_code=401,
        )
    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
    except ImportError:
        return JSONResponse(
            {"ok": False, "error": "google-api-python-client not installed"},
            status_code=500,
        )
    try:
        creds = Credentials(
            token=session["token"],
            refresh_token=session.get("refresh_token"),
            token_uri=session["token_uri"],
            client_id=session["client_id"],
            client_secret=session["client_secret"],
            scopes=session["scopes"],
        )
        sheets_svc = build("sheets", "v4", credentials=creds)

        date_str = datetime.datetime.now().strftime("%d %B %Y")
        title    = f"Copy Bank — {body.profile_name} — {date_str}"

        # Create the spreadsheet
        spreadsheet = sheets_svc.spreadsheets().create(body={
            "properties": {"title": title},
            "sheets": [{"properties": {"title": "Copy Bank"}}],
        }).execute()
        sheet_id     = spreadsheet["spreadsheetId"]
        sheet_tab_id = spreadsheet["sheets"][0]["properties"]["sheetId"]

        # Build rows: header + one row per piece of copy
        rows = [["Territory", "Industry", "Channel", "Type", "#", "Content"]]
        for sec in body.sections:
            for i, s in enumerate(sec.subjects, 1):
                rows.append([sec.territory, sec.industry, sec.channel, "Subject", i, s])
            label = "Step" if sec.channel == "LinkedIn" else "Variation"
            for i, b in enumerate(sec.bodies, 1):
                rows.append([sec.territory, sec.industry, sec.channel, label, i, b])

        # Write data
        sheets_svc.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range="Copy Bank!A1",
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()

        # Format header row: bold + purple background
        sheets_svc.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={"requests": [
                {"repeatCell": {
                    "range": {"sheetId": sheet_tab_id, "startRowIndex": 0, "endRowIndex": 1},
                    "cell": {"userEnteredFormat": {
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                        "backgroundColor": {"red": 0.486, "green": 0.227, "blue": 0.929},
                    }},
                    "fields": "userEnteredFormat(textFormat,backgroundColor)",
                }},
                # Freeze header row
                {"updateSheetProperties": {
                    "properties": {"sheetId": sheet_tab_id, "gridProperties": {"frozenRowCount": 1}},
                    "fields": "gridProperties.frozenRowCount",
                }},
                # Auto-resize all columns
                {"autoResizeDimensions": {
                    "dimensions": {"sheetId": sheet_tab_id, "dimension": "COLUMNS",
                                   "startIndex": 0, "endIndex": 6},
                }},
            ]},
        ).execute()

        url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
        return JSONResponse({"ok": True, "url": url})
    except Exception as exc:
        log.error("Google Sheets creation failed: %s", exc)
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=500)
