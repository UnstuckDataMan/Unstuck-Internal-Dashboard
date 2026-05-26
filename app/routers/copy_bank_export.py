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
    "https://www.googleapis.com/auth/documents",
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


# ── Google Doc creation ────────────────────────────────────────────────────────

class ExportSection(BaseModel):
    territory: str
    industry:  str
    channel:   str
    subjects:  List[str] = []
    bodies:    List[str] = []


class CreateDocBody(BaseModel):
    session_id:   str
    profile_name: str
    sections:     List[ExportSection]


@router.post("/api/export/google/create-doc")
def create_google_doc(body: CreateDocBody):
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
        docs_svc = build("docs",  "v1", credentials=creds)

        title = f"Copy Bank Export — {body.profile_name}"
        doc   = docs_svc.documents().create(body={"title": title}).execute()
        doc_id = doc["documentId"]

        reqs = _build_doc_requests(body)
        if reqs:
            docs_svc.documents().batchUpdate(
                documentId=doc_id,
                body={"requests": reqs},
            ).execute()

        return JSONResponse({"ok": True, "url": f"https://docs.google.com/document/d/{doc_id}/edit"})
    except Exception as exc:
        log.error("Google Doc creation failed: %s", exc)
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=500)


# ── Google Doc formatting ──────────────────────────────────────────────────────

def _build_doc_requests(body: CreateDocBody) -> list:
    """
    Build a Google Docs batchUpdate request list that populates the document
    with styled content reflecting the Unstuck brand colours.
    """
    PURPLE = {"red": 0.486, "green": 0.227, "blue": 0.929}   # #7c3aed
    GREY   = {"red": 0.420, "green": 0.396, "blue": 0.475}
    DARK   = {"red": 0.122, "green": 0.098, "blue": 0.176}

    date_str = datetime.datetime.now().strftime("%-d %B %Y")

    # Build list of (text, style_tag)
    segs = []
    segs.append((f"Copy Bank Export — {body.profile_name}\n", "doc_title"))
    segs.append((f"Generated {date_str}\n\n",               "doc_subtitle"))

    for sec in body.sections:
        segs.append((f"{sec.territory} — {sec.industry}\n", "section_h1"))
        segs.append((f"{sec.channel}\n",                    "channel_h2"))

        if sec.subjects:
            segs.append(("Subject Lines\n", "label"))
            for i, s in enumerate(sec.subjects, 1):
                segs.append((f"{i}. {s}\n", "body"))
            segs.append(("\n", "body"))

        if sec.bodies:
            segs.append(("Variations\n", "label"))
            for i, b in enumerate(sec.bodies, 1):
                segs.append((f"Variation {i}\n", "variation_label"))
                segs.append((f"{b}\n\n",         "body"))

    full_text = "".join(s[0] for s in segs)
    if not full_text.strip():
        return []

    requests = [{"insertText": {"location": {"index": 1}, "text": full_text}}]

    pos = 1
    for text, style in segs:
        length   = len(text)
        end      = pos + length
        text_end = end - 1 if text.endswith("\n") else end   # exclude trailing newline from style range

        if style == "doc_title":
            requests.append({"updateTextStyle": {
                "range": {"startIndex": pos, "endIndex": text_end},
                "textStyle": {"bold": True,
                              "fontSize": {"magnitude": 22, "unit": "PT"},
                              "foregroundColor": {"color": {"rgbColor": PURPLE}}},
                "fields": "bold,fontSize,foregroundColor",
            }})
        elif style == "doc_subtitle":
            requests.append({"updateTextStyle": {
                "range": {"startIndex": pos, "endIndex": text_end},
                "textStyle": {"italic": True,
                              "fontSize": {"magnitude": 10, "unit": "PT"},
                              "foregroundColor": {"color": {"rgbColor": GREY}}},
                "fields": "italic,fontSize,foregroundColor",
            }})
        elif style == "section_h1":
            requests.append({"updateTextStyle": {
                "range": {"startIndex": pos, "endIndex": text_end},
                "textStyle": {"bold": True,
                              "fontSize": {"magnitude": 15, "unit": "PT"},
                              "foregroundColor": {"color": {"rgbColor": DARK}}},
                "fields": "bold,fontSize,foregroundColor",
            }})
            requests.append({"updateParagraphStyle": {
                "range": {"startIndex": pos, "endIndex": end},
                "paragraphStyle": {
                    "spaceAbove": {"magnitude": 18, "unit": "PT"},
                    "spaceBelow": {"magnitude": 2, "unit": "PT"},
                    "borderBottom": {
                        "color": {"color": {"rgbColor": PURPLE}},
                        "width": {"magnitude": 1, "unit": "PT"},
                        "padding": {"magnitude": 4, "unit": "PT"},
                        "dashStyle": "SOLID",
                    }
                },
                "fields": "spaceAbove,spaceBelow,borderBottom",
            }})
        elif style == "channel_h2":
            requests.append({"updateTextStyle": {
                "range": {"startIndex": pos, "endIndex": text_end},
                "textStyle": {"bold": True,
                              "fontSize": {"magnitude": 12, "unit": "PT"},
                              "foregroundColor": {"color": {"rgbColor": PURPLE}}},
                "fields": "bold,fontSize,foregroundColor",
            }})
            requests.append({"updateParagraphStyle": {
                "range": {"startIndex": pos, "endIndex": end},
                "paragraphStyle": {"spaceBelow": {"magnitude": 8, "unit": "PT"}},
                "fields": "spaceBelow",
            }})
        elif style == "label":
            requests.append({"updateTextStyle": {
                "range": {"startIndex": pos, "endIndex": text_end},
                "textStyle": {"bold": True, "smallCaps": True,
                              "fontSize": {"magnitude": 9, "unit": "PT"},
                              "foregroundColor": {"color": {"rgbColor": GREY}}},
                "fields": "bold,smallCaps,fontSize,foregroundColor",
            }})
        elif style == "variation_label":
            requests.append({"updateTextStyle": {
                "range": {"startIndex": pos, "endIndex": text_end},
                "textStyle": {"bold": True, "italic": True,
                              "fontSize": {"magnitude": 10, "unit": "PT"},
                              "foregroundColor": {"color": {"rgbColor": PURPLE}}},
                "fields": "bold,italic,fontSize,foregroundColor",
            }})
        pos = end

    return requests
