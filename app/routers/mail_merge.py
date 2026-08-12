"""
Mail Merge API endpoints — FastAPI port of mail_merge/app.py.

Reuses the existing utility modules at mail_merge/utils/ (merge engine,
scheduler, Excel reader/writer) so no business logic is duplicated.
"""
from __future__ import annotations

import re
import sys
import uuid
from pathlib import Path

import requests as http_req
from fastapi import APIRouter, Depends, File, Request, UploadFile
from fastapi.responses import JSONResponse

from app import auth

# ── Ensure mail_merge/utils is importable ─────────────────────────────────────
_MM_DIR = Path(__file__).resolve().parents[2] / "mail_merge"
if str(_MM_DIR) not in sys.path:
    sys.path.insert(0, str(_MM_DIR))

from utils.excel_reader import parse_prospect_file      # noqa: E402
from utils.merge import validate_templates, perform_merge, reassign_templates  # noqa: E402
from utils.scheduler import generate_schedule, _dvariance  # noqa: E402
from utils.excel_writer import write_merge_output          # noqa: E402

from app.utils.supabase import (                            # noqa: E402
    SUPABASE_URL,
    sb_headers as _sb_headers,
    sb_configured as _sb_configured,
)

# ── Config ────────────────────────────────────────────────────────────────────
router = APIRouter()

UPLOAD_DIR = _MM_DIR / "uploads"
OUTPUT_DIR = _MM_DIR / "outputs"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ALLOWED_EXTENSIONS = {"xlsx", "xls"}
MAX_UPLOAD_BYTES   = 32 * 1024 * 1024  # 32 MB

_SAFE_RE = re.compile(r"[^\w\s\-.]", re.ASCII)


def _secure_name(name: str) -> str:
    """Simple filename sanitiser (avoids werkzeug dependency)."""
    name = name.replace("/", "_").replace("\\", "_").replace("..", "_")
    return _SAFE_RE.sub("", name).strip("._") or "file"


def _allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ── 1. Upload Prospects ───────────────────────────────────────────────────────

@router.post("/api/merge/upload-prospects")
async def upload_prospects(file: UploadFile = File(...)):
    if not file.filename or not _allowed(file.filename):
        return JSONResponse(
            {"error": "Invalid file type. Please upload .xlsx or .xls"}, status_code=400
        )

    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        return JSONResponse({"error": "File too large (max 32 MB)"}, status_code=400)

    safe_name = f"{uuid.uuid4().hex}_{_secure_name(file.filename)}"
    filepath  = UPLOAD_DIR / safe_name

    with open(filepath, "wb") as f:
        f.write(contents)

    try:
        headers, all_rows, total_rows = parse_prospect_file(str(filepath))
        preview = [dict(row) for row in all_rows[:5]]
        return {
            "file_id":    safe_name,
            "headers":    headers,
            "preview":    preview,
            "total_rows": total_rows,
        }
    except Exception as exc:
        return JSONResponse(
            {"error": f"Failed to parse file: {exc}"}, status_code=400
        )


# ── 2. Validate Templates ────────────────────────────────────────────────────

@router.post("/api/merge/validate-templates")
async def validate_templates_route(request: Request):
    data = await request.json()
    headers   = data.get("headers", [])
    templates = (
        data.get("subject_templates", [])
        + data.get("body_templates", [])
        + ([data["chaser_body"]] if data.get("chaser_body") else [])
    )
    errors = validate_templates(templates, headers)
    return {"valid": len(errors) == 0, "errors": errors}


# ── 3. Generate Merge ────────────────────────────────────────────────────────

@router.post("/api/merge/generate")
async def generate_merge(request: Request):
    data = await request.json()

    file_id           = data.get("file_id")
    subject_templates = data.get("subject_templates", [])
    body_templates    = data.get("body_templates", [])
    chaser_body       = data.get("chaser_body", "")
    sender_emails     = data.get("sender_emails", [])
    missing_value     = data.get("missing_value", "[MISSING]")
    email_column      = data.get("email_column", "")
    recipient_tz      = data.get("recipient_tz", "Europe/London")
    sender_tz         = data.get("sender_tz", "Europe/London")
    sends_per_day_raw = data.get("sends_per_day", 10)
    sends_per_day     = sends_per_day_raw if sends_per_day_raw in (5, 10, 15) else 10
    window_start      = data.get("window_start", "08:30")
    window_end        = data.get("window_end", "15:30")

    # ── Validation ────────────────────────────────────────────────────────
    if not file_id:
        return JSONResponse({"error": "No prospect file specified"}, status_code=400)
    if not subject_templates:
        return JSONResponse(
            {"error": "At least one subject line template is required"}, status_code=400
        )
    if not body_templates:
        return JSONResponse(
            {"error": "At least one email body template is required"}, status_code=400
        )
    if not sender_emails:
        return JSONResponse(
            {"error": "At least one sender email address is required"}, status_code=400
        )

    # Same traversal guard as download_id in upload_to_sheets — file_id is
    # client-supplied and must never resolve outside UPLOAD_DIR.
    if "/" in file_id or "\\" in file_id or ".." in file_id:
        return JSONResponse({"error": "Invalid file_id."}, status_code=400)

    filepath = UPLOAD_DIR / file_id
    if not filepath.exists():
        return JSONResponse(
            {"error": "Prospect file not found. Please re-upload."}, status_code=404
        )

    try:
        headers, all_rows, _ = parse_prospect_file(str(filepath))

        # Validate all templates
        all_templates = subject_templates + body_templates
        if chaser_body:
            all_templates.append(chaser_body)

        errors = validate_templates(all_templates, headers)
        if errors:
            return JSONResponse(
                {"error": "Template validation failed", "details": errors},
                status_code=400,
            )

        # Merge
        merged_rows = perform_merge(
            all_rows, headers, subject_templates, body_templates,
            chaser_body, sender_emails, missing_value, email_column,
        )

        # Schedule
        campaign_seed = uuid.uuid4().hex[:8]
        schedule = generate_schedule(
            len(merged_rows), sender_emails,
            campaign_seed=campaign_seed,
            recipient_tz=recipient_tz, sender_tz=sender_tz,
            max_per_sender_per_day=sends_per_day,
            window_start=window_start, window_end=window_end,
        )

        for entry in schedule:
            row = merged_rows[entry["prospect_id"] - 1]
            row["__send_date__"]      = entry["date"]
            row["__send_time__"]      = entry["send_time"]
            row["__send_instant__"]   = entry.get("send_instant", 0.0)
            row["__sender_number__"]  = entry["sender_number"]
            row["__sender_account__"] = entry["sender"]
            chaser_offset = _dvariance(
                f"chaser-{entry['date']}-p{entry['prospect_id']}", 4, 5
            )
            send_h, send_m = map(int, entry["send_time"].split(":"))
            total_chaser_mins = send_h * 60 + send_m + chaser_offset
            row["__chaser_send_time__"] = (
                f"{total_chaser_mins // 60:02d}:{total_chaser_mins % 60:02d}"
            )

        # Sort: date → sender profile order → chronological send instant.
        # Ordering on the absolute instant, not the "HH:MM" text: when the
        # sender's zone is far enough ahead of the recipient's, the day's last
        # sends fall after midnight on the sender's clock, and sorting those
        # strings put "00:13"/"01:07" at the TOP of a block that should end
        # with them.
        sender_order = {email: idx for idx, email in enumerate(sender_emails)}
        merged_rows.sort(key=lambda r: (
            r.get("__send_date__", ""),
            sender_order.get(r.get("__sender_account__", ""), 999),
            r.get("__send_instant__", 0.0),
        ))
        merged_rows = [r for r in merged_rows if r.get("__send_date__")]

        # Re-assign templates in final sorted order → output shows 1-2-3-1-2-3
        reassign_templates(
            merged_rows, subject_templates, body_templates, headers, missing_value
        )

        # ── Missing-value guard ───────────────────────────────────────────
        # Reject the merge if any prospect still contains the missing sentinel
        # after all substitutions — this means a {{placeholder}} referenced a
        # column that had no value for that row.
        has_chaser = bool(chaser_body)
        fields_to_check = [
            ("__subject_line__", "Subject Line"),
            ("__email_body__",   "Email Body"),
        ]
        if has_chaser:
            fields_to_check.append(("__chaser_body__", "Chaser Body"))

        missing_rows = []
        for idx, row in enumerate(merged_rows, 1):
            bad_fields = [
                label
                for key, label in fields_to_check
                if missing_value in (row.get(key) or "")
            ]
            if bad_fields:
                missing_rows.append({
                    "row":    idx,
                    "email":  row.get("__recipient_email__", f"Row {idx}"),
                    "fields": bad_fields,
                })

        if missing_rows:
            return JSONResponse(
                {
                    "error": (
                        f"{len(missing_rows)} prospect(s) have missing data — "
                        "fix the spreadsheet and regenerate."
                    ),
                    "missing_rows": missing_rows,
                },
                status_code=422,
            )

        # Write Excel output
        output_filename = f"outreach_merge_{uuid.uuid4().hex[:8]}.xlsx"
        output_path = str(OUTPUT_DIR / output_filename)
        write_merge_output(
            output_path, headers, merged_rows, has_chaser, email_column,
            has_schedule=True,
        )

        return {
            "download_id": output_filename,
            "total_rows": len(merged_rows),
            "scheduled_count": len(merged_rows),
            "preview": [
                {
                    "sender":       r.get("__sender_account__", ""),
                    "recipient":    r.get("__recipient_email__", ""),
                    "subject":      r.get("__subject_line__", ""),
                    "body_preview": (
                        r.get("__email_body__", "")[:120] + "…"
                        if len(r.get("__email_body__", "")) > 120
                        else r.get("__email_body__", "")
                    ),
                    "variant":   r.get("__template_variant__", ""),
                    "send_date": r.get("__send_date__", ""),
                    "send_time": r.get("__send_time__", ""),
                }
                for r in merged_rows[:5]
            ],
        }
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)


# ── 4. Sender Profiles — GET ─────────────────────────────────────────────────

@router.get("/api/merge/sender-profiles")
async def get_sender_profiles():
    if not _sb_configured():
        return JSONResponse(
            {"error": "Supabase is not configured on this server."}, status_code=503
        )
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/sender_profiles",
            headers=_sb_headers(),
            params={"select": "profile_name,emails", "order": "profile_name.asc"},
            timeout=10,
        )
        r.raise_for_status()
        return r.json()
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)


# ── 5. Sender Profiles — POST (upsert) ───────────────────────────────────────

@router.post("/api/merge/sender-profiles")
async def save_sender_profile(request: Request):
    if not _sb_configured():
        return JSONResponse(
            {"error": "Supabase is not configured on this server."}, status_code=503
        )
    data = await request.json()
    profile_name = (data.get("profile_name") or "").strip()
    emails       = data.get("emails", [])

    if not profile_name:
        return JSONResponse({"error": "Profile name is required."}, status_code=400)
    if not emails:
        return JSONResponse(
            {"error": "At least one email address is required."}, status_code=400
        )

    try:
        r = http_req.post(
            f"{SUPABASE_URL}/rest/v1/sender_profiles",
            headers=_sb_headers("resolution=merge-duplicates"),
            params={"on_conflict": "profile_name"},
            json={"profile_name": profile_name, "emails": emails},
            timeout=10,
        )
        r.raise_for_status()
        return {"ok": True}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)


# ── 6. Sender Profiles — DELETE ───────────────────────────────────────────────

@router.delete("/api/merge/sender-profiles/{profile_name:path}", dependencies=[Depends(auth.require_admin)])
async def delete_sender_profile(profile_name: str):
    if not _sb_configured():
        return JSONResponse(
            {"error": "Supabase is not configured on this server."}, status_code=503
        )
    try:
        r = http_req.delete(
            f"{SUPABASE_URL}/rest/v1/sender_profiles",
            headers=_sb_headers(),
            params={"profile_name": f"eq.{profile_name}"},
            timeout=10,
        )
        r.raise_for_status()
        return {"ok": True}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)


# ── 7. Google Sheets — upload outreach sheet ─────────────────────────────────

@router.post("/api/merge/upload-to-sheets")
async def upload_to_sheets(request: Request):
    """
    Read a previously-generated merge xlsx (identified by download_id),
    write it to a new Google Sheet, share as "anyone with link can edit",
    save a campaign record to Supabase, and return the sheet URL + ID.
    """
    from datetime import date as _date
    from app.utils.google_sheets import is_configured, create_outreach_sheet

    data                = await request.json()
    download_id         = (data.get("download_id") or "").strip()
    client_name         = (data.get("client_name") or "Client").strip() or "Client"
    client_id           = (data.get("client_id") or "").strip()
    campaign            = (data.get("campaign_name") or "").strip()
    sender_profile_name = (data.get("sender_profile_name") or "").strip()
    total_prospects     = int(data.get("total_prospects") or 0)
    copy_territory      = (data.get("copy_territory") or "").strip()
    copy_industry       = (data.get("copy_industry") or "").strip()

    if not is_configured():
        return JSONResponse(
            {"error": "Google Sheets is not configured on this server. "
                      "Ask your administrator to set GOOGLE_SHEETS_SA_JSON."},
            status_code=400,
        )

    if not download_id or "/" in download_id or "\\" in download_id or ".." in download_id:
        return JSONResponse({"error": "Invalid download_id."}, status_code=400)

    xlsx_path = OUTPUT_DIR / download_id
    if not xlsx_path.exists():
        return JSONResponse(
            {"error": "Merge file not found — please regenerate the outreach sheet first."},
            status_code=404,
        )

    # Build sheet title: "ClientName – Campaign – YYYY-MM-DD"
    parts = [client_name]
    if campaign:
        parts.append(campaign)
    parts.append(_date.today().isoformat())
    title = " – ".join(parts)

    try:
        result = create_outreach_sheet(title, str(xlsx_path))
    except Exception as exc:
        return JSONResponse({"error": f"Google Sheets error: {exc}"}, status_code=500)

    # ── Save campaign record to Supabase (best-effort — don't block on failure) ──
    if _sb_configured():
        base_row = {
            "campaign_name":       campaign or title,
            "sender_profile_name": sender_profile_name,
            "client_id":           client_id or None,
            "client_name":         client_name,
            "sheet_id":            result["sheet_id"],
            "sheet_url":           result["sheet_url"],
            "total_prospects":     total_prospects,
            "sent_count":          0,
        }
        has_copy_source = bool(copy_territory or copy_industry)
        row = dict(base_row)
        if has_copy_source:
            row["copy_territory"] = copy_territory or None
            row["copy_industry"]  = copy_industry or None

        def _insert(payload: dict) -> bool:
            try:
                r = http_req.post(
                    f"{SUPABASE_URL}/rest/v1/campaigns",
                    headers=_sb_headers("return=minimal"),
                    json=payload,
                    timeout=10,
                )
                return r.status_code < 400
            except Exception:
                return False

        # Copy-source columns may not be migrated yet; PostgREST rejects the whole
        # insert on any unknown column, so retry without them on failure.
        if not _insert(row) and has_copy_source:
            _insert(base_row)

    return result  # {"sheet_id": ..., "sheet_url": ..., "title": ...}


@router.get("/api/merge/sheets-status")
async def sheets_status():
    """Return whether Google Sheets integration is configured."""
    from app.utils.google_sheets import is_configured
    return {"configured": is_configured()}


@router.post("/api/merge/cleanup-drive", dependencies=[Depends(auth.require_admin)])
async def cleanup_drive():
    """Delete ALL spreadsheets from the service account's Drive to free quota."""
    from app.utils.google_sheets import is_configured, cleanup_service_account_drive
    if not is_configured():
        return JSONResponse({"error": "Google Sheets not configured."}, status_code=503)
    try:
        result = cleanup_service_account_drive(older_than_days=0)
        return result
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)


