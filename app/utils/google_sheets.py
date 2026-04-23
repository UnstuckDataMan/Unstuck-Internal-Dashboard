"""
Google Sheets integration helper.

Requires:
  - GOOGLE_SHEETS_SA_JSON env var: base64-encoded service account JSON key
    (Sheets API + Drive API must both be enabled on the Cloud project)
  - GOOGLE_DRIVE_FOLDER_ID env var: ID of a Shared Drive shared with the
    service account as Contributor. Files in a Shared Drive are owned by the
    drive, so they never count against individual or SA storage quotas.

One-time setup:
  1. Google Cloud Console → enable Sheets API + Drive API
  2. IAM & Admin → Service Accounts → Create → download JSON key
  3. Base64-encode: base64 -w0 service-account.json
  4. Add as GOOGLE_SHEETS_SA_JSON env var in Render
  5. Create a Shared Drive → add the service account EMAIL as Contributor
  6. Add the Shared Drive ID as GOOGLE_DRIVE_FOLDER_ID env var in Render
"""
from __future__ import annotations

import base64
import json
import os
import re

import gspread
import openpyxl
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_DRIVE_FILES_URL      = "https://www.googleapis.com/drive/v3/files"
_SHEETS_BASE_URL      = "https://sheets.googleapis.com/v4/spreadsheets"
_SEPARATOR_TEXT       = "No More Emails For Today."


# ── Auth helpers ──────────────────────────────────────────────────────────────

def is_configured() -> bool:
    """Return True if the Google Sheets env var is present."""
    return bool(os.environ.get("GOOGLE_SHEETS_SA_JSON", "").strip())


def _decode_sa_json() -> dict:
    """Decode the base64 service account JSON from the env var."""
    raw = os.environ.get("GOOGLE_SHEETS_SA_JSON", "").strip()
    if not raw:
        raise RuntimeError(
            "GOOGLE_SHEETS_SA_JSON env var is not set. "
            "Follow the setup instructions to configure Google Sheets integration."
        )
    try:
        return json.loads(base64.b64decode(raw + "=="))
    except Exception as exc:
        raise RuntimeError(f"Could not decode GOOGLE_SHEETS_SA_JSON: {exc}") from exc


def _client() -> gspread.Client:
    """Return an authenticated gspread Client (gspread v6 native method)."""
    info = _decode_sa_json()
    return gspread.service_account_from_dict(info, scopes=SCOPES)


def _authed_session() -> AuthorizedSession:
    """Return an AuthorizedSession for direct Sheets/Drive REST calls."""
    info = _decode_sa_json()
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return AuthorizedSession(creds)


# ── Drive cleanup (admin utility) ─────────────────────────────────────────────

def cleanup_service_account_drive(older_than_days: int = 0) -> dict:
    """
    Delete spreadsheets owned by the service account.
    older_than_days=0 deletes ALL sheets; >0 deletes only older ones.
    Returns {"deleted": int, "errors": int}
    """
    session = _authed_session()
    q = "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    if older_than_days > 0:
        from datetime import datetime, timezone, timedelta
        cutoff = (datetime.now(timezone.utc) - timedelta(days=older_than_days)).strftime("%Y-%m-%dT%H:%M:%S")
        q += f" and createdTime < '{cutoff}'"

    files, page_token = [], None
    while True:
        params: dict = {"q": q, "fields": "nextPageToken,files(id,name)", "pageSize": 100}
        if page_token:
            params["pageToken"] = page_token
        resp = session.get(_DRIVE_FILES_URL, params=params)
        if not resp.ok:
            break
        data = resp.json()
        files.extend(data.get("files", []))
        page_token = data.get("nextPageToken")
        if not page_token:
            break

    deleted = errors = 0
    for f in files:
        r = session.delete(f"{_DRIVE_FILES_URL}/{f['id']}")
        if r.ok or r.status_code == 204:
            deleted += 1
        else:
            errors += 1
    return {"deleted": deleted, "errors": errors}


# ── Sheet creation ────────────────────────────────────────────────────────────

def _create_spreadsheet(title: str) -> tuple[str, str]:
    """
    Create a blank Google Spreadsheet inside the configured Shared Drive.

    Files in a Shared Drive are owned by the drive, not the service account,
    so they never count against individual or SA storage quotas.

    Returns (spreadsheet_id, spreadsheet_url).
    """
    drive_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "").strip()
    if not drive_id:
        raise RuntimeError(
            "GOOGLE_DRIVE_FOLDER_ID env var is not set. "
            "Set it to your Shared Drive ID."
        )

    session = _authed_session()
    resp = session.post(
        _DRIVE_FILES_URL,
        params={"supportsAllDrives": "true"},
        json={
            "name":     title,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents":  [drive_id],
        },
    )
    if not resp.ok:
        raise RuntimeError(
            f"Drive create failed (HTTP {resp.status_code}): {resp.text[:400]}"
        )
    sid = resp.json()["id"]
    url = f"https://docs.google.com/spreadsheets/d/{sid}/edit"
    return sid, url


# ── Colour / formatting helpers ───────────────────────────────────────────────

def _rgb(hex_str: str) -> dict:
    """Convert a 6-char hex colour string to a Sheets API colour object."""
    h = hex_str.lstrip("#")
    return {
        "red":   int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue":  int(h[4:6], 16) / 255,
    }


def _apply_sheet_formatting(
    session: AuthorizedSession,
    spreadsheet_id: str,
    sheet_gid: int,
    all_rows: list[list[str]],   # header row + data rows (already str-converted)
) -> None:
    """
    Apply data-validation, conditional formatting, and per-row styles to the
    Outreach List sheet via a single Sheets API batchUpdate call.

    Replicates the Excel writer's formatting:
      - Send Status column    → checkbox (TRUE/FALSE)
      - Lead Status column    → dropdown (Lead / Reply / Unsubscribe)
      - First-of-sender rows  → pale yellow background (#FFFDE7)
      - Separator rows        → merged, orange-tinted, bold brown text
      - Lead Status "Lead"    → green cell
      - Lead Status "Reply"   → orange cell
      - Lead Status "Unsubscribe" → red cell
      - Send Status = TRUE    → whole-row light green
    """
    if not all_rows:
        return

    headers   = all_rows[0]
    data_rows = all_rows[1:]
    n_cols    = len(headers)

    def col_idx(name: str) -> int:
        try:
            return headers.index(name)
        except ValueError:
            return -1

    send_status_col = col_idx("Send Status")
    lead_status_col = col_idx("Lead Status")
    sender_col      = col_idx("Sender Account")
    n_data          = len(data_rows)
    last_row_idx    = 1 + n_data   # exclusive end (0-based, row 0 = header)

    requests: list[dict] = []

    # ── Per-row styles (separator merge + yellow first-sender stripe) ─────
    prev_sender: str | None = None
    for ri, row in enumerate(data_rows):
        row_idx = ri + 1   # 0-based sheet row index

        # Detect separator row: only first cell has content = _SEPARATOR_TEXT
        is_sep = (
            row[0].strip() == _SEPARATOR_TEXT
            and all(v == "" for v in row[1:])
        )

        if is_sep:
            # Merge all cells in the separator row
            requests.append({"mergeCells": {
                "range": {
                    "sheetId":          sheet_gid,
                    "startRowIndex":    row_idx,
                    "endRowIndex":      row_idx + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex":   n_cols,
                },
                "mergeType": "MERGE_ALL",
            }})
            # Style: warm orange background, bold brown text, centred
            requests.append({"repeatCell": {
                "range": {
                    "sheetId":          sheet_gid,
                    "startRowIndex":    row_idx,
                    "endRowIndex":      row_idx + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex":   n_cols,
                },
                "cell": {"userEnteredFormat": {
                    "backgroundColor": _rgb("FFF3E0"),
                    "textFormat": {
                        "bold":            True,
                        "foregroundColor": _rgb("5D4037"),
                        "fontSize":        9,
                    },
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment":   "MIDDLE",
                }},
                "fields": (
                    "userEnteredFormat(backgroundColor,textFormat,"
                    "horizontalAlignment,verticalAlignment)"
                ),
            }})
            prev_sender = None   # reset so first row of next day gets stripe

        else:
            # Yellow stripe for first row of each new sender block
            if sender_col >= 0 and sender_col < len(row):
                curr = row[sender_col]
                if curr != prev_sender:
                    requests.append({"repeatCell": {
                        "range": {
                            "sheetId":          sheet_gid,
                            "startRowIndex":    row_idx,
                            "endRowIndex":      row_idx + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex":   n_cols,
                        },
                        "cell": {"userEnteredFormat": {
                            "backgroundColor": _rgb("FFFDE7"),
                        }},
                        "fields": "userEnteredFormat.backgroundColor",
                    }})
                    prev_sender = curr

    # ── Data validation: Send Status = checkbox ───────────────────────────
    if send_status_col >= 0:
        requests.append({"setDataValidation": {
            "range": {
                "sheetId":          sheet_gid,
                "startRowIndex":    1,
                "endRowIndex":      last_row_idx,
                "startColumnIndex": send_status_col,
                "endColumnIndex":   send_status_col + 1,
            },
            "rule": {
                "condition":   {"type": "BOOLEAN"},
                "showCustomUi": True,
            },
        }})

    # ── Data validation: Lead Status = dropdown ───────────────────────────
    if lead_status_col >= 0:
        requests.append({"setDataValidation": {
            "range": {
                "sheetId":          sheet_gid,
                "startRowIndex":    1,
                "endRowIndex":      last_row_idx,
                "startColumnIndex": lead_status_col,
                "endColumnIndex":   lead_status_col + 1,
            },
            "rule": {
                "condition": {
                    "type":   "ONE_OF_LIST",
                    "values": [
                        {"userEnteredValue": "Lead"},
                        {"userEnteredValue": "Reply"},
                        {"userEnteredValue": "Unsubscribe"},
                    ],
                },
                "showCustomUi": True,
                "strict":       False,
            },
        }})

    # ── Conditional formatting: Lead Status cell colours ─────────────────
    if lead_status_col >= 0:
        ls_range = {
            "sheetId":          sheet_gid,
            "startRowIndex":    1,
            "endRowIndex":      last_row_idx,
            "startColumnIndex": lead_status_col,
            "endColumnIndex":   lead_status_col + 1,
        }
        for value, hex_color in [
            ("Lead",        "D4EDD6"),
            ("Reply",       "FFD9B3"),
            ("Unsubscribe", "FFCDD2"),
        ]:
            requests.append({"addConditionalFormatRule": {
                "rule": {
                    "ranges": [ls_range],
                    "booleanRule": {
                        "condition": {
                            "type":   "TEXT_EQ",
                            "values": [{"userEnteredValue": value}],
                        },
                        "format": {"backgroundColor": _rgb(hex_color)},
                    },
                },
                "index": 0,
            }})

    # ── Conditional formatting: Send Status checked → whole-row light green
    if send_status_col >= 0:
        ss_letter = _col_letter(send_status_col + 1)
        requests.append({"addConditionalFormatRule": {
            "rule": {
                "ranges": [{
                    "sheetId":          sheet_gid,
                    "startRowIndex":    1,
                    "endRowIndex":      last_row_idx,
                    "startColumnIndex": 0,
                    "endColumnIndex":   n_cols,
                }],
                "booleanRule": {
                    "condition": {
                        "type":   "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": f"=${ss_letter}2=TRUE"}],
                    },
                    "format": {"backgroundColor": _rgb("E8F5E9")},
                },
            },
            "index": 0,
        }})

    if not requests:
        return

    resp = session.post(
        f"{_SHEETS_BASE_URL}/{spreadsheet_id}:batchUpdate",
        json={"requests": requests},
    )
    if not resp.ok:
        # Non-fatal — sheet is usable without formatting
        print(f"[google_sheets] batchUpdate warning {resp.status_code}: {resp.text[:300]}")


# ── Public: create outreach sheet ─────────────────────────────────────────────

def create_outreach_sheet(title: str, xlsx_path: str) -> dict:
    """
    Read the 'Outreach List' sheet from an openpyxl xlsx file
    (generated by mail_merge/utils/excel_writer.py), write all rows to a new
    Google Sheet in the configured Shared Drive, apply Excel-equivalent
    formatting (dropdowns, conditional formatting, sender-change stripes,
    separator rows), freeze the header, and share as "anyone with link can edit".

    Returns: {"sheet_id": str, "sheet_url": str, "title": str}
    """
    # ── Read Excel ────────────────────────────────────────────────────────
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheet_name = "Outreach List" if "Outreach List" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    raw_rows = [[cell.value for cell in row] for row in ws.iter_rows()]
    if not raw_rows:
        raise ValueError("The merge output file is empty.")

    str_rows  = [[str(v) if v is not None else "" for v in row] for row in raw_rows]
    headers   = str_rows[0]
    data_rows = str_rows[1:]
    all_rows  = [headers] + data_rows

    # ── Create spreadsheet in Shared Drive ───────────────────────────────
    file_id, sheet_url = _create_spreadsheet(title)

    # ── Write data via gspread ───────────────────────────────────────────
    gc     = _client()
    sh     = gc.open_by_key(file_id)
    gsheet = sh.sheet1
    gsheet.update_title("Outreach List")
    gsheet.update(all_rows, "A1")

    # Bold header + freeze
    gsheet.format(
        f"A1:{_col_letter(len(headers))}1",
        {"textFormat": {"bold": True}},
    )
    gsheet.freeze(rows=1)

    # ── Apply full formatting via batchUpdate ─────────────────────────────
    try:
        session = _authed_session()
        _apply_sheet_formatting(session, file_id, gsheet.id, all_rows)
    except Exception as fmt_err:
        print(f"[google_sheets] formatting skipped: {fmt_err}")

    # ── Share: anyone with link can edit ─────────────────────────────────
    sh.share("", perm_type="anyone", role="writer")

    return {
        "sheet_id":  file_id,
        "sheet_url": sheet_url,
        "title":     title,
    }


# ── Public: read sheet data ───────────────────────────────────────────────────

def extract_sheet_id(url_or_id: str) -> str:
    """
    Accept either a full Google Sheets URL or a bare spreadsheet ID
    and return just the ID component.
    """
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9\-_]+)", url_or_id)
    return m.group(1) if m else url_or_id.strip()


def read_sheet_status(sheet_id: str) -> dict:
    """
    Count total prospect rows and how many have Send Status ticked / = "Sent".
    Supports both checkbox (TRUE boolean) and legacy dropdown ("Sent" text).

    Returns: {"total": int, "sent": int, "is_complete": bool}
    """
    gc = _client()
    sh = gc.open_by_key(sheet_id)
    gsheet = sh.sheet1
    records = gsheet.get_all_records()

    data_rows = [r for r in records if str(r.get("Recipient Email", "")).strip()]
    total     = len(data_rows)
    sent      = sum(
        1 for r in data_rows
        if (r.get("Send Status") is True
            or str(r.get("Send Status", "")).strip().upper() in ("TRUE", "SENT"))
    )
    return {"total": total, "sent": sent, "is_complete": (total > 0 and sent >= total)}


def read_sent_emails(sheet_id: str) -> list[str]:
    """
    Return the list of Recipient Email values where Send Status is ticked
    (checkbox = TRUE) or equals "Sent".  Skips separator rows (no email).
    """
    gc = _client()
    sh = gc.open_by_key(sheet_id)
    gsheet = sh.sheet1
    records = gsheet.get_all_records()

    emails: list[str] = []
    for r in records:
        status = r.get("Send Status", "")
        email  = str(r.get("Recipient Email", "")).strip().lower()
        if (
            (status is True or str(status).strip().upper() in ("TRUE", "SENT"))
            and email and "@" in email
        ):
            emails.append(email)
    return emails


def read_leads(sheet_id: str) -> list[dict]:
    """
    Return {"email", "status"} pairs where Lead Status is "Lead" or "Unsubscribe".
    Rows with status "Reply" or blank are ignored.
    """
    gc = _client()
    sh = gc.open_by_key(sheet_id)
    gsheet = sh.sheet1
    records = gsheet.get_all_records()

    results: list[dict] = []
    for record in records:
        status = str(record.get("Lead Status", "")).strip()
        email  = str(record.get("Recipient Email", "")).strip().lower()
        if status in ("Lead", "Unsubscribe") and email and "@" in email:
            results.append({"email": email, "status": status})
    return results


# ── Internal helpers ──────────────────────────────────────────────────────────

def _col_letter(n: int) -> str:
    """Convert a 1-based column index to an Excel-style letter (1→A, 26→Z, 27→AA)."""
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result or "A"
