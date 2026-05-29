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
import threading

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


# Module-level gspread client cache.
# Creating a new client for every call triggers a full OAuth token fetch each
# time (≈300–600 ms per call).  gspread's internal AuthorizedSession refreshes
# the token automatically when it nears expiry, so a long-lived client is safe.
_gc_cache: gspread.Client | None = None
_gc_lock  = threading.Lock()


def _client() -> gspread.Client:
    """Return a cached, authenticated gspread Client (gspread v6 native method).

    The client is created once per process and reused for all subsequent calls.
    gspread's AuthorizedSession refreshes the OAuth access token transparently
    when it expires, so no manual TTL management is needed.
    """
    global _gc_cache
    if _gc_cache is not None:
        return _gc_cache
    with _gc_lock:
        # Double-checked locking: another thread may have initialised it while
        # we were waiting for the lock.
        if _gc_cache is None:
            info = _decode_sa_json()
            _gc_cache = gspread.service_account_from_dict(info, scopes=SCOPES)
    return _gc_cache


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
    email_col       = col_idx("Recipient Email")
    div_col         = col_idx("__divider__")
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
                            "textFormat": {"bold": True},
                        }},
                        "fields": "userEnteredFormat(backgroundColor,textFormat.bold)",
                    }})
                    prev_sender = curr

    # ── Divider column (grey separator before prospect columns) ──────────
    # Matches the light-grey divider column written by excel_writer.py.
    # Header cell text is cleared; the whole column gets a grey fill and a
    # narrow pixel width so it reads as a visual boundary, not a data column.
    if div_col >= 0:
        # Header cell: dark grey fill, clear the "__divider__" text
        requests.append({"updateCells": {
            "rows": [{"values": [{"userEnteredFormat": {
                "backgroundColor": _rgb("BDBDBD"),
            }}]}],
            "fields": "userEnteredFormat.backgroundColor",
            "start": {
                "sheetId":     sheet_gid,
                "rowIndex":    0,
                "columnIndex": div_col,
            },
        }})
        requests.append({"updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": ""}}]}],
            "fields": "userEnteredValue",
            "start": {
                "sheetId":     sheet_gid,
                "rowIndex":    0,
                "columnIndex": div_col,
            },
        }})
        # Data rows: lighter grey fill
        if n_data > 0:
            requests.append({"repeatCell": {
                "range": {
                    "sheetId":          sheet_gid,
                    "startRowIndex":    1,
                    "endRowIndex":      last_row_idx,
                    "startColumnIndex": div_col,
                    "endColumnIndex":   div_col + 1,
                },
                "cell": {"userEnteredFormat": {
                    "backgroundColor": _rgb("EEEEEE"),
                }},
                "fields": "userEnteredFormat.backgroundColor",
            }})
        # Narrow column width (~20 px ≈ 3-char wide visual separator)
        requests.append({"updateDimensionProperties": {
            "range": {
                "sheetId":    sheet_gid,
                "dimension":  "COLUMNS",
                "startIndex": div_col,
                "endIndex":   div_col + 1,
            },
            "properties": {"pixelSize": 20},
            "fields": "pixelSize",
        }})

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
            "rule": {"condition": {"type": "BOOLEAN"}, "showCustomUi": True},
        }})

    # ── Data validation: Chaser Sent? = checkbox ──────────────────────────
    chaser_col = col_idx("Chaser Sent?")
    if chaser_col >= 0:
        requests.append({"setDataValidation": {
            "range": {
                "sheetId":          sheet_gid,
                "startRowIndex":    1,
                "endRowIndex":      last_row_idx,
                "startColumnIndex": chaser_col,
                "endColumnIndex":   chaser_col + 1,
            },
            "rule": {"condition": {"type": "BOOLEAN"}, "showCustomUi": True},
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
                        {"userEnteredValue": "Interested"},
                        {"userEnteredValue": "Unsubscribe"},
                    ],
                },
                "showCustomUi": True,
                "strict":       False,
            },
        }})

    # ── Conditional formatting ─────────────────────────────────────────────
    # Rules are added with index=0. Each insertion shifts previous rules down,
    # so the LAST rule added ends up at the highest priority (index 0).
    #
    # Priority order (low → high):
    #   1. Domain-grey  whole-row grey  (lowest  — overridden by send & lead)
    #   2. Send Status  whole-row blue  (middle)
    #   3. Lead Status  cell colours    (highest — narrow range, always wins)

    # 1. Domain-grey (lowest priority — added first).
    # When any row has Lead Status="Lead", ALL rows sharing the same email
    # domain are highlighted grey.  The formula extracts the domain from the
    # Recipient Email column directly via RIGHT/FIND and matches it with a
    # wildcard COUNTIFS — no helper column required.
    if lead_status_col >= 0 and email_col >= 0:
        ls_letter = _col_letter(lead_status_col + 1)
        em_letter = _col_letter(email_col + 1)
        # Formula anchors the column ($) but leaves the row relative so it
        # shifts correctly as Sheets evaluates it for each row in the range.
        formula = (
            f'=AND(${ls_letter}2<>"Lead",'
            f'ISNUMBER(FIND("@",${em_letter}2)),'
            f'COUNTIFS(${ls_letter}$2:${ls_letter}${last_row_idx},"Lead",'
            f'${em_letter}$2:${em_letter}${last_row_idx},'
            f'"*@"&RIGHT(${em_letter}2,LEN(${em_letter}2)-FIND("@",${em_letter}2)))>0)'
        )
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
                        "values": [{"userEnteredValue": formula}],
                    },
                    "format": {"backgroundColor": _rgb("E0E0E0")},  # medium grey
                },
            },
            "index": 0,
        }})

    # 3. Send Status checked → whole-row light BLUE  (medium priority)
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
                    "format": {"backgroundColor": _rgb("E3F2FD")},  # light blue
                },
            },
            "index": 0,
        }})

    # 3.5. Lead Status = "Lead" → whole-row grey EXCEPT the LS column (higher than sent-blue).
    # Ensures a lead row is always grey even when it is also marked as sent.
    # The LS column is excluded from the two ranges so the green Lead cell colour
    # (rule 4, highest priority) is never shadowed by this grey.
    if lead_status_col >= 0:
        ls_letter  = _col_letter(lead_status_col + 1)
        lead_formula = f'=${ls_letter}2="Lead"'
        lead_ranges: list[dict] = []
        if lead_status_col > 0:
            lead_ranges.append({
                "sheetId":          sheet_gid,
                "startRowIndex":    1,
                "endRowIndex":      last_row_idx,
                "startColumnIndex": 0,
                "endColumnIndex":   lead_status_col,   # up to (not including) LS col
            })
        if lead_status_col + 1 < n_cols:
            lead_ranges.append({
                "sheetId":          sheet_gid,
                "startRowIndex":    1,
                "endRowIndex":      last_row_idx,
                "startColumnIndex": lead_status_col + 1,  # one after LS col
                "endColumnIndex":   n_cols,
            })
        if lead_ranges:
            requests.append({"addConditionalFormatRule": {
                "rule": {
                    "ranges": lead_ranges,
                    "booleanRule": {
                        "condition": {
                            "type":   "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": lead_formula}],
                        },
                        "format": {"backgroundColor": _rgb("E0E0E0")},  # same grey as domain-grey
                    },
                },
                "index": 0,
            }})

    # 4. Lead Status cell colours (highest priority — narrow range, LS cell only).
    if lead_status_col >= 0:
        ls_range = {
            "sheetId":          sheet_gid,
            "startRowIndex":    1,
            "endRowIndex":      last_row_idx,
            "startColumnIndex": lead_status_col,
            "endColumnIndex":   lead_status_col + 1,
        }
        for value, bg_hex, fg_hex in [
            ("Lead",        "66BB6A", "FFFFFF"),   # vivid green / white
            ("Reply",       "FFA726", "FFFFFF"),   # vivid amber / white
            ("Interested",  "AB47BC", "FFFFFF"),   # vivid purple / white
            ("Unsubscribe", "EF5350", "FFFFFF"),   # vivid red / white
        ]:
            requests.append({"addConditionalFormatRule": {
                "rule": {
                    "ranges": [ls_range],
                    "booleanRule": {
                        "condition": {
                            "type":   "TEXT_EQ",
                            "values": [{"userEnteredValue": value}],
                        },
                        "format": {
                            "backgroundColor": _rgb(bg_hex),
                            "textFormat": {
                                "foregroundColor": _rgb(fg_hex),
                                "bold": True,
                            },
                        },
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
        print(f"[google_sheets] batchUpdate FAILED {resp.status_code}: {resp.text[:600]}")
    else:
        print(f"[google_sheets] batchUpdate OK — {len(requests)} requests applied")


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

    str_rows = [[str(v) if v is not None else "" for v in row] for row in raw_rows]

    headers   = str_rows[0]
    data_rows = str_rows[1:]
    all_rows  = str_rows

    # Log detected column positions for debugging
    _div_idx = headers.index("__divider__") if "__divider__" in headers else -1
    _ls_idx  = headers.index("Lead Status") if "Lead Status" in headers else -1
    _em_idx  = headers.index("Recipient Email") if "Recipient Email" in headers else -1
    print(
        f"[google_sheets] columns — __divider__:{_div_idx}  "
        f"LeadStatus:{_ls_idx}  RecipientEmail:{_em_idx}  total:{len(headers)}"
    )

    # Append "Sent Date" helper column (tracks the date the prospect was marked sent).
    # Each data row gets an empty cell; the auto_sync will fill in the date when syncing.
    if "Sent Date" not in headers:
        for row in str_rows:
            row.append("")
        str_rows[0][-1] = "Sent Date"
        headers   = str_rows[0]
        data_rows = str_rows[1:]
        all_rows  = str_rows

    # Append "Chaser Sent?" column — checkbox to track follow-up / chaser emails.
    if "Chaser Sent?" not in headers:
        for row in str_rows:
            row.append("")
        str_rows[0][-1] = "Chaser Sent?"
        headers   = str_rows[0]
        data_rows = str_rows[1:]
        all_rows  = str_rows

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
        print(f"[google_sheets] applying formatting (sheet_gid={gsheet.id}) …")
        _apply_sheet_formatting(session, file_id, gsheet.id, all_rows)
        print("[google_sheets] formatting applied OK")
    except Exception as fmt_err:
        print(f"[google_sheets] formatting FAILED: {fmt_err}")

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


def _is_sent(value) -> bool:
    """Return True if a Send Status cell value counts as 'sent'.

    Handles all representations gspread may return:
      • Python bool True  (UNFORMATTED_VALUE render)
      • String "TRUE"     (FORMATTED_VALUE render)
      • String "Sent"     (legacy dropdown before checkbox migration)
      • Integer 1         (edge case numeric true)
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value == 1
    return str(value).strip().upper() in ("TRUE", "SENT")


def read_sheet_status(sheet_id: str) -> dict:
    """
    Count total prospect rows and how many have Send Status ticked / = "Sent".
    Supports both checkbox (TRUE boolean) and legacy dropdown ("Sent" text).

    Returns: {"total": int, "sent": int, "is_complete": bool}
    """
    gc = _client()
    sh = gc.open_by_key(sheet_id)
    gsheet = sh.sheet1
    # UNFORMATTED_VALUE returns Python booleans for checkbox cells
    records = gsheet.get_all_records(value_render_option="UNFORMATTED_VALUE")

    data_rows = [r for r in records if str(r.get("Recipient Email", "")).strip()]
    total     = len(data_rows)
    sent      = sum(1 for r in data_rows if _is_sent(r.get("Send Status", "")))
    return {"total": total, "sent": sent, "is_complete": (total > 0 and sent >= total)}


def read_sent_emails(sheet_id: str) -> list[str]:
    """
    Return the list of Recipient Email values where Send Status is ticked.
    Uses UNFORMATTED_VALUE so checkbox cells come back as Python booleans.
    Skips separator rows (no email address).
    """
    gc = _client()
    sh = gc.open_by_key(sheet_id)
    gsheet = sh.sheet1
    records = gsheet.get_all_records(value_render_option="UNFORMATTED_VALUE")

    emails: list[str] = []
    for r in records:
        email = str(r.get("Recipient Email", "")).strip().lower()
        if _is_sent(r.get("Send Status", "")) and email and "@" in email:
            emails.append(email)
    return emails


def read_sent_with_dates(sheet_id: str) -> list[dict]:
    """
    Return [{email, sent_date, row_num}] for all rows where Send Status is ticked.

    sent_date is the value from the 'Sent Date' column (ISO date string, may be
    "" if the column doesn't exist yet or has not been populated for that row).
    row_num is 1-based sheet row number (row 1 = header, row 2 = first data row).
    """
    gc = _client()
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1
    records = ws.get_all_records(value_render_option="UNFORMATTED_VALUE")

    results: list[dict] = []
    for i, r in enumerate(records, start=2):   # row 2 = first data row
        email = str(r.get("Recipient Email", "")).strip().lower()
        if not _is_sent(r.get("Send Status", "")) or not email or "@" not in email:
            continue
        sent_date    = str(r.get("Sent Date", "")).strip()
        chaser_sent  = _is_sent(r.get("Chaser Sent?", ""))
        results.append({"email": email, "sent_date": sent_date,
                         "row_num": i, "chaser_sent": chaser_sent})
    return results


def write_sent_dates(sheet_id: str, row_date_pairs: list[tuple[int, str]]) -> None:
    """
    Batch-write (row_num, date_str) pairs to the 'Sent Date' column.

    If 'Sent Date' doesn't exist in the header row it is appended automatically.
    row_num values are 1-based sheet row numbers (1 = header).
    """
    if not row_date_pairs:
        return

    gc = _client()
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1

    headers = ws.row_values(1)
    if "Sent Date" not in headers:
        new_col_idx = len(headers) + 1          # 1-based
        ws.update_cell(1, new_col_idx, "Sent Date")
        sent_date_col = new_col_idx
    else:
        sent_date_col = headers.index("Sent Date") + 1   # 1-based

    updates = [
        {"range": f"{_col_letter(sent_date_col)}{row_num}", "values": [[date_str]]}
        for row_num, date_str in row_date_pairs
    ]
    if updates:
        ws.batch_update(updates)


def read_ab_stats(sheet_id: str) -> list[dict]:
    """
    Group sheet rows by 'Template Variant' and count Lead Status outcomes.

    Only rows with a valid email address and a non-empty Template Variant
    are counted; separator rows are skipped.

    Returns a list of dicts sorted by positive responses (lead + interested)
    descending, then by reply count:

    [
      {
        "variant":    "S2/B2",
        "total":      45,        # total prospects allocated to this variant
        "lead":       10,
        "interested": 2,
        "reply":      5,
        "unsubscribe": 0,
        "positive":   12,        # lead + interested
      },
      ...
    ]
    """
    from collections import defaultdict

    gc     = _client()
    sh     = gc.open_by_key(sheet_id)
    gsheet = sh.sheet1
    records = gsheet.get_all_records(value_render_option="UNFORMATTED_VALUE")

    stats: dict = defaultdict(
        lambda: {"total": 0, "lead": 0, "interested": 0, "reply": 0, "unsubscribe": 0}
    )

    for r in records:
        email   = str(r.get("Recipient Email", "")).strip()
        variant = str(r.get("A/B Variant", "")).strip()
        if not email or "@" not in email or not variant:
            continue  # separator rows or rows without variant assignment

        status = str(r.get("Lead Status", "")).strip()
        stats[variant]["total"] += 1
        if status == "Lead":
            stats[variant]["lead"] += 1
        elif status == "Interested":
            stats[variant]["interested"] += 1
        elif status == "Reply":
            stats[variant]["reply"] += 1
        elif status == "Unsubscribe":
            stats[variant]["unsubscribe"] += 1

    result = []
    for variant, s in stats.items():
        positive = s["lead"] + s["interested"]
        result.append({
            "variant":     variant,
            "total":       s["total"],
            "lead":        s["lead"],
            "interested":  s["interested"],
            "reply":       s["reply"],
            "unsubscribe": s["unsubscribe"],
            "positive":    positive,
        })

    result.sort(key=lambda x: (x["positive"], x["reply"]), reverse=True)
    return result


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
        if status in ("Lead", "Unsubscribe", "Interested", "Reply") and email and "@" in email:
            results.append({"email": email, "status": status})
    return results


# ── Lead-grey CF rule helper ──────────────────────────────────────────────────

def _apply_lead_grey_cf(
    sh: gspread.Spreadsheet,
    ws: gspread.Worksheet,
    ls_col_idx: int,   # 1-based
    n_cols: int,
) -> None:
    """
    Add a conditional-format rule to *sh* that turns the whole row grey (except
    the Lead Status column) whenever Lead Status = "Lead".

    This is called from mark_email_in_sheet for EXISTING sheets that were
    created before this rule was added to _apply_sheet_formatting.  It runs
    inside the background thread so the extra Sheets API round-trip is free
    from the HTTP response's perspective.

    The rule is inserted at index 0 (highest priority) so it overrides the
    sent-blue rule.  Because the LS column is excluded from the rule's ranges,
    the existing green Lead cell colour still shows on that cell.

    Google Sheets allows up to 400 CF rules per sheet; adding this rule once
    per manual lead-marking event is safe in all realistic usage scenarios.
    """
    sheet_gid = ws.id
    ls_0      = ls_col_idx - 1          # 0-based column index
    ls_letter = _col_letter(ls_col_idx)
    formula   = f'=${ls_letter}2="Lead"'
    row_count = ws.row_count or 1000

    ranges: list[dict] = []
    if ls_0 > 0:
        ranges.append({
            "sheetId":          sheet_gid,
            "startRowIndex":    1,
            "endRowIndex":      row_count,
            "startColumnIndex": 0,
            "endColumnIndex":   ls_0,
        })
    if ls_0 + 1 < n_cols:
        ranges.append({
            "sheetId":          sheet_gid,
            "startRowIndex":    1,
            "endRowIndex":      row_count,
            "startColumnIndex": ls_0 + 1,
            "endColumnIndex":   n_cols,
        })
    if not ranges:
        return

    sh.batch_update({"requests": [{
        "addConditionalFormatRule": {
            "rule": {
                "ranges": ranges,
                "booleanRule": {
                    "condition": {
                        "type":   "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": formula}],
                    },
                    "format": {"backgroundColor": _rgb("E0E0E0")},
                },
            },
            "index": 0,
        }
    }]})


# ── DNC reason → Lead Status mapping ─────────────────────────────────────────

_REASON_TO_LEAD_STATUS: dict[str, str] = {
    "lead":       "Lead",
    "interested": "Interested",
    "reply":      "Reply",
    "opt_out":    "Unsubscribe",
    "manual":     "Unsubscribe",
}


def mark_email_in_sheet(sheet_id: str, email_or_domain: str, reason: str = "manual") -> int:
    """
    Find every matching row in the sheet and set 'Lead Status' to the value
    that corresponds to *reason*.

    *email_or_domain* can be:
      • A full email address (``user@company.com``) — only that exact row is updated.
      • A bare domain (``company.com``, no "@") — ALL rows whose Recipient Email
        belongs to that domain are updated.  This is used for "lead" reason when
        a whole company domain should be stamped as "Lead".

    Returns the number of rows updated (0 if no match or required columns absent).
    """
    lead_status   = _REASON_TO_LEAD_STATUS.get(reason.lower(), "Unsubscribe")
    target        = email_or_domain.lower().strip()
    is_domain_key = "@" not in target   # True → domain-level match

    gc = _client()
    sh = gc.open_by_key(sheet_id)
    ws = sh.sheet1

    # Find column positions from the header row
    headers = ws.row_values(1)
    try:
        email_col_idx = headers.index("Recipient Email") + 1   # 1-based
        ls_col_idx    = headers.index("Lead Status")    + 1
    except ValueError:
        return 0   # sheet doesn't have the expected columns

    # Fetch the entire email column (header included at position 0)
    col_values = ws.col_values(email_col_idx)

    # Collect the A1 ranges that need updating (skip header row 1)
    updates = []
    for row_num, cell_val in enumerate(col_values, start=1):
        if row_num == 1:
            continue   # header
        cell_email = str(cell_val).lower().strip()
        if is_domain_key:
            # Domain-level: match any row whose email belongs to this domain
            if "@" in cell_email and cell_email.split("@")[1] == target:
                updates.append({"range": f"{_col_letter(ls_col_idx)}{row_num}",
                                 "values": [[lead_status]]})
        else:
            # Exact email match
            if cell_email == target:
                updates.append({"range": f"{_col_letter(ls_col_idx)}{row_num}",
                                 "values": [[lead_status]]})

    if updates:
        ws.batch_update(updates)

        # For lead rows: ensure the grey-row CF rule exists on this sheet.
        # This handles sheets that were created before the rule was added to
        # _apply_sheet_formatting — it runs in the background thread so the
        # extra batchUpdate call does not block the HTTP response.
        if lead_status == "Lead":
            try:
                _apply_lead_grey_cf(sh, ws, ls_col_idx, len(headers))
            except Exception:
                pass   # non-fatal; cell values are correct regardless

    return len(updates)


# ── Internal helpers ──────────────────────────────────────────────────────────

def _col_letter(n: int) -> str:
    """Convert a 1-based column index to an Excel-style letter (1→A, 26→Z, 27→AA)."""
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result or "A"
