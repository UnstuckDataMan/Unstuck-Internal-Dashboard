from __future__ import annotations

import io
import os
import threading
import time
import uuid
from pathlib import Path
from typing import Optional

import pandas as pd
import requests as http_req
from fastapi import APIRouter, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.responses import Response

from app.deps import templates

router = APIRouter()

SUPABASE_URL      = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")

MAX_UPLOAD_BYTES = 20 * 1024 * 1024
CHUNK_SIZE       = 500   # emails per Supabase batch query
PAGE_SIZE        = 50    # rows per page in Manage tab

_EMAIL_CANDIDATES = [
    "email", "email address", "e-mail", "e_mail", "emailaddress",
    "work email", "email_address", "mail", "email addr",
]

# ── in-memory token store (same pattern as gender.py / city.py) ───────────────

_store: dict[str, dict] = {}
_store_lock = threading.Lock()
_TOKEN_TTL  = 3600


def _evict_expired() -> None:
    cutoff = time.monotonic() - _TOKEN_TTL
    with _store_lock:
        for k in [k for k, v in _store.items() if v["created_at"] < cutoff]:
            del _store[k]


def _store_result(data: bytes, mime: str, filename: str) -> str:
    _evict_expired()
    token = str(uuid.uuid4())
    with _store_lock:
        _store[token] = {
            "data": data,
            "mime": mime,
            "filename": filename,
            "created_at": time.monotonic(),
        }
    return token


# ── Supabase helpers ───────────────────────────────────────────────────────────

def _sb_headers(prefer: Optional[str] = None) -> dict:
    h = {
        "apikey":        SUPABASE_ANON_KEY,
        "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
        "Content-Type":  "application/json",
    }
    if prefer:
        h["Prefer"] = prefer
    return h


def _sb_configured() -> bool:
    return bool(SUPABASE_URL and SUPABASE_ANON_KEY)


# ── file helpers ───────────────────────────────────────────────────────────────

def _read_file(contents: bytes, filename: str) -> tuple[pd.DataFrame, str]:
    name = filename.lower()
    if name.endswith(".csv"):
        return pd.read_csv(
            io.BytesIO(contents), dtype=str, keep_default_na=False, na_values=[]
        ), "csv"
    if name.endswith((".xls", ".xlsx")):
        return pd.read_excel(
            io.BytesIO(contents), dtype=str, engine="openpyxl",
            keep_default_na=False, na_values=[]
        ), "xlsx"
    raise ValueError("Unsupported file type. Upload a .csv or .xlsx file.")


def _detect_email_column(headers: list[str]) -> list[str]:
    """Return all headers that look like an email column."""
    def norm(h: str) -> str:
        return " ".join(str(h).strip().lower().replace("_", " ").replace("-", " ").split())

    matches = []
    for h in headers:
        n = norm(h)
        for cand in _EMAIL_CANDIDATES:
            if n == cand or cand in n or n in cand:
                matches.append(h)
                break
    return matches


def _write_output(df: pd.DataFrame, ext: str, base: str, suffix: str) -> tuple[bytes, str, str]:
    out_name = f"{base}_{suffix}.{ext}"
    if ext == "csv":
        return df.to_csv(index=False).encode("utf-8"), "text/csv", out_name
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return (
        buf.getvalue(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        out_name,
    )


# ── DNC lookup (batched) ───────────────────────────────────────────────────────

def _fetch_dnc_set(client_id: str, emails: list[str]) -> set[str]:
    """
    Query Supabase in chunks of CHUNK_SIZE.
    Returns set of lowercased emails that are on the DNC list.
    Emails in dnc_entries are always stored lowercase, so plain `in` filter works.
    """
    dnc: set[str] = set()
    for i in range(0, len(emails), CHUNK_SIZE):
        chunk = emails[i : i + CHUNK_SIZE]
        email_filter = "(" + ",".join(chunk) + ")"
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/dnc_entries",
            headers=_sb_headers(),
            params={
                "select":    "email",
                "client_id": f"eq.{client_id}",
                "email":     f"in.{email_filter}",
            },
            timeout=30,
        )
        r.raise_for_status()
        for row in r.json():
            dnc.add(row["email"])
    return dnc


# ── routes ─────────────────────────────────────────────────────────────────────

@router.get("/dnc-removal")
async def dnc_page(request: Request):
    return templates.TemplateResponse("dnc.html", {"request": request, "active": "dnc"})


@router.get("/api/dnc/clients")
async def get_clients(request: Request):
    if not _sb_configured():
        return templates.TemplateResponse(
            "partials/dnc_clients_options.html",
            {"request": request, "clients": [], "error": "Supabase not configured."},
        )
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/clients",
            headers=_sb_headers(),
            params={"select": "id,name", "order": "name.asc"},
            timeout=10,
        )
        r.raise_for_status()
        clients = r.json()
    except Exception as e:
        clients = []
    return templates.TemplateResponse(
        "partials/dnc_clients_options.html",
        {"request": request, "clients": clients},
    )


@router.post("/api/dnc/scrub")
async def api_dnc_scrub(
    request:   Request,
    file:      UploadFile = File(...),
    client_id: str        = Form(...),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_scrub_result.html",
            {"request": request, "error": msg},
        )

    if not _sb_configured():
        return error("Supabase is not configured. Set SUPABASE_URL and SUPABASE_ANON_KEY.")

    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        return error("File too large. Maximum 20 MB.")

    try:
        df, ext = _read_file(contents, file.filename or "upload.csv")
    except Exception as e:
        return error(f"Could not read file: {e}")

    if df.empty:
        return error("The uploaded file contains no rows.")

    matches = _detect_email_column(list(df.columns))
    if len(matches) == 0:
        return error(
            "No email column found. Ensure the file has a column whose header "
            "contains the word 'email'."
        )
    if len(matches) > 1:
        return error(
            f"Multiple possible email columns found: {', '.join(matches)}. "
            "Rename the file so there is exactly one email column."
        )
    email_col = matches[0]

    # Normalise emails for comparison (lowercase + strip)
    df["_norm_email"] = df[email_col].astype(str).str.lower().str.strip()
    uploaded_count = len(df)

    try:
        dnc_set = _fetch_dnc_set(client_id, df["_norm_email"].tolist())
    except Exception as e:
        return error(f"Supabase error during DNC lookup: {e}")

    removed_mask    = df["_norm_email"].isin(dnc_set)
    df_clean        = df[~removed_mask].drop(columns=["_norm_email"])
    df_removed      = df[ removed_mask].drop(columns=["_norm_email"])
    removed_count   = len(df_removed)
    remaining_count = len(df_clean)

    # Log scrub (non-fatal)
    try:
        http_req.post(
            f"{SUPABASE_URL}/rest/v1/scrub_logs",
            headers=_sb_headers("return=minimal"),
            json={
                "client_id":       client_id,
                "uploaded_count":  uploaded_count,
                "removed_count":   removed_count,
                "remaining_count": remaining_count,
                "performed_by":    "dashboard_user",
            },
            timeout=10,
        ).raise_for_status()
    except Exception:
        pass

    base = Path(file.filename or "prospects").stem
    clean_bytes, clean_mime, clean_name       = _write_output(df_clean,   ext, base, "clean")
    removed_bytes, removed_mime, removed_name = _write_output(df_removed, ext, base, "removed")

    clean_token   = _store_result(clean_bytes,   clean_mime,   clean_name)
    removed_token = _store_result(removed_bytes, removed_mime, removed_name)

    removed_emails_preview = df_removed[email_col].tolist()[:200]

    return templates.TemplateResponse(
        "partials/dnc_scrub_result.html",
        {
            "request":              request,
            "uploaded_count":       uploaded_count,
            "removed_count":        removed_count,
            "remaining_count":      remaining_count,
            "clean_token":          clean_token,
            "clean_filename":       clean_name,
            "removed_token":        removed_token,
            "removed_filename":     removed_name,
            "removed_emails":       removed_emails_preview,
            "removed_emails_total": removed_count,
        },
    )


@router.get("/api/dnc/download/{token}")
async def dnc_download_clean(token: str):
    with _store_lock:
        entry = _store.pop(token, None)
    if entry is None:
        raise HTTPException(404, "Download link has expired or already been used.")
    return Response(
        content=entry["data"],
        media_type=entry["mime"],
        headers={"Content-Disposition": f'attachment; filename="{entry["filename"]}"'},
    )


@router.get("/api/dnc/download-removed/{token}")
async def dnc_download_removed(token: str):
    with _store_lock:
        entry = _store.pop(token, None)
    if entry is None:
        raise HTTPException(404, "Download link has expired or already been used.")
    return Response(
        content=entry["data"],
        media_type=entry["mime"],
        headers={"Content-Disposition": f'attachment; filename="{entry["filename"]}"'},
    )


@router.get("/api/dnc/entries")
async def get_dnc_entries(
    request:   Request,
    client_id: str = Query(...),
    offset:    int = Query(0, ge=0),
    search:    str = Query(""),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_entries_table.html",
            {"request": request, "error": msg, "entries": [], "total": 0,
             "offset": 0, "page_size": PAGE_SIZE, "has_prev": False,
             "has_next": False, "client_id": client_id, "search": search},
        )

    if not _sb_configured():
        return error("Supabase is not configured.")

    params: dict = {
        "select":    "id,email,reason,added_by,notes,created_at",
        "client_id": f"eq.{client_id}",
        "order":     "created_at.desc",
        "limit":     str(PAGE_SIZE),
        "offset":    str(offset),
    }
    if search.strip():
        params["email"] = f"ilike.*{search.strip()}*"

    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/dnc_entries",
            headers={**_sb_headers(), "Prefer": "count=exact"},
            params=params,
            timeout=15,
        )
        r.raise_for_status()
    except Exception as e:
        return error(f"Supabase error: {e}")

    total = 0
    cr = r.headers.get("Content-Range", "")
    if "/" in cr:
        try:
            total = int(cr.split("/")[1])
        except ValueError:
            pass

    entries  = r.json()
    has_prev = offset > 0
    has_next = (offset + PAGE_SIZE) < total

    return templates.TemplateResponse(
        "partials/dnc_entries_table.html",
        {
            "request":   request,
            "entries":   entries,
            "total":     total,
            "offset":    offset,
            "page_size": PAGE_SIZE,
            "has_prev":  has_prev,
            "has_next":  has_next,
            "client_id": client_id,
            "search":    search,
        },
    )


@router.post("/api/dnc/entries")
async def add_dnc_entry(
    request:   Request,
    client_id: str = Form(...),
    email:     str = Form(...),
    reason:    str = Form("manual"),
    notes:     str = Form(""),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_entries_table.html",
            {"request": request, "error": msg, "entries": [], "total": 0,
             "offset": 0, "page_size": PAGE_SIZE, "has_prev": False,
             "has_next": False, "client_id": client_id, "search": ""},
        )

    if not _sb_configured():
        return error("Supabase is not configured.")

    email_norm = email.lower().strip()
    if not email_norm or "@" not in email_norm:
        return error("Please enter a valid email address.")

    payload: dict = {
        "client_id": client_id,
        "email":     email_norm,
        "reason":    reason.strip() or "manual",
        "added_by":  "dashboard_user",
    }
    if notes.strip():
        payload["notes"] = notes.strip()

    try:
        r = http_req.post(
            f"{SUPABASE_URL}/rest/v1/dnc_entries",
            headers=_sb_headers("return=minimal"),
            json=payload,
            timeout=10,
        )
        if r.status_code == 409:
            return error(f"{email_norm} is already on the DNC list for this client.")
        r.raise_for_status()
    except Exception as e:
        return error(f"Supabase error: {e}")

    return await get_dnc_entries(request, client_id=client_id, offset=0, search="")


@router.post("/api/dnc/entries/bulk")
async def bulk_import_dnc(
    request:   Request,
    client_id: str        = Form(...),
    file:      UploadFile = File(...),
    reason:    str        = Form("bulk_import"),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_entries_table.html",
            {"request": request, "error": msg, "entries": [], "total": 0,
             "offset": 0, "page_size": PAGE_SIZE, "has_prev": False,
             "has_next": False, "client_id": client_id, "search": ""},
        )

    if not _sb_configured():
        return error("Supabase is not configured.")

    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        return error("File too large. Maximum 20 MB.")

    try:
        df, _ = _read_file(contents, file.filename or "bulk.csv")
    except Exception as e:
        return error(f"Could not read file: {e}")

    matches = _detect_email_column(list(df.columns))
    if not matches:
        return error("No email column found in the uploaded file.")
    email_col = matches[0]

    emails = (
        df[email_col].astype(str)
        .str.lower().str.strip()
        .drop_duplicates()
        .tolist()
    )
    emails = [e for e in emails if e and "@" in e]

    if not emails:
        return error("No valid email addresses found in the file.")

    reason_clean = reason.strip() or "bulk_import"

    for i in range(0, len(emails), CHUNK_SIZE):
        chunk = emails[i : i + CHUNK_SIZE]
        rows = [
            {
                "client_id": client_id,
                "email":     e,
                "reason":    reason_clean,
                "added_by":  "dashboard_user",
            }
            for e in chunk
        ]
        try:
            r = http_req.post(
                f"{SUPABASE_URL}/rest/v1/dnc_entries",
                headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
                params={"on_conflict": "client_id,email"},
                json=rows,
                timeout=30,
            )
            r.raise_for_status()
        except Exception as e:
            return error(f"Supabase error during import: {e}")

    return await get_dnc_entries(request, client_id=client_id, offset=0, search="")


@router.delete("/api/dnc/entries/{entry_id}")
async def delete_dnc_entry(
    request:   Request,
    entry_id:  str,
    client_id: str = Query(...),
):
    if not _sb_configured():
        return templates.TemplateResponse(
            "partials/dnc_entries_table.html",
            {"request": request, "error": "Supabase is not configured.", "entries": [],
             "total": 0, "offset": 0, "page_size": PAGE_SIZE, "has_prev": False,
             "has_next": False, "client_id": client_id, "search": ""},
        )
    try:
        http_req.delete(
            f"{SUPABASE_URL}/rest/v1/dnc_entries",
            headers=_sb_headers(),
            params={"id": f"eq.{entry_id}"},
            timeout=10,
        ).raise_for_status()
    except Exception as e:
        return templates.TemplateResponse(
            "partials/dnc_entries_table.html",
            {"request": request, "error": f"Supabase error: {e}", "entries": [],
             "total": 0, "offset": 0, "page_size": PAGE_SIZE, "has_prev": False,
             "has_next": False, "client_id": client_id, "search": ""},
        )

    return await get_dnc_entries(request, client_id=client_id, offset=0, search="")
