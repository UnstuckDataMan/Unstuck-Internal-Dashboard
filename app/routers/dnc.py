from __future__ import annotations

import csv
import hmac
import io
import os
import threading
import time
import uuid
from datetime import date, timedelta
from pathlib import Path
from typing import List, Optional

import pandas as pd
import requests as http_req
from fastapi import APIRouter, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.responses import HTMLResponse, Response, StreamingResponse

from app.deps import templates

router = APIRouter()

SUPABASE_URL      = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")

MAX_UPLOAD_BYTES = 20 * 1024 * 1024
CHUNK_SIZE       = 500   # emails per Supabase batch query
PAGE_SIZE        = 25    # rows per page in Manage tab

_EMAIL_CANDIDATES = [
    "email", "email address", "e-mail", "e_mail", "emailaddress",
    "work email", "email_address", "mail", "email addr",
]

# ── in-memory token store ──────────────────────────────────────────────────────

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


def _parse_total(headers: dict) -> int:
    cr = headers.get("Content-Range", "")
    if "/" in cr:
        try:
            return int(cr.split("/")[1])
        except ValueError:
            pass
    return 0


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


def _normalize_industry_name(name: str) -> str:
    """Title-case and trim an industry name: '  financial services  ' → 'Financial Services'"""
    return " ".join(word.capitalize() for word in name.strip().split())


# ── DNC lookup (batched, supports full emails and domain blocks) ───────────────

def _batch_query_dnc(client_id: str, values: list[str]) -> set[str]:
    """Query dnc_entries for an exact match against a list of values."""
    matched: set[str] = set()
    for i in range(0, len(values), CHUNK_SIZE):
        chunk = values[i : i + CHUNK_SIZE]
        val_filter = "(" + ",".join(chunk) + ")"
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/dnc_entries",
            headers=_sb_headers(),
            params={
                "select":    "email",
                "client_id": f"eq.{client_id}",
                "email":     f"in.{val_filter}",
            },
            timeout=30,
        )
        r.raise_for_status()
        for row in r.json():
            matched.add(row["email"])
    return matched


def _fetch_dnc_matches(client_id: str, norm_emails: list[str]) -> set[str]:
    """
    Returns the subset of norm_emails that should be removed by DNC check.
    Checks both exact email matches and domain-level blocks.
    """
    dnc_emails = _batch_query_dnc(client_id, norm_emails)
    unique_domains = list({e.split("@")[1] for e in norm_emails if "@" in e})
    dnc_domains = _batch_query_dnc(client_id, unique_domains) if unique_domains else set()

    to_remove: set[str] = set()
    for email in norm_emails:
        if email in dnc_emails:
            to_remove.add(email)
        elif "@" in email and email.split("@")[1] in dnc_domains:
            to_remove.add(email)
    return to_remove


def _fetch_contacted_matches(
    client_id: str,
    norm_emails: list[str],
    cutoff_date: str,
    industry_id: str = "",
    location: str = "",
) -> set[str]:
    """
    Returns the subset of norm_emails contacted on or after cutoff_date.
    Optionally scoped to a specific industry.
    """
    contacted: set[str] = set()
    for i in range(0, len(norm_emails), CHUNK_SIZE):
        chunk = norm_emails[i : i + CHUNK_SIZE]
        email_filter = "(" + ",".join(chunk) + ")"
        params: list[tuple[str, str]] = [
            ("select",       "email"),
            ("client_id",    f"eq.{client_id}"),
            ("contacted_at", f"gte.{cutoff_date}"),
            ("email",        f"in.{email_filter}"),
        ]
        if industry_id:
            params.append(("client_industry_id", f"eq.{industry_id}"))
        if location.strip():
            params.append(("location", f"ilike.*{location.strip()}*"))
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/contacted_prospects",
            headers=_sb_headers(),
            params=params,
            timeout=30,
        )
        r.raise_for_status()
        for row in r.json():
            contacted.add(row["email"])
    return contacted


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
    except Exception:
        clients = []
    return templates.TemplateResponse(
        "partials/dnc_clients_options.html",
        {"request": request, "clients": clients},
    )


@router.post("/api/dnc/clients")
async def create_client(
    request: Request,
    name:    str = Form(...),
):
    """Create a new client and return OOB-swap HTML to refresh all client selects."""

    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_client_created.html",
            {"request": request, "clients": [], "new_client_id": None, "error": msg},
        )

    if not _sb_configured():
        return error("Supabase is not configured.")

    name_norm = name.strip()
    if not name_norm:
        return error("Please enter a client name.")

    try:
        r = http_req.post(
            f"{SUPABASE_URL}/rest/v1/clients",
            headers=_sb_headers("return=representation"),
            json={"name": name_norm},
            timeout=10,
        )
        if r.status_code == 409:
            return error(f"A client named \"{name_norm}\" already exists.")
        r.raise_for_status()
        new_client    = r.json()[0]
        new_client_id = new_client["id"]
    except Exception as exc:
        return error(f"Could not create client: {exc}")

    # Re-fetch full sorted list for OOB select updates
    try:
        rc = http_req.get(
            f"{SUPABASE_URL}/rest/v1/clients",
            headers=_sb_headers(),
            params={"select": "id,name", "order": "name.asc"},
            timeout=10,
        )
        rc.raise_for_status()
        clients = rc.json()
    except Exception:
        clients = [{"id": new_client_id, "name": name_norm}]

    return templates.TemplateResponse(
        "partials/dnc_client_created.html",
        {"request": request, "clients": clients, "new_client_id": new_client_id, "error": None},
    )


# ── scrub ──────────────────────────────────────────────────────────────────────

@router.post("/api/dnc/scrub")
async def api_dnc_scrub(
    request:             Request,
    file:                UploadFile = File(...),
    client_id:           str        = Form(...),
    industry_id:         str        = Form(""),
    remove_contacted:    str        = Form(""),        # "on" when checkbox checked
    lookback_days_raw:   str        = Form("30"),      # "30", "60", "90", or "custom"
    lookback_custom_from: str       = Form(""),        # YYYY-MM-DD for custom range
    location:            str        = Form(""),         # location filter for contacted check
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

    df["_norm_email"] = df[email_col].astype(str).str.lower().str.strip()
    uploaded_count = len(df)

    # ── Pass 1: DNC check ─────────────────────────────────────────
    try:
        dnc_removed_set = _fetch_dnc_matches(client_id, df["_norm_email"].tolist())
    except Exception as e:
        return error(f"Supabase error during DNC lookup: {e}")

    # ── Pass 2: Recently contacted check (optional) ───────────────
    contacted_removed_set: set[str] = set()
    cutoff_date: Optional[str] = None

    if remove_contacted == "on":
        if lookback_days_raw == "custom":
            if not lookback_custom_from:
                return error("Please select a start date for the custom lookback range.")
            cutoff_date = lookback_custom_from
        else:
            try:
                days = int(lookback_days_raw)
            except ValueError:
                days = 30
            cutoff_date = (date.today() - timedelta(days=days)).isoformat()

        try:
            contacted_removed_set = _fetch_contacted_matches(
                client_id, df["_norm_email"].tolist(), cutoff_date, industry_id, location
            )
        except Exception as e:
            return error(f"Supabase error during recently-contacted lookup: {e}")

    # DNC takes priority — emails in both sets count only as DNC
    contacted_only = contacted_removed_set - dnc_removed_set
    all_removed    = dnc_removed_set | contacted_only

    # ── Split and annotate ────────────────────────────────────────
    removed_mask    = df["_norm_email"].isin(all_removed)
    df_clean        = df[~removed_mask].drop(columns=["_norm_email"])

    df_rem = df[removed_mask].copy()
    df_rem["removal_reason"] = df_rem["_norm_email"].map(
        lambda e: "DNC" if e in dnc_removed_set else "Recently Contacted"
    )
    df_rem = df_rem.drop(columns=["_norm_email"])

    dnc_removed_count       = int((df_rem["removal_reason"] == "DNC").sum())
    contacted_removed_count = int((df_rem["removal_reason"] == "Recently Contacted").sum())
    removed_count           = len(df_rem)
    remaining_count         = len(df_clean)

    # ── Log scrub (non-fatal) ─────────────────────────────────────
    try:
        log_payload: dict = {
            "client_id":               client_id,
            "uploaded_count":          uploaded_count,
            "removed_count":           removed_count,
            "remaining_count":         remaining_count,
            "performed_by":            "dashboard_user",
            "contacted_removed_count": contacted_removed_count,
        }
        if cutoff_date and lookback_days_raw != "custom":
            log_payload["lookback_days"] = int(lookback_days_raw)
        if industry_id:
            log_payload["industry_filter"] = industry_id
        http_req.post(
            f"{SUPABASE_URL}/rest/v1/scrub_logs",
            headers=_sb_headers("return=minimal"),
            json=log_payload,
            timeout=10,
        ).raise_for_status()
    except Exception:
        pass

    # ── Serialize outputs ─────────────────────────────────────────
    base = Path(file.filename or "prospects").stem
    clean_bytes, clean_mime, clean_name       = _write_output(df_clean, ext, base, "clean")
    removed_bytes, removed_mime, removed_name = _write_output(df_rem,   ext, base, "removed")

    clean_token   = _store_result(clean_bytes,   clean_mime,   clean_name)
    removed_token = _store_result(removed_bytes, removed_mime, removed_name)

    # Separate preview lists (first 200 each)
    dnc_preview       = df_rem[df_rem["removal_reason"] == "DNC"][email_col].tolist()[:200]
    contacted_preview = df_rem[df_rem["removal_reason"] == "Recently Contacted"][email_col].tolist()[:200]

    return templates.TemplateResponse(
        "partials/dnc_scrub_result.html",
        {
            "request":                  request,
            "uploaded_count":           uploaded_count,
            "dnc_removed_count":        dnc_removed_count,
            "contacted_removed_count":  contacted_removed_count,
            "removed_count":            removed_count,
            "remaining_count":          remaining_count,
            "clean_token":              clean_token,
            "clean_filename":           clean_name,
            "removed_token":            removed_token,
            "removed_filename":         removed_name,
            "dnc_preview":              dnc_preview,
            "contacted_preview":        contacted_preview,
            "show_contacted":           remove_contacted == "on",
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


# ── industries ─────────────────────────────────────────────────────────────────

def _get_industries_for_client(client_id: str) -> list[dict]:
    r = http_req.get(
        f"{SUPABASE_URL}/rest/v1/client_industries",
        headers=_sb_headers(),
        params={"select": "id,name", "client_id": f"eq.{client_id}", "order": "name.asc"},
        timeout=10,
    )
    r.raise_for_status()
    return r.json()


@router.get("/api/dnc/industries")
async def get_industry_options(request: Request, client_id: str = Query("")):
    """Returns <option> elements for industry dropdowns (scrub tab + manage tab upload/add forms)."""
    industries = []
    if client_id and _sb_configured():
        try:
            industries = _get_industries_for_client(client_id)
        except Exception:
            pass
    return templates.TemplateResponse(
        "partials/dnc_industries_options.html",
        {"request": request, "industries": industries},
    )


@router.get("/api/dnc/industries/manage")
async def get_industries_managed(request: Request, client_id: str = Query("")):
    """Returns chips HTML + OOB option updates for all industry selects in the manage tab."""
    industries = []
    if client_id and _sb_configured():
        try:
            industries = _get_industries_for_client(client_id)
        except Exception:
            pass
    return templates.TemplateResponse(
        "partials/dnc_industries_managed.html",
        {"request": request, "industries": industries, "client_id": client_id},
    )


@router.post("/api/dnc/industries")
async def add_industry(
    request:   Request,
    client_id: str = Form(...),
    name:      str = Form(...),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_industries_managed.html",
            {"request": request, "industries": [], "client_id": client_id, "error": msg},
        )

    if not _sb_configured():
        return error("Supabase is not configured.")

    name_norm = _normalize_industry_name(name)
    if not name_norm:
        return error("Please enter an industry name.")

    try:
        r = http_req.post(
            f"{SUPABASE_URL}/rest/v1/client_industries",
            headers=_sb_headers("return=minimal"),
            json={"client_id": client_id, "name": name_norm},
            timeout=10,
        )
        if r.status_code == 409:
            return error(f"'{name_norm}' already exists for this client.")
        r.raise_for_status()
    except Exception as e:
        return error(f"Supabase error: {e}")

    industries = _get_industries_for_client(client_id)
    return templates.TemplateResponse(
        "partials/dnc_industries_managed.html",
        {"request": request, "industries": industries, "client_id": client_id},
    )


@router.delete("/api/dnc/industries/{industry_id}")
async def delete_industry(
    request:     Request,
    industry_id: str,
    client_id:   str = Query(...),
):
    if _sb_configured():
        try:
            http_req.delete(
                f"{SUPABASE_URL}/rest/v1/client_industries",
                headers=_sb_headers(),
                params={"id": f"eq.{industry_id}"},
                timeout=10,
            ).raise_for_status()
        except Exception:
            pass

    industries = _get_industries_for_client(client_id) if _sb_configured() else []
    return templates.TemplateResponse(
        "partials/dnc_industries_managed.html",
        {"request": request, "industries": industries, "client_id": client_id},
    )


# ── contacted prospects ────────────────────────────────────────────────────────

@router.get("/api/dnc/contacted")
async def get_contacted(
    request:     Request,
    client_id:   str = Query(...),
    industry_id: str = Query(""),
    search:      str = Query(""),
    date_from:   str = Query(""),
    date_to:     str = Query(""),
    location:    str = Query(""),
    offset:      int = Query(0, ge=0),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_contacted_table.html",
            {"request": request, "error": msg, "entries": [], "total": 0,
             "offset": 0, "page_size": PAGE_SIZE, "has_prev": False, "has_next": False,
             "client_id": client_id, "industry_id": industry_id,
             "search": search, "date_from": date_from, "date_to": date_to, "location": location},
        )

    if not _sb_configured():
        return error("Supabase is not configured.")

    params: list[tuple[str, str]] = [
        ("select",    "id,email,location,contacted_at,campaign_name,source,created_at,client_industries(name)"),
        ("client_id", f"eq.{client_id}"),
        ("order",     "contacted_at.desc,created_at.desc"),
        ("limit",     str(PAGE_SIZE)),
        ("offset",    str(offset)),
    ]
    if industry_id:
        params.append(("client_industry_id", f"eq.{industry_id}"))
    if search.strip():
        params.append(("email", f"ilike.*{search.strip()}*"))
    if date_from:
        params.append(("contacted_at", f"gte.{date_from}"))
    if date_to:
        params.append(("contacted_at", f"lte.{date_to}"))
    if location.strip():
        params.append(("location", f"ilike.*{location.strip()}*"))

    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/contacted_prospects",
            headers={**_sb_headers(), "Prefer": "count=exact"},
            params=params,
            timeout=15,
        )
        r.raise_for_status()
    except Exception as e:
        return error(f"Supabase error: {e}")

    total    = _parse_total(r.headers)
    entries  = r.json()
    has_prev = offset > 0
    has_next = (offset + PAGE_SIZE) < total

    return templates.TemplateResponse(
        "partials/dnc_contacted_table.html",
        {
            "request":     request,
            "entries":     entries,
            "total":       total,
            "offset":      offset,
            "page_size":   PAGE_SIZE,
            "has_prev":    has_prev,
            "has_next":    has_next,
            "client_id":   client_id,
            "industry_id": industry_id,
            "search":      search,
            "date_from":   date_from,
            "date_to":     date_to,
            "location":    location,
        },
    )


@router.post("/api/dnc/contacted")
async def add_contacted(
    request:       Request,
    client_id:     str = Form(...),
    industry_id:   str = Form(...),
    email:         str = Form(...),
    contacted_at:  str = Form(...),
    campaign_name: str = Form(""),
    location:      str = Form(""),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_contacted_table.html",
            {"request": request, "error": msg, "entries": [], "total": 0,
             "offset": 0, "page_size": PAGE_SIZE, "has_prev": False, "has_next": False,
             "client_id": client_id, "industry_id": industry_id,
             "search": "", "date_from": "", "date_to": "", "location": ""},
        )

    if not _sb_configured():
        return error("Supabase is not configured.")

    email_norm = email.lower().strip()
    if not email_norm or "@" not in email_norm:
        return error("Please enter a valid email address.")
    if not contacted_at:
        return error("Please select a contact date.")
    if not industry_id:
        return error("Please select an industry.")

    payload: dict = {
        "client_id":          client_id,
        "client_industry_id": industry_id,
        "email":              email_norm,
        "contacted_at":       contacted_at,
        "source":             "manual",
    }
    if campaign_name.strip():
        payload["campaign_name"] = campaign_name.strip()
    if location.strip():
        payload["location"] = location.strip()

    try:
        r = http_req.post(
            f"{SUPABASE_URL}/rest/v1/contacted_prospects",
            headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
            params={"on_conflict": "client_id,client_industry_id,email,contacted_at"},
            json=payload,
            timeout=10,
        )
        r.raise_for_status()
    except Exception as e:
        return error(f"Supabase error: {e}")

    return await get_contacted(request, client_id=client_id, industry_id=industry_id,
                               search="", date_from="", date_to="", location="", offset=0)


@router.post("/api/dnc/contacted/upload")
async def upload_contacted(
    request:       Request,
    client_id:     str        = Form(...),
    industry_id:   str        = Form(...),
    file:          UploadFile = File(...),
    contacted_at:  str        = Form(...),
    campaign_name: str        = Form(""),
    location:      str        = Form(""),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_contacted_table.html",
            {"request": request, "error": msg, "entries": [], "total": 0,
             "offset": 0, "page_size": PAGE_SIZE, "has_prev": False, "has_next": False,
             "client_id": client_id, "industry_id": industry_id,
             "search": "", "date_from": "", "date_to": "", "location": ""},
        )

    if not _sb_configured():
        return error("Supabase is not configured.")
    if not industry_id:
        return error("Please select an industry before uploading.")
    if not contacted_at:
        return error("Please select a contact date.")

    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        return error("File too large. Maximum 20 MB.")

    try:
        df, _ = _read_file(contents, file.filename or "contacts.csv")
    except Exception as e:
        return error(f"Could not read file: {e}")

    col_matches = _detect_email_column(list(df.columns))
    if not col_matches:
        return error("No email column found in the uploaded file.")
    email_col = col_matches[0]

    emails = (
        df[email_col].astype(str)
        .str.lower().str.strip()
        .drop_duplicates()
        .tolist()
    )
    emails = [e for e in emails if e and "@" in e]
    if not emails:
        return error("No valid email addresses found in the file.")

    campaign = campaign_name.strip() or None

    for i in range(0, len(emails), CHUNK_SIZE):
        chunk = emails[i : i + CHUNK_SIZE]
        rows = [
            {
                "client_id":          client_id,
                "client_industry_id": industry_id,
                "email":              e,
                "contacted_at":       contacted_at,
                "source":             "csv_upload",
                **({"campaign_name": campaign} if campaign else {}),
                **({"location": location.strip()} if location.strip() else {}),
            }
            for e in chunk
        ]
        try:
            r = http_req.post(
                f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
                params={"on_conflict": "client_id,client_industry_id,email,contacted_at"},
                json=rows,
                timeout=30,
            )
            r.raise_for_status()
        except Exception as e:
            return error(f"Supabase error during upload: {e}")

    return await get_contacted(request, client_id=client_id, industry_id=industry_id,
                               search="", date_from="", date_to="", location="", offset=0)


@router.delete("/api/dnc/contacted/{entry_id}")
async def delete_contacted(
    request:     Request,
    entry_id:    str,
    client_id:   str = Query(...),
    industry_id: str = Query(""),
    location:    str = Query(""),
):
    if _sb_configured():
        try:
            http_req.delete(
                f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                headers=_sb_headers(),
                params={"id": f"eq.{entry_id}"},
                timeout=10,
            ).raise_for_status()
        except Exception:
            pass

    return await get_contacted(request, client_id=client_id, industry_id=industry_id,
                               search="", date_from="", date_to="", location=location, offset=0)


@router.post("/api/dnc/contacted/bulk-delete")
async def bulk_delete_contacted(
    request:     Request,
    client_id:   str       = Form(...),
    industry_id: str       = Form(""),
    location:    str       = Form(""),
    ids:         List[str] = Form(default=[]),
):
    if ids and _sb_configured():
        id_filter = "(" + ",".join(ids) + ")"
        try:
            http_req.delete(
                f"{SUPABASE_URL}/rest/v1/contacted_prospects",
                headers=_sb_headers(),
                params={"id": f"in.{id_filter}"},
                timeout=15,
            ).raise_for_status()
        except Exception:
            pass

    return await get_contacted(request, client_id=client_id, industry_id=industry_id,
                               search="", date_from="", date_to="", location=location, offset=0)


# ── purge old contacted records ───────────────────────────────────────────────

@router.post("/api/dnc/contacted/purge")
async def purge_contacted(
    request:       Request,
    client_id:     str = Form(...),
    cutoff_months: str = Form("6"),
    cutoff_custom: str = Form(""),
    password:      str = Form(...),
):
    def error(msg: str):
        return templates.TemplateResponse(
            "partials/dnc_purge_result.html",
            {"request": request, "error": msg},
        )

    purge_pw = os.environ.get("PURGE_PASSWORD", "")
    if not purge_pw:
        return error("PURGE_PASSWORD is not configured on this server. Ask your administrator.")
    if not hmac.compare_digest(password.encode(), purge_pw.encode()):
        return error("Incorrect password.")
    if not client_id:
        return error("No client selected. Select a client before purging.")
    if not _sb_configured():
        return error("Supabase is not configured.")

    if cutoff_months == "custom":
        if not cutoff_custom:
            return error("Please select a custom cutoff date.")
        cutoff = cutoff_custom
    else:
        try:
            months = int(cutoff_months)
        except ValueError:
            months = 6
        cutoff = (date.today() - timedelta(days=months * 30)).isoformat()

    try:
        r = http_req.delete(
            f"{SUPABASE_URL}/rest/v1/contacted_prospects",
            headers={**_sb_headers(), "Prefer": "count=exact"},
            params={
                "client_id":    f"eq.{client_id}",
                "contacted_at": f"lt.{cutoff}",
            },
            timeout=30,
        )
        r.raise_for_status()
        deleted = _parse_total(r.headers)
    except Exception as e:
        return error(f"Supabase error: {e}")

    return templates.TemplateResponse(
        "partials/dnc_purge_result.html",
        {"request": request, "deleted": deleted, "cutoff": cutoff, "client_id": client_id},
    )


# ── DNC entries (existing — unchanged) ────────────────────────────────────────

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

    total    = _parse_total(r.headers)
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


@router.get("/api/dnc/entries/export")
async def export_dnc_entries(
    client_id: str = Query(...),
):
    """Download all DNC entries for a client as a CSV file."""
    if not _sb_configured():
        raise HTTPException(400, "Supabase is not configured.")
    if not client_id:
        raise HTTPException(400, "client_id is required.")

    # Resolve client name for a friendly filename
    client_name = "client"
    try:
        rc = http_req.get(
            f"{SUPABASE_URL}/rest/v1/clients",
            headers=_sb_headers(),
            params={"select": "name", "id": f"eq.{client_id}"},
            timeout=10,
        )
        rc.raise_for_status()
        rows_c = rc.json()
        if rows_c:
            client_name = rows_c[0]["name"].replace(" ", "_")
    except Exception:
        pass

    # Paginated fetch — gather all rows
    all_rows: list = []
    fetch_offset = 0
    batch = 1000
    while True:
        try:
            r = http_req.get(
                f"{SUPABASE_URL}/rest/v1/dnc_entries",
                headers=_sb_headers(),
                params={
                    "select":    "email,reason,added_by,notes,created_at",
                    "client_id": f"eq.{client_id}",
                    "order":     "created_at.desc",
                    "limit":     str(batch),
                    "offset":    str(fetch_offset),
                },
                timeout=30,
            )
            r.raise_for_status()
        except Exception as exc:
            raise HTTPException(502, f"Supabase error: {exc}")
        page = r.json()
        all_rows.extend(page)
        if len(page) < batch:
            break
        fetch_offset += batch

    # Build CSV in memory
    today_str = date.today().isoformat()
    filename  = f"dnc_{client_name}_{today_str}.csv"
    buf       = io.StringIO()
    writer    = csv.DictWriter(
        buf,
        fieldnames=["email", "reason", "added_by", "notes", "created_at"],
        extrasaction="ignore",
        lineterminator="\n",
    )
    writer.writeheader()
    writer.writerows(all_rows)

    return Response(
        content=buf.getvalue().encode("utf-8"),
        media_type="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
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
        # Retarget the swap to #add-entry-error so the table is left untouched
        return HTMLResponse(
            content=f'<div class="err-box" style="margin-top:0">{msg}</div>',
            headers={
                "HX-Retarget": "#add-entry-error",
                "HX-Reswap":   "innerHTML",
                "HX-Add-Error": "true",
            },
        )

    if not _sb_configured():
        return error("Supabase is not configured.")

    email_norm = email.lower().strip()
    is_domain = "@" not in email_norm
    if not email_norm:
        return error("Please enter an email address or domain.")
    if is_domain and "." not in email_norm:
        return error("Enter a valid domain (e.g. company.com) or full email address.")
    if not is_domain and email_norm.count("@") != 1:
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
    # Accept full emails (user@domain.com) AND bare domains (spam.com)
    emails = [e for e in emails if e and ("@" in e or "." in e)]
    if not emails:
        return error("No valid email addresses or domains found in the file.")

    reason_clean = reason.strip() or "bulk_import"

    for i in range(0, len(emails), CHUNK_SIZE):
        chunk = emails[i : i + CHUNK_SIZE]
        rows = [
            {"client_id": client_id, "email": e, "reason": reason_clean, "added_by": "dashboard_user"}
            for e in chunk
        ]
        try:
            r = http_req.post(
                f"{SUPABASE_URL}/rest/v1/dnc_entries",
                headers=_sb_headers("resolution=ignore-duplicates,return=minimal"),
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
