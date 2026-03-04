from __future__ import annotations

import io
import threading
import time
import uuid
from pathlib import Path
from typing import List, Optional

import pandas as pd
from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import Response

from app.deps import templates
from logic.city_classifier import load_city_whitelist, choose_location_output

router = APIRouter()

MAX_UPLOAD_BYTES = 20 * 1024 * 1024

_WHITELIST_PATH = Path(__file__).resolve().parents[2] / "data" / "city_whitelist.json"

_CITY_CANDIDATES = ["city", "company city", "town", "locality", "location city"]
_STATE_CANDIDATES = ["state", "state/region", "region", "province", "county", "state or region"]

# ── in-memory token store ──────────────────────────────────────────────────────

_store: dict[str, dict] = {}
_store_lock = threading.Lock()
_TOKEN_TTL = 3600


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


# ── helpers (ported from tools/city_app.py) ───────────────────────────────────

def _norm_header(h: str) -> str:
    return " ".join(str(h).strip().lower().replace("_", " ").replace("-", " ").split())


def detect_column(headers: List[str], candidates: List[str]) -> Optional[str]:
    norm_map = {h: _norm_header(h) for h in headers}
    cand_norms = [_norm_header(c) for c in candidates]

    for cand in cand_norms:
        for original, normed in norm_map.items():
            if normed == cand:
                return original

    for cand in cand_norms:
        toks = cand.split()
        for original, normed in norm_map.items():
            parts = normed.split()
            for i in range(len(parts) - len(toks) + 1):
                if parts[i : i + len(toks)] == toks:
                    return original

    return None


def read_input(contents: bytes, filename: str) -> tuple[pd.DataFrame, str]:
    name = filename.lower()
    if name.endswith(".csv"):
        return (
            pd.read_csv(io.BytesIO(contents), dtype=str, keep_default_na=False, na_values=[]),
            "csv",
        )
    if name.endswith(".xlsx"):
        return (
            pd.read_excel(
                io.BytesIO(contents), dtype=str, engine="openpyxl",
                keep_default_na=False, na_values=[],
            ),
            "xlsx",
        )
    raise ValueError("Unsupported file type. Please upload a CSV or XLSX file.")


def build_output(df: pd.DataFrame, city_col: str, state_col: str, whitelist_key_set) -> pd.DataFrame:
    out = df.copy(deep=True)
    city_series = out[city_col] if city_col in out.columns else pd.Series([""] * len(out), index=out.index, dtype="object")
    state_series = out[state_col] if state_col in out.columns else pd.Series([""] * len(out), index=out.index, dtype="object")
    out_city = [
        choose_location_output(c, s, whitelist_key_set)
        for c, s in zip(city_series.tolist(), state_series.tolist())
    ]
    if "city" in out.columns:
        out["city"] = out_city
    else:
        out.insert(len(out.columns), "city", out_city)
    return out


def write_output(df: pd.DataFrame, ext: str) -> tuple[bytes, str]:
    if ext == "csv":
        return df.to_csv(index=False).encode("utf-8"), "text/csv"
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return bio.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ── routes ─────────────────────────────────────────────────────────────────────

@router.get("/city-state")
async def city_page(request: Request):
    return templates.TemplateResponse("city.html", {"request": request, "active": "city"})


@router.post("/api/normalize")
async def api_normalize(
    request: Request,
    file: UploadFile = File(...),
    city_col: str = Form(""),
    state_col: str = Form(""),
):
    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "File too large. Maximum 20 MB.")

    def error(msg: str):
        return templates.TemplateResponse(
            "partials/city_result.html", {"request": request, "error": msg}
        )

    try:
        _, whitelist_key_set = load_city_whitelist(_WHITELIST_PATH)
    except Exception as e:
        return error(f"Could not load city whitelist: {e}")

    try:
        df, ext = read_input(contents, file.filename or "upload.csv")
    except Exception as e:
        return error(str(e))

    if df.empty:
        return error("The uploaded file contains no rows.")

    headers = list(df.columns)

    resolved_city = (city_col.strip() if city_col.strip() in headers else None) \
        or detect_column(headers, _CITY_CANDIDATES)
    resolved_state = (state_col.strip() if state_col.strip() in headers else None) \
        or detect_column(headers, _STATE_CANDIDATES)

    if not resolved_city or not resolved_state:
        missing = []
        if not resolved_city:
            missing.append("city")
        if not resolved_state:
            missing.append("state/region")
        return error(
            f"Could not detect column(s): {', '.join(missing)}. Please specify column names manually."
        )

    try:
        out = build_output(df, resolved_city, resolved_state, whitelist_key_set)
    except Exception as e:
        return error(f"Normalisation failed: {e}")

    file_bytes, mime = write_output(out, ext)
    base_name = Path(file.filename or "output").stem
    out_filename = f"{base_name}_normalised.{ext}"

    token = _store_result(file_bytes, mime, out_filename)
    return templates.TemplateResponse(
        "partials/city_result.html",
        {
            "request": request,
            "preview": out.head(20).fillna("").to_dict(orient="records"),
            "columns": out.columns.tolist(),
            "token": token,
            "filename": out_filename,
            "row_count": len(out),
        },
    )


@router.get("/api/normalize/download/{token}")
async def city_download(token: str):
    with _store_lock:
        entry = _store.pop(token, None)
    if entry is None:
        raise HTTPException(404, "Download link has expired or already been used.")
    return Response(
        content=entry["data"],
        media_type=entry["mime"],
        headers={"Content-Disposition": f'attachment; filename="{entry["filename"]}"'},
    )
