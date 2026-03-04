from __future__ import annotations

import io
import threading
import time
import uuid
from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import Response

from app.deps import templates
from logic.gender_classifier import classify

router = APIRouter()

MAX_UPLOAD_BYTES = 20 * 1024 * 1024

_OVERRIDES_PATH = Path(__file__).resolve().parents[2] / "overrides.csv"

_NAME_CANDIDATES = [
    "first_name", "firstname", "first", "given", "name", "first name", "given name",
]

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


# ── helpers ────────────────────────────────────────────────────────────────────

def _detect_name_column(headers: list[str]) -> Optional[str]:
    """Deterministic first-name column detection (exact then contains match)."""
    def norm(h: str) -> str:
        return " ".join(str(h).strip().lower().replace("_", " ").replace("-", " ").split())

    norm_map = {h: norm(h) for h in headers}
    cand_norms = [norm(c) for c in _NAME_CANDIDATES]

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

    return headers[0] if headers else None


def _read_input(contents: bytes, filename: str) -> tuple[pd.DataFrame, str]:
    name = filename.lower()
    if name.endswith((".xls", ".xlsx")):
        return pd.read_excel(io.BytesIO(contents), engine="openpyxl"), "xlsx"
    return pd.read_csv(io.BytesIO(contents)), "csv"


def _apply_overrides(df: pd.DataFrame, name_col: str, new_col: str) -> pd.DataFrame:
    if not _OVERRIDES_PATH.exists():
        return df
    o = pd.read_csv(_OVERRIDES_PATH)
    if not {"name", new_col}.issubset(o.columns):
        return df
    key = (
        df[name_col].astype(str).str.split().str[0]
        .str.replace(r"[^A-Za-z\-]", "", regex=True)
        .str.lower()
    )
    df = df.copy()
    df["_ov_key"] = key
    o["_ov_key"] = o["name"].astype(str).str.lower()
    ov_map = dict(zip(o["_ov_key"], o[new_col]))
    df["_ov_val"] = df["_ov_key"].map(ov_map)
    df[new_col] = df["_ov_val"].fillna(df[new_col])
    df.drop(columns=["_ov_key", "_ov_val"], inplace=True)
    return df


# ── routes ─────────────────────────────────────────────────────────────────────

@router.get("/gender")
async def gender_page(request: Request):
    return templates.TemplateResponse("gender.html", {"request": request, "active": "gender"})


@router.post("/api/gender")
async def api_gender(
    request: Request,
    file: UploadFile = File(...),
    name_col: str = Form(""),
    new_col: str = Form("gender_mf"),
):
    contents = await file.read()
    if len(contents) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "File too large. Maximum 20 MB.")

    def error(msg: str):
        return templates.TemplateResponse(
            "partials/gender_result.html", {"request": request, "error": msg}
        )

    try:
        df, ext = _read_input(contents, file.filename or "upload.csv")
    except Exception as e:
        return error(f"Could not read file: {e}")

    if df.empty:
        return error("The uploaded file contains no rows.")

    col = (name_col.strip() if name_col.strip() in df.columns else None) or _detect_name_column(list(df.columns))
    if not col:
        return error("Could not detect a first-name column. Please specify one.")

    try:
        df[new_col] = [classify(v) for v in df[col]]
    except Exception as e:
        return error(str(e))

    df = _apply_overrides(df, col, new_col)

    if ext == "xlsx":
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="data")
        data, mime, out_filename = buf.getvalue(), \
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "classified.xlsx"
    else:
        data, mime, out_filename = (
            df.to_csv(index=False).encode("utf-8"), "text/csv", "classified.csv"
        )

    token = _store_result(data, mime, out_filename)
    return templates.TemplateResponse(
        "partials/gender_result.html",
        {
            "request": request,
            "preview": df.head(20).fillna("").to_dict(orient="records"),
            "columns": df.columns.tolist(),
            "token": token,
            "filename": out_filename,
            "row_count": len(df),
        },
    )


@router.get("/api/gender/download/{token}")
async def gender_download(token: str):
    with _store_lock:
        entry = _store.pop(token, None)
    if entry is None:
        raise HTTPException(404, "Download link has expired or already been used.")
    return Response(
        content=entry["data"],
        media_type=entry["mime"],
        headers={"Content-Disposition": f'attachment; filename="{entry["filename"]}"'},
    )
