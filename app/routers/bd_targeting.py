from __future__ import annotations

import os
import re
from collections import defaultdict
from datetime import datetime, timezone, timedelta
from typing import Optional

import requests as http_req
from fastapi import APIRouter, Query, Request

from app.deps import templates

router = APIRouter()

SUPABASE_URL      = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")


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


# ── Campaign name parser ───────────────────────────────────────────────────────

KNOWN_TERRITORIES = {
    "UK", "US", "AU", "NZ", "CA", "IE", "ZA", "SG",
    "UAE", "SA", "IN", "DE", "FR", "EU",
}
TERRITORY_ALIASES = {"AUS": "AU", "USA": "US"}
_HEADCOUNT_RE = re.compile(r"^\d+\s*[-–—]\s*\d+$")


def _parse_campaign_name(name: str) -> dict:
    """
    Parse a campaign name like:
      'Chris (Devlin Sender) - PR - CSuite-Marketing - 8-200 - UK'
    into structured fields: sdr, sender, territory, headcount, industry.
    """
    result: dict = {
        "sdr": None, "sender": None, "territory": None,
        "headcount": None, "industry": None, "parse_ok": False,
    }
    if not name:
        return result

    paren_open = name.find("(")
    if paren_open == -1:
        return result

    sdr_raw = name[:paren_open].strip().rstrip("-").strip()
    result["sdr"] = sdr_raw or None

    paren_close = name.find(")", paren_open)
    if paren_close == -1:
        return result

    inner = name[paren_open + 1: paren_close].strip()
    m = re.match(r"^(.*?)\s+[Ss]enders?$", inner)
    if m:
        result["sender"] = m.group(1).strip() or None

    payload = name[paren_close + 1:].lstrip().lstrip("-–—").lstrip()
    segments = [s.strip() for s in re.split(r"\s*[-–—]\s*", payload) if s.strip()]

    industry_parts: list[str] = []
    for seg in segments:
        upper = seg.upper().strip()
        normalised = TERRITORY_ALIASES.get(upper, upper)
        if normalised in KNOWN_TERRITORIES:
            if result["territory"] is None:
                result["territory"] = normalised
        elif _HEADCOUNT_RE.match(seg):
            if result["headcount"] is None:
                result["headcount"] = seg
        else:
            industry_parts.append(seg)

    result["industry"]  = " / ".join(industry_parts) if industry_parts else None
    result["parse_ok"]  = bool(result["territory"] or industry_parts)
    return result


# ── Time helpers ───────────────────────────────────────────────────────────────

def _recency_class(iso: Optional[str]) -> str:
    if not iso:
        return "red"
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        weeks = (datetime.now(timezone.utc) - dt).days / 7
        if weeks < 4:
            return "green"
        elif weeks < 12:
            return "amber"
        else:
            return "red"
    except Exception:
        return "red"


def _human_recency(iso: Optional[str]) -> str:
    if not iso:
        return "Never"
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        days = (datetime.now(timezone.utc) - dt).days
        if days == 0:
            return "Today"
        elif days == 1:
            return "Yesterday"
        elif days < 7:
            return f"{days}d ago"
        elif days < 30:
            weeks = days // 7
            return f"{weeks} week{'s' if weeks != 1 else ''} ago"
        else:
            months = days // 30
            return f"{months} month{'s' if months != 1 else ''} ago"
    except Exception:
        return "Unknown"


def _lead_rate(lead_count, sent_count) -> Optional[float]:
    try:
        sc = int(sent_count or 0)
        if sc > 0:
            return int(lead_count or 0) / sc
    except Exception:
        pass
    return None


def _best_date(row: dict) -> Optional[str]:
    """Prefer completed_at, fall back to created_at."""
    return row.get("completed_at") or row.get("created_at") or None


# ── String normalisation for fuzzy matching ────────────────────────────────────

def _norm(s: str) -> str:
    """Lower-case, collapse whitespace/underscores/dashes/slashes to a single space."""
    return re.sub(r"[\s_\-/]+", " ", (s or "")).strip().lower()


# ── Copy Bank profile fetch ────────────────────────────────────────────────────

def _fetch_cb_profiles(
    client_id: str,
) -> tuple[list[str], list[str], dict[str, str]]:
    """
    Returns (territories, industries, industry_labels) for the given client_id.
    Industries are the raw profile keys; industry_labels maps key → display name.
    Falls back to empty lists on any error.
    """
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/copy_bank_templates",
            headers=_sb_headers(),
            params={"key": "eq.__cb_profiles__", "select": "content"},
            timeout=10,
        )
        if not r.ok:
            return [], [], {}
        data = r.json()
        if not data or not isinstance(data[0].get("content"), list):
            return [], [], {}

        profiles = [
            p for p in data[0]["content"]
            if isinstance(p, dict) and p.get("client_id") == client_id
        ]

        territories: set[str] = set()
        industries:  set[str] = set()
        labels:      dict[str, str] = {}

        for p in profiles:
            for t in (p.get("territories") or []):
                if t:
                    territories.add(str(t).upper())
            for ind in (p.get("industries") or []):
                if ind:
                    industries.add(str(ind))
            for k, v in (p.get("industryLabels") or {}).items():
                labels[k] = v

        sorted_t = sorted(territories)
        sorted_i = sorted(industries, key=lambda x: labels.get(x, x).lower())
        return sorted_t, sorted_i, labels

    except Exception:
        return [], [], {}


def _industry_display(key: str, labels: dict) -> str:
    return labels.get(key) or key.replace("_", " ")


# ── Routes ─────────────────────────────────────────────────────────────────────

@router.get("/bd-targeting")
async def bd_targeting_page(request: Request):
    return templates.TemplateResponse(
        "bd_targeting.html",
        {"request": request, "active": "bd_targeting"},
    )


@router.get("/api/bd-targeting/clients")
async def bd_targeting_clients(request: Request):
    if not _sb_configured():
        return templates.TemplateResponse(
            "partials/bd_targeting_clients_options.html",
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
        "partials/bd_targeting_clients_options.html",
        {"request": request, "clients": clients},
    )


@router.get("/api/bd-targeting/data")
async def bd_targeting_data(
    request:   Request,
    client_id: str = Query(""),
):
    _EMPTY_CTX = {
        "request":              request,
        "no_client":            False,
        "error":                None,
        "sdr_cards":            [],
        "sdr_tabs":             [],
        "map_rows":             [],
        "map_industry_headers": [],
        "unparseable":          [],
    }

    def _error(msg: str):
        return templates.TemplateResponse(
            "partials/bd_targeting_data.html",
            {**_EMPTY_CTX, "error": msg},
        )

    if not client_id:
        return templates.TemplateResponse(
            "partials/bd_targeting_data.html",
            {**_EMPTY_CTX, "no_client": True},
        )

    if not _sb_configured():
        return _error("Supabase is not configured.")

    # ── 1. Fetch campaigns ─────────────────────────────────────────
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={
                "select":    "id,created_at,campaign_name,sender_profile_name,"
                             "total_prospects,sent_count,"
                             "lead_count,reply_count,interested_count,unsubscribe_count,"
                             "completed,completed_at,paused",
                "client_id": f"eq.{client_id}",
                "order":     "created_at.desc",
            },
            timeout=15,
        )
        r.raise_for_status()
        raw = r.json()
    except Exception as exc:
        return _error(f"Could not load campaigns: {exc}")

    # ── 2. Parse + annotate ────────────────────────────────────────
    campaigns: list[dict] = []
    for c in raw:
        parsed = _parse_campaign_name(c.get("campaign_name") or "")
        row = {**c, **parsed}
        row["lead_rate"] = _lead_rate(c.get("lead_count"), c.get("sent_count"))
        campaigns.append(row)

    # ── 3. Partition ───────────────────────────────────────────────
    active_camps = [c for c in campaigns if not c.get("completed") and not c.get("paused")]
    paused_camps = [c for c in campaigns if c.get("paused")]
    completed_camps = [
        c for c in campaigns
        if c.get("completed") or (
            int(c.get("sent_count") or 0) >= int(c.get("total_prospects") or 1) > 0
        )
    ]
    live_and_paused = active_camps + paused_camps

    # ── 4. SDR cards (Team tab) ────────────────────────────────────
    sdr_card_map: dict[str, list] = defaultdict(list)
    for c in live_and_paused:
        sdr_card_map[c.get("sdr") or "Unassigned"].append(c)

    sdr_cards = [
        {
            "sdr":  sdr,
            "rows": sorted(rows, key=lambda x: (x.get("territory") or "", x.get("industry") or "")),
        }
        for sdr, rows in sorted(sdr_card_map.items())
    ]

    # ── 5. Individual SDR tabs ─────────────────────────────────────
    all_sdrs = sorted({c.get("sdr") for c in campaigns if c.get("sdr")})
    sdr_tabs: list[dict] = []

    for sdr in all_sdrs:
        sdr_camps     = [c for c in campaigns if c.get("sdr") == sdr]
        sdr_live      = [c for c in sdr_camps if c in live_and_paused]
        sdr_completed = [c for c in sdr_camps if c in completed_camps]

        # Build per-SDR history groups
        grp_map: dict[tuple, dict] = {}
        for c in sdr_completed:
            key = (c.get("territory") or "Unknown", c.get("industry") or "Uncategorised")
            g = grp_map.setdefault(key, {
                "territory":      key[0],
                "industry":       key[1],
                "campaign_count": 0,
                "total_sent":     0,
                "total_leads":    0,
                "_dates":         [],
                "headcounts":     set(),
            })
            g["campaign_count"] += 1
            g["total_sent"]     += int(c.get("sent_count") or 0)
            g["total_leads"]    += int(c.get("lead_count") or 0)
            d = _best_date(c)
            if d:
                g["_dates"].append(d)
            if c.get("headcount"):
                g["headcounts"].add(c["headcount"])

        history_groups: list[dict] = []
        for g in sorted(
            grp_map.values(),
            key=lambda x: max(x["_dates"]) if x["_dates"] else "",
            reverse=True,
        ):
            last     = max(g["_dates"]) if g["_dates"] else None
            agg_rate = g["total_leads"] / g["total_sent"] if g["total_sent"] > 0 else None
            history_groups.append({
                "territory":           g["territory"],
                "industry":            g["industry"],
                "headcounts":          ", ".join(sorted(g["headcounts"])) or "—",
                "last_targeted":       last,
                "human_recency":       _human_recency(last),
                "recency_class":       _recency_class(last),
                "aggregate_lead_rate": agg_rate,
                "campaign_count":      g["campaign_count"],
            })

        sdr_tabs.append({
            "sdr":            sdr,
            "tab_id":         "bdt-tab-" + re.sub(r"[^a-z0-9]", "-", sdr.lower()),
            "current":        sdr_live,
            "history_groups": history_groups,
        })

    # ── 6. Targeting map ───────────────────────────────────────────
    map_territories, map_industries, ind_labels = _fetch_cb_profiles(client_id)

    # Fallback: derive universe from campaign data if no CB profiles found
    if not map_territories:
        map_territories = sorted({c.get("territory") for c in campaigns if c.get("territory")})
    if not map_industries:
        map_industries = sorted({c.get("industry") for c in campaigns if c.get("industry")})

    today      = datetime.now(timezone.utc)
    cutoff_14  = today - timedelta(days=14)

    map_rows: list[dict] = []
    for ter in map_territories:
        cells: list[dict] = []
        for ind_key in map_industries:
            ind_disp = _industry_display(ind_key, ind_labels)
            n_ter    = _norm(ter)
            n_ind    = _norm(ind_disp)

            # Active: live/paused campaign whose territory+industry match
            active_sdrs = [
                c.get("sdr") or "?"
                for c in live_and_paused
                if _norm(c.get("territory") or "") == n_ter
                and _norm(c.get("industry")  or "") == n_ind
            ]

            # Recent: completed within 14 days
            recent_info: list[dict] = []
            if not active_sdrs:
                for c in completed_camps:
                    if (
                        _norm(c.get("territory") or "") != n_ter
                        or _norm(c.get("industry")  or "") != n_ind
                    ):
                        continue
                    d = _best_date(c)
                    if not d:
                        continue
                    try:
                        dt = datetime.fromisoformat(d.replace("Z", "+00:00"))
                        if dt.tzinfo is None:
                            dt = dt.replace(tzinfo=timezone.utc)
                        if dt >= cutoff_14:
                            recent_info.append({
                                "sdr":      c.get("sdr") or "?",
                                "days_ago": (today - dt).days,
                            })
                    except Exception:
                        pass

            if active_sdrs:
                status = "active"
                label  = "LIVE"
                sub    = ", ".join(sorted(set(active_sdrs)))
            elif recent_info:
                min_d  = min(x["days_ago"] for x in recent_info)
                status = "recent"
                label  = f"{min_d}d ago"
                sub    = ", ".join(sorted({x["sdr"] for x in recent_info}))
            else:
                status = "available"
                label  = ""
                sub    = ""

            cells.append({
                "industry_key":     ind_key,
                "industry_display": ind_disp,
                "status":           status,
                "label":            label,
                "sub":              sub,
            })

        map_rows.append({"territory": ter, "cells": cells})

    map_industry_headers = [
        {"key": k, "display": _industry_display(k, ind_labels)}
        for k in map_industries
    ]

    unparseable = [c for c in campaigns if not c.get("parse_ok")]

    return templates.TemplateResponse(
        "partials/bd_targeting_data.html",
        {
            "request":              request,
            "no_client":            False,
            "error":                None,
            "sdr_cards":            sdr_cards,
            "sdr_tabs":             sdr_tabs,
            "map_rows":             map_rows,
            "map_industry_headers": map_industry_headers,
            "unparseable":          unparseable,
        },
    )
