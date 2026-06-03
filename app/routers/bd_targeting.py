from __future__ import annotations

import os
import re
from collections import defaultdict
from datetime import datetime, timezone, timedelta
from typing import Optional

import requests as http_req
from fastapi import APIRouter, Request

from app.deps import templates

router = APIRouter()

SUPABASE_URL      = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_ANON_KEY = os.environ.get("SUPABASE_ANON_KEY", "")

# ── Team members ───────────────────────────────────────────────────────────────
# Maps lowercase tag → display name (add/remove as team grows)
TEAM_MEMBERS: dict[str, str] = {
    "chris":  "Chris",
    "robyn":  "Robyn",
    "glen":   "Glen",
    "leo":    "Leo",
    "devlin": "Devlin",
}
# Display order in Tab 1 and toggles
TEAM_ORDER = ["Chris", "Robyn", "Glen", "Leo", "Devlin"]

# Colour hex per person (used in heatmap chips and history table)
TEAM_COLOURS: dict[str, str] = {
    "Chris":  "#3b82f6",   # blue
    "Robyn":  "#f472b6",   # pink
    "Glen":   "#22c55e",   # green
    "Leo":    "#a78bfa",   # purple
    "Devlin": "#fb923c",   # orange
}

# ── Territory constants ────────────────────────────────────────────────────────
KNOWN_TERRITORIES = {
    "UK", "US", "AU", "NZ", "CA", "IE", "ZA", "SG",
    "UAE", "SA", "IN", "DE", "FR", "EU",
}
TERRITORY_ALIASES = {"AUS": "AU", "USA": "US"}


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


# ── SDR from tags ──────────────────────────────────────────────────────────────

def _sdr_from_tags(tags) -> Optional[str]:
    """Find the first tag that matches a known team member."""
    if not tags:
        return None
    for tag in (tags if isinstance(tags, list) else []):
        key = str(tag).strip().lower()
        if key in TEAM_MEMBERS:
            return TEAM_MEMBERS[key]
    return None


# ── Campaign name parser ───────────────────────────────────────────────────────
# Handles both:
#   "Leo – Accounting – Founder & Marketing Roles – 4 – 200 – UK"  (en-dash, space-padded)
#   "Chris (Devlin Sender) - PR - CSuite-Marketing - 8-200 - UK"   (legacy hyphen format)
#
# Strategy: split on dash/en-dash/em-dash that is SURROUNDED by spaces.
# This preserves "CSuite-Marketing" (no surrounding spaces) and "8-200" as single tokens,
# while correctly splitting on " – " and " - ".
_SPLIT_RE      = re.compile(r" [–—\-] ")
_LONE_INT_RE   = re.compile(r"^\d+$")
_RANGE_RE      = re.compile(r"^\d+\s*[-–—]\s*\d+$")


def _parse_campaign_name(name: str) -> dict:
    """
    Returns {territory, headcount, industry, parse_ok}.
    SDR is resolved separately from tags — not from the name.
    """
    result: dict = {
        "territory": None, "headcount": None,
        "industry": None, "parse_ok": False,
    }
    if not name:
        return result

    segments = [s.strip() for s in _SPLIT_RE.split(name) if s.strip()]

    numbers:        list[str] = []
    industry_parts: list[str] = []

    for seg in segments:
        upper = seg.upper()
        normalised = TERRITORY_ALIASES.get(upper, upper)

        if normalised in KNOWN_TERRITORIES:
            if result["territory"] is None:
                result["territory"] = normalised

        elif _RANGE_RE.match(seg):
            # Already a combined range like "8-200" or "50–200"
            if result["headcount"] is None:
                result["headcount"] = re.sub(r"\s*[–—]\s*", "-", seg)

        elif _LONE_INT_RE.match(seg):
            numbers.append(seg)

        else:
            # Skip segments that are just a team member's name
            if seg.lower() not in TEAM_MEMBERS:
                industry_parts.append(seg)

    # Combine lone integers into headcount (e.g. "4" + "200" → "4-200")
    if result["headcount"] is None:
        if len(numbers) >= 2:
            result["headcount"] = f"{numbers[0]}-{numbers[1]}"
        elif len(numbers) == 1:
            result["headcount"] = numbers[0]

    result["industry"]  = " · ".join(industry_parts) if industry_parts else None
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
            w = days // 7
            return f"{w} week{'s' if w != 1 else ''} ago"
        else:
            mo = days // 30
            return f"{mo} month{'s' if mo != 1 else ''} ago"
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
    return row.get("completed_at") or row.get("created_at") or None


def _norm(s: str) -> str:
    return re.sub(r"[\s_\-/·]+", " ", (s or "")).strip().lower()


# ── Copy Bank profile fetch ────────────────────────────────────────────────────

def _fetch_cb_profiles_for_unstuck() -> tuple[list[str], list[str], dict[str, str]]:
    """
    Pull the BizDev profile territories and industries from Copy Bank profiles.
    Returns (territories, industries, industry_labels).
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

        # Use all bizdev profiles (type == 'bizdev')
        profiles = [
            p for p in data[0]["content"]
            if isinstance(p, dict) and p.get("type") == "bizdev"
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


@router.get("/api/bd-targeting/data")
async def bd_targeting_data(request: Request):
    def _error(msg: str):
        return templates.TemplateResponse(
            "partials/bd_targeting_data.html",
            {
                "request":              request,
                "error":                msg,
                "team_rows":            [],
                "team_order":           TEAM_ORDER,
                "team_colours":         TEAM_COLOURS,
                "map_rows":             [],
                "map_industry_headers": [],
                "sdr_tabs":             [],
                "unparseable":          [],
            },
        )

    if not _sb_configured():
        return _error("Supabase is not configured.")

    # ── 1. Fetch all Unstuck Agency campaigns ──────────────────────
    try:
        r = http_req.get(
            f"{SUPABASE_URL}/rest/v1/campaigns",
            headers=_sb_headers(),
            params={
                "select":      "id,created_at,campaign_name,tags,"
                               "total_prospects,sent_count,"
                               "lead_count,reply_count,interested_count,unsubscribe_count,"
                               "completed,completed_at,paused",
                "client_name": "ilike.*Unstuck*",
                "order":       "created_at.desc",
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
        sdr    = _sdr_from_tags(c.get("tags"))
        row = {**c, **parsed, "sdr": sdr}
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
    live = active_camps + paused_camps

    # ── 4. Tab 1 — Team table ──────────────────────────────────────
    # For each person in TEAM_ORDER, list their live campaigns.
    # Include any person found in campaign tags who isn't in TEAM_ORDER.
    extra_sdrs = sorted({
        c["sdr"] for c in live if c["sdr"] and c["sdr"] not in TEAM_ORDER
    })
    ordered_sdrs = [s for s in TEAM_ORDER if s in TEAM_MEMBERS.values()] + extra_sdrs

    team_rows: list[dict] = []
    for person in ordered_sdrs:
        person_live = sorted(
            [c for c in live if c.get("sdr") == person],
            key=lambda x: (x.get("territory") or "", x.get("industry") or ""),
        )
        team_rows.append({"sdr": person, "campaigns": person_live})

    # ── 5. Tab 2 — Heatmap ────────────────────────────────────────
    map_territories, map_industries, ind_labels = _fetch_cb_profiles_for_unstuck()

    # Fallback: derive from campaign data if no CB profiles
    if not map_territories:
        map_territories = sorted({c.get("territory") for c in campaigns if c.get("territory")})
    if not map_industries:
        map_industries  = sorted({c.get("industry") for c in campaigns if c.get("industry")})

    today     = datetime.now(timezone.utc)
    cutoff_14 = today - timedelta(days=14)

    map_rows: list[dict] = []
    for ter in map_territories:
        cells: list[dict] = []
        for ind_key in map_industries:
            ind_disp = _industry_display(ind_key, ind_labels)
            n_ter    = _norm(ter)
            n_ind    = _norm(ind_disp)

            # Active persons
            active_persons = sorted({
                c.get("sdr") or "?"
                for c in live
                if _norm(c.get("territory") or "") == n_ter
                and _norm(c.get("industry")  or "") == n_ind
            })

            # Recent persons (completed within 14 days)
            recent_persons: list[str] = []
            if not active_persons:
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
                            p = c.get("sdr") or "?"
                            if p not in recent_persons:
                                recent_persons.append(p)
                    except Exception:
                        pass

            status = "active" if active_persons else ("recent" if recent_persons else "available")
            cells.append({
                "industry_key":     ind_key,
                "industry_display": ind_disp,
                "status":           status,
                "active_persons":   active_persons,
                "recent_persons":   sorted(recent_persons),
            })

        map_rows.append({"territory": ter, "cells": cells})

    map_industry_headers = [
        {"key": k, "display": _industry_display(k, ind_labels)}
        for k in map_industries
    ]

    # ── 6. Tab 3 — Individual breakdowns ──────────────────────────
    sdr_tabs: list[dict] = []
    all_sdrs = [s for s in TEAM_ORDER if s in TEAM_MEMBERS.values()] + sorted({
        c["sdr"] for c in campaigns if c["sdr"] and c["sdr"] not in TEAM_ORDER
    })

    for sdr in all_sdrs:
        sdr_camps     = [c for c in campaigns if c.get("sdr") == sdr]
        sdr_live      = [c for c in sdr_camps if c in live]
        sdr_completed = [c for c in sdr_camps if c in completed_camps]

        # History groups
        grp_map: dict[tuple, dict] = {}
        for c in sdr_completed:
            key = (c.get("territory") or "Unknown", c.get("industry") or "Uncategorised")
            g = grp_map.setdefault(key, {
                "territory": key[0], "industry": key[1],
                "campaign_count": 0, "total_sent": 0,
                "total_leads": 0, "_dates": [], "headcounts": set(),
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
                "human_recency":       _human_recency(last),
                "recency_class":       _recency_class(last),
                "aggregate_lead_rate": agg_rate,
                "campaign_count":      g["campaign_count"],
            })

        sdr_tabs.append({
            "sdr":            sdr,
            "tab_id":         re.sub(r"[^a-z0-9]", "-", sdr.lower()),
            "colour":         TEAM_COLOURS.get(sdr, "#8b5cf6"),
            "current":        sdr_live,
            "history_groups": history_groups,
            "any_data":       bool(sdr_live or history_groups),
        })

    unparseable = [c for c in campaigns if not c.get("parse_ok")]

    return templates.TemplateResponse(
        "partials/bd_targeting_data.html",
        {
            "request":              request,
            "error":                None,
            "team_rows":            team_rows,
            "team_order":           TEAM_ORDER,
            "team_colours":         TEAM_COLOURS,
            "map_rows":             map_rows,
            "map_industry_headers": map_industry_headers,
            "sdr_tabs":             sdr_tabs,
            "unparseable":          unparseable,
        },
    )
