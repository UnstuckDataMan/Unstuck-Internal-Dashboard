"""Which copy variation won, per client, for both channels we run.

TWO SOURCES THAT DO NOT MEASURE THE SAME THING
Smartlead splits a sequence step into variants and computes their stats itself
(`/campaigns/{id}/sequence-analytics`), so a Smartlead winner is decided on
Smartlead's own `positive_reply_count` and reply rate.

A manual campaign draws its subject and body from the Copy Bank and records the
combination on the sheet as an "A/B Variant" ("S2/B1"), so a manual winner is
decided on the team's own Lead and Interested statuses.

Those are different measures, and this module keeps them apart rather than
adding them up: a combined "winner" across the two would be a number with no
meaning. Every row carries `basis` so the view can say which one it is.

DATES
Smartlead's endpoint takes a window, so its figures match whatever period is
being looked at. The campaign sheet has no per-row date, so a manual variant is
always all-time for that campaign. Views must say so — a reader comparing a
month of Smartlead against the lifetime of a manual campaign would otherwise
draw a false conclusion.
"""
from __future__ import annotations

import logging
from datetime import date

from app.utils.pulse import smartlead

logger = logging.getLogger(__name__)

# A variant with barely any volume will top the table on noise alone. Below
# this it is reported but never called a winner.
MIN_SENDS_FOR_A_WINNER = 40


def _rate(value: int, base: int) -> float | None:
    return round(value * 100.0 / base, 1) if base else None


def _decide(variants: list[dict]) -> dict:
    """Rank variants and name a winner, or say why there isn't one.

    Sorted by the positive count, then by replies. A tie or too little volume
    is reported as such: "we cannot tell yet" is a useful answer and a coin
    flip dressed as a result is not.
    """
    for row in variants:
        row["positive_rate"] = _rate(row["positive"], row["sent"])
        row["reply_rate"] = _rate(row["reply"], row["sent"])
    variants.sort(key=lambda v: (v["positive"], v["reply"], v["sent"]), reverse=True)

    total_sent = sum(v["sent"] for v in variants)
    out = {
        "variants":   variants,
        "total_sent": total_sent,
        "winner":     None,
        "reason":     "",
    }
    if len(variants) < 2:
        out["reason"] = "Only one variation has been sent, so there is nothing to compare."
        return out
    if total_sent < MIN_SENDS_FOR_A_WINNER:
        out["reason"] = (f"Too little volume to call — {total_sent} sent across "
                         f"{len(variants)} variations.")
        return out
    best = variants[0]
    if best["positive"] == 0:
        out["reason"] = "No positive responses yet, so no variation is ahead."
        return out
    joint = [v for v in variants
             if v["positive"] == best["positive"] and v["reply"] == best["reply"]]
    if len(joint) > 1:
        out["reason"] = f"{len(joint)} variations are level on positive responses."
        return out
    out["winner"] = best["variant"]
    return out


# ── Smartlead ─────────────────────────────────────────────────────────────────

def smartlead_variants(external_campaign_id: str, start: date, end: date) -> list[dict]:
    """Per-step variant stats for one Smartlead campaign, over a window.

    Smartlead computes this, so nothing here re-derives it from events — which
    would have meant a variant column on the events and outcomes tables and a
    rollup to aggregate it.
    """
    payload = smartlead.request_json(
        f"/campaigns/{external_campaign_id}/sequence-analytics",
        {
            "start_date": f"{start.isoformat()}T00:00:00.000Z",
            "end_date":   f"{end.isoformat()}T23:59:59.999Z",
            "time_zone":  "UTC",
        },
    )
    data = payload.get("data") if isinstance(payload, dict) else None
    sequences = (data or {}).get("sequences") if isinstance(data, dict) else None
    if not isinstance(sequences, list):
        return []

    steps = []
    for sequence in sequences:
        if not isinstance(sequence, dict):
            continue
        variants = []
        for variant in sequence.get("variants") or []:
            if not isinstance(variant, dict):
                continue
            stats = variant.get("stats") or {}
            sent = int(stats.get("sent_count") or 0)
            if not sent:
                continue          # a variant that never went out is not a test
            variants.append({
                "variant":     str(variant.get("variant_label") or variant.get("id") or "?"),
                "is_baseline": bool(variant.get("is_baseline")),
                "sent":        sent,
                "reply":       int(stats.get("reply_count") or 0),
                "positive":    int(stats.get("positive_reply_count") or 0),
                "unsubscribe": int(stats.get("unsubscribed_count") or 0),
            })
        if len(variants) < 1:
            continue
        step = _decide(variants)
        step.update({
            "step":    int(sequence.get("seq_number") or 0),
            "subject": str(sequence.get("subject_line") or "").strip(),
            "basis":   "Smartlead's own positive replies",
        })
        steps.append(step)
    steps.sort(key=lambda s: s["step"])
    return steps


# ── Manual ────────────────────────────────────────────────────────────────────

def manual_variants(sheet_ids: list[str]) -> dict:
    """Variant stats across a client's manual campaign sheets, combined.

    All-time, not windowed: the sheet records a prospect's current status with
    no date against it, so there is no honest way to bound this by a period.
    A sheet that cannot be read is skipped rather than failing the panel —
    one revoked share should not hide every other campaign's result.
    """
    from app.utils.google_sheets import read_ab_stats

    totals: dict[str, dict] = {}
    read, skipped = 0, 0
    for sheet_id in sheet_ids:
        if not sheet_id:
            continue
        try:
            rows = read_ab_stats(sheet_id)
        except Exception as exc:
            skipped += 1
            logger.warning("Pulse A/B: sheet %s unreadable: %s", sheet_id, exc)
            continue
        read += 1
        for row in rows:
            agg = totals.setdefault(row["variant"], {
                "variant": row["variant"], "is_baseline": False,
                "sent": 0, "reply": 0, "positive": 0, "unsubscribe": 0,
                "lead": 0, "interested": 0,
            })
            agg["sent"] += row.get("total", 0)
            agg["reply"] += row.get("reply", 0)
            agg["positive"] += row.get("positive", 0)
            agg["unsubscribe"] += row.get("unsubscribe", 0)
            agg["lead"] += row.get("lead", 0)
            agg["interested"] += row.get("interested", 0)

    result = _decide(list(totals.values()))
    result.update({
        "sheets_read":    read,
        "sheets_skipped": skipped,
        "basis":          "your Lead and Interested statuses",
    })
    return result
