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


def winners_for_report(steps: list[dict],
                       manual: dict | list[dict] | None = None) -> list[dict]:
    """The decided winners only, trimmed for a client-facing report.

    A client gets the result, not the method: which message worked and by how
    much. The full per-variant table stays internal — publishing every version
    we tested, including the ones that did badly, tells them more about our
    process than about their campaign.

    Steps with no decided winner are left out entirely rather than shown as
    "too early to tell". On a client's report that reads as an excuse; on the
    internal panel, where it is actionable, it is still said in full.
    """
    out: list[dict] = []
    for step in steps:
        if not step.get("winner"):
            continue
        best = next((v for v in step["variants"] if v["variant"] == step["winner"]), None)
        if best is None:
            continue
        copy = best.get("variant_copy") or {}
        out.append({
            "label":         f"Version {step['winner']}",
            "campaign":      step.get("campaign", ""),
            # Written-down copy wins over the step's subject: Smartlead reports
            # one subject for the whole step, so it is the same string for
            # every variant and cannot distinguish the winner.
            "subject":       copy.get("subject") or step.get("subject", ""),
            "body":          copy.get("body", ""),
            "reply_rate":    best.get("reply_rate"),
            "positive_rate": best.get("positive_rate"),
            "sent":          best.get("sent", 0),
            "basis":         step.get("basis", ""),
        })
    # A dict is one combined verdict, a list is one per campaign. Both are
    # accepted because reports published before the per-campaign breakdown
    # existed are replayed through here when they are re-published.
    blocks = [manual] if isinstance(manual, dict) else list(manual or [])
    for block in blocks:
        if not block or not block.get("winner"):
            continue
        best = next((v for v in block["variants"]
                     if v["variant"] == block["winner"]), None)
        if best is None:
            continue
        copy = best.get("variant_copy") or {}
        out.append({
            "label":         f"Version {block['winner']}",
            # Named per campaign now. "Manual campaigns" was accurate when
            # there was one combined verdict and would be a lie against a
            # figure that covers one campaign.
            "campaign":      block.get("name") or "Manual campaigns",
            "subject":       copy.get("subject", ""),
            "body":          copy.get("body", ""),
            "reply_rate":    best.get("reply_rate"),
            "positive_rate": best.get("positive_rate"),
            "sent":          best.get("sent", 0),
            "basis":         block.get("basis", ""),
        })
    return out


# ── Manual ────────────────────────────────────────────────────────────────────

_MANUAL_BASIS = "your Lead and Interested statuses"


def _blank_variant(key: str) -> dict:
    return {"variant": key, "is_baseline": False, "sent": 0, "reply": 0,
            "positive": 0, "unsubscribe": 0, "lead": 0, "interested": 0}


def _add_row(totals: dict[str, dict], row: dict) -> None:
    agg = totals.setdefault(row["variant"], _blank_variant(row["variant"]))
    agg["sent"] += row.get("total", 0)
    for field in ("reply", "positive", "unsubscribe", "lead", "interested"):
        agg[field] += row.get(field, 0)


def manual_breakdown(campaigns: list[dict]) -> dict:
    """Variant stats per manual campaign, and across all of them.

    `campaigns` carries whatever identifies a campaign to the rest of the app:
    `sheet_id` to read, `ref` to key its written-down copy against, `name` to
    print, and the Copy Bank `territory`/`industry` the merge recorded.

    PER CAMPAIGN IS THE REAL TEST. A campaign is one send of one Copy Bank
    entry, so "S1/B1" means one specific message inside it and the winner is a
    verdict about that message. Across campaigns the same label can be two
    different messages, which is why the combined figure below carries
    `shared_copy` rather than being presented as equivalent.

    COMBINED IS STILL WORTH HAVING, because a single campaign often has too
    little volume to call and several campaigns of the same copy do not. It is
    computed from the same reads, so it costs no extra requests.

    All-time, not windowed: the sheet records a prospect's current status with
    no date against it, so there is no honest way to bound this by a period.
    A sheet that cannot be read is skipped rather than failing the panel —
    one revoked share should not hide every other campaign's result.
    """
    from app.utils.google_sheets import read_ab_stats

    blocks: list[dict] = []
    combined: dict[str, dict] = {}
    sources: list[tuple[str, str]] = []
    read, skipped = 0, 0

    for campaign in campaigns:
        sheet_id = str(campaign.get("sheet_id") or "")
        if not sheet_id:
            continue
        try:
            rows = read_ab_stats(sheet_id)
        except Exception as exc:
            skipped += 1
            logger.warning("Pulse A/B: sheet %s unreadable: %s", sheet_id, exc)
            continue
        read += 1

        totals: dict[str, dict] = {}
        for row in rows:
            _add_row(totals, row)
            _add_row(combined, row)
        if not totals:
            continue        # a sheet with no variant column is not a test

        territory = str(campaign.get("territory") or "")
        industry = str(campaign.get("industry") or "")
        if territory and industry and (territory, industry) not in sources:
            sources.append((territory, industry))

        block = _decide(list(totals.values()))
        block.update({
            "ref":       str(campaign.get("ref") or sheet_id),
            "name":      str(campaign.get("name") or "").strip() or "Untitled campaign",
            "territory": territory,
            "industry":  industry,
            "basis":     _MANUAL_BASIS,
        })
        blocks.append(block)

    # Biggest first: the campaign with the most behind it is the one a reader
    # should weigh most, and it is the one that most often has a verdict.
    blocks.sort(key=lambda b: b["total_sent"], reverse=True)

    result = _decide(list(combined.values()))
    result.update({
        "basis": _MANUAL_BASIS,
        # One Copy Bank entry behind every campaign means a label means the
        # same message throughout, and only then can the combined figures be
        # traced back to copy.
        "shared_copy": len(sources) == 1,
        "copy_sources": sources,
    })
    return {
        "campaigns":      blocks,
        "combined":       result,
        "sheets_read":    read,
        "sheets_skipped": skipped,
        "basis":          _MANUAL_BASIS,
    }


def manual_variants(sheet_ids: list[str]) -> dict:
    """The combined verdict alone, for callers with only sheet ids to hand."""
    out = manual_breakdown([{"sheet_id": s} for s in sheet_ids])
    return {**out["combined"],
            "sheets_read":    out["sheets_read"],
            "sheets_skipped": out["sheets_skipped"]}
