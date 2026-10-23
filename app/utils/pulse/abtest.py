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

# The total above is a sum ACROSS variants, so the volume it implies per arm
# shrinks as arms are added — three arms clear 40 on 14 sends each, which is
# nothing. Ranking on a rate makes that actively dangerous (1 reply from 3 is
# 33%), so the leader and the runner-up must each clear a floor of their own.
# Only those two: a third arm somebody sent to a handful of people must not
# block a verdict between two arms with hundreds.
MIN_SENDS_PER_VARIANT = 25

# And the leader needs enough of whatever decided it. Two positives is a
# coincidence.
MIN_WINS_FOR_A_WINNER = 3

# How far ahead the leader has to be, as a share of the runner-up's rate.
# Without this, 10 positives against 9 reads as a decisive result.
MIN_RELATIVE_LEAD = 0.20


def variant_sort_key(label: str) -> tuple[int, int, int, str]:
    """Order a variant label the way the copy behind it is ordered.

    "S2/B1" sorts by its Copy Bank indexes, so S1/B1 < S1/B2 < S2/B1 regardless
    of how the strings compare. A Smartlead letter has no indexes and sorts
    alphabetically after them. Anything else sorts last, on its raw text.
    """
    from app.utils.copy_bank import parse_variant

    subject_idx, body_idx = parse_variant(label)
    text = str(label or "")
    if subject_idx is not None and body_idx is not None:
        return (0, subject_idx, body_idx, text)
    if text.strip():
        return (1, 0, 0, text.strip().upper())
    return (2, 0, 0, text)


def number_variants(labels) -> dict[str, int]:
    """{label: 1-based number}, in variant_sort_key order.

    "S2/B1" tells a reader which Copy Bank row to open and nothing about which
    message it is; "Copy 2" is what the team actually compares. The number is
    DERIVED from the label rather than stored, so it needs no table and cannot
    drift — and it is derived from the label rather than from the ranking, so
    the copy that was Copy 2 yesterday is still Copy 2 when it starts winning.

    Dense and 1-based over the labels present. Scoped by the caller to one
    comparison unit — a Smartlead step, or one manual campaign — because
    those are the only sets whose members are ever measured against each other.
    """
    unique = sorted(set(str(label) for label in labels), key=variant_sort_key)
    return {label: i for i, label in enumerate(unique, start=1)}


def _rate(value: int, base: int) -> float | None:
    return round(value * 100.0 / base, 1) if base else None


def _decide(variants: list[dict]) -> dict:
    """Rank variants and name a winner, or say why there isn't one.

    RANKED ON RATES, NOT COUNTS. The old rule sorted on the raw positive count,
    which hands the verdict to whichever variation happened to be sent to more
    people — the one thing an A/B test is supposed to control for. Counts
    remain as tiebreakers, so equal rates still order by volume.

    DECIDED ON POSITIVES WHERE THERE ARE ANY, OTHERWISE ON REPLIES. A manual
    campaign is graded by hand, and an account manager marks Reply long before
    anyone goes back and marks Lead or Interested. The old rule bailed out
    whenever every positive count was zero, so those campaigns reported "no
    variation is ahead" indefinitely while sitting on a perfectly clear
    difference in replies. `decided_on` says which measure was used.

    A tie, too little volume, or a lead inside the noise is reported as such:
    "we cannot tell yet" is a useful answer and a coin flip dressed as a result
    is not.
    """
    numbers = number_variants(v["variant"] for v in variants)
    for row in variants:
        row["positive_rate"] = _rate(row["positive"], row["sent"])
        row["reply_rate"] = _rate(row["reply"], row["sent"])
        row["copy_no"] = numbers[str(row["variant"])]
        row["copy_label"] = f"Copy {row['copy_no']}"

    total_sent = sum(v["sent"] for v in variants)
    # Whichever measure decides, every gate below reads the same two fields.
    graded = any(v["positive"] for v in variants)
    measure = "positive" if graded else "reply"
    variants.sort(reverse=True, key=lambda v: (
        v[f"{measure}_rate"] or 0.0, v[measure], v["sent"]))

    out = {
        "variants":     variants,
        "total_sent":   total_sent,
        "winner":       None,
        "winner_label": "",
        "decided_on":   measure,
        "reason":       "",
    }
    if len(variants) < 2:
        out["reason"] = "Only one variation has been sent, so there is nothing to compare."
        return out
    if total_sent < MIN_SENDS_FOR_A_WINNER:
        out["reason"] = (f"Too little volume to call — {total_sent} sent across "
                         f"{len(variants)} variations.")
        return out

    best, runner = variants[0], variants[1]
    if not best[measure]:
        out["reason"] = "No replies yet, so no variation is ahead."
        return out
    if best[measure] < MIN_WINS_FOR_A_WINNER:
        noun = "positive response" if measure == "positive" else "reply"
        out["reason"] = (f"{best[measure]} {noun}"
                         f"{'' if best[measure] == 1 else 's'} is not enough to "
                         f"call a winner.")
        return out
    thin = [v for v in (best, runner) if v["sent"] < MIN_SENDS_PER_VARIANT]
    if thin:
        out["reason"] = (f"The two leading variations need at least "
                         f"{MIN_SENDS_PER_VARIANT} sends each before this can be "
                         f"called; one has {min(v['sent'] for v in thin)}.")
        return out

    best_rate = best[f"{measure}_rate"] or 0.0
    runner_rate = runner[f"{measure}_rate"] or 0.0
    if best_rate == runner_rate:
        joint = [v for v in variants if (v[f"{measure}_rate"] or 0.0) == best_rate]
        out["reason"] = f"{len(joint)} variations are level."
        return out
    if runner_rate and (best_rate - runner_rate) / runner_rate < MIN_RELATIVE_LEAD:
        out["reason"] = (f"{best['copy_label']} and {runner['copy_label']} are "
                         f"within {best_rate - runner_rate:.1f} points of each "
                         f"other — too close to call.")
        return out

    out["winner"] = best["variant"]
    out["winner_label"] = best["copy_label"]
    return out


# ── Smartlead ─────────────────────────────────────────────────────────────────

# A campaign Smartlead reports as finished. Anything else that is not skipped
# outright counts as live, PAUSED included: a paused campaign has not finished
# and its result is still the current one.
FINISHED_STATUSES = frozenset({"COMPLETED", "STOPPED"})


def smartlead_bucket(status: str) -> str:
    """"hidden" | "past" | "live" for a Smartlead campaign status.

    ANYTHING UNRECOGNISED IS LIVE. The column holds whatever Smartlead sent,
    plus "imported" and "unknown" written by our own code, so an allow-list of
    live statuses would quietly hide campaigns whenever Smartlead added a word
    — which is the exact failure this split was introduced to fix. Erring
    the other way shows a finished campaign one bucket too early, which a
    reader can see and dismiss.
    """
    from app.utils.pulse.sync import SKIP_STATUSES

    value = str(status or "").upper()
    if value in SKIP_STATUSES:        # DRAFTED, ARCHIVED: never ran, or filed away
        return "hidden"
    return "past" if value in FINISHED_STATUSES else "live"


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
                # Smartlead reports what went out, so there is no allocation
                # behind it that has not been sent yet. Carried all the same,
                # so every variant dict has the same shape.
                "allocated":   sent,
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
            # The number the team uses for it. The `or` covers a block handed
            # in by a caller that never went through _decide — the replay
            # path below does exactly that — which keeps the old wording
            # rather than inventing a number out of nothing.
            "label":         step.get("winner_label") or f"Version {step['winner']}",
            # The underlying code, so a future reader of a frozen snapshot can
            # tell which variation "Copy 2" was. Reports published before this
            # carry no such key and render exactly as they did.
            "variant":       step["winner"],
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
            "label":         block.get("winner_label") or f"Version {block['winner']}",
            "variant":       block["winner"],
            # Named per campaign now. The old fallback was "Manual campaigns",
            # accurate when there was one combined verdict — but it also
            # told a client which of our tools sent their email, which is a
            # distinction they have no use for. It is all email to them.
            "campaign":      block.get("name") or "Email campaigns",
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
    return {"variant": key, "is_baseline": False, "sent": 0, "allocated": 0,
            "reply": 0, "positive": 0, "unsubscribe": 0, "lead": 0,
            "interested": 0}


def _add_row(totals: dict[str, dict], row: dict) -> None:
    """Fold one sheet's variant row into a running total.

    SENT AND ALLOCATED ARE DIFFERENT NUMBERS WHILE A CAMPAIGN IS RUNNING. The
    merge writes every prospect up front and ticks Send Status as each one goes
    out, so a live campaign's allocation is its eventual size, not its reach.
    `total` was used as `sent` for as long as the panel only ever showed
    finished campaigns, where the two agree; it stops being true the moment a
    live one is on screen.

    Falls back to `total` for a sheet with no Send Status column, which is the
    old behaviour and the right answer when nothing better is available.
    """
    agg = totals.setdefault(row["variant"], _blank_variant(row["variant"]))
    allocated = row.get("total", 0) or 0
    agg["sent"] += row.get("sent", allocated) or 0
    agg["allocated"] += allocated
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
