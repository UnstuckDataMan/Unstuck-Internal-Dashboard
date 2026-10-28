"""Evidence labels around Pulse's existing A/B decision, not another winner model."""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json

from app.utils.pulse.abtest import _decide


def _count(value) -> int:
    try:
        return max(0, int(value or 0))
    except (TypeError, ValueError):
        return 0


def manual_variant_id(label: str) -> str:
    """Canonical identifier for a manual label (still scoped by campaign/step)."""
    return hashlib.sha256(json.dumps(str(label), ensure_ascii=False, separators=(",", ":")).encode()).hexdigest()[:24]


def comparison(variants: list[dict], *, source: str, campaign_id: str,
               step: int = 1, period: dict | None = None,
               observed_at: str | None = None, complete: bool = True,
               campaign_status: str = "", campaign_name: str = "",
               stale: bool = False) -> dict:
    """Keep all arms, their attribution and the denominator behind a leader.

    Pulse's heuristic is reused verbatim. A heuristic leader is never called
    statistically significant. Ungraded reply fallback, missing sent counts,
    unresolved attribution and incomplete coverage cannot become a supported
    positive-response reference.
    """
    rows = []
    for item in variants:
        row = deepcopy(item)
        if not str(row.get("variant", "")).strip():
            continue
        for field in ("sent", "allocated", "positive", "reply", "unsubscribe"):
            row[field] = _count(row.get(field))
        row.setdefault("attribution", "unresolved")
        row.setdefault("copy_source_keys", [])
        row.setdefault("denominator", "sent")
        rows.append(row)
    decision = _decide(rows)
    excluded = str(campaign_status).upper() in {"DRAFT", "DRAFTED", "TEST", "DELETED", "ARCHIVED"}
    # Name alone is insufficient to prove that an arbitrary campaign is real;
    # obvious tests must nevertheless not get promoted automatically.
    name_words = campaign_name.lower().replace("_", " ").replace("-", " ").split()
    excluded |= any(w in {"test", "testing", "dummy"} for w in name_words)
    for row in decision["variants"]:
        leader = row["variant"] == decision["winner"]
        problems = []
        if not complete:
            problems.append("Source coverage is partial.")
        if stale:
            problems.append("Source data is stale.")
        if excluded:
            problems.append("Test, draft or excluded campaign.")
        if row.get("attribution") != "exact":
            problems.append("Results cannot be assigned to one exact sent copy version.")
        if row.get("denominator") != "sent":
            problems.append("Denominator is allocation, not verified sends.")
        if row["positive"] > row["sent"] or row["reply"] > row["sent"]:
            problems.append("Outcome count exceeds the send denominator.")
        if row.get("unverified_outcome_rows"):
            problems.append("Some outcome statuses belong to rows without verified sends.")
        if not row.get("copy_source_keys"):
            problems.append("No non-empty sent copy was recovered.")
        if decision["decided_on"] != "positive":
            problems.append("Reply-rate fallback; positive outcomes are not graded.")
        row["is_leader"] = leader
        row["reference_eligible"] = bool(leader and not problems)
        row["evidence_level"] = (
            "observed_positive_leader" if row["reference_eligible"] else
            "reply_rate_leader" if leader and decision["decided_on"] == "reply" else
            "promising" if leader else "comparison_arm")
        row["limitations"] = problems
    decision.update({
        "source": source, "channel": "email", "campaign_id": str(campaign_id),
        "step": step, "campaign_name": campaign_name,
        "basis": "Smartlead native positive replies" if source == "smartlead" else "Manual Lead + Interested statuses",
        "period": period or {"kind": "all_time_status_snapshot"},
        "observed_at": observed_at or datetime.now(timezone.utc).isoformat(),
        "coverage_complete": complete, "stale": stale,
        "statistical_significance": "not_tested", "proven_winner": False,
        "interpretation": "Observed comparison using Pulse's sample and relative-rate gates; not proof that wording caused the result.",
        "comparison_unit": "campaign + sequence step + subject/body pair",
        "confounders": ["audience", "deliverability", "offer", "timing"],
    })
    return decision


def manual_variants(stats: list[dict], observations: dict[str, dict]) -> list[dict]:
    """Apply the existing manual counts without pooling unrelated S1/B1 labels."""
    rows = []
    for item in stats:
        label = str(item.get("variant") or "")
        if not label:
            continue
        observation = observations.get(label, {})
        versions = observation.get("versions", {})
        sent = _count(item.get("sent", item.get("total")))
        exact = (len(versions) == 1 and observation.get("with_copy", 0) == sent
                 and observation.get("sent", 0) == sent and "sent" in item
                 and not observation.get("unverified_outcome_rows"))
        rows.append({
            **item, "variant": label, "variant_id": manual_variant_id(label),
            "sequence_id": "manual", "sent": sent, "allocated": _count(item.get("total")),
            "positive": _count(item.get("positive", _count(item.get("lead")) + _count(item.get("interested")))),
            "attribution": "exact" if exact else "unresolved",
            "denominator": "sent" if "sent" in item else "allocated_legacy",
            "copy_source_keys": sorted(versions),
            "observed_versions": [{"source_key": key, "sent": count} for key, count in sorted(versions.items())],
            "meeting_requests": _count(observation.get("meeting_requests")),
            "unverified_outcome_rows": _count(observation.get("unverified_outcome_rows")),
            "meeting_booked": 0, "meeting_held": 0,
        })
    return rows


def refresh_evidence(report: dict, *, now: datetime | None = None, max_age_days: int = 7) -> dict:
    """Decorate an immutable observation at read time without rewriting history.

    An old report is still a valid historical observation, but a stale snapshot
    must not silently remain an eligible current performance reference forever.
    """
    result = deepcopy(report)
    now = now or datetime.now(timezone.utc)
    try:
        observed = datetime.fromisoformat(str(result.get("observed_at") or "").replace("Z", "+00:00"))
        if observed.tzinfo is None:
            raise ValueError("Missing timezone")
        age = (now - observed).total_seconds()
        fresh = 0 <= age <= max_age_days * 86400
    except (ValueError, TypeError):
        fresh = False
    result["stale"] = bool(result.get("stale") or not fresh)
    if result["stale"]:
        for variant in result.get("variants", []):
            variant["reference_eligible"] = False
            variant["limitations"] = list(dict.fromkeys([*variant.get("limitations", []), "Source data is stale or its observation time is unavailable."]))
    return result
