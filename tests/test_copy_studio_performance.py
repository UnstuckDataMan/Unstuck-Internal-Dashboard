from copy import deepcopy
from datetime import datetime, timezone

from app.utils.copy_studio.performance import comparison, manual_variants, refresh_evidence
from app.utils.pulse.abtest import _decide


def arm(name, sent, positive, reply=0, **extra):
    return {"variant": name, "sent": sent, "positive": positive, "reply": reply,
            "unsubscribe": 0, "allocated": sent, "attribution": "exact",
            "copy_source_keys": [f"sent:{name}"], **extra}


def test_reuses_pulse_rate_winner_and_keeps_losing_arms():
    variants = [arm("A", 1000, 30), arm("B", 100, 12)]
    pulse = _decide(deepcopy(variants))
    studio = comparison(variants, source="smartlead", campaign_id="123")
    assert studio["winner"] == pulse["winner"] == "B"
    assert len(studio["variants"]) == 2
    assert studio["variants"][0]["reference_eligible"] is True
    assert studio["proven_winner"] is False
    assert studio["statistical_significance"] == "not_tested"
    assert variants[0].get("positive_rate") is None  # no caller mutation


def test_reply_fallback_is_not_positive_performance_evidence():
    result = comparison([arm("A", 100, 0, 20), arm("B", 100, 0, 5)], source="manual", campaign_id="m")
    assert result["decided_on"] == "reply"
    assert result["variants"][0]["evidence_level"] == "reply_rate_leader"
    assert not any(row["reference_eligible"] for row in result["variants"])


def test_thin_arms_and_incompatible_denominators_are_not_high_confidence():
    thin = comparison([arm("A", 3, 3), arm("B", 100, 5)], source="manual", campaign_id="m")
    assert thin["winner"] is None
    for kwargs in ({"complete": False}, {"stale": True}, {"campaign_status": "DRAFTED"}, {"campaign_name": "Testing campaign"}):
        result = comparison([arm("A", 100, 12), arm("B", 100, 3)], source="manual", campaign_id="m", **kwargs)
        assert not result["variants"][0]["reference_eligible"]
    result = comparison([arm("A", 100, 12, denominator="allocated_legacy"), arm("B", 100, 3)], source="manual", campaign_id="m")
    assert not result["variants"][0]["reference_eligible"]


def test_attribution_is_required_independently_of_a_high_reply_rate():
    result = comparison([arm("A", 100, 20, attribution="unresolved"), arm("B", 100, 5)], source="smartlead", campaign_id="a")
    assert result["winner"] == "A"
    assert not result["variants"][0]["reference_eligible"]
    assert result["variants"][0]["evidence_level"] == "promising"


def test_manual_snapshot_native_definition_and_legacy_exposure():
    rows = manual_variants([{"variant": "S1/B1", "total": 100, "lead": 5, "interested": 2}], {})
    assert rows[0]["sent"] == 100
    assert rows[0]["denominator"] == "allocated_legacy"
    assert rows[0]["positive"] == 7
    result = comparison(rows, source="manual", campaign_id="m")
    assert result["period"]["kind"] == "all_time_status_snapshot"
    assert result["basis"] == "Manual Lead + Interested statuses"


def test_meeting_request_never_becomes_booked_or_held():
    rows = manual_variants([{"variant": "S1/B1", "sent": 1, "total": 1}], {
        "S1/B1": {"sent": 1, "with_copy": 1, "versions": {"copy": 1}, "meeting_requests": 1}})
    assert rows[0]["meeting_requests"] == 1
    assert rows[0]["meeting_booked"] == rows[0]["meeting_held"] == 0


def test_counts_exceeding_sends_are_ineligible():
    result = comparison([arm("A", 30, 45), arm("B", 30, 5)], source="manual", campaign_id="m")
    assert not result["variants"][0]["reference_eligible"]


def test_kelton_style_fixture_is_an_observed_comparison_not_proven():
    rows = [arm("S1/B1", 138, 7, 5, allocated=139), arm("S2/B2", 137, 9, 3, allocated=138)]
    result = comparison(rows, source="manual", campaign_id="kelton-fixture")
    by_label = {row["variant"]: row for row in result["variants"]}
    assert by_label["S1/B1"]["positive_rate"] == 5.1
    assert by_label["S2/B2"]["positive_rate"] == 6.6
    assert result["proven_winner"] is False
    assert result["statistical_significance"] == "not_tested"


def test_stale_historical_leader_is_not_a_current_reference():
    report = comparison([arm("A", 100, 12), arm("B", 100, 3)], source="manual", campaign_id="m",
                        observed_at="2026-10-01T10:00:00Z")
    assert report["variants"][0]["reference_eligible"]
    decorated = refresh_evidence(report, now=datetime(2026, 10, 28, tzinfo=timezone.utc))
    assert decorated["stale"] and not decorated["variants"][0]["reference_eligible"]
    assert report["variants"][0]["reference_eligible"]  # immutable observation preserved
    assert decorated["winner"] == "A"  # retain historical comparison


def test_missing_or_future_observation_time_is_not_trusted():
    for observed in (None, "2026-10-01", "broken", "2027-01-01T00:00:00Z"):
        result = refresh_evidence({"observed_at": observed, "variants": [{"reference_eligible": True}]},
                                 now=datetime(2026, 10, 28, tzinfo=timezone.utc))
        assert result["stale"]
        assert not result["variants"][0]["reference_eligible"]
