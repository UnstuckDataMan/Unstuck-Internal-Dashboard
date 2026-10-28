"""Offline evaluation-pack behavior; these fixtures are not writing-quality evidence."""
from copy import deepcopy
import csv
import json
from pathlib import Path

import pytest

from scripts import copy_studio_evaluate as evaluate
from app.utils.copy_studio import providers
from app.utils.copy_studio.store import SQLiteStore


def draft(text="Could simpler reporting help your finance team?"):
    return {"angles": [{"name": "Operational simplicity", "rationale": "Focus on the reporting burden.",
                        "subjects": ["Reporting"], "steps": [{"kind": "initial", "body": text, "evidence_ids": ["offer"]}],
                        "evidence_ids": ["offer"], "claims": []}],
            "provenance": {"provider": "openai", "model": "gpt-fixture", "quality_evaluation": "pending"}}


def case(case_id="real-id-supplied", client_id="client-a", campaign_family="heldout-campaign"):
    brief = {"id": "brief-a", "client_id": client_id, "kind": "brief", "status": "approved", "content_hash": "brief-hash",
             "content": {"client_id": client_id, "coverage": {"complete": True}, "constraints": ["Keep language restrained."],
                         "excluded_sectors": ["gambling"], "forbidden_names": ["Private Corp"], "forbidden_phrases": ["guaranteed results"],
                         "facts": [{"id": "audience", "category": "audience", "text": "UK software finance leaders", "evidence": []},
                                   {"id": "offer", "category": "offer", "text": "Help software finance teams simplify reporting.", "evidence": []}],
                         "contradictions": []}}
    example = {"id": "example-a", "client_id": client_id, "kind": "example", "status": "approved", "content_hash": "example-hash",
               "lineage": {"client_family": client_id, "campaign_family": "training-campaign"},
               "content": {"source": "copy_bank", "channel": "email", "subjects": ["Finance reporting"],
                           "steps": [{"kind": "initial", "body": "Hi {{first_name}}, could a reporting review help your finance team?", "evidence_ids": ["offer"]}],
                           "evidence_ids": ["offer"], "reusable": True}}
    return {"case_id": case_id, "client_id": client_id, "origin": "synthetic_fixture",
            "lineage": {"client_family": client_id, "campaign_family": campaign_family},
            "brief": brief, "examples": [example], "methodology": [],
            "request": {"channel": "email", "mode": "draft", "angles": 1,
                        "audience": {"subsector": "software", "geography": "UK", "company_size": "50–200", "role": "Finance director"}},
            "baseline": draft("Would a review of your reporting process be useful?"),
            "outputs": [{"candidate_id": "model-a", "provider": "openai", "model": "gpt-fixture", "draft": draft()},
                        {"candidate_id": "model-b", "provider": "anthropic", "model": "claude-fixture", "draft": draft("Would simpler reporting free up your finance team's time?")}]}


class FixtureModel:
    def __init__(self, name, model):
        self.name, self.model, self.calls = name, model, []

    def complete_json(self, system, user, **kwargs):
        self.calls.append({"system": system, "user": json.loads(user), **kwargs})
        return {"issues": []} if kwargs["schema_name"] == "copy_studio_review" else {"angles": draft()["angles"]}


def test_default_cli_prepares_supplied_copy_without_provider_or_network(monkeypatch, tmp_path):
    cases = tmp_path / "cases.jsonl"
    cases.write_text(json.dumps(case()) + "\n")
    monkeypatch.setattr(evaluate, "_provider", lambda *_: pytest.fail("Default preparation must not construct a provider"))
    monkeypatch.setattr(providers.requests, "post", lambda *a, **k: pytest.fail("No API call allowed"))
    assert evaluate.main(["--input", str(cases), "--out-dir", str(tmp_path / "pack"), "--seed", "7"]) == 0
    result = json.loads((tmp_path / "pack" / "ratings-summary.json").read_text())
    assert result["status"] == "pending"
    assert result["quality_pass"] is None
    assert result["by_candidate"] == {}


def test_generation_requires_explicit_paid_flag_before_factory_is_called():
    prepared = evaluate.validate_cases([case()])
    with pytest.raises(evaluate.EvaluationError, match="disabled"):
        evaluate.generate_outputs(prepared, ["openai=model"], provider_factory=lambda _: pytest.fail("No provider without flag"))


def test_cli_generate_without_paid_flag_fails_before_reading_sources(tmp_path):
    assert evaluate.main(["--input", str(tmp_path / "not-opened"), "--out-dir", str(tmp_path / "pack"), "--generate"]) == 2
    assert not (tmp_path / "pack").exists()


def test_group_leakage_filters_every_related_version_and_missing_lineage():
    item = case()
    same_family = deepcopy(item["examples"][0])
    same_family["id"] = "related-version"
    same_family["lineage"]["campaign_family"] = item["lineage"]["campaign_family"]
    missing = deepcopy(same_family)
    missing["id"] = "no-lineage"
    missing.pop("lineage")
    item["examples"] += [same_family, missing]
    result = evaluate.validate_cases([item])
    assert [r["id"] for r in result["cases"][0]["examples"]] == ["example-a"]
    assert {r["reason"] for r in result["excluded_references"]} == {"heldout_family_overlap", "missing_reference_lineage"}


def test_other_heldout_case_family_also_excludes_reference():
    one = case()
    two = case("second", campaign_family="training-campaign")
    result = evaluate.validate_cases([one, two])
    assert all(not c["examples"] for c in result["cases"])


def test_client_family_holdout_excludes_same_client_other_campaigns():
    result = evaluate.validate_cases([case()], "client_family")
    assert result["cases"][0]["examples"] == []
    assert result["excluded_references"][0]["reason"] == "heldout_family_overlap"


def test_exact_baseline_copy_is_excluded_even_with_different_family_tag():
    item = case()
    item["examples"][0]["content"]["steps"] = deepcopy(item["baseline"]["angles"][0]["steps"])
    result = evaluate.validate_cases([item])
    assert result["cases"][0]["examples"] == []
    assert result["excluded_references"][0]["reason"] == "matches_heldout_baseline_copy"


def test_linkedin_requires_copy_bank_examples_after_holdout_filter():
    item = case()
    item["request"]["channel"] = "linkedin"
    item["examples"][0]["content"].update(channel="linkedin", source="smartlead")
    with pytest.raises(evaluate.EvaluationError, match="LinkedIn"):
        evaluate.validate_cases([item])
    item["examples"][0]["content"]["source"] = "copy_bank"
    result = evaluate.validate_cases([item])
    assert len(result["cases"][0]["examples"]) == 1


def test_cross_client_examples_cannot_enter_context():
    item = case()
    item["examples"][0]["client_id"] = "other-client"
    result = evaluate.validate_cases([item])
    assert result["cases"][0]["examples"] == []


def test_both_providers_receive_identical_writing_context_without_baseline(tmp_path):
    first, second = FixtureModel("openai", "fixture-a"), FixtureModel("anthropic", "fixture-b")
    models = iter([first, second])
    item = case()
    item["outputs"] = []
    result, outcomes = evaluate.generate_outputs(evaluate.validate_cases([item]), ["openai=fixture-a", "anthropic=fixture-b"],
                                                 allow_paid_calls=True, provider_factory=lambda _: next(models),
                                                 meter_store=SQLiteStore(str(tmp_path / "costs.sqlite")))
    assert first.calls[0] == second.calls[0]
    assert "Keep language restrained." in first.calls[0]["user"]["mandatory_constraints"]
    assert "Do not target: gambling" in first.calls[0]["user"]["mandatory_constraints"]
    assert item["baseline"]["angles"][0]["steps"][0]["body"] not in json.dumps(first.calls[0])
    assert len(result[0]["outputs"]) == 2
    assert {r["status"] for r in outcomes} == {"generated"}
    assert len({r["writing_prompt_hash"] for r in outcomes}) == 1
    assert first.calls[1]["schema_name"] == second.calls[1]["schema_name"] == "copy_studio_review"


def test_provider_failure_creates_no_substitute_copy(tmp_path):
    class Broken(FixtureModel):
        def complete_json(self, *a, **kw):
            raise providers.ProviderError("Upstream failed with private text")
    item = case()
    item["outputs"] = []
    result, outcomes = evaluate.generate_outputs(evaluate.validate_cases([item]), ["openai=fixture"], allow_paid_calls=True,
                                                 provider_factory=lambda _: Broken("openai", "fixture"),
                                                 meter_store=SQLiteStore(str(tmp_path / "costs.sqlite")))
    assert result[0]["outputs"] == []
    assert outcomes[0]["status"] == "failed"
    assert "private text" not in json.dumps(outcomes)


def test_blinded_export_contains_copy_but_not_provider_or_provenance(tmp_path):
    manifest = evaluate.build_review_pack(evaluate.validate_cases([case()]), tmp_path / "pack", seed=9)
    review = (tmp_path / "pack" / "blind-review.csv").read_text()
    for forbidden in ("openai", "anthropic", "gpt-fixture", "claude-fixture", "current_human_baseline", "provenance", "model-a", "model-b"):
        assert forbidden not in review
    assert "Could simpler reporting help" in review
    assert manifest["human_review_status"] == "pending"
    assert manifest["quality_pass"] is None
    assert manifest["target_met"] is False
    key = json.loads((tmp_path / "pack" / "answer-key.private.json").read_text())
    assert len(key["candidates"]) == 3
    assert {x["candidate_id"] for x in key["candidates"]} == {"model-a", "model-b", "current_human_baseline"}
    assert (tmp_path / "pack" / "answer-key.private.json").stat().st_mode & 0o777 == 0o600
    with pytest.raises(evaluate.EvaluationError, match="already exist"):
        evaluate.build_review_pack(evaluate.validate_cases([case()]), tmp_path / "pack", seed=9)


def test_provider_self_identification_blocks_blind_pack_instead_of_rewriting(tmp_path):
    item = case()
    item["outputs"][0]["draft"] = draft("As GPT-4o I recommend this reporting process.")
    with pytest.raises(evaluate.EvaluationError, match="identity marker"):
        evaluate.build_review_pack(evaluate.validate_cases([item]), tmp_path / "pack")
    assert not (tmp_path / "pack").exists()


def test_blank_and_partial_ratings_never_imply_quality_pass(tmp_path):
    evaluate.build_review_pack(evaluate.validate_cases([case()]), tmp_path / "pack", seed=9)
    with (tmp_path / "pack" / "blind-review.csv").open(newline="") as handle:
        rows = list(csv.DictReader(handle))
    key = json.loads((tmp_path / "pack" / "answer-key.private.json").read_text())
    assert evaluate.summarise_ratings(rows, key)["status"] == "pending"
    rows[0].update(reviewer_id="reviewer-one", **{field: "4" for field in evaluate.SCORES},
                   editing_minutes="3.5", preference_rank="1", **{field: "yes" for field in evaluate.GATES})
    summary = evaluate.summarise_ratings(rows, key)
    assert summary["status"] == "partial" and summary["rated_candidates"] == 1
    assert summary["quality_pass"] is None and summary["automatic_winner"] is None
    assert summary["release_gate_status"] == "pending"


def test_complete_human_ratings_expose_failures_without_auto_selecting_winner(tmp_path):
    evaluate.build_review_pack(evaluate.validate_cases([case()]), tmp_path / "pack", seed=9)
    with (tmp_path / "pack" / "blind-review.csv").open(newline="") as handle:
        rows = list(csv.DictReader(handle))
    key = json.loads((tmp_path / "pack" / "answer-key.private.json").read_text())
    for i, row in enumerate(rows, 1):
        row.update(reviewer_id="reviewer-one", **{field: "5" for field in evaluate.SCORES},
                   editing_minutes=str(i), preference_rank=str(i), **{field: "yes" for field in evaluate.GATES})
    rows[1]["claims_supported"] = "no"
    summary = evaluate.summarise_ratings(rows, key)
    assert summary["status"] == "rated"
    assert summary["gate_failures"] == 1
    assert summary["release_gate_status"] == "failed"
    assert summary["quality_pass"] is None
    assert summary["benchmark_target_met"] is False


def test_empty_provided_outputs_do_not_create_drafts_or_ratings(tmp_path):
    item = case()
    item["outputs"], item["baseline"] = [], None
    manifest = evaluate.build_review_pack(evaluate.validate_cases([item]), tmp_path / "pack")
    assert manifest["review_candidate_count"] == 0
    assert json.loads((tmp_path / "pack" / "ratings-summary.json").read_text())["status"] == "pending"


def test_origin_and_lineage_are_explicit_and_input_limits_do_not_truncate(tmp_path):
    item = case()
    item.pop("origin")
    with pytest.raises(evaluate.EvaluationError, match="origin"):
        evaluate.validate_cases([item])
    path = tmp_path / "cases.jsonl"
    path.write_text(json.dumps(case()) + "\n" + json.dumps(case("second")) + "\n")
    with pytest.raises(evaluate.EvaluationError, match="limit"):
        evaluate.read_cases(path, maximum=1)


def test_csv_formula_cells_are_escaped():
    assert evaluate._csv_safe("=1+1") == "'=1+1"
    assert evaluate._csv_safe("ordinary text") == "ordinary text"


def test_opt_in_evaluation_calls_record_usage_in_tool_ledger(tmp_path):
    from app.utils.copy_studio import costs
    class MeteredFixture(FixtureModel):
        def complete_json(self, *args, **kwargs):
            costs.record_usage("openai", self.model, {"input_tokens": 100, "output_tokens": 20},
                               request_id=f"fixture-request-{len(self.calls)}")
            return super().complete_json(*args, **kwargs)
    storage = SQLiteStore(str(tmp_path / "costs.sqlite"))
    item = case()
    item["outputs"] = []
    cases, outcomes = evaluate.generate_outputs(evaluate.validate_cases([item]), ["openai=fixture"], allow_paid_calls=True,
                                               provider_factory=lambda _: MeteredFixture("openai", "fixture"),
                                               meter_store=storage, actor="test-operator")
    runs = list(storage.iter_records("client-a", kind="evaluation_run"))
    usage = list(storage.iter_records("client-a", kind="usage"))
    assert len(runs) == 1
    assert len({row["content"]["call_id"] for row in usage}) == 2  # writing and independent review
    assert len([row for row in usage if row["content"]["phase"] == "recorded"]) == 2
    assert runs[0]["created_by"] == "test-operator"
    assert cases[0]["outputs"][0]["evaluation_job_id"] == outcomes[0]["evaluation_job_id"]
    assert all(row["content"]["job_id"] == outcomes[0]["evaluation_job_id"] for row in usage)


def test_partial_rating_failure_is_not_hidden_and_bad_ranks_are_rejected(tmp_path):
    evaluate.build_review_pack(evaluate.validate_cases([case()]), tmp_path / "pack", seed=9)
    with (tmp_path / "pack" / "blind-review.csv").open(newline="") as handle:
        rows = list(csv.DictReader(handle))
    key = json.loads((tmp_path / "pack" / "answer-key.private.json").read_text())
    rows[0]["no_client_leakage"] = "no"
    summary = evaluate.summarise_ratings(rows, key)
    assert summary["status"] == "pending"
    assert summary["gate_failures"] == 1
    assert summary["release_gate_status"] == "failed"
    for row in rows:
        row.update(reviewer_id="reviewer", **{field: "4" for field in evaluate.SCORES},
                   editing_minutes="5", preference_rank="1", **{field: "yes" for field in evaluate.GATES})
    with pytest.raises(evaluate.EvaluationError, match="distinct"):
        evaluate.summarise_ratings(rows, key)


@pytest.mark.parametrize("provider,preferred,fallback", [
    ("openai", "COPY_STUDIO_OPENAI_API_KEY", "OPENAI_API_KEY"),
    ("anthropic", "COPY_STUDIO_ANTHROPIC_API_KEY", "ANTHROPIC_API_KEY"),
])
def test_company_key_alias_preferred_then_legacy_fallback(monkeypatch, provider, preferred, fallback):
    monkeypatch.setenv(preferred, "fixture-tool-company-key")
    monkeypatch.setenv(fallback, "fixture-existing-company-key")
    configured = evaluate._provider(f"{provider}=fixture-model")
    assert configured._key == "fixture-tool-company-key"
    assert configured.model == "fixture-model"
    monkeypatch.delenv(preferred)
    assert evaluate._provider(f"{provider}=fixture-model")._key == "fixture-existing-company-key"
    monkeypatch.delenv(fallback)
    with pytest.raises(evaluate.EvaluationError) as error:
        evaluate._provider(f"{provider}=fixture-model")
    assert preferred in str(error.value) and fallback in str(error.value)
    assert "fixture-tool-company-key" not in str(error.value)


def test_occupied_output_directory_blocks_before_paid_generation(monkeypatch, tmp_path):
    cases = tmp_path / "cases.jsonl"
    cases.write_text(json.dumps(case()) + "\n")
    output = tmp_path / "pack"
    output.mkdir()
    (output / "blind-review.csv").write_text("existing human ratings")
    monkeypatch.setattr(evaluate, "generate_outputs", lambda *a, **k: pytest.fail("Must reject occupied pack before any paid call"))
    assert evaluate.main(["--input", str(cases), "--out-dir", str(output), "--generate", "--allow-paid-calls",
                          "--model", "openai=fixture-model"]) == 2
    assert (output / "blind-review.csv").read_text() == "existing human ratings"
