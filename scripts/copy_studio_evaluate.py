#!/usr/bin/env python3
"""Prepare a private, blind comparison pack; never call a model by default.

Inputs and outputs can contain client evidence. Keep the pack in an access-
controlled directory outside Git. See docs/copy-studio-evaluation.md.
"""
from __future__ import annotations

import argparse
from copy import deepcopy
import csv
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import random
import re
import secrets
import statistics
import sys
import time
import uuid

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.utils.copy_studio import providers, writer
from app.utils.copy_studio.store import StoreError

SCHEMA_VERSION = "copy-studio-evaluation-1"
SCORES = ("factual_accuracy", "audience_relevance", "voice", "specificity", "channel_fit", "strategic_variety")
GATES = ("claims_supported", "no_client_leakage", "exclusions_respected", "variables_and_format_valid")
RATING_FIELDS = ("reviewer_id", *SCORES, "editing_minutes", "preference_rank", *GATES, "review_notes")
IDENTITY_MARKER = re.compile(r"\b(?:openai|anthropic|claude(?:[- ]\d[\w.-]*)?|chatgpt|gpt[- ]?\d[\w.-]*)\b", re.I)


class EvaluationError(ValueError):
    pass


def _hash(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False, allow_nan=False).encode()).hexdigest()


def _text(value):
    return isinstance(value, str) and bool(value.strip())


def read_cases(path: Path, maximum: int = 100) -> list[dict]:
    cases = []
    with Path(path).open(encoding="utf-8") as handle:
        for line_number, line in enumerate(handle, 1):
            if not line.strip():
                continue
            if len(line.encode()) > 4_000_000:
                raise EvaluationError(f"Input line {line_number} exceeds 4 MB.")
            try:
                item = json.loads(line)
            except ValueError as exc:
                raise EvaluationError(f"Input line {line_number} is not valid JSON.") from exc
            cases.append(item)
            if len(cases) > maximum:
                raise EvaluationError(f"Pack exceeds the explicit {maximum}-case limit; split the input or raise --max-cases.")
    return cases


def _copy_parts(draft):
    if not isinstance(draft, dict):
        return []
    parts = []
    for angle in draft.get("angles", []):
        parts.extend(str(s).strip() for s in angle.get("subjects", []) if str(s).strip())
        parts.extend(str(s.get("body", "")).strip() for s in angle.get("steps", []) if str(s.get("body", "")).strip())
    return parts


def _reference_parts(record):
    c = record.get("content", {})
    steps = c.get("steps") or [{"body": b} for b in c.get("bodies", [])]
    if not steps and c.get("body"):
        steps = [{"body": c["body"]}]
    return [str(x).strip() for x in c.get("subjects", []) if str(x).strip()] + [
        str(x if isinstance(x, str) else x.get("body", "")).strip() for x in steps]


def _normalise(text):
    return " ".join(text.lower().split())


def validate_cases(cases: list[dict], holdout_level: str = "campaign_family") -> dict:
    if holdout_level not in {"campaign_family", "client_family"}:
        raise EvaluationError("Holdout level must be campaign_family or client_family.")
    if not cases:
        raise EvaluationError("At least one supplied case is required; no cases are invented.")
    seen = set()
    for index, case in enumerate(cases, 1):
        if not isinstance(case, dict) or not _text(case.get("case_id")):
            raise EvaluationError(f"Case {index} needs a case_id.")
        if case["case_id"] in seen:
            raise EvaluationError("Case IDs must be unique.")
        seen.add(case["case_id"])
        if case.get("origin") not in {"real", "synthetic_fixture"}:
            raise EvaluationError(f"Case {index} must declare origin real or synthetic_fixture.")
        lineage = case.get("lineage") or {}
        if not all(_text(lineage.get(key)) for key in ("client_family", "campaign_family")):
            raise EvaluationError(f"Case {index} needs stable client_family and campaign_family lineage.")
        if not _text(case.get("client_id")) or not isinstance(case.get("brief"), dict):
            raise EvaluationError(f"Case {index} needs a client ID and approved brief record.")
        request = case.get("request") or {}
        if request.get("channel") not in {"email", "linkedin"}:
            raise EvaluationError(f"Case {index} needs an email or LinkedIn request.")
        if request.get("mode", "draft") not in {"draft", "new_draft"}:
            raise EvaluationError("Provider comparisons use new-draft mode; assembly is a separate workflow.")
        # Validate client ownership, approval, required facts and contradictions
        # using the same checks as the live writer, before any optional model call.
        try:
            writer._permission_guard(case["brief"], case["client_id"])
            writer.effective_limits(request["channel"], request.get("limits"))
        except writer.WritingError as exc:
            raise EvaluationError(f"Case {index}: {exc}") from exc
        if not isinstance(case.get("examples", []), list) or not isinstance(case.get("methodology", []), list):
            raise EvaluationError(f"Case {index}: examples and methodology must be record lists.")
        if not isinstance(case.get("outputs", []), list):
            raise EvaluationError(f"Case {index}: outputs must be a list of supplied drafts.")
        for candidate in case.get("outputs", []):
            if not isinstance(candidate, dict) or not _text(candidate.get("candidate_id")) or not isinstance(candidate.get("draft"), dict):
                raise EvaluationError(f"Case {index}: provided outputs need candidate_id and draft.")
        if case.get("baseline") is not None and not isinstance(case["baseline"], dict):
            raise EvaluationError(f"Case {index}: baseline must be a structured draft object.")
    held_out = {case["lineage"][holdout_level] for case in cases}
    heldout_parts = {_normalise(text) for case in cases for text in _copy_parts(case.get("baseline")) if len(text) >= 20}
    prepared, exclusions = [], []
    for case in cases:
        item = deepcopy(case)
        safe = []
        for record in item.get("examples", []):
            if not isinstance(record, dict):
                raise EvaluationError("Every example must be an immutable example record.")
            lineage = record.get("lineage") or {}
            reason = None
            if not all(_text(lineage.get(k)) for k in ("client_family", "campaign_family")):
                reason = "missing_reference_lineage"
            elif lineage[holdout_level] in held_out:
                reason = "heldout_family_overlap"
            elif any(_normalise(text) in heldout_parts for text in _reference_parts(record) if len(text) >= 20):
                reason = "matches_heldout_baseline_copy"
            if reason:
                exclusions.append({"case_id": case["case_id"], "record_id": record.get("id"), "reason": reason})
            else:
                safe.append(record)
        selected = writer.retrieve_examples(item["client_id"], item["request"]["channel"], safe,
                                            item["request"].get("audience"))
        selected_ids = {r["id"] for r in selected}
        for record in safe:
            if record.get("id") not in selected_ids:
                exclusions.append({"case_id": case["case_id"], "record_id": record.get("id"),
                                   "reason": "writer_scope_channel_approval_or_retrieval_filter"})
        item["examples"] = selected
        if item["request"]["channel"] == "linkedin" and not selected:
            raise EvaluationError(f"Case {case['case_id']} has no approved LinkedIn Copy Bank example after holdout filtering.")
        prepared.append(item)
    real = [c for c in cases if c["origin"] == "real"]
    real_clients = {c["lineage"]["client_family"] for c in real}
    target_met = 50 <= len(real) <= 100 and 5 <= len(real_clients) <= 10
    return {"cases": prepared, "excluded_references": exclusions,
            "manifest": {"schema_version": SCHEMA_VERSION, "holdout_level": holdout_level,
                         "case_count": len(cases), "real_case_count": len(real), "synthetic_case_count": len(cases) - len(real),
                         "real_client_family_count": len(real_clients), "target": "50–100 held-out real briefs across 5–10 pilot clients",
                         "target_met": target_met, "human_review_status": "pending", "quality_pass": None,
                         "baseline_count": sum(c.get("baseline") is not None for c in cases),
                         "source_authenticity": "Origin and lineage are supplied by the pack curator; this utility does not authenticate source systems."}}


def _provider(spec):
    try:
        name, model = spec.split("=", 1)
    except ValueError as exc:
        raise EvaluationError("Model choices must be provider=model, for example openai=<configured-model>.") from exc
    if name not in {"openai", "anthropic"} or not model.strip():
        raise EvaluationError("Choose an explicit OpenAI or Anthropic model.")
    key_name = "COPY_STUDIO_OPENAI_API_KEY" if name == "openai" else "COPY_STUDIO_ANTHROPIC_API_KEY"
    fallback = "OPENAI_API_KEY" if name == "openai" else "ANTHROPIC_API_KEY"
    key = (os.environ.get(key_name) or os.environ.get(fallback) or "").strip()
    if not key:
        raise EvaluationError(f"Missing company server credential: {key_name} (or existing {fallback}).")
    return providers.OpenAIProvider(model.strip(), key) if name == "openai" else providers.AnthropicProvider(model.strip(), key)


class _CapturedProvider:
    def __init__(self, wrapped):
        self.wrapped, self.name, self.model = wrapped, wrapped.name, wrapped.model
        self.writing_prompt_hash = None

    def complete_json(self, system, user, **kwargs):
        if kwargs.get("schema_name") == "copy_studio_draft":
            self.writing_prompt_hash = _hash({"system": system, "user": json.loads(user), **kwargs})
        return self.wrapped.complete_json(system, user, **kwargs)


def generate_outputs(prepared: dict, model_specs: list[str], *, allow_paid_calls: bool = False,
                     provider_factory=None, meter_store=None, actor="copy-studio-evaluation") -> tuple[list[dict], list[dict]]:
    # This guard comes before provider construction and all network-capable code.
    if not allow_paid_calls:
        raise EvaluationError("Model calls are disabled. Use --generate AND --allow-paid-calls for explicit paid evaluation.")
    if not model_specs or len(set(model_specs)) != len(model_specs) or len(model_specs) > 3:
        raise EvaluationError("Select one to three distinct explicit provider=model choices.")
    from app.utils.copy_studio import costs
    from app.utils.copy_studio.store import get_store
    # Paid evaluations belong in the same private cost ledger as UI jobs.
    # Never fall back to unmetered generation if storage is unconfigured.
    meter_store = meter_store if meter_store is not None else get_store()
    factory = provider_factory or _provider
    chosen = [factory(spec) for spec in model_specs]
    cases, outcomes = deepcopy(prepared["cases"]), []
    for case in cases:
        first_prompt = None
        outputs = list(case.get("outputs", []))
        for provider in chosen:
            captured = _CapturedProvider(provider)
            started = time.monotonic()
            candidate_id = f"{provider.name}:{provider.model}"
            evaluation_job = {"id": str(uuid.uuid4()), "client_id": case["client_id"],
                              "created_by": actor, "kind": "evaluation", "agency_id": meter_store.agency_id}
            try:
                meter_store.save("evaluation_run", case["client_id"], evaluation_job["id"],
                                 {"case_id": case["case_id"], "candidate_id": candidate_id,
                                  "brief_id": case["brief"]["id"], "job_id": evaluation_job["id"],
                                  "heldout_level": prepared["manifest"]["holdout_level"]}, actor)
                with costs.metering(meter_store, evaluation_job):
                    draft = writer.generate(case["client_id"], deepcopy(case["brief"]), deepcopy(case["examples"]),
                                            deepcopy(case["request"]), provider=captured, methodology=deepcopy(case.get("methodology", [])))
                if first_prompt is not None and first_prompt != captured.writing_prompt_hash:
                    raise EvaluationError("Compared writing prompts differ; discard this run and repeat with a frozen brief/corpus.")
                first_prompt = captured.writing_prompt_hash
                outputs.append({"candidate_id": candidate_id, "provider": provider.name, "model": provider.model,
                                "draft": draft, "writing_prompt_hash": captured.writing_prompt_hash,
                                "elapsed_seconds": round(time.monotonic() - started, 3), "cost": None,
                                "evaluation_job_id": evaluation_job["id"]})
                outcomes.append({"case_id": case["case_id"], "candidate_id": candidate_id, "status": "generated",
                                 "writing_prompt_hash": captured.writing_prompt_hash, "evaluation_job_id": evaluation_job["id"]})
            except (providers.ProviderError, writer.WritingError) as exc:
                outcomes.append({"case_id": case["case_id"], "candidate_id": candidate_id, "status": "failed",
                                 "error_type": type(exc).__name__, "evaluation_job_id": evaluation_job["id"],
                                 "message": "Generation failed; inspect server configuration or brief validity. No substitute copy was created."})
        case["outputs"] = outputs
    return cases, outcomes


def _copy_text(draft):
    providers.validate_json({"angles": draft.get("angles")}, writer.DRAFT_SCHEMA)
    lines = []
    for i, angle in enumerate(draft["angles"], 1):
        lines.append(f"ANGLE {i}")
        for subject in angle.get("subjects", []):
            lines.append(f"Subject: {subject}")
        for step in angle["steps"]:
            lines.append(f"{step['kind']}: {step['body']}")
    text = "\n\n".join(lines)
    if IDENTITY_MARKER.search(text):
        raise EvaluationError("A draft contains a provider/model identity marker; resolve this before blind review rather than silently editing its copy.")
    return text


def _csv_safe(value):
    value = str(value)
    return "'" + value if value.startswith(("=", "+", "-", "@", "\t", "\r")) else value


def _write_private(path: Path, text: str):
    # Refuse overwrite of an existing pack so identities cannot drift under ratings.
    with path.open("x", encoding="utf-8", newline="") as handle:
        os.chmod(path, 0o600)
        handle.write(text)


def _pack_paths(out_dir: Path):
    out_dir = Path(out_dir)
    paths = [out_dir / "blind-review.csv", out_dir / "answer-key.private.json", out_dir / "manifest.json", out_dir / "ratings-summary.json"]
    if out_dir.exists() and not out_dir.is_dir():
        raise EvaluationError("Output destination must be a directory.")
    if any(path.exists() for path in paths):
        raise EvaluationError("Output files already exist. Use a new directory to preserve review identities.")
    return paths


def build_review_pack(prepared: dict, out_dir: Path, *, seed=None, generated_cases=None, generation_outcomes=None):
    rng = random.Random(seed if seed is not None else secrets.randbits(128))
    cases = generated_cases if generated_cases is not None else prepared["cases"]
    order = list(cases)
    rng.shuffle(order)
    rows, key = [], []
    for case_number, case in enumerate(order, 1):
        blind_case = f"Case-{case_number:03d}"
        candidates = list(case.get("outputs", []))
        if case.get("baseline") is not None:
            candidates.append({"candidate_id": "current_human_baseline", "provider": "human", "model": None, "draft": case["baseline"]})
        if len({c["candidate_id"] for c in candidates}) != len(candidates):
            raise EvaluationError("A case contains duplicate candidate IDs; keep one output per model/run or give unique provided-output IDs.")
        rng.shuffle(candidates)
        for position, candidate in enumerate(candidates):
            if position >= 26:
                raise EvaluationError("A blind comparison supports at most 26 candidates per case.")
            label = chr(65 + position)
            content = case["brief"]["content"]
            facts = deepcopy(content.get("facts", []))
            constraints = {key: content.get(key, []) for key in
                           ("constraints", "excluded_sectors", "forbidden_names", "forbidden_phrases", "contradictions")}
            row = {"blind_case_id": blind_case, "candidate": label, "channel": case["request"]["channel"],
                   "audience": json.dumps(case["request"].get("audience", {}), ensure_ascii=False),
                   "approved_facts_and_sources": json.dumps(facts, ensure_ascii=False),
                   "mandatory_constraints": json.dumps(constraints, ensure_ascii=False),
                   "brief_coverage": json.dumps(content.get("coverage", {}), ensure_ascii=False),
                   "copy": _copy_text(candidate["draft"]), **{field: "" for field in RATING_FIELDS}}
            rows.append(row)
            key.append({"blind_case_id": blind_case, "candidate": label, "case_id": case["case_id"],
                        "client_id": case["client_id"], "origin": case["origin"], "lineage": case["lineage"],
                        "candidate_id": candidate["candidate_id"], "provider": candidate.get("provider"), "model": candidate.get("model"),
                        "content_hash": _hash(candidate["draft"]), "writing_prompt_hash": candidate.get("writing_prompt_hash"),
                        "elapsed_seconds": candidate.get("elapsed_seconds"), "cost": candidate.get("cost"),
                        "evaluation_job_id": candidate.get("evaluation_job_id"),
                        "draft_provenance": candidate["draft"].get("provenance", {}),
                        "automated_review": candidate["draft"].get("review", {}),
                        "source_versions": {r["id"]: r.get("content_hash") for r in [case["brief"], *case["examples"]]}})
    out_dir = Path(out_dir)
    paths = _pack_paths(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    os.chmod(out_dir, 0o700)
    import io
    csv_buffer = io.StringIO(newline="")
    fields = ("blind_case_id", "candidate", "channel", "audience", "approved_facts_and_sources", "mandatory_constraints", "brief_coverage", "copy", *RATING_FIELDS)
    csv_writer = csv.DictWriter(csv_buffer, fieldnames=fields)
    csv_writer.writeheader()
    csv_writer.writerows({k: _csv_safe(v) for k, v in row.items()} for row in rows)
    manifest = {**prepared["manifest"], "created_at": datetime.now(timezone.utc).isoformat(),
                "review_candidate_count": len(rows), "cases_with_outputs": len({r["blind_case_id"] for r in rows}),
                "excluded_references": prepared["excluded_references"], "generation_outcomes": generation_outcomes or [],
                "model_calls_requested": generated_cases is not None,
                "cost_status": "Paid API calls are recorded in the tool cost ledger under evaluation_job_id; missing usage or pricing remains unknown. Supplied outputs have no inferred cost.",
                "quality_pass": None, "human_review_status": "pending"}
    answer_key = {"schema_version": SCHEMA_VERSION, "candidates": key, "manifest": manifest}
    _write_private(paths[0], csv_buffer.getvalue())
    _write_private(paths[1], json.dumps(answer_key, indent=2, ensure_ascii=False))
    _write_private(paths[2], json.dumps(manifest, indent=2, ensure_ascii=False))
    summary = summarise_ratings(rows, answer_key)
    _write_private(paths[3], json.dumps(summary, indent=2, ensure_ascii=False))
    return manifest


def summarise_ratings(rows: list[dict], answer_key: dict):
    expected = {(item["blind_case_id"], item["candidate"]): item for item in answer_key["candidates"]}
    per_case = {}
    for case_id, label in expected:
        per_case.setdefault(case_id, set()).add(label)
    seen, completed, aggregates, failures, ranks = set(), 0, {}, 0, {}
    for row in rows:
        identity = (row.get("blind_case_id"), row.get("candidate"))
        if identity not in expected or identity in seen:
            raise EvaluationError("Ratings contain an unknown or duplicate blind candidate.")
        seen.add(identity)
        failed = any(str(row.get(field, "")).strip().lower() == "no" for field in GATES)
        failures += int(failed)
        if not all(str(row.get(field, "")).strip() for field in RATING_FIELDS if field != "review_notes"):
            continue
        try:
            scores = {field: int(row[field]) for field in SCORES}
            if any(not 1 <= score <= 5 for score in scores.values()):
                raise ValueError
            minutes, rank = float(row["editing_minutes"]), int(row["preference_rank"])
            if minutes < 0 or not minutes < 100000 or not 1 <= rank <= len(per_case[identity[0]]):
                raise ValueError
            gates = {field: str(row[field]).strip().lower() for field in GATES}
            if any(value not in {"yes", "no"} for value in gates.values()):
                raise ValueError
        except (TypeError, ValueError) as exc:
            raise EvaluationError("Completed ratings need integer1–5 scores, nonnegative editing minutes, positive rank and yes/no gates.") from exc
        completed += 1
        candidate_id = expected[identity]["candidate_id"]
        bucket = aggregates.setdefault(candidate_id, {"rated_count": 0, "first_preferences": 0, "gate_failures": 0,
                                                       "editing_minutes": [], **{k: [] for k in SCORES}})
        bucket["rated_count"] += 1
        ranks.setdefault(identity[0], []).append((candidate_id, rank))
        bucket["gate_failures"] += int(failed)
        bucket["editing_minutes"].append(minutes)
        for field, score in scores.items():
            bucket[field].append(score)
    paired_cases = 0
    for case_id, values in ranks.items():
        if len({rank for _, rank in values}) != len(values):
            raise EvaluationError("Preference ranks must be distinct within a case; explain an unresolved tie in review notes before final ranking.")
        if len(values) == len(per_case[case_id]) and len(values) > 1:
            paired_cases += 1
            for candidate_id, rank in values:
                aggregates[candidate_id]["first_preferences"] += int(rank == 1)
    for bucket in aggregates.values():
        for field in (*SCORES, "editing_minutes"):
            bucket[field] = round(statistics.mean(bucket[field]), 3)
    full = bool(expected) and completed == len(expected)
    target_met = answer_key.get("manifest", {}).get("target_met", False)
    return {"status": "rated" if full else "partial" if completed else "pending", "rated_candidates": completed,
            "expected_candidates": len(expected), "unrated_candidates": len(expected) - completed,
            "completed_paired_cases": paired_cases,
            "gate_failures": failures, "by_candidate": aggregates, "benchmark_target_met": target_met,
            "quality_pass": None, "automatic_winner": None,
            "decision": "Human comparison available; assess paired client/channel results, editing effort and factual gates before choosing a provider. Deployment still requires approval." if full
            else "Human blind ratings are pending or incomplete. Automated checks do not establish writing quality.",
            "release_gate_status": "failed" if failures else "reviewed_sample_clear" if full else "pending"}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, help="Curated JSONL brief cases; required except when scoring an existing pack.")
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--holdout-level", choices=("campaign_family", "client_family"), default="campaign_family")
    parser.add_argument("--max-cases", type=int, default=100)
    parser.add_argument("--seed", type=int, help="Optional reproducible random ordering; keep private from reviewers.")
    parser.add_argument("--generate", action="store_true", help="Request actual model outputs in addition to supplied outputs.")
    parser.add_argument("--allow-paid-calls", action="store_true", help="Explicit opt-in to company API calls and charges.")
    parser.add_argument("--model", action="append", default=[], help="Repeat for provider=model choices; no model quality default is assumed.")
    parser.add_argument("--actor", default="copy-studio-evaluation", help="Operator identity for the private evaluation cost ledger; never a review approval.")
    parser.add_argument("--ratings", type=Path, help="Score a completed CSV against out-dir/answer-key.private.json; makes no API calls.")
    args = parser.parse_args(argv)
    try:
        if args.ratings:
            if args.generate or args.allow_paid_calls or args.model:
                raise EvaluationError("Scoring ratings cannot request generation.")
            with args.ratings.open(encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
            key = json.loads((args.out_dir / "answer-key.private.json").read_text())
            result = summarise_ratings(rows, key)
            path = args.out_dir / f"ratings-summary-{int(time.time())}.json"
            _write_private(path, json.dumps(result, indent=2, ensure_ascii=False))
            print(json.dumps({"status": result["status"], "rated_candidates": result["rated_candidates"], "summary": str(path)}))
            return 0
        if not args.input or not 1 <= args.max_cases <= 500:
            raise EvaluationError("Supply --input and a --max-cases limit between1 and500.")
        if args.generate and not args.allow_paid_calls:
            raise EvaluationError("Generation requires both --generate and --allow-paid-calls.")
        if not args.generate and (args.allow_paid_calls or args.model):
            raise EvaluationError("Model flags only apply with --generate; default preparation never calls APIs.")
        prepared = validate_cases(read_cases(args.input, args.max_cases), args.holdout_level)
        # Fail before charging for model calls if this review destination is
        # already occupied. build_review_pack repeats the check against races.
        _pack_paths(args.out_dir)
        generated, outcomes = (generate_outputs(prepared, args.model, allow_paid_calls=True, actor=args.actor) if args.generate else (None, None))
        manifest = build_review_pack(prepared, args.out_dir, seed=args.seed, generated_cases=generated, generation_outcomes=outcomes)
        print(json.dumps({"status": "prepared", "cases": manifest["case_count"], "review_candidates": manifest["review_candidate_count"],
                          "human_review_status": "pending", "output_directory": str(args.out_dir)}))
        return 0
    except (EvaluationError, providers.ProviderError, StoreError, OSError, ValueError) as exc:
        # Expected errors contain instructions or validation metadata, not source bodies.
        print(f"Evaluation stopped: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
