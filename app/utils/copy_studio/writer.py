"""Grounded retrieval, controlled assembly, drafting and independent review.

Output is always a pending immutable draft. Automated checks aid a reviewer;
they are not evidence that copy quality or every semantic claim is correct.
"""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import re

from .methodology import SYSTEM_RULES, VERSION, approved_methodology
from .providers import ProviderError, get_provider, validate_json

DEFAULT_LIMITS = {
    "email": {"subject": 120, "initial": 1200, "follow_up": 800, "closing": 500},
    "linkedin": {"connection": 200, "initial": 600, "follow_up": 600, "closing": 400},
}
ALLOWED_VARIABLES = {"first_name", "last_name", "company", "company_name", "job_title",
                     "sender_name", "sender_company"}
_VARIABLE = re.compile(r"{{\s*([A-Za-z_][A-Za-z0-9_]*)\s*}}")
_TOKENS = re.compile(r"[a-z0-9]+")
_PRIOR_CONTACT = re.compile(r"\b(?:I|we) (?:emailed|called|spoke to|met|contacted|sent you)|\b(?:following up on|as discussed|as promised|our (?:call|meeting|conversation))\b", re.I)
_TRAVEL = re.compile(r"\b(?:I(?:'m| am)|we(?:'re| are)|I'll|we'll) (?:visiting|travelling|traveling|attending|in (?:London|New York|Paris|Berlin)|coming to)|\b(?:at your (?:event|conference)|in your (?:area|city|office))\b", re.I)
_PROOF = re.compile(r"\b(?:we|I|our (?:team|clients)) (?:helped|grew|increased|generated|delivered|achieved|worked with|work with|partnered with)|\b(?:case study|trusted by|clients include|guarantee[ds]?)\b", re.I)
_NUMERIC_CLAIM = re.compile(r"(?:[$£€]\s?\d[\d,.]*|\d[\d,.]*\s?%|\b\d[\d,.]*\s+(?:leads|clients|customers|sales|meetings|revenue|ROI)\b)", re.I)
_DENIED_SUPPORT = re.compile(r"\b(?:cannot|can't|must not|may not|not allowed|not permitted|do not|don't|never|no permission|without permission|no prior|no previous|no travel|not attending|not visiting)\b", re.I)

_STRINGS = {"type": "array", "items": {"type": "string", "maxLength": 300}, "maxItems": 30}
_STEP = {"type": "object", "additionalProperties": False, "required": ["kind", "body", "evidence_ids"],
         "properties": {"kind": {"type": "string", "enum": ["connection", "initial", "follow_up", "closing"]},
                        "body": {"type": "string", "minLength": 1, "maxLength": 10000}, "evidence_ids": _STRINGS}}
_CLAIM = {"type": "object", "additionalProperties": False, "required": ["text", "fact_ids"],
          "properties": {"text": {"type": "string", "minLength": 1, "maxLength": 2000}, "fact_ids": _STRINGS}}
_ANGLE = {"type": "object", "additionalProperties": False,
          "required": ["name", "rationale", "subjects", "steps", "evidence_ids", "claims"],
          "properties": {"name": {"type": "string", "minLength": 1, "maxLength": 100},
                         "rationale": {"type": "string", "maxLength": 1200},
                         "subjects": {"type": "array", "items": {"type": "string", "maxLength": 500}, "maxItems": 3},
                         "steps": {"type": "array", "items": _STEP, "minItems": 1, "maxItems": 6},
                         "evidence_ids": _STRINGS, "claims": {"type": "array", "items": _CLAIM, "maxItems": 40}}}
DRAFT_SCHEMA = {"type": "object", "additionalProperties": False, "required": ["angles"],
                "properties": {"angles": {"type": "array", "items": _ANGLE, "minItems": 1, "maxItems": 3}}}
REVIEW_SCHEMA = {"type": "object", "additionalProperties": False, "required": ["issues"],
                 "properties": {"issues": {"type": "array", "maxItems": 40,
                    "items": {"type": "object", "additionalProperties": False,
                              "required": ["code", "severity", "message", "path"],
                              "properties": {"code": {"type": "string", "maxLength": 80},
                                             "severity": {"type": "string", "enum": ["error", "warning"]},
                                             "message": {"type": "string", "maxLength": 1500},
                                             "path": {"type": "string", "maxLength": 200}}}}}}


class WritingError(ValueError):
    pass


def _hash(value) -> str:
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def effective_limits(channel: str, supplied: dict | None = None) -> dict:
    if channel not in DEFAULT_LIMITS:
        raise WritingError("Choose email or LinkedIn.")
    supplied = supplied or {}
    if not isinstance(supplied, dict) or set(supplied) - set(DEFAULT_LIMITS[channel]):
        raise WritingError("One or more message limits are not supported for this channel.")
    limits = dict(DEFAULT_LIMITS[channel])
    for key, value in supplied.items():
        maximum = 300 if key == "connection" else 3000 if channel == "linkedin" else 10000
        if type(value) is not int or value < 20 or value > maximum:
            raise WritingError(f"The {key} limit must be between 20 and {maximum} characters.")
        limits[key] = value
    return limits


def _audience(request: dict) -> dict:
    audience = request.get("audience") or {}
    return {"subsector": str(audience.get("subsector", audience.get("sector", ""))).strip(),
            "geography": str(audience.get("geography", "")).strip(),
            "company_size": str(audience.get("company_size", audience.get("headcount", ""))).strip(),
            "role": str(audience.get("role", "")).strip(), "offer": str(audience.get("offer", "")).strip()}


def _facts(brief: dict) -> list[dict]:
    today = datetime.now(timezone.utc).date().isoformat()
    return [f for f in brief.get("content", {}).get("facts", []) if f.get("decision") != "exclude"
            and (not f.get("valid_until") or str(f["valid_until"])[:10] >= today)
            and (not f.get("valid_from") or str(f["valid_from"])[:10] <= today)]


def _active_constraints(brief: dict) -> list[str]:
    content = brief.get("content", {})
    values = [str(v) for v in content.get("constraints", [])]
    values += [str(f.get("text", "")) for f in _facts(brief) if f.get("category") in {"constraint", "exclusion"}]
    for key, label in (("excluded_sectors", "Do not target"), ("excluded_audiences", "Do not target"),
                       ("forbidden_names", "Do not name"), ("forbidden_phrases", "Do not use this phrase")):
        if isinstance(content.get(key), list):
            values.extend(f"{label}: {value}" for value in content[key] if isinstance(value, str) and value.strip())
    return list(dict.fromkeys(v for v in values if v.strip()))


def _permission_guard(brief: dict, client_id: str):
    if brief.get("client_id") != client_id or brief.get("kind") != "brief" or brief.get("status") != "approved":
        raise WritingError("Use an approved brief belonging to the selected client.")
    content = brief.get("content", {})
    if content.get("client_id") not in (None, "", client_id):
        raise WritingError("The brief content belongs to a different client.")
    from .briefs import find_contradictions
    active = _facts(brief)
    active_ids = {f.get("id") for f in active}
    recorded_conflicts = [c for c in content.get("contradictions", [])
                          if len(active_ids.intersection(c.get("fact_ids", []))) > 1]
    if find_contradictions(active) or recorded_conflicts:
        raise WritingError("Resolve contradictory onboarding statements before generating copy.")
    if not any(f.get("category") == "audience" for f in _facts(brief)):
        raise WritingError("The approved brief needs audience facts before generating copy.")
    if not any(f.get("category") == "offer" for f in _facts(brief)):
        raise WritingError("The approved brief needs an offer before generating copy.")


def _source(c):
    return c.get("source") or c.get("provider") or ""


def performance_evidence(example: dict, performance: list[dict] | None) -> list[dict]:
    """Join to the latest exact sent-version observation, never an A/B letter alone."""
    from .performance import refresh_evidence

    content = example.get("content", {})
    if content.get("channel") != "email" or not content.get("original_sent_source_key"):
        return []
    def ordering(record):
        raw = record.get("content", {}).get("observed_at") or record.get("created_at", "")
        try:
            stamp = datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
            value = stamp.timestamp() if stamp.tzinfo else float("-inf")
        except (TypeError, ValueError, OverflowError):
            value = float("-inf")
        return (value, record.get("created_at", ""), record.get("id", ""))
    latest = {}
    for record in performance or []:
        if record.get("kind") != "performance" or record.get("client_id") != example.get("client_id"):
            continue
        c = record.get("content", {})
        # Report source keys include their observation window. Rank the latest
        # window per physical comparison unit, not an older favourable window.
        key = (c.get("source"), str(c.get("campaign_id")), str(c.get("step")), str(c.get("sequence_id", "")))
        if key not in latest or ordering(record) > ordering(latest[key]):
            latest[key] = record
    matches = []
    for record in latest.values():
        report = refresh_evidence(record.get("content", {}))
        if report.get("channel") != "email" or str(report.get("campaign_id")) != str(content.get("campaign_id")):
            continue
        if content.get("step") is not None and str(content["step"]) != str(report.get("step")):
            continue
        if content.get("sequence_id") is not None and str(content["sequence_id"]) != str(report.get("sequence_id")):
            continue
        for variant in report.get("variants", []):
            if content["original_sent_source_key"] not in variant.get("copy_source_keys", []):
                continue
            expected = content.get("variant_id")
            if expected is not None and variant.get("variant_id") is not None and str(expected) != str(variant["variant_id"]):
                continue
            supported = (variant.get("reference_eligible") is True and variant.get("attribution") == "exact"
                         and report.get("decided_on") == "positive" and not report.get("stale")
                         and report.get("coverage_complete") is True and variant.get("sent", 0) > 0)
            matches.append({"record_id": record["id"], "version": record.get("content_hash"),
                            "campaign_id": report["campaign_id"], "step": report.get("step"),
                            "variant_id": variant.get("variant_id", variant.get("variant")),
                            "basis": report.get("basis"), "period": report.get("period"),
                            "observed_at": report.get("observed_at"), "sent": variant.get("sent"),
                            "positive": variant.get("positive"), "reply": variant.get("reply"),
                            "evidence_level": variant.get("evidence_level"), "eligible_positive_leader": supported,
                            "attribution": variant.get("attribution"), "limitations": variant.get("limitations", []),
                            "stale": bool(report.get("stale")), "statistical_significance": "not_tested",
                            "interpretation": "Historical results belong to the exact original sent version; this reviewed reusable wording is not proven to reproduce them."})
    return matches


def retrieve_examples(client_id: str, channel: str, records: list[dict], audience: dict | None = None,
                      *, limit: int = 6, performance: list[dict] | None = None) -> list[dict]:
    """Filter before ranking; no client fallback or winner-only reference bias.

    Shared structure requires an explicit agency record, sanitisation and approval.
    Recipient-specific sent_copy rows are never examples, regardless of status.
    """
    query = set(_TOKENS.findall(json.dumps(audience or {}).lower()))
    candidates = []
    for record in records:
        c = record.get("content", {})
        shared = record.get("client_id") == "__agency__" and c.get("shared_structure") is True and c.get("sanitized") is True
        if record.get("client_id") != client_id and not shared:
            continue
        if record.get("kind") != "example" or record.get("status") != "approved" or c.get("channel") != channel:
            continue
        if c.get("test") or c.get("is_test") or c.get("draft") or c.get("obsolete") or c.get("stale"):
            continue
        if channel == "linkedin" and _source(c) != "copy_bank":
            continue
        if _source(c) != "copy_bank" and c.get("reusable") is not True:
            continue
        if not _steps(c) or not any(s["body"].strip() for s in _steps(c)):
            continue
        text = json.dumps({"audience": c.get("audience"), "offer": c.get("offer"), "step": c.get("step"),
                           "industry": c.get("industry"), "territory": c.get("territory")}).lower()
        score = len(query & set(_TOKENS.findall(text)))
        # A matched reference is useful even if it lost. Metrics are displayed
        # separately; raw reply-rate ranking is deliberately not a quality score.
        evidence = performance_evidence(record, performance) if channel == "email" else []
        if evidence:
            record = {**record, "retrieval_evidence": evidence}
        supported = any(e["eligible_positive_leader"] for e in evidence)
        candidates.append((score, supported, record))
    candidates.sort(key=lambda item: (-item[0], -item[1], item[2].get("id", "")))
    output, fingerprints = [], set()
    for _, _, record in candidates:
        fp = _hash({"subjects": record["content"].get("subjects", []), "steps": _steps(record["content"])})
        if fp in fingerprints:
            continue
        fingerprints.add(fp)
        output.append(record)
        if len(output) >= limit:
            break
    return output


def _steps(content: dict) -> list[dict]:
    steps = content.get("steps") or []
    if not steps and content.get("bodies"):
        steps = [{"kind": "initial" if i == 0 else "follow_up", "body": body}
                 for i, body in enumerate(content["bodies"])]
    if not steps and content.get("body"):
        steps = [{"kind": "initial", "body": content["body"]}]
    result = []
    for i, step in enumerate(steps):
        if isinstance(step, str):
            step = {"body": step}
        if isinstance(step, dict) and isinstance(step.get("body"), str):
            kind = step.get("kind") or ("initial" if i == 0 else "follow_up")
            result.append({"kind": kind, "body": step["body"], "evidence_ids": list(step.get("evidence_ids") or [])})
    return result


def _strings(draft):
    for ai, angle in enumerate(draft.get("angles", [])):
        for si, subject in enumerate(angle.get("subjects", [])):
            yield f"angles[{ai}].subjects[{si}]", "subject", subject, angle.get("evidence_ids", [])
        for si, step in enumerate(angle.get("steps", [])):
            yield f"angles[{ai}].steps[{si}].body", step.get("kind"), step.get("body", ""), step.get("evidence_ids", [])


def _review(issues: list[dict]) -> dict:
    errors = {i["code"] for i in issues if i["severity"] == "error"}
    return {"issues": issues, "can_approve": not errors, "human_review_required": True,
            "checks": {"schema": not bool(errors & {"schema", "channel", "empty_copy", "angle_count", "duplicate_angles", "incomplete_sequence"}),
                       "client_scope": "client_scope" not in errors,
                       "evidence": not bool(errors & {"unknown_evidence", "unsupported_proof", "unsupported_number", "forbidden_name", "unsupported_claim"}),
                       "channel_limits": not bool(errors & {"channel_limits", "channel", "variables"}),
                       "constraints": not bool(errors & {"excluded_audience", "forbidden_phrase", "unsupported_travel", "fake_prior_contact", "contradiction"})}}


def _affirmative_support(category: str, evidence: list[dict]) -> bool:
    """A negative permission/contact/travel statement cannot license its opposite."""
    return any(f.get("category") == category and not _DENIED_SUPPORT.search(f.get("text", ""))
               for f in evidence)


def validate_draft(client_id: str, brief: dict, examples: list[dict], request: dict, draft: dict) -> dict:
    issues = []

    def issue(code, message, path="", severity="error"):
        issues.append({"code": code, "severity": severity, "message": message, "path": path})

    try:
        _permission_guard(brief, client_id)
    except WritingError as exc:
        issue("client_scope" if brief.get("client_id") != client_id else "contradiction", str(exc))
    channel = request.get("channel", draft.get("channel", "email"))
    try:
        limits = effective_limits(channel, request.get("limits", draft.get("provenance", {}).get("limits")))
    except WritingError as exc:
        issue("channel", str(exc))
        return _review(issues)
    try:
        validate_json({"angles": draft.get("angles")}, DRAFT_SCHEMA)
    except ProviderError as exc:
        issue("schema", str(exc))
        return _review(issues)
    facts = {f["id"]: f for f in _facts(brief) if f.get("id")}
    content = brief.get("content", {})
    audience = _audience(request)
    target = " ".join(audience.values()).casefold()
    exclusions = content.get("excluded_sectors", []) + content.get("excluded_audiences", [])
    # Structured exclusion values are enforceable before a model call. The
    # complete free-text constraints are always supplied to independent review.
    for fact in facts.values():
        if fact.get("category") == "exclusion":
            raw = fact.get("value") or re.sub(r"^(?:exclude|avoid|do not target|don't target|not targeting)\s+", "", fact["text"], flags=re.I)
            exclusions.append(raw.rstrip("."))
    for excluded in exclusions:
        if len(str(excluded).strip()) >= 3 and str(excluded).casefold().strip() in target:
            issue("excluded_audience", f"This audience conflicts with a client exclusion: {excluded}.", "audience")
    forbidden = [str(v) for v in content.get("forbidden_phrases", []) if str(v).strip()]
    forbidden_names = [str(v) for v in content.get("forbidden_names", []) if str(v).strip()]
    allowed_vars = ALLOWED_VARIABLES | set(content.get("allowed_variables", []))
    selected = retrieve_examples(client_id, channel, examples, audience)
    if channel == "linkedin" and not selected:
        issue("linkedin_reference", "Approve a LinkedIn Copy Bank example before generating LinkedIn copy.")
    if channel == "linkedin" and any(a["subjects"] for a in draft["angles"]):
        issue("channel", "LinkedIn sequences do not have email subject lines.")
    if channel == "email" and any(s["kind"] == "connection" for a in draft["angles"] for s in a["steps"]):
        issue("channel", "Connection requests belong to LinkedIn, not email.")
    mode = {"assembly": "assemble", "new_draft": "draft"}.get(request.get("mode"), request.get("mode", "draft"))
    if mode == "draft":
        expected = {"initial", "follow_up", "closing"} | ({"connection"} if channel == "linkedin" else set())
        for ai, angle in enumerate(draft["angles"]):
            missing = expected - {s["kind"] for s in angle["steps"]}
            if missing:
                issue("incomplete_sequence", "The sequence needs these steps: " + ", ".join(sorted(missing)) + ".", f"angles[{ai}].steps")
            if channel == "email" and len(angle["subjects"]) < 2:
                issue("incomplete_sequence", "Provide at least two email subject variants.", f"angles[{ai}].subjects")
    if len({a["name"].casefold().strip() for a in draft["angles"]}) < len(draft["angles"]):
        issue("duplicate_angles", "Give each strategic angle a distinct name.")
    if len({_hash(a["steps"]) for a in draft["angles"]}) < len(draft["angles"]):
        issue("duplicate_angles", "The strategic angles repeat the same sequence.")
    for ai, angle in enumerate(draft["angles"]):
        references = list(angle["evidence_ids"])
        for claim in angle["claims"]:
            references.extend(claim["fact_ids"])
            if not claim["fact_ids"]:
                issue("unsupported_claim", "A factual claim has no approved brief evidence.", f"angles[{ai}].claims")
        references.extend(f for step in angle["steps"] for f in step["evidence_ids"])
        if any(ref not in facts for ref in references):
            issue("unknown_evidence", "Copy refers to a fact outside this approved client brief.", f"angles[{ai}]")
    for path, kind, text, ids in _strings(draft):
        if not text.strip():
            issue("empty_copy", "A message is empty.", path)
        if kind not in limits or len(text) > limits.get(kind, 0):
            issue("channel_limits", f"This {kind} exceeds its character limit.", path)
        names = _VARIABLE.findall(text)
        if any(name not in allowed_vars for name in names) or "{{" in _VARIABLE.sub("", text) or "}}" in _VARIABLE.sub("", text):
            issue("variables", "Use valid, approved merge variables only.", path)
        if any(v.casefold() in text.casefold() for v in forbidden):
            issue("forbidden_phrase", "The copy includes a prohibited phrase.", path)
        if any(re.search(r"(?<!\w)" + re.escape(v) + r"(?!\w)", text, re.I) for v in forbidden_names):
            issue("forbidden_name", "The copy names a client or organisation without permission.", path)
        evidence = [facts[f] for f in ids if f in facts]
        categories = {f["category"] for f in evidence}
        if _PRIOR_CONTACT.search(text) and not _affirmative_support("prior_contact", evidence):
            issue("fake_prior_contact", "Prior contact is claimed without supporting client evidence.", path)
        if _TRAVEL.search(text) and not _affirmative_support("travel", evidence):
            issue("unsupported_travel", "Travel or attendance is claimed without current approved evidence.", path)
        if _PROOF.search(text) and not ("proof" in categories and _affirmative_support("permission", evidence)):
            issue("unsupported_proof", "Results or named-client proof needs both approved proof and permission.", path)
        supported_text = " ".join(f["text"] + " " + " ".join(e.get("quote", "") for e in f.get("evidence", [])) for f in evidence)
        for claim in _NUMERIC_CLAIM.findall(text):
            if claim.casefold() not in supported_text.casefold():
                issue("unsupported_number", "A result or numeric claim is not supported by the cited facts.", path)
            if not _affirmative_support("permission", evidence):
                issue("unsupported_proof", "Using a numeric client result requires approved permission, not a prohibited or unresolved permission statement.", path)
        if re.search(r"\[(?:fact_|source_|\d+\])", text) or "app.read.ai/analytics/meetings/" in text:
            issue("evidence_in_copy", "Keep citations in the evidence panel, outside outreach text.", path)
    if not content.get("coverage", {}).get("complete"):
        issue("incomplete_sources", "Source-call completeness has not been confirmed; review the coverage before using this brief.", severity="warning")
    if not any(f.get("category") == "proof" for f in facts.values()):
        issue("missing_proof", "No approved proof is available. The copy must use the offer without inventing achievements.", severity="warning")
    if channel == "linkedin":
        issue("linkedin_unmeasured", "LinkedIn examples are approved writing references, not performance-proven examples.", severity="warning")
    reviewed_hash = draft.get("provenance", {}).get("model_reviewed_content_hash")
    model_review_current = bool(reviewed_hash and reviewed_hash == _hash({"angles": draft["angles"]}))
    if model_review_current:
        issues.extend(i for i in draft.get("review", {}).get("issues", []) if i.get("origin") == "model_review")
    elif reviewed_hash:
        issue("model_review_stale", "The automated semantic review predates these edits. Run a fresh copy check before approving this version.")
    result = _review(issues)
    result["checks"]["model_review_current"] = model_review_current
    return result


def review_draft(client_id: str, brief: dict, examples: list[dict], request: dict, draft: dict, *,
                 provider=None, methodology=None, performance=None) -> dict:
    """Recheck an edited immutable version without rewriting its outreach text."""
    _permission_guard(brief, client_id)
    result = deepcopy(draft)
    request = deepcopy(request)
    channel = request.get("channel", result.get("channel", "email"))
    limits = effective_limits(channel, request.get("limits"))
    request.update(channel=channel, audience=_audience(request), limits=limits)
    validate_json({"angles": result.get("angles")}, DRAFT_SCHEMA)
    # An earlier review is historical evidence, not a current validation result.
    result.pop("review", None)
    provenance = result.setdefault("provenance", {})
    provenance.pop("model_reviewed_content_hash", None)
    deterministic = validate_draft(client_id, brief, examples, request, result)
    mode = {"assembly": "assemble", "new_draft": "draft"}.get(request.get("mode"), request.get("mode", "draft"))
    if mode == "assemble":
        result["review"] = deterministic
        result["review"]["checks"]["model_review_current"] = False
        provenance["model_reviewed_content_hash"] = None
        provenance["review_provenance"] = {"provider": "none", "model": "deterministic-approved-assembly",
                                           "reviewed_at": datetime.now(timezone.utc).isoformat()}
        return result
    selected = retrieve_examples(client_id, channel, examples, request["audience"], performance=performance)
    if channel == "linkedin" and not selected:
        raise WritingError("LinkedIn requires an approved LinkedIn example from Copy Bank.")
    method = approved_methodology(methodology if isinstance(methodology, list) else [methodology] if methodology else [])
    context = {"client_id": client_id, "brief_id": brief["id"], "facts": _facts(brief),
               "as_of": datetime.now(timezone.utc).date().isoformat(),
               "mandatory_constraints": _active_constraints(brief), "audience": request["audience"],
               "resolved_fact_history": [{"key": c.get("key"), "resolution": c.get("resolution", "")}
                                         for c in brief.get("content", {}).get("contradictions", [])],
               "channel": channel, "limits": limits, "methodology": method,
               "references": [{"id": r["id"], "source": _source(r["content"]), "channel": channel,
                               "subjects": r["content"].get("subjects", []), "steps": _steps(r["content"]),
                               "structure_only": r.get("client_id") != client_id,
                               "performance_evidence": r.get("retrieval_evidence", [])} for r in selected],
               "draft": {"angles": result["angles"]}}
    system = SYSTEM_RULES + "\nIndependently review this edited copy without rewriting it. " \
        "Find unsupported/exaggerated claims, prohibited audiences/names, missing permission, fake prior contact, " \
        "unsupported travel, leaked client facts, citations in outreach and poor channel fit. " \
        "Check actual claim meaning against exact evidence, not merely that a fact ID is present. " \
        "Use errors for factual/constraint violations, warnings for style and weak angle diversity. " \
        "Reference copy is not client evidence. Do not claim statistical significance."
    provider = provider or get_provider()
    critique = provider.complete_json(system, json.dumps(context, ensure_ascii=False), schema=REVIEW_SCHEMA,
                                      schema_name="copy_studio_review", max_tokens=4000)
    validate_json(critique, REVIEW_SCHEMA)
    result["review"] = _review(deterministic["issues"] + [{**issue, "origin": "model_review"} for issue in critique["issues"]])
    result["review"]["checks"]["model_review_current"] = True
    provenance["model_reviewed_content_hash"] = _hash({"angles": result["angles"]})
    provenance["limits"] = limits
    provenance["review_provenance"] = {"provider": provider.name, "model": provider.model,
                                       "prompt_hash": _hash({"system": system, "context": context}),
                                       "source_versions": {r["id"]: r.get("content_hash") or _hash(r.get("content")) for r in [brief] + selected},
                                       "reviewed_at": datetime.now(timezone.utc).isoformat()}
    return result


def _assembly(template: dict, brief: dict, request: dict) -> dict:
    c = template["content"]
    variables = request.get("variables") or {}
    if not isinstance(variables, dict) or set(variables) - ALLOWED_VARIABLES:
        raise WritingError("Assembly only substitutes approved identity merge variables.")
    for key, value in variables.items():
        if not isinstance(value, str) or len(value) > 120 or any(ch in value for ch in ("\n", "\r", "{", "}")):
            raise WritingError("Assembly variable values must be short, single-line identity values.")
    facts = _facts(brief)
    # Attribution is inherited only when template facts refer to this exact
    # approved brief. Otherwise the reviewer must ground its claims afresh.
    valid = {f["id"] for f in facts}
    evidence = [f for f in c.get("evidence_ids", []) if f in valid]

    def substitute(text):
        return _VARIABLE.sub(lambda m: variables.get(m.group(1), m.group(0)), text)

    steps = [{**step, "body": substitute(step["body"]),
              "evidence_ids": [f for f in (step["evidence_ids"] or evidence) if f in valid]}
             for step in _steps(c)]
    return {"angles": [{"name": c.get("title") or "Approved wording", "rationale": "Assembled from approved wording; only requested identity variables were substituted.",
                        "subjects": [substitute(s) for s in c.get("subjects", [])] if request["channel"] == "email" else [],
                        "steps": steps, "evidence_ids": evidence, "claims": deepcopy(c.get("claims", []))}]}


def generate(client_id: str, brief: dict, examples: list[dict], request: dict, *, provider=None,
             methodology=None, performance=None) -> dict:
    _permission_guard(brief, client_id)
    request = deepcopy(request)
    channel = request.get("channel", "email")
    limits = effective_limits(channel, request.get("limits"))
    mode = {"assembly": "assemble", "new_draft": "draft"}.get(request.get("mode"), request.get("mode", "draft"))
    if mode not in {"assemble", "draft"}:
        raise WritingError("Choose approved-template assembly or a new draft.")
    count = request.get("angles", 3)
    if type(count) is not int or not 1 <= count <= 3:
        raise WritingError("Choose between one and three strategic angles.")
    request.update(channel=channel, mode=mode, audience=_audience(request), limits=limits)
    candidates = retrieve_examples(client_id, channel, examples, request["audience"], performance=performance)
    if channel == "linkedin" and not candidates:
        raise WritingError("LinkedIn requires an approved LinkedIn example from Copy Bank.")
    method = approved_methodology(methodology if isinstance(methodology, list) else [methodology] if methodology else [])
    facts = _facts(brief)
    context = {"client_id": client_id, "brief_id": brief["id"], "facts": facts,
               "as_of": datetime.now(timezone.utc).date().isoformat(),
               "mandatory_constraints": _active_constraints(brief), "audience": request["audience"],
               "resolved_fact_history": [{"key": c.get("key"), "resolution": c.get("resolution", "")}
                                         for c in brief.get("content", {}).get("contradictions", [])],
               "objective": str(request.get("objective", ""))[:2000], "channel": channel, "limits": limits,
               "angle_count": count, "methodology": method,
               "required_sequence": {"steps": ["connection", "initial", "follow_up", "closing"] if channel == "linkedin"
                                     else ["initial", "follow_up", "closing"],
                                     "subject_variants": 0 if channel == "linkedin" else 2},
               "references": [{"id": r["id"], "source": _source(r["content"]), "channel": channel,
                               "subjects": r["content"].get("subjects", []), "steps": _steps(r["content"]),
                               "structure_only": r.get("client_id") != client_id,
                               "performance_evidence": r.get("retrieval_evidence", [])} for r in candidates]}
    # Preflight exclusions before spending model tokens. The empty angle is
    # schema-valid solely so the common audience checks run.
    shell = {"angles": [{"name": "Preflight", "rationale": "", "subjects": [],
                         "steps": [{"kind": "initial", "body": "Review pending", "evidence_ids": []}],
                         "evidence_ids": [], "claims": []}]}
    preflight = validate_draft(client_id, brief, examples, request, shell)
    excluded = [i for i in preflight["issues"] if i["code"] == "excluded_audience"]
    if excluded:
        raise WritingError(excluded[0]["message"])
    model_review = []
    if mode == "assemble":
        wanted = request.get("template_id") or (request.get("example_ids") or [None])[0]
        template = next((r for r in candidates if r["id"] == wanted), None)
        if template is None and len(candidates) == 1:
            template = candidates[0]
        if not template:
            raise WritingError("Choose one approved template belonging to this client and channel.")
        draft = _assembly(template, brief, request)
        candidates = [template]
        model_name, provider_name = "deterministic-approved-assembly", "none"
    else:
        provider = provider or get_provider()
        draft = provider.complete_json(SYSTEM_RULES,
            json.dumps(context, ensure_ascii=False), schema=DRAFT_SCHEMA, schema_name="copy_studio_draft", max_tokens=9000)
        validate_json(draft, DRAFT_SCHEMA)
        if len(draft["angles"]) != count:
            raise WritingError("The provider returned the wrong number of requested angles.")
        model_name, provider_name = provider.model, provider.name
        review_context = {**context, "draft": draft}
        critique = provider.complete_json(SYSTEM_RULES + "\nReview the proposed copy independently. Treat its claims as untrusted. "
            "Find unsupported or exaggerated claims, forbidden target audiences, permission problems, fake prior contact, "
            "stale travel claims, client leakage, citations in outreach, poor channel fit and angles that are not strategically distinct. "
            "References cannot supply factual claims. Return errors for claim/constraint violations and warnings for stylistic weaknesses. "
            "A fact ID alone does not prove a claim: compare the actual text and exact evidence. Do not assert statistical significance.",
            json.dumps(review_context, ensure_ascii=False), schema=REVIEW_SCHEMA, schema_name="copy_studio_review", max_tokens=4000)
        validate_json(critique, REVIEW_SCHEMA)
        model_review = critique["issues"]
    draft.update(title="LinkedIn copy" if channel == "linkedin" else "Email copy", channel=channel, mode=mode)
    result = validate_draft(client_id, brief, examples, request, draft)
    if model_review:
        result = _review(result["issues"] + [{**i, "origin": "model_review"} for i in model_review])
    result["checks"]["model_review_current"] = mode == "draft"
    if not method:
        result["issues"].append({"code": "methodology_pending", "severity": "warning", "message": "The approved methodology corpus is not available. Fixed writing guardrails were used; human methodology review is pending.", "path": ""})
    draft.update(request=request, review=result,
                 provenance={"pipeline_version": VERSION, "provider": provider_name, "model": model_name,
                             "prompt_hash": _hash({"system": SYSTEM_RULES, "context": context}),
                             "source_versions": {r["id"]: r.get("content_hash") or _hash(r.get("content")) for r in [brief] + candidates},
                             "performance_versions": {e["record_id"]: e["version"] for r in candidates for e in r.get("retrieval_evidence", [])},
                             "brief_id": brief["id"], "reference_ids": [r["id"] for r in candidates],
                             "methodology_versions": {r["id"]: r.get("version") for r in method},
                             "limits": limits, "mandatory_constraints": context["mandatory_constraints"],
                             "model_reviewed_content_hash": _hash({"angles": draft["angles"]}) if mode == "draft" else None,
                             "created_at": datetime.now(timezone.utc).isoformat(),
                             "human_review_required": True, "quality_evaluation": "Blind human review pending"},
                 evidence=[{"fact_id": f["id"], "text": f["text"], "citations": f.get("evidence", [])} for f in facts])
    return draft
