from copy import deepcopy
import json

import pytest

from app.utils.copy_studio import proofreading, writer


def brief(*facts):
    return {"id": "brief-a", "client_id": "a", "kind": "brief", "status": "approved",
            "content": {"facts": [
                {"id": "audience", "category": "audience", "text": "UK finance teams", "evidence": []},
                {"id": "offer", "category": "offer", "text": "Simplify reporting", "evidence": []}, *facts],
                        "coverage": {"complete": True}}}


def preference(text, quote=None, **extra):
    return {"id": "language", "category": "tone", "text": text, "origin": "extracted",
            "evidence": [{"source_id": "call", "source_version": "v1", "block": 0, "quote": quote or text}], **extra}


def draft(text="Could we organise a call?"):
    return {"channel": "email", "mode": "draft", "request": {"locale": "en-US"},
            "angles": [{"name": "Reporting", "rationale": "Reduce admin", "subjects": ["Reporting", "Less admin"],
                        "steps": [{"kind": "initial", "body": text, "evidence_ids": ["offer"]},
                                  {"kind": "follow_up", "body": "Could simpler reporting help?", "evidence_ids": ["offer"]},
                                  {"kind": "closing", "body": "Happy to leave this here.", "evidence_ids": []}],
                        "evidence_ids": ["offer"], "claims": []}]}


def correction(original="organise", replacement="organize", start=9, **extra):
    return {"category": "locale", "severity": "warning", "message": "Use the selected US spelling.",
            "path": "angles[0].steps[0].body", "original": original, "replacement": replacement, "start": start, **extra}


class Model:
    name, model = "fixture", "fixture-proofreader"

    def __init__(self, issues=None):
        self.issues, self.calls = issues or [], []

    def complete_json(self, system, user, **kwargs):
        self.calls.append({"system": system, "context": json.loads(user), **kwargs})
        return {"issues": deepcopy(self.issues)}


@pytest.mark.parametrize("phrase,locale", [("Please use UK English.", "en-GB"),
    ("Use British spelling in our copy.", "en-GB"), ("We prefer American English.", "en-US"),
    ("Use US spelling.", "en-US"), ("Don't use American English; use British English.", "en-GB")])
def test_explicit_preference_is_grounded_in_approved_quote(phrase, locale):
    out = proofreading.infer_locale(brief(preference(phrase)))
    assert out["locale"] == locale and out["status"] == "explicit"
    assert out["evidence"][0]["quote"] == phrase


@pytest.mark.parametrize("text", ["We don't want American English.", "We could use American English, but haven't decided."])
def test_negated_or_undecided_language_is_not_a_confirmed_preference(text):
    assert proofreading.infer_locale(brief(preference(text)))["status"] == "tentative"


def test_geography_is_only_a_tentative_hint_not_approved_language_preference():
    out = proofreading.infer_locale(brief(preference("Use American English", "We target the UK.")))
    assert out["locale"] == "en-GB" and out["status"] == "tentative"
    pending = brief(preference("Use American English"))
    pending["status"] = "pending"
    assert proofreading.infer_locale(pending)["status"] == "unset"
    b = brief()
    b["content"]["facts"][0]["text"] = "Businesses that want to work with us."
    assert proofreading.infer_locale(b)["status"] == "unset"


def test_conflicting_preferences_require_manual_selection_and_excluded_fact_is_ignored():
    b = brief(preference("Use British English"), preference("Use American English", id="us-language"))
    assert proofreading.infer_locale(b)["status"] == "ambiguous"
    assert proofreading.infer_locale(b)["locale"] is None
    assert proofreading.selected_locale(b, "en-US") == "en-US"
    b["content"]["facts"][-1]["decision"] = "exclude"
    assert proofreading.infer_locale(b)["locale"] == "en-GB"


def test_verified_transcript_versions_and_quotes_are_checked_when_supplied():
    b = brief(preference("Use US English"))
    source = {"id": "call", "kind": "transcript", "client_id": "a", "content_hash": "v1",
              "content": {"transcript": [{"text": "Use US English"}]}}
    assert proofreading.infer_locale(b, [source])["status"] == "explicit"
    source["content_hash"] = "v2"
    assert proofreading.infer_locale(b, [source])["status"] == "tentative"
    source["content_hash"] = "v1"
    source["client_id"] = "other-client"
    assert proofreading.infer_locale(b, [source])["status"] == "tentative"


def test_interpretation_and_expired_facts_do_not_become_language_instructions():
    b = brief(preference("Use US English", interpretation=True),
              preference("Use US English", id="expired", valid_until="2000-01-01"))
    assert proofreading.infer_locale(b)["status"] == "tentative"
    b["content"]["facts"] += [preference("Use US English", id="human", origin="human", edited_by="reviewer", evidence=[])]
    assert proofreading.infer_locale(b)["status"] == "explicit"


def test_check_returns_exact_suggestions_without_changing_original_copy():
    d = draft()
    before = deepcopy(d)
    model = Model([correction()])
    report = proofreading.check(d, "en-US", model)
    assert d == before
    assert report["complete"] is True and report["copy_hash"] == proofreading.copy_hash(d)
    assert report["issues"][0]["replacement"] == "organize"
    assert report["issues"][0]["start"] == 9 and report["issues"][0]["end"] == 17
    assert report["issues"][0]["applicable"] is True
    assert len(model.calls[0]["context"]["texts"]) == 5
    assert "not a rewrite or a source of factual approval" in model.calls[0]["system"]


@pytest.mark.parametrize("locale,original,replacement", [("en-US", "organise", "organize"), ("en-GB", "organize", "organise")])
def test_both_english_locales_reach_provider(locale, original, replacement):
    model = Model([correction(original=original, replacement=replacement)])
    report = proofreading.check(draft("Could we " + original + " a call?"), locale, model)
    assert report["locale"] == locale
    assert model.calls[0]["context"]["locale"] == locale


@pytest.mark.parametrize("text,original,replacement", [
    ("Hi {{first_name}}, hello.", "{{first_name}}", "{{firstname}}"),
    ("We are ColourCo.", "ColourCo", "ColorCo"),
    ("Meet Acme Labs.", "Acme Labs", "Acme Lab"),
    ("Contact help@colour.co.uk.", "help@colour.co.uk", "help@color.co.uk"),
    ("Read https://colour.co.uk.", "https://colour.co.uk", "https://color.co.uk"),
    ("We generated 10 leads.", "10 leads", "100 leads"),
    ("We support reporting.", "support reporting", "generated outstanding results"),
])
def test_protected_values_and_new_claims_cannot_be_changed_by_spelling(text, original, replacement):
    report = proofreading.check(draft(text), "en-US", Model([correction(original, replacement, text.index(original))]))
    assert report["issues"][0]["severity"] == "error"
    assert report["issues"][0]["applicable"] is False
    assert report["issues"][0]["replacement"] == ""


def test_incorrect_offsets_and_non_outreach_paths_are_blocked():
    report = proofreading.check(draft(), "en-US", Model([correction(start=0), correction(path="request.objective")]))
    assert all(i["severity"] == "error" for i in report["issues"])
    assert all(not i["applicable"] for i in report["issues"])


def test_unicode_offsets_use_code_points_and_subjects_are_checked():
    d = draft("🙂 organise this")
    report = proofreading.check(d, "en-US", Model([correction(start=2)]))
    assert report["issues"][0]["end"] == 10
    assert report["issues"][0]["applicable"] is True


def checked_draft():
    d = draft()
    d["review"] = {"can_approve": True, "issues": []}
    d["provenance"] = {"model_reviewed_content_hash": writer._hash({"angles": d["angles"]})}
    report = proofreading.check(d, "en-US", Model())
    d["final_check"] = {"complete": True, "locale": "en-US", "copy_hash": report["copy_hash"],
                        "checked_angles": deepcopy(d["angles"]), "report": report}
    return d


def test_final_check_binds_exact_immutable_copy_locale_semantic_review_and_snapshot():
    d = checked_draft()
    assert proofreading.final_check_current(d) is True
    changes = [lambda x: x["angles"][0]["steps"][0].update(body="An edited message"),
               lambda x: x["request"].update(locale="en-GB"),
               lambda x: x["provenance"].update(model_reviewed_content_hash="stale"),
               lambda x: x["final_check"]["report"].update(complete=False),
               lambda x: x["final_check"].update(checked_angles=[]),
               lambda x: x["final_check"]["report"].update(copy_hash="stale"),
               lambda x: x["review"].update(can_approve=False)]
    for mutate in changes:
        edited = deepcopy(d)
        mutate(edited)
        assert proofreading.final_check_current(edited) is False


def test_warning_suggestions_can_be_acknowledged_but_blockers_cannot():
    d = checked_draft()
    d["final_check"]["report"]["issues"] = [{"severity": "warning", "message": "Confirm spelling."}]
    assert proofreading.final_check_current(d) is True
    d["final_check"]["report"]["issues"][0]["severity"] = "error"
    assert proofreading.final_check_current(d) is False


def test_edited_assembly_needs_semantic_review_as_well_as_spelling():
    d = checked_draft()
    d.update(mode="assemble", human_edited=True)
    d["provenance"]["model_reviewed_content_hash"] = None
    assert proofreading.final_check_current(d) is False


def test_invalid_locale_is_rejected_before_model_call():
    model = Model()
    with pytest.raises(ValueError, match="UK or US"):
        proofreading.check(draft(), "auto", model)
    assert model.calls == []
