from copy import deepcopy
import json

import pytest

from app.utils.copy_studio import briefs
from app.utils.copy_studio.providers import ProviderError


def source(text=None, **content):
    return {"id": "transcript-a", "client_id": "client-a", "kind": "transcript", "content_hash": "hash-a",
            "content": {"transcript": [{"text": text or "We target software companies.", "start_ms": 1200}],
                        "source_url": "https://app.read.ai/analytics/meetings/example", **content}}


class Extractor:
    name, model = "fixture", "fixture-extractor"

    def __init__(self, fn=None):
        self.calls = []
        self.fn = fn

    def complete_json(self, system, user, **kwargs):
        data = json.loads(user)
        self.calls.append(data)
        if self.fn:
            return self.fn(data)
        block = data["blocks"][0]
        return {"facts": [{"category": "audience", "key": "target", "text": block["text"].strip(),
                           "block": block["i"], "quote": block["text"][:30], "interpretation": False}]}


def test_complete_call_includes_tail_and_oversized_single_block():
    text = "A" * 25000 + "Do not target gambling businesses."
    provider = Extractor(lambda data: {"facts": [{"category": "exclusion", "key": "gambling", "text": "Do not target gambling businesses.",
                 "block": data["blocks"][-1]["i"], "quote": "Do not target gambling businesses.", "interpretation": False}]}
                 if "gambling" in data["blocks"][-1]["text"] else {"facts": []})
    result = briefs.extract_brief("client-a", [source(text, call_complete=True)], provider=provider)
    assert len(provider.calls) == 3
    assert "".join(b["text"] for c in provider.calls for b in c["blocks"]) == text
    assert result["coverage"]["complete"] is True
    assert result["constraints"] == ["Do not target gambling businesses."]
    evidence = result["facts"][0]["evidence"][0]
    assert evidence["char_start"] == 25000
    assert evidence["start_ms"] == 1200
    assert text[evidence["char_start"]:evidence["char_end"]] == evidence["quote"]


def test_unknown_upstream_completeness_is_not_full_call_complete():
    result = briefs.extract_brief("client-a", [source(source_complete=True)], provider=Extractor())
    assert result["coverage"]["stored_text_complete"] is True
    assert result["coverage"]["complete"] is False
    assert result["coverage"]["sources"][0]["call_complete"] is False
    reviewed = briefs.extract_brief("client-a", [source(source_complete=True, source_completeness_review={"actor": "reviewer"})], provider=Extractor())
    assert reviewed["coverage"]["complete"] is True


def test_partial_model_failure_retains_checkpoint_and_resume_processes_tail_only():
    transcript = source("A" * 12000 + "B" * 12000 + "C" * 300, call_complete=True)
    checkpoints = []
    provider = Extractor(lambda data: (_ for _ in ()).throw(ProviderError("Retry later."))
                         if data["section"] == 2 else {"facts": []})
    result = briefs.extract_brief("client-a", [transcript], provider=provider, progress=checkpoints.append)
    assert result["coverage"]["complete"] is False
    assert result["coverage"]["processed_chunks"] == 1
    assert result["errors"][0]["chunk"] == 1
    resumed = Extractor(lambda _: {"facts": []})
    final = briefs.extract_brief("client-a", [transcript], provider=resumed, checkpoint=checkpoints[-1])
    assert [c["section"] for c in resumed.calls] == [2, 3]
    assert final["coverage"]["complete"] is True
    transcript["content_hash"] = "changed"
    with pytest.raises(briefs.BriefError, match="versions changed"):
        briefs.extract_brief("client-a", [transcript], provider=resumed, checkpoint=checkpoints[-1])


def test_hallucinated_quote_is_rejected_and_coverage_never_called_complete():
    provider = Extractor(lambda _: {"facts": [{"category": "proof", "key": "result", "text": "We doubled revenue.",
          "block": 0, "quote": "We doubled revenue.", "interpretation": False}]})
    result = briefs.extract_brief("client-a", [source(call_complete=True)], provider=provider)
    assert result["facts"] == []
    assert result["errors"]
    assert result["coverage"]["complete"] is False


def test_resume_cannot_launder_rejected_quote_into_complete_coverage():
    transcript = source("A" * 12000 + "We do not target gambling businesses.", call_complete=True)
    checkpoints = []
    def bad(data):
        if data["section"] == 1:
            return {"facts": []}
        return {"facts": [{"category": "exclusion", "key": "gambling", "text": "Exclude gambling businesses.",
                           "block": 0, "quote": "We never work with gaming.", "interpretation": False}]}
    result = briefs.extract_brief("client-a", [transcript], provider=Extractor(bad), progress=checkpoints.append)
    assert result["coverage"]["complete"] is False
    assert checkpoints[-1]["completed_chunk_ids"] == ["transcript-a:0"]
    retried = Extractor(bad)
    again = briefs.extract_brief("client-a", [transcript], provider=retried, checkpoint=checkpoints[-1])
    assert [c["section"] for c in retried.calls] == [2]
    assert again["coverage"]["complete"] is False
    repaired = Extractor(lambda _: {"facts": [{"category": "exclusion", "key": "gambling", "text": "Exclude gambling businesses.",
                       "block": 0, "quote": "We do not target gambling businesses.", "interpretation": False}]})
    final = briefs.extract_brief("client-a", [transcript], provider=repaired, checkpoint=checkpoints[-1])
    assert final["coverage"]["complete"] is True
    assert final["constraints"] == ["Exclude gambling businesses."]


def test_transcript_dedup_cross_client_denial_and_human_edit_preservation():
    original = source(call_complete=True)
    duplicate = deepcopy(original)
    duplicate["id"] = "duplicate"
    provider = Extractor()
    human = {"id": "human-fact", "category": "tone", "key": "voice", "text": "Keep it plain.",
             "origin": "human", "evidence": [], "edited_by": "reviewer"}
    result = briefs.extract_brief("client-a", [original, duplicate], provider=provider,
                                  previous={"id": "brief-before", "content": {"facts": [human]}})
    assert len(provider.calls) == 1
    assert human in result["facts"]
    assert result["supersedes"] == "brief-before"
    duplicate["client_id"] = "client-b"
    with pytest.raises(briefs.BriefError, match="selected client"):
        briefs.extract_brief("client-a", [duplicate], provider=provider)


def test_contradictions_remain_unresolved_for_review():
    facts = [{"id": "a", "category": "permission", "key": "name Acme", "text": "We may name Acme."},
             {"id": "b", "category": "permission", "key": "name Acme", "text": "We cannot name Acme."},
             {"id": "c", "category": "audience", "key": "target", "text": "We target gambling businesses."},
             {"id": "d", "category": "exclusion", "key": "gambling", "text": "Do not target gambling businesses."}]
    conflicts = briefs.find_contradictions(facts)
    assert len(conflicts) == 2
    assert all(c["resolution"] == "" for c in conflicts)
    assert {frozenset(c["fact_ids"]) for c in conflicts} == {frozenset(["a", "b"]), frozenset(["c", "d"])}
