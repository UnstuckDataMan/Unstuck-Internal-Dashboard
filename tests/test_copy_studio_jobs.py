"""Durable ingestion behaviour with offline source fixtures, no provider calls."""
import pytest

from app.utils.copy_studio import sources, worker
from app.utils.copy_studio.integrations import IntegrationUnavailable
from app.utils.copy_studio.store import SQLiteStore, LeaseLost


@pytest.fixture
def job_store(tmp_path, monkeypatch):
    monkeypatch.setattr(sources, "reconcile_smartlead", lambda client_id: {
        "client_mappings_missing_from_provider": [], "provider_without_agency_mapping_count": 0})
    return SQLiteStore(str(tmp_path / "studio.db"), agency_id="agency-1")


def claim(store, source="manual"):
    row = store.enqueue("client-1", "sync", {"source": source}, "reviewer@example.invalid", "request-1")
    return store.claim_job("worker-1")


def result(key="source-1", *, complete=True, cursor=None):
    return {"items": [{"kind": "example", "client_id": "client-1", "source_key": key,
                       "content": {"source": "fixture", "body": "Fixture text"}}],
            "complete": complete, "next_cursor": cursor, "coverage": {"complete": complete}, "warnings": []}


def listing(*ids):
    return {"campaigns": [{"id": identifier} for identifier in ids], "complete": True, "next_offset": None}


def test_failed_sheet_does_not_block_remaining_campaigns(job_store, monkeypatch):
    job = claim(job_store)
    monkeypatch.setattr(sources, "list_campaigns", lambda *a, **k: listing("missing-share", "readable"))
    def sync(client_id, source, campaign, **kwargs):
        if campaign["id"] == "missing-share":
            raise IntegrationUnavailable("Share the campaign sheet as Viewer.")
        return result("readable")
    monkeypatch.setattr(sources, "sync_campaign", sync)
    outcome = worker._sync(job_store, job, "worker-1")
    assert outcome["complete"] is False
    report = outcome["sources"][0]
    assert report["status"] == "partial"
    assert "last_successful_sync_at" not in report
    assert len(report["coverage"]["campaigns"]) == 2
    assert len(job_store.list("client-1", kind="example")) == 1
    assert job_store.list("client-1", kind="sync_report")[0]["content"]["status"] == "partial"


def test_all_source_job_continues_after_missing_credentials(job_store, monkeypatch):
    job = claim(job_store, "all")
    def fail(*a, **k):
        raise IntegrationUnavailable("This source needs its server credential.")
    monkeypatch.setattr(sources, "sync_copy_bank", fail)
    monkeypatch.setattr(sources, "sync_onboarding", lambda client_id: result("onboarding"))
    def campaigns(client_id, source, **kwargs):
        if source == "smartlead":
            raise IntegrationUnavailable("Smartlead is not configured.")
        return listing("manual-working")
    monkeypatch.setattr(sources, "list_campaigns", campaigns)
    monkeypatch.setattr(sources, "sync_campaign", lambda *a, **k: result("manual-working"))
    outcome = worker._sync(job_store, job, "worker-1")
    assert not outcome["complete"]
    assert {r["source"]: r["status"] for r in outcome["sources"]} == {
        "copy_bank": "partial", "onboarding": "complete", "smartlead": "partial", "manual": "complete"}
    assert len(job_store.list("client-1", kind="example")) == 2


def test_checkpoint_resume_keeps_page_identity_and_previously_saved_versions(job_store, monkeypatch):
    job = claim(job_store, "smartlead")
    calls, pages = [], []
    def campaigns(*args, **kwargs):
        pages.append(kwargs)
        return listing("first", "second")
    monkeypatch.setattr(sources, "list_campaigns", campaigns)
    def sync(client_id, source, campaign, cursor=None, **kwargs):
        calls.append((campaign["id"], cursor))
        if cursor is None:
            return result("first-page", complete=False, cursor={"offset": 100})
        raise SystemExit("worker interrupted")
    monkeypatch.setattr(sources, "sync_campaign", sync)
    with pytest.raises(SystemExit):
        worker._sync(job_store, job, "worker-1")
    saved = job_store.get_job(job["id"], "client-1")
    assert saved["checkpoint"]["campaigns"]["cursor"] == {"offset": 100}
    assert len(job_store.list("client-1", kind="example")) == 1
    def resumed(client_id, source, campaign, cursor=None, **kwargs):
        calls.append((campaign["id"], cursor))
        return result(campaign["id"]+"-done")
    monkeypatch.setattr(sources, "sync_campaign", resumed)
    outcome = worker._sync(job_store, saved, "worker-1")
    assert outcome["complete"]
    assert len(pages) == 1  # no mutable re-enumeration at the saved offset
    assert ("first", {"offset": 100}) in calls
    assert ("second", None) in calls
    assert len(job_store.list("client-1", kind="example")) == 3


def test_nonadvancing_source_cursor_marks_partial_and_continues(job_store, monkeypatch):
    job = claim(job_store, "smartlead")
    monkeypatch.setattr(sources, "list_campaigns", lambda *a, **k: listing("bad", "good"))
    calls = []
    def sync(client_id, source, campaign, **kwargs):
        calls.append(campaign["id"])
        return result("bad", complete=False, cursor={"offset": 100}) if campaign["id"] == "bad" else result("good")
    monkeypatch.setattr(sources, "sync_campaign", sync)
    outcome = worker._sync(job_store, job, "worker-1")
    assert not outcome["complete"]
    assert calls == ["bad", "bad", "good"]
    assert len(job_store.list("client-1", kind="example")) == 2  # duplicate replay is immutable/idempotent


def test_missing_continuation_and_stuck_listing_are_not_complete(job_store, monkeypatch):
    job = claim(job_store, "smartlead")
    monkeypatch.setattr(sources, "list_campaigns", lambda *a, **k: {**listing("bad"), "complete": False, "next_offset": 0})
    monkeypatch.setattr(sources, "sync_campaign", lambda *a, **k: result("partial", complete=False, cursor=None))
    outcome = worker._sync(job_store, job, "worker-1")
    assert not outcome["complete"]
    assert len(outcome["sources"][0]["coverage"]["campaigns"]) == 2


def test_lease_loss_after_network_prevents_source_writes(job_store, monkeypatch):
    job = claim(job_store, "copy_bank")
    original = job_store.heartbeat_job
    def source(client_id):
        monkeypatch.setattr(job_store, "heartbeat_job", lambda *a, **k: (_ for _ in ()).throw(LeaseLost("Lease lost")))
        return result()
    monkeypatch.setattr(sources, "sync_copy_bank", source)
    with pytest.raises(LeaseLost):
        worker._sync(job_store, job, "worker-1")
    assert job_store.list("client-1", kind="example") == []


def test_foreign_client_source_is_refused_not_logged_as_a_partial_success(job_store, monkeypatch):
    job = claim(job_store, "copy_bank")
    bad = result()
    bad["items"][0]["client_id"] = "client-2"
    monkeypatch.setattr(sources, "sync_copy_bank", lambda *a: bad)
    with pytest.raises(ValueError, match="different client"):
        worker._sync(job_store, job, "worker-1")
    assert job_store.list("client-1", kind="example") == []


def test_missing_campaign_reconciliation_is_visible(job_store, monkeypatch):
    job = claim(job_store, "smartlead")
    monkeypatch.setattr(sources, "list_campaigns", lambda *a, **k: listing())
    monkeypatch.setattr(sources, "reconcile_smartlead", lambda client_id: {
        "client_mappings_missing_from_provider": ["missing-123"], "provider_without_agency_mapping_count": 3})
    outcome = worker._sync(job_store, job, "worker-1")
    assert not outcome["complete"]
    assert outcome["sources"][0]["reconciliation"]["client_mappings_missing_from_provider"] == ["missing-123"]
    assert len(outcome["sources"][0]["warnings"]) == 2
