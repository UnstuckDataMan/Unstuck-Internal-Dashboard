"""Persistence invariants: tenant isolation, immutable history, resumable leases."""
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
import json
import sqlite3

import pytest

from app.utils.copy_studio import store
from tests.conftest import FakeResponse

AGENCY = "11111111-1111-1111-1111-111111111111"
OTHER_AGENCY = "22222222-2222-2222-2222-222222222222"


@pytest.fixture
def db(tmp_path):
    return store.SQLiteStore(str(tmp_path / "studio.sqlite"), AGENCY)


def save(db, **changes):
    args = dict(kind="brief", client_id="client-a", source_key="onboarding:one",
                content={"offer": "Accounting", "proof": []}, actor="staff@example.test")
    args.update(changes)
    return db.save(**args)


def enqueue(db, **changes):
    args = dict(client_id="client-a", kind="sync", payload={"source": "smartlead"},
                actor="staff@example.test", idempotency_key="sync:2026-10-28")
    args.update(changes)
    return db.enqueue(**args)


def alter_job(db, job_id, **changes):
    with sqlite3.connect(db.path) as conn:
        conn.execute("UPDATE copy_studio_jobs SET " + ",".join(f"{k}=?" for k in changes) + " WHERE id=?", (*changes.values(), job_id))


def test_records_are_idempotent_versions_not_mutable_slots(db):
    first = save(db)
    again = save(db, actor="second@example.test")
    assert again == first
    assert first["status"] == "pending"
    changed = save(db, content={"offer": "Tax", "proof": []}, parent_id=first["id"])
    assert changed["id"] != first["id"]
    assert changed["version"] != first["version"]
    assert changed["parent_id"] == first["id"]
    assert db.get(first["id"], "client-a")["content"]["offer"] == "Accounting"
    assert len(db.list("client-a", "brief")) == 2
    assert db.get(first["id"], "client-a")["created_by"] == "staff@example.test"


def test_content_mutation_after_save_does_not_change_database(db):
    content = {"facts": ["One"]}
    first = save(db, content=content)
    content["facts"].append("Two")
    assert db.get(first["id"], "client-a")["content"] == {"facts": ["One"]}


def test_reviews_are_audit_events_and_do_not_change_the_source(db):
    record = save(db)
    approved = db.review(record["id"], "client-a", "reviewer@example.test", "approved", "Verified against transcript")
    assert approved["status"] == "approved"
    assert approved["review"]["reviewed_by"] == "reviewer@example.test"
    assert db.list("client-a", status="approved")[0]["id"] == record["id"]
    revised = db.review(record["id"], "client-a", "reviewer@example.test", "needs_changes", "Evidence superseded")
    assert revised["status"] == "needs_changes"
    assert revised["content"] == record["content"]
    assert db.list("client-a", status="approved") == []
    with sqlite3.connect(db.path) as conn:
        assert conn.execute("SELECT COUNT(*) FROM copy_studio_reviews").fetchone()[0] == 2


@pytest.mark.parametrize("table", ["copy_studio_records", "copy_studio_reviews"])
@pytest.mark.parametrize("verb", ["UPDATE", "DELETE"])
def test_history_is_immutable_even_with_direct_database_writes(db, table, verb):
    record = save(db)
    db.review(record["id"], "client-a", "reviewer@example.test", "approved")
    with sqlite3.connect(db.path) as conn, pytest.raises(sqlite3.IntegrityError, match="immutable"):
        statement = f"UPDATE {table} SET created_at='2020-01-01'" if verb == "UPDATE" else f"DELETE FROM {table}"
        conn.execute(statement)


def test_agency_and_client_isolation_for_reads_reviews_and_parent_links(db):
    record = save(db)
    other = store.SQLiteStore(db.path, OTHER_AGENCY)
    assert other.get(record["id"], "client-a") is None
    assert db.get(record["id"], "client-b") is None
    assert other.list("client-a") == []
    assert db.list("client-b") == []
    for target, client in [(other, "client-a"), (db, "client-b")]:
        with pytest.raises(store.StoreError, match="not available"):
            target.review(record["id"], client, "reviewer@example.test", "approved")
        with pytest.raises(store.StoreError, match="Parent"):
            save(target, client_id=client, parent_id=record["id"])
    assert save(other)["id"] != record["id"]
    assert save(db, client_id="client-b")["id"] != record["id"]


def test_reserved_scopes_are_separate_from_client_data(db):
    global_record = save(db, client_id="__agency__", kind="methodology")
    assert db.list("client-a") == []
    assert db.get(global_record["id"], "__agency__") is not None
    assert db.get(global_record["id"], "__workspace__") is None


def test_sql_like_client_ids_are_bound_parameters(db):
    save(db)
    assert db.list("client-a' OR 1=1 --") == []


def test_non_json_and_nonfinite_data_rejected(db):
    for value in ({"n": float("nan")}, {"n": object()}, []):
        with pytest.raises(store.StoreError):
            save(db, content=value)


def test_jobs_idempotent_and_key_cannot_change_request(db):
    job = enqueue(db)
    assert enqueue(db)["id"] == job["id"]
    with pytest.raises(store.StoreError, match="different job"):
        enqueue(db, payload={"source": "different"})
    assert len(db.list_jobs("client-a")) == 1
    assert db.get_job(job["id"], "client-b") is None
    assert store.SQLiteStore(db.path, OTHER_AGENCY).get_job(job["id"], "client-a") is None
    assert store.SQLiteStore(db.path, OTHER_AGENCY).claim_job("other-worker") is None


@pytest.mark.parametrize('status', ['queued', 'running'])
@pytest.mark.parametrize('source', ['context', 'onboarding', 'all'])
def test_active_context_refresh_cannot_hide_behind_newer_completed_jobs(db, status, source):
    from app.utils.copy_studio import service
    active = enqueue(db, payload={'source': source}, idempotency_key='older-context')
    alter_job(db, active['id'], status=status, created_at='2025-01-01T00:00:00+00:00')
    for i in range(35):
        newer = enqueue(db, kind='generate', payload={}, idempotency_key=f'newer:{i}')
        alter_job(db, newer['id'], status='succeeded')
    assert len(db.list_jobs('client-a')) == 30
    assert active['id'] not in {job['id'] for job in db.list_jobs('client-a')}
    assert db.context_refresh_active('client-a')
    assert 'refreshing' in service.brief_context_error(db, 'client-a', {'content': {}})
    assert not db.context_refresh_active('client-b')
    assert not store.SQLiteStore(db.path, OTHER_AGENCY).context_refresh_active('client-a')
    alter_job(db, active['id'], status='succeeded')
    assert not db.context_refresh_active('client-a')


def test_unrelated_or_finished_jobs_do_not_block_context(db):
    enqueue(db, payload={'source': 'copy_bank'}, idempotency_key='copy-bank')
    enqueue(db, payload={'source': 'smartlead'}, idempotency_key='smartlead')
    enqueue(db, kind='generate', payload={'source': 'context'}, idempotency_key='generation')
    completed = enqueue(db, payload={'source': 'all'}, idempotency_key='done')
    failed = enqueue(db, payload={'source': 'onboarding'}, idempotency_key='failed')
    alter_job(db, completed['id'], status='succeeded')
    alter_job(db, failed['id'], status='failed')
    assert not db.context_refresh_active('client-a')


def test_supabase_active_context_query_filters_before_limiting_and_remains_scoped(monkeypatch):
    requests = []
    responses = iter([[{'id': 'active-context'}], []])
    def get(url, **kwargs):
        requests.append((url, kwargs))
        return FakeResponse(200, next(responses))
    monkeypatch.setattr(store.requests, 'get', get)
    remote = store.SupabaseStore('https://db.test', 'test-service-only', AGENCY)
    assert remote.context_refresh_active('client-a')
    assert not remote.context_refresh_active('client-b')
    for client_id, (url, request) in zip(('client-a', 'client-b'), requests):
        assert url.endswith('/copy_studio_jobs')
        assert request['params'] == {'agency_id': f'eq.{AGENCY}', 'client_id': f'eq.{client_id}',
            'select': 'id', 'kind': 'eq.sync', 'status': 'in.(queued,running)', 'limit': 1,
            'payload->>source': 'in.(context,onboarding,all)'}


def test_concurrent_workers_cannot_claim_same_live_lease(db):
    job = enqueue(db)
    with ThreadPoolExecutor(max_workers=8) as pool:
        claimed = list(pool.map(lambda i: db.claim_job(f"worker-{i}"), range(8)))
    successes = [j for j in claimed if j is not None]
    assert len(successes) == 1
    assert successes[0]["id"] == job["id"]
    assert successes[0]["attempts"] == 1


def test_checkpoint_and_expired_lease_survive_new_store_process(db):
    job = enqueue(db)
    db.claim_job("worker-one")
    db.checkpoint_job(job["id"], "worker-one", {"campaign": "123", "page": 7})
    alter_job(db, job["id"], lease_expires_at="2000-01-01T00:00:00+00:00")
    resumed = store.SQLiteStore(db.path, AGENCY).claim_job("worker-two")
    assert resumed["checkpoint"] == {"campaign": "123", "page": 7}
    assert resumed["attempts"] == 2
    with pytest.raises(store.LeaseLost):
        db.complete_job(job["id"], "worker-one", {"wrong": True})
    db.complete_job(job["id"], "worker-two", {"records": 5})
    complete = db.get_job(job["id"], "client-a")
    assert complete["status"] == "succeeded"
    assert complete["result"] == {"records": 5}
    assert complete["worker_id"] is None
    assert db.claim_job("worker-three") is None


@pytest.mark.parametrize("operation", ["checkpoint_job", "heartbeat_job", "complete_job", "fail_job"])
def test_wrong_worker_cannot_modify_a_job(db, operation):
    job = enqueue(db)
    db.claim_job("owner")
    extras = {"checkpoint_job": ({"p": 1},), "heartbeat_job": (), "complete_job": ({"ok": True},), "fail_job": ("failed",)}
    with pytest.raises(store.LeaseLost):
        getattr(db, operation)(job["id"], "other", *extras[operation])


def test_retry_backoff_exhaustion_and_checkpoints(db):
    job = enqueue(db, max_attempts=2)
    db.claim_job("one")
    db.checkpoint_job(job["id"], "one", {"page": 4})
    db.fail_job(job["id"], "one", "Provider rate limit")
    queued = db.get_job(job["id"], "client-a")
    assert queued["status"] == "queued"
    assert queued["checkpoint"] == {"page": 4}
    assert datetime.fromisoformat(queued["run_after"]) > datetime.now(timezone.utc)
    assert db.claim_job("two") is None
    alter_job(db, job["id"], run_after="2000-01-01T00:00:00+00:00")
    assert db.claim_job("two")["attempts"] == 2
    db.fail_job(job["id"], "two", "Provider still unavailable")
    assert db.get_job(job["id"], "client-a")["status"] == "failed"
    assert db.claim_job("three") is None


def test_expired_final_attempt_does_not_remain_running_forever(db):
    job = enqueue(db, max_attempts=1)
    db.claim_job("one")
    alter_job(db, job["id"], lease_expires_at="2000-01-01T00:00:00+00:00")
    assert db.claim_job("two") is None
    assert db.get_job(job["id"], "client-a")["status"] == "failed"


def test_heartbeat_extends_live_lease_and_nonretry_failure_is_terminal(db):
    job = enqueue(db)
    claimed = db.claim_job("one", lease_seconds=10)
    db.heartbeat_job(job["id"], "one", lease_seconds=300)
    assert db.get_job(job["id"], "client-a")["lease_expires_at"] > claimed["lease_expires_at"]
    db.fail_job(job["id"], "one", "Review required", retry=False)
    assert db.get_job(job["id"], "client-a")["status"] == "failed"


def test_factory_requires_explicit_configuration_never_falls_back(monkeypatch, tmp_path):
    from app.utils import supabase
    for key in ("COPY_STUDIO_SQLITE_PATH", "COPY_STUDIO_AGENCY_ID", "PULSE_AGENCY_ID", "SUPABASE_SERVICE_ROLE_KEY"):
        monkeypatch.delenv(key, raising=False)
    monkeypatch.setattr(supabase, "SUPABASE_SERVICE_ROLE_KEY", "")
    with pytest.raises(store.StoreNotReady):
        store.get_store()
    monkeypatch.setenv("COPY_STUDIO_SQLITE_PATH", str(tmp_path / "explicit.sqlite"))
    assert isinstance(store.get_store(), store.SQLiteStore)
    monkeypatch.delenv("COPY_STUDIO_SQLITE_PATH")
    monkeypatch.setenv("SUPABASE_SERVICE_ROLE_KEY", "test-service")
    monkeypatch.setenv("COPY_STUDIO_AGENCY_ID", AGENCY)
    assert isinstance(store.get_store(), store.SupabaseStore)


def test_supabase_reads_are_service_authenticated_and_client_scoped(monkeypatch):
    calls = []
    def get(url, **kwargs):
        calls.append((url, kwargs))
        return FakeResponse(200, [])
    monkeypatch.setattr(store.requests, "get", get)
    db = store.SupabaseStore("https://db.test", "test-service-only", AGENCY)
    assert db.get("record", "client-a") is None
    assert db.list("client-a", status="approved") == []
    assert db.get_job("job", "client-a") is None
    assert db.list_jobs("client-a") == []
    for _, call in calls:
        assert call["headers"]["apikey"] == "test-service-only"
        assert call["params"]["agency_id"] == f"eq.{AGENCY}"
        assert call["params"]["client_id"] == "eq.client-a"
    assert calls[1][1]["params"]["status"] == "eq.approved"


def test_supabase_job_leases_use_atomic_rpc_and_guard_false(monkeypatch):
    calls = []
    def post(url, **kwargs):
        calls.append((url, kwargs))
        return FakeResponse(200, [] if url.endswith("claim_job") else False)
    monkeypatch.setattr(store.requests, "post", post)
    db = store.SupabaseStore("https://db.test", "test-service-only", AGENCY)
    assert db.claim_job("worker") is None
    with pytest.raises(store.LeaseLost):
        db.complete_job("job", "worker", {"ok": True})
    assert calls[0][0].endswith("/rpc/copy_studio_claim_job")
    assert calls[0][1]["json"]["p_agency_id"] == AGENCY
    assert calls[1][1]["json"]["p_worker_id"] == "worker"


def test_supabase_errors_do_not_expose_provider_bodies_or_secrets(monkeypatch):
    monkeypatch.setattr(store.requests, "get", lambda *a, **k: FakeResponse(403, {"message": "secret token and personal data"}))
    db = store.SupabaseStore("https://db.test", "test-service-only", AGENCY)
    with pytest.raises(store.StoreNotReady) as error:
        db.get("record", "client-a")
    assert "personal data" not in str(error.value)
    assert "test-service-only" not in str(error.value)


def test_sqlite_file_is_private(db):
    from pathlib import Path
    assert Path(db.path).stat().st_mode & 0o777 == 0o600


def test_keyset_scan_reads_more_than_one_thousand_versions(db):
    # Force identical timestamps to exercise the ID tie-breaker on every page.
    from unittest.mock import patch
    with patch.object(store, "now_iso", return_value="2026-10-28T12:00:00+00:00"):
        originals = [save(db, content={"revision": i}) for i in range(1013)]
    found = list(db.iter_records("client-a", kind="brief", page_size=73))
    assert len(found) == 1013
    assert len({r["id"] for r in found}) == 1013
    assert {r["id"] for r in found} == {r["id"] for r in originals}
    assert len(db.list("client-a", limit=1000)) == 1000


def test_new_versions_between_pages_do_not_shift_older_results(db):
    originals = [save(db, content={"revision": i}) for i in range(7)]
    first = db.list_page("client-a", limit=3)
    assert first["has_more"]
    newest = save(db, content={"revision": "new after first page"})
    second = db.list_page("client-a", limit=3, cursor=first["next_cursor"])
    third = db.list_page("client-a", limit=3, cursor=second["next_cursor"])
    assert third["has_more"] is False
    assert third["next_cursor"] is None
    all_ids = [r["id"] for page in (first, second, third) for r in page["records"]]
    assert len(all_ids) == len(set(all_ids)) == 7
    assert set(all_ids) == {r["id"] for r in originals}
    assert newest["id"] not in all_ids


def test_cursor_cannot_change_client_agency_or_filter(db):
    for i in range(3):
        save(db, content={"revision": i})
    first = db.list_page("client-a", limit=1)
    cursor = first["next_cursor"]
    with pytest.raises(store.StoreError, match="cursor"):
        db.list_page("client-b", cursor=cursor)
    with pytest.raises(store.StoreError, match="cursor"):
        db.list_page("client-a", kind="brief", cursor=cursor)
    with pytest.raises(store.StoreError, match="cursor"):
        store.SQLiteStore(db.path, OTHER_AGENCY).list_page("client-a", cursor=cursor)
    with pytest.raises(store.StoreError, match="cursor"):
        db.list_page("client-a", status="approved", cursor=cursor)
    with pytest.raises(store.StoreError, match="cursor"):
        db.list_page("client-a", cursor="not a valid cursor!")


def test_supabase_keyset_page_uses_safe_scoped_filter(monkeypatch):
    records = [
        {"id": "00000000-0000-0000-0000-000000000003", "created_at": "2026-10-28T12:00:00+00:00"},
        {"id": "00000000-0000-0000-0000-000000000002", "created_at": "2026-10-28T12:00:00+00:00"},
        {"id": "00000000-0000-0000-0000-000000000001", "created_at": "2026-10-28T12:00:00+00:00"},
    ]
    calls = []
    def get(url, **kwargs):
        calls.append(kwargs)
        return FakeResponse(200, records if len(calls) == 1 else records[-1:])
    monkeypatch.setattr(store.requests, "get", get)
    db = store.SupabaseStore("https://db.test", "test-service", AGENCY)
    first = db.list_page("client-a", limit=2)
    second = db.list_page("client-a", limit=2, cursor=first["next_cursor"])
    assert first["has_more"] and len(first["records"]) == 2
    assert not second["has_more"] and len(second["records"]) == 1
    assert calls[0]["params"]["limit"] == 3
    assert calls[1]["params"]["agency_id"] == f"eq.{AGENCY}"
    assert calls[1]["params"]["client_id"] == "eq.client-a"
    assert "id.lt.00000000-0000-0000-0000-000000000002" in calls[1]["params"]["or"]
