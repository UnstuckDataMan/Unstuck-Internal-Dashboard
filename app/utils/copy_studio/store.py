"""Private, immutable Copy Studio records and a leased, durable work queue.

Production uses service-role PostgREST tables installed by ai_copy_studio.sql.
SQLite is an explicit local-development option, never a fallback when Supabase
fails. Both stores scope every operation to one agency and, for user-facing
operations, one client. HTTP handlers must additionally authorize that client
and derive ``actor`` from the authenticated session.
"""
from __future__ import annotations

from contextlib import contextmanager
import base64
from datetime import datetime, timedelta, timezone
import hashlib
import json
import os
from pathlib import Path
import sqlite3
import uuid

import requests


class StoreError(RuntimeError):
    pass


class StoreNotReady(StoreError):
    pass


class LeaseLost(StoreError):
    pass


DECISIONS = {"approved", "rejected", "needs_changes"}
_NAMESPACE = uuid.UUID("27dc068e-d044-4f4b-9e02-df67b941e202")


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _json(value: dict) -> str:
    if not isinstance(value, dict):
        raise StoreError("Stored content must be a JSON object.")
    try:
        result = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (ValueError, TypeError) as exc:
        raise StoreError("Content must contain valid JSON values.") from exc
    if len(result.encode()) > 16_000_000:
        raise StoreError("Record exceeds the 16 MB storage limit.")
    return result


def _hash(value: dict) -> str:
    return hashlib.sha256(_json(value).encode()).hexdigest()


def _text(value, name: str, maximum: int = 500) -> str:
    if not isinstance(value, str) or not value.strip() or len(value) > maximum or "\x00" in value:
        raise StoreError(f"A valid {name} is required.")
    return value.strip()


def _limit(value, maximum=1000):
    return max(1, min(maximum, int(value)))


def _page_cursor(agency, client, kind, status, record):
    payload = {"agency": agency, "client": client, "kind": kind, "status": status,
               "created_at": record["created_at"], "id": record["id"]}
    return base64.urlsafe_b64encode(_json(payload).encode()).decode().rstrip("=")


def _parse_cursor(cursor, agency, client, kind, status):
    if not cursor:
        return None
    try:
        if not isinstance(cursor, str) or len(cursor) > 4000:
            raise ValueError
        value = json.loads(base64.urlsafe_b64decode(cursor + "=" * (-len(cursor) % 4)))
        if {k: value[k] for k in ("agency", "client", "kind", "status")} != {
            "agency": agency, "client": client, "kind": kind, "status": status
        }:
            raise ValueError
        # Cursor values also enter PostgREST's logical filter grammar. Parse and
        # canonicalise them rather than interpolating arbitrary caller text.
        value["id"] = str(uuid.UUID(value["id"]))
        stamp = datetime.fromisoformat(value["created_at"])
        if stamp.tzinfo is None:
            raise ValueError
        value["created_at"] = stamp.astimezone(timezone.utc).isoformat()
        return value
    except (ValueError, TypeError, KeyError, UnicodeError) as exc:
        raise StoreError("Invalid record cursor or cursor scope.") from exc


class _RecordPagination:
    def list(self, client_id, kind=None, limit=200, status=None):
        # Compatibility helper for bounded UI views. Full scans use iter_records.
        wanted = _limit(limit)
        output = []
        for record in self.iter_records(client_id, kind=kind, status=status, page_size=min(wanted, 500)):
            output.append(record)
            if len(output) == wanted:
                break
        return output

    def iter_records(self, client_id, kind=None, status=None, page_size=200):
        """Read all matching versions with a stable keyset, without a 1,000-row cap.

        Newer versions arriving after the first page do not shift/skip older
        versions. This is not a transaction snapshot of changing review states.
        """
        cursor = None
        while True:
            page = self.list_page(client_id, kind=kind, limit=page_size, status=status, cursor=cursor)
            yield from page["records"]
            cursor = page["next_cursor"]
            if not cursor:
                return


def _record(agency_id, kind, client_id, source_key, content, actor, parent_id=None):
    kind, client_id = _text(kind, "record kind", 80), _text(client_id, "client ID", 200)
    source_key, actor = _text(source_key, "source key", 1000), _text(actor, "actor", 320)
    digest = _hash(content)
    identity = _json({"agency": agency_id, "client": client_id, "kind": kind, "source": source_key,
                      "content": digest, "parent": parent_id})
    return {"id": str(uuid.uuid5(_NAMESPACE, identity)), "agency_id": agency_id, "client_id": client_id,
            "kind": kind, "source_key": source_key, "content": content, "content_hash": digest,
            "version": digest, "parent_id": parent_id, "created_by": actor, "created_at": now_iso()}


def _job(agency_id, client_id, kind, payload, actor, idempotency_key, max_attempts):
    client_id, kind = _text(client_id, "client ID", 200), _text(kind, "job kind", 80)
    actor, key = _text(actor, "actor", 320), _text(idempotency_key, "idempotency key", 500)
    if not 1 <= int(max_attempts) <= 10:
        raise StoreError("Maximum job attempts must be between 1 and 10.")
    payload_hash = _hash(payload)
    identity = _json({"agency": agency_id, "client": client_id, "kind": kind, "key": key})
    timestamp = now_iso()
    return {"id": str(uuid.uuid5(_NAMESPACE, identity)), "agency_id": agency_id, "client_id": client_id,
            "kind": kind, "payload": payload, "payload_hash": payload_hash, "created_by": actor,
            "idempotency_key": key, "status": "queued", "attempts": 0, "max_attempts": int(max_attempts),
            "result": {}, "checkpoint": {}, "last_error": "", "worker_id": None,
            "lease_expires_at": None, "run_after": timestamp, "created_at": timestamp,
            "updated_at": timestamp, "completed_at": None}


def _review(agency, client_id, record_id, actor, decision, notes):
    if decision not in DECISIONS:
        raise StoreError("Review decision must be approved, rejected or needs_changes.")
    if not isinstance(notes, str) or len(notes) > 10000:
        raise StoreError("Review notes exceed the permitted length.")
    return {"id": str(uuid.uuid4()), "agency_id": agency, "client_id": client_id, "record_id": record_id,
            "decision": decision, "reviewed_by": _text(actor, "reviewer", 320), "notes": notes,
            "created_at": now_iso()}


_SQLITE_SCHEMA = """
CREATE TABLE IF NOT EXISTS copy_studio_records (
 id TEXT PRIMARY KEY, agency_id TEXT NOT NULL, client_id TEXT NOT NULL, kind TEXT NOT NULL,
 source_key TEXT NOT NULL, content TEXT NOT NULL, content_hash TEXT NOT NULL, version TEXT NOT NULL,
 parent_id TEXT, created_by TEXT NOT NULL, created_at TEXT NOT NULL,
 UNIQUE(agency_id, client_id, id),
 FOREIGN KEY(agency_id,client_id,parent_id) REFERENCES copy_studio_records(agency_id,client_id,id)
);
CREATE INDEX IF NOT EXISTS copy_studio_records_scope ON copy_studio_records(agency_id,client_id,kind,created_at);
CREATE TABLE IF NOT EXISTS copy_studio_reviews (
 id TEXT PRIMARY KEY, agency_id TEXT NOT NULL, client_id TEXT NOT NULL, record_id TEXT NOT NULL,
 decision TEXT NOT NULL CHECK(decision IN ('approved','rejected','needs_changes')),
 reviewed_by TEXT NOT NULL, notes TEXT NOT NULL, created_at TEXT NOT NULL,
 FOREIGN KEY(agency_id,client_id,record_id) REFERENCES copy_studio_records(agency_id,client_id,id)
);
CREATE INDEX IF NOT EXISTS copy_studio_reviews_record ON copy_studio_reviews(agency_id,client_id,record_id,created_at);
CREATE TABLE IF NOT EXISTS copy_studio_jobs (
 id TEXT PRIMARY KEY, agency_id TEXT NOT NULL, client_id TEXT NOT NULL, kind TEXT NOT NULL,
 payload TEXT NOT NULL, payload_hash TEXT NOT NULL, created_by TEXT NOT NULL, idempotency_key TEXT NOT NULL,
 status TEXT NOT NULL CHECK(status IN ('queued','running','succeeded','failed')),
 attempts INTEGER NOT NULL, max_attempts INTEGER NOT NULL, result TEXT NOT NULL,
 checkpoint TEXT NOT NULL, last_error TEXT NOT NULL, worker_id TEXT, lease_expires_at TEXT,
 run_after TEXT NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL, completed_at TEXT,
 UNIQUE(agency_id,client_id,kind,idempotency_key)
);
CREATE INDEX IF NOT EXISTS copy_studio_jobs_claim ON copy_studio_jobs(agency_id,status,run_after);
CREATE TRIGGER IF NOT EXISTS copy_studio_records_immutable_update BEFORE UPDATE ON copy_studio_records BEGIN SELECT RAISE(ABORT,'Copy Studio records are immutable'); END;
CREATE TRIGGER IF NOT EXISTS copy_studio_records_immutable_delete BEFORE DELETE ON copy_studio_records BEGIN SELECT RAISE(ABORT,'Copy Studio records are immutable'); END;
CREATE TRIGGER IF NOT EXISTS copy_studio_reviews_immutable_update BEFORE UPDATE ON copy_studio_reviews BEGIN SELECT RAISE(ABORT,'Copy Studio reviews are immutable'); END;
CREATE TRIGGER IF NOT EXISTS copy_studio_reviews_immutable_delete BEFORE DELETE ON copy_studio_reviews BEGIN SELECT RAISE(ABORT,'Copy Studio reviews are immutable'); END;
"""


class SQLiteStore(_RecordPagination):
    """Explicit persistent local store. Each operation gets its own connection."""
    def __init__(self, path: str, agency_id: str = "local"):
        self.agency_id = _text(agency_id, "agency ID", 200)
        if path == ":memory:":
            raise StoreNotReady("Use a temporary file for SQLite so jobs survive connection lifetimes.")
        self.path = str(Path(path).expanduser().resolve())
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        with self._db() as db:
            db.execute("PRAGMA journal_mode=WAL")
            db.executescript(_SQLITE_SCHEMA)
        os.chmod(self.path, 0o600)

    @contextmanager
    def _db(self, *, write=False):
        db = sqlite3.connect(self.path, timeout=15, isolation_level=None)
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys=ON")
        try:
            if write:
                db.execute("BEGIN IMMEDIATE")
            yield db
            if write:
                db.commit()
        except Exception:
            if write:
                db.rollback()
            raise
        finally:
            db.close()

    @staticmethod
    def _decode(row):
        if row is None:
            return None
        data = dict(row)
        for key in ("content", "payload", "result", "checkpoint"):
            if key in data:
                data[key] = json.loads(data[key])
        return data

    def _with_review(self, db, record):
        if record is None:
            return None
        review = db.execute("SELECT * FROM copy_studio_reviews WHERE agency_id=? AND client_id=? AND record_id=? ORDER BY created_at DESC,rowid DESC LIMIT 1",
                            (self.agency_id, record["client_id"], record["id"])).fetchone()
        record["status"] = review["decision"] if review else "pending"
        record["review"] = dict(review) if review else None
        return record

    def save(self, kind, client_id, source_key, content, actor, parent_id=None):
        record = _record(self.agency_id, kind, client_id, source_key, content, actor, parent_id)
        with self._db(write=True) as db:
            if parent_id and not db.execute("SELECT 1 FROM copy_studio_records WHERE id=? AND agency_id=? AND client_id=?",
                                           (parent_id, self.agency_id, client_id)).fetchone():
                raise StoreError("Parent record is not available to this client.")
            values = {**record, "content": _json(content)}
            db.execute(f"INSERT OR IGNORE INTO copy_studio_records ({','.join(values)}) VALUES ({','.join('?' for _ in values)})", tuple(values.values()))
        return self.get(record["id"], client_id)

    def get(self, record_id, client_id):
        with self._db() as db:
            row = db.execute("SELECT * FROM copy_studio_records WHERE id=? AND agency_id=? AND client_id=?",
                             (record_id, self.agency_id, _text(client_id, "client ID", 200))).fetchone()
            return self._with_review(db, self._decode(row))

    def list_page(self, client_id, kind=None, limit=200, status=None, cursor=None):
        if status is not None and status not in DECISIONS | {"pending"}:
            raise StoreError("Invalid record status.")
        limit = _limit(limit, 500)
        cursor_data = _parse_cursor(cursor, self.agency_id, client_id, kind, status)
        clauses, params = ["r.agency_id=?", "r.client_id=?"], [self.agency_id, _text(client_id, "client ID", 200)]
        if kind:
            clauses.append("r.kind=?")
            params.append(kind)
        if status:
            clauses.append("COALESCE((SELECT decision FROM copy_studio_reviews v WHERE v.agency_id=r.agency_id AND v.client_id=r.client_id AND v.record_id=r.id ORDER BY v.created_at DESC,v.rowid DESC LIMIT 1),'pending')=?")
            params.append(status)
        if cursor_data:
            clauses.append("(r.created_at < ? OR (r.created_at = ? AND r.id < ?))")
            params.extend((cursor_data["created_at"], cursor_data["created_at"], cursor_data["id"]))
        with self._db() as db:
            rows = db.execute("SELECT r.* FROM copy_studio_records r WHERE " + " AND ".join(clauses) + " ORDER BY r.created_at DESC,r.id DESC LIMIT ?",
                              (*params, limit + 1)).fetchall()
            more = len(rows) > limit
            records = [self._with_review(db, self._decode(row)) for row in rows[:limit]]
            next_cursor = _page_cursor(self.agency_id, client_id, kind, status, records[-1]) if more else None
            return {"records": records, "has_more": more, "next_cursor": next_cursor}

    def review(self, record_id, client_id, actor, decision, notes=""):
        review = _review(self.agency_id, client_id, record_id, actor, decision, notes)
        with self._db(write=True) as db:
            if not db.execute("SELECT 1 FROM copy_studio_records WHERE id=? AND agency_id=? AND client_id=?",
                              (record_id, self.agency_id, client_id)).fetchone():
                raise StoreError("Record is not available to this client.")
            db.execute(f"INSERT INTO copy_studio_reviews ({','.join(review)}) VALUES ({','.join('?' for _ in review)})", tuple(review.values()))
        return self.get(record_id, client_id)

    def enqueue(self, client_id, kind, payload, actor, idempotency_key, max_attempts=4):
        job = _job(self.agency_id, client_id, kind, payload, actor, idempotency_key, max_attempts)
        values = {**job, **{k: _json(job[k]) for k in ("payload", "result", "checkpoint")}}
        with self._db(write=True) as db:
            existing = db.execute("SELECT payload_hash FROM copy_studio_jobs WHERE id=? AND agency_id=? AND client_id=?", (job["id"], self.agency_id, client_id)).fetchone()
            if existing and existing["payload_hash"] != job["payload_hash"]:
                raise StoreError("This idempotency key already belongs to a different job request.")
            db.execute(f"INSERT OR IGNORE INTO copy_studio_jobs ({','.join(values)}) VALUES ({','.join('?' for _ in values)})", tuple(values.values()))
        return self.get_job(job["id"], client_id)

    def get_job(self, job_id, client_id):
        with self._db() as db:
            return self._decode(db.execute("SELECT * FROM copy_studio_jobs WHERE id=? AND agency_id=? AND client_id=?",
                                           (job_id, self.agency_id, _text(client_id, "client ID", 200))).fetchone())

    def list_jobs(self, client_id, limit=30):
        with self._db() as db:
            return [self._decode(row) for row in db.execute("SELECT * FROM copy_studio_jobs WHERE agency_id=? AND client_id=? ORDER BY created_at DESC,rowid DESC LIMIT ?",
                                                         (self.agency_id, _text(client_id, "client ID", 200), _limit(limit))).fetchall()]

    def context_refresh_active(self, client_id):
        """Publication guards must not depend on the bounded UI job history."""
        with self._db() as db:
            return db.execute(
                "SELECT 1 FROM copy_studio_jobs WHERE agency_id=? AND client_id=? "
                "AND kind='sync' AND status IN ('queued','running') "
                "AND json_extract(payload, '$.source') IN ('context','onboarding','all') LIMIT 1",
                (self.agency_id, _text(client_id, "client ID", 200))).fetchone() is not None

    def claim_job(self, worker_id, lease_seconds=300):
        worker_id = _text(worker_id, "worker ID", 200)
        lease_seconds = max(10, min(3600, int(lease_seconds)))
        timestamp = now_iso()
        expires = (datetime.now(timezone.utc) + timedelta(seconds=lease_seconds)).isoformat()
        with self._db(write=True) as db:
            db.execute("UPDATE copy_studio_jobs SET status='failed',worker_id=NULL,lease_expires_at=NULL,last_error='Worker lease expired; retry limit reached.',updated_at=?,completed_at=? WHERE agency_id=? AND status='running' AND lease_expires_at<=? AND attempts>=max_attempts",
                       (timestamp, timestamp, self.agency_id, timestamp))
            row = db.execute("SELECT id FROM copy_studio_jobs WHERE agency_id=? AND attempts<max_attempts AND ((status='queued' AND run_after<=?) OR (status='running' AND lease_expires_at<=?)) ORDER BY run_after,created_at,id LIMIT 1",
                             (self.agency_id, timestamp, timestamp)).fetchone()
            if row is None:
                return None
            db.execute("UPDATE copy_studio_jobs SET status='running',attempts=attempts+1,worker_id=?,lease_expires_at=?,updated_at=? WHERE id=? AND agency_id=?",
                       (worker_id, expires, timestamp, row["id"], self.agency_id))
            return self._decode(db.execute("SELECT * FROM copy_studio_jobs WHERE id=? AND agency_id=?", (row["id"], self.agency_id)).fetchone())

    def _owned(self, db, job_id, worker_id):
        row = db.execute("SELECT * FROM copy_studio_jobs WHERE id=? AND agency_id=? AND worker_id=? AND status='running' AND lease_expires_at>?",
                         (job_id, self.agency_id, worker_id, now_iso())).fetchone()
        if row is None:
            raise LeaseLost("Job lease expired or belongs to another worker.")
        return row

    def checkpoint_job(self, job_id, worker_id, checkpoint):
        with self._db(write=True) as db:
            self._owned(db, job_id, worker_id)
            db.execute("UPDATE copy_studio_jobs SET checkpoint=?,updated_at=? WHERE id=? AND agency_id=?", (_json(checkpoint), now_iso(), job_id, self.agency_id))

    def heartbeat_job(self, job_id, worker_id, lease_seconds=300):
        expires = (datetime.now(timezone.utc) + timedelta(seconds=max(10, min(3600, int(lease_seconds))))).isoformat()
        with self._db(write=True) as db:
            self._owned(db, job_id, worker_id)
            db.execute("UPDATE copy_studio_jobs SET lease_expires_at=?,updated_at=? WHERE id=? AND agency_id=?", (expires, now_iso(), job_id, self.agency_id))

    def complete_job(self, job_id, worker_id, result):
        with self._db(write=True) as db:
            self._owned(db, job_id, worker_id)
            timestamp = now_iso()
            db.execute("UPDATE copy_studio_jobs SET status='succeeded',result=?,last_error='',worker_id=NULL,lease_expires_at=NULL,updated_at=?,completed_at=? WHERE id=? AND agency_id=?",
                       (_json(result), timestamp, timestamp, job_id, self.agency_id))

    def fail_job(self, job_id, worker_id, message, retry=True):
        # The caller supplies a sanitised operational message, never provider response bodies.
        message = str(message)[:1000]
        with self._db(write=True) as db:
            row = self._owned(db, job_id, worker_id)
            again = retry and row["attempts"] < row["max_attempts"]
            timestamp = now_iso()
            after = (datetime.now(timezone.utc) + timedelta(seconds=min(3600, 30 * 2 ** (row["attempts"] - 1)))).isoformat()
            db.execute("UPDATE copy_studio_jobs SET status=?,last_error=?,worker_id=NULL,lease_expires_at=NULL,run_after=?,updated_at=?,completed_at=? WHERE id=? AND agency_id=?",
                       ("queued" if again else "failed", message, after, timestamp, None if again else timestamp, job_id, self.agency_id))


class SupabaseStore(_RecordPagination):
    """Service-role-only PostgREST adapter; atomic leases are database functions."""
    def __init__(self, url: str, service_key: str, agency_id: str):
        if not url or not service_key or not agency_id:
            raise StoreNotReady("Copy Studio needs SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY and COPY_STUDIO_AGENCY_ID (or PULSE_AGENCY_ID).")
        self.url = url.rstrip("/")
        self._key = service_key
        self.agency_id = _text(agency_id, "agency ID", 200)
        try:
            uuid.UUID(self.agency_id)
        except ValueError as exc:
            raise StoreNotReady("Copy Studio agency ID must be a UUID.") from exc

    def _request(self, method, table, *, params=None, json=None, prefer=""):
        headers = {"apikey": self._key, "Authorization": f"Bearer {self._key}", "Content-Type": "application/json"}
        if prefer:
            headers["Prefer"] = prefer
        try:
            response = getattr(requests, method)(f"{self.url}/rest/v1/{table}", headers=headers,
                                                params=params, json=json, timeout=(5, 30))
        except requests.RequestException as exc:
            raise StoreNotReady("Copy Studio storage could not be reached.") from exc
        if response.status_code in (401, 403):
            raise StoreNotReady("Copy Studio storage denied access; check the server service credential and migration.")
        if response.status_code == 404:
            raise StoreNotReady("Copy Studio tables or functions are missing; apply the reviewed migration.")
        if response.status_code >= 400:
            raise StoreError(f"Copy Studio storage operation failed (HTTP {response.status_code}).")
        try:
            return response.json()
        except ValueError as exc:
            raise StoreError("Copy Studio storage returned an invalid response.") from exc

    def _scope(self, client_id, **params):
        return {"agency_id": f"eq.{self.agency_id}", "client_id": f"eq.{_text(client_id, 'client ID', 200)}", **params}

    def get(self, record_id, client_id):
        rows = self._request("get", "copy_studio_record_status", params=self._scope(client_id, id=f"eq.{record_id}", limit=1))
        return rows[0] if rows else None

    def list_page(self, client_id, kind=None, limit=200, status=None, cursor=None):
        if status is not None and status not in DECISIONS | {"pending"}:
            raise StoreError("Invalid record status.")
        # Request one sentinel row but remain below PostgREST's default 1,000 cap.
        limit = _limit(limit, 500)
        cursor_data = _parse_cursor(cursor, self.agency_id, client_id, kind, status)
        params = self._scope(client_id, order="created_at.desc,id.desc", limit=limit + 1)
        if kind:
            params["kind"] = f"eq.{kind}"
        if status:
            params["status"] = f"eq.{status}"
        if cursor_data:
            stamp, row_id = cursor_data["created_at"], cursor_data["id"]
            params["or"] = f"(created_at.lt.{stamp},and(created_at.eq.{stamp},id.lt.{row_id}))"
        rows = self._request("get", "copy_studio_record_status", params=params)
        more, records = len(rows) > limit, rows[:limit]
        next_cursor = _page_cursor(self.agency_id, client_id, kind, status, records[-1]) if more else None
        return {"records": records, "has_more": more, "next_cursor": next_cursor}

    def save(self, kind, client_id, source_key, content, actor, parent_id=None):
        row = _record(self.agency_id, kind, client_id, source_key, content, actor, parent_id)
        if parent_id and self.get(parent_id, client_id) is None:
            raise StoreError("Parent record is not available to this client.")
        self._request("post", "copy_studio_records", json=row, params={"on_conflict": "id"},
                      prefer="resolution=ignore-duplicates,return=representation")
        return self.get(row["id"], client_id)

    def review(self, record_id, client_id, actor, decision, notes=""):
        if self.get(record_id, client_id) is None:
            raise StoreError("Record is not available to this client.")
        self._request("post", "copy_studio_reviews", json=_review(self.agency_id, client_id, record_id, actor, decision, notes), prefer="return=representation")
        return self.get(record_id, client_id)

    def enqueue(self, client_id, kind, payload, actor, idempotency_key, max_attempts=4):
        row = _job(self.agency_id, client_id, kind, payload, actor, idempotency_key, max_attempts)
        self._request("post", "copy_studio_jobs", json=row, params={"on_conflict": "id"},
                      prefer="resolution=ignore-duplicates,return=representation")
        result = self.get_job(row["id"], client_id)
        if not result or result["payload_hash"] != row["payload_hash"]:
            raise StoreError("This idempotency key already belongs to a different job request.")
        return result

    def get_job(self, job_id, client_id):
        rows = self._request("get", "copy_studio_jobs", params=self._scope(client_id, id=f"eq.{job_id}", limit=1))
        return rows[0] if rows else None

    def list_jobs(self, client_id, limit=30):
        return self._request("get", "copy_studio_jobs", params=self._scope(client_id, order="created_at.desc,id.desc", limit=_limit(limit)))

    def context_refresh_active(self, client_id):
        rows = self._request("get", "copy_studio_jobs", params=self._scope(
            client_id, select="id", kind="eq.sync", status="in.(queued,running)", limit=1,
            **{"payload->>source": "in.(context,onboarding,all)"}))
        return bool(rows)

    def _rpc(self, name, **params):
        return self._request("post", f"rpc/copy_studio_{name}", json={"p_agency_id": self.agency_id, **params})

    def claim_job(self, worker_id, lease_seconds=300):
        rows = self._rpc("claim_job", p_worker_id=_text(worker_id, "worker ID", 200), p_lease_seconds=max(10, min(3600, int(lease_seconds))))
        return rows[0] if rows else None

    def _owned_rpc(self, name, job_id, worker_id, **params):
        result = self._rpc(name, p_job_id=job_id, p_worker_id=_text(worker_id, "worker ID", 200), **params)
        if result is not True:
            raise LeaseLost("Job lease expired or belongs to another worker.")

    def checkpoint_job(self, job_id, worker_id, checkpoint):
        _json(checkpoint)
        self._owned_rpc("checkpoint_job", job_id, worker_id, p_checkpoint=checkpoint)

    def heartbeat_job(self, job_id, worker_id, lease_seconds=300):
        self._owned_rpc("heartbeat_job", job_id, worker_id, p_lease_seconds=max(10, min(3600, int(lease_seconds))))

    def complete_job(self, job_id, worker_id, result):
        _json(result)
        self._owned_rpc("complete_job", job_id, worker_id, p_result=result)

    def fail_job(self, job_id, worker_id, message, retry=True):
        self._owned_rpc("fail_job", job_id, worker_id, p_message=str(message)[:1000], p_retry=bool(retry))


def get_store():
    """Read configuration at call time; no DB/network work occurs during import."""
    agency = os.environ.get("COPY_STUDIO_AGENCY_ID") or os.environ.get("PULSE_AGENCY_ID") or ""
    path = os.environ.get("COPY_STUDIO_SQLITE_PATH", "").strip()
    if path:
        return SQLiteStore(path, agency_id=agency or "local")
    from app.utils import supabase
    return SupabaseStore(os.environ.get("SUPABASE_URL") or supabase.SUPABASE_URL,
                         os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or supabase.SUPABASE_SERVICE_ROLE_KEY,
                         agency)
