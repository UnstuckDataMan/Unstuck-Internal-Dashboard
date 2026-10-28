-- AI Copy Studio: private append-only evidence/version corpus and durable jobs.
-- Review and apply in staging first, then production only with release approval.
-- Existing Copy Bank, Pulse, clients and app_users permissions are not changed.
-- Requires the existing public.agencies table. Application users never receive a
-- database grant: FastAPI authorizes each client and uses the server service key.
-- Safe to re-run. Rollback: disable the Copy Studio routes/worker first; retain
-- these tables for audit. Destructive DROP is deliberately not automated.
BEGIN;

CREATE TABLE IF NOT EXISTS public.copy_studio_records (
    id uuid PRIMARY KEY,
    agency_id uuid NOT NULL REFERENCES public.agencies(id),
    client_id text NOT NULL CHECK (length(client_id) BETWEEN 1 AND 200),
    kind text NOT NULL CHECK (length(kind) BETWEEN 1 AND 80),
    source_key text NOT NULL CHECK (length(source_key) BETWEEN 1 AND 1000),
    content jsonb NOT NULL CHECK (jsonb_typeof(content) = 'object'),
    content_hash text NOT NULL CHECK (length(content_hash) = 64),
    version text NOT NULL CHECK (version = content_hash),
    parent_id uuid,
    created_by text NOT NULL CHECK (length(created_by) BETWEEN 1 AND 320),
    created_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE (agency_id, client_id, id),
    FOREIGN KEY (agency_id, client_id, parent_id)
        REFERENCES public.copy_studio_records(agency_id, client_id, id)
);
CREATE INDEX IF NOT EXISTS copy_studio_records_scope
    ON public.copy_studio_records(agency_id, client_id, kind, created_at DESC);
CREATE INDEX IF NOT EXISTS copy_studio_records_source
    ON public.copy_studio_records(agency_id, client_id, source_key, created_at DESC);

CREATE TABLE IF NOT EXISTS public.copy_studio_reviews (
    id uuid PRIMARY KEY,
    agency_id uuid NOT NULL REFERENCES public.agencies(id),
    client_id text NOT NULL,
    record_id uuid NOT NULL,
    decision text NOT NULL CHECK (decision IN ('approved', 'rejected', 'needs_changes')),
    reviewed_by text NOT NULL CHECK (length(reviewed_by) BETWEEN 1 AND 320),
    notes text NOT NULL DEFAULT '' CHECK (length(notes) <= 10000),
    created_at timestamptz NOT NULL DEFAULT now(),
    FOREIGN KEY (agency_id, client_id, record_id)
        REFERENCES public.copy_studio_records(agency_id, client_id, id)
);
CREATE INDEX IF NOT EXISTS copy_studio_reviews_record
    ON public.copy_studio_reviews(agency_id, client_id, record_id, created_at DESC, id DESC);

-- Every immutable version has a separately auditable sequence of review
-- decisions. Edits create a new record; neither source nor review is overwritten.
CREATE OR REPLACE FUNCTION public.copy_studio_immutable()
RETURNS trigger LANGUAGE plpgsql SET search_path = pg_catalog, public AS $$
BEGIN
    RAISE EXCEPTION 'Copy Studio evidence and reviews are immutable';
END;
$$;
DROP TRIGGER IF EXISTS copy_studio_records_immutable ON public.copy_studio_records;
CREATE TRIGGER copy_studio_records_immutable BEFORE UPDATE OR DELETE ON public.copy_studio_records
    FOR EACH ROW EXECUTE FUNCTION public.copy_studio_immutable();
DROP TRIGGER IF EXISTS copy_studio_reviews_immutable ON public.copy_studio_reviews;
CREATE TRIGGER copy_studio_reviews_immutable BEFORE UPDATE OR DELETE ON public.copy_studio_reviews
    FOR EACH ROW EXECUTE FUNCTION public.copy_studio_immutable();

CREATE OR REPLACE VIEW public.copy_studio_record_status
    WITH (security_invoker = true) AS
SELECT r.*, COALESCE(v.decision, 'pending') AS status,
       CASE WHEN v.id IS NULL THEN NULL ELSE to_jsonb(v) END AS review
FROM public.copy_studio_records r
LEFT JOIN LATERAL (
    SELECT a.* FROM public.copy_studio_reviews a
    WHERE a.agency_id = r.agency_id AND a.client_id = r.client_id AND a.record_id = r.id
    ORDER BY a.created_at DESC, a.id DESC LIMIT 1
) v ON true;

CREATE TABLE IF NOT EXISTS public.copy_studio_jobs (
    id uuid PRIMARY KEY,
    agency_id uuid NOT NULL REFERENCES public.agencies(id),
    client_id text NOT NULL CHECK (length(client_id) BETWEEN 1 AND 200),
    kind text NOT NULL CHECK (length(kind) BETWEEN 1 AND 80),
    payload jsonb NOT NULL CHECK (jsonb_typeof(payload) = 'object'),
    payload_hash text NOT NULL CHECK (length(payload_hash) = 64),
    created_by text NOT NULL CHECK (length(created_by) BETWEEN 1 AND 320),
    idempotency_key text NOT NULL CHECK (length(idempotency_key) BETWEEN 1 AND 500),
    status text NOT NULL DEFAULT 'queued' CHECK (status IN ('queued', 'running', 'succeeded', 'failed')),
    attempts integer NOT NULL DEFAULT 0 CHECK (attempts >= 0),
    max_attempts integer NOT NULL DEFAULT 4 CHECK (max_attempts BETWEEN 1 AND 10),
    result jsonb NOT NULL DEFAULT '{}' CHECK (jsonb_typeof(result) = 'object'),
    checkpoint jsonb NOT NULL DEFAULT '{}' CHECK (jsonb_typeof(checkpoint) = 'object'),
    last_error text NOT NULL DEFAULT '' CHECK (length(last_error) <= 1000),
    worker_id text,
    lease_expires_at timestamptz,
    run_after timestamptz NOT NULL DEFAULT now(),
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now(),
    completed_at timestamptz,
    UNIQUE (agency_id, client_id, kind, idempotency_key)
);
CREATE INDEX IF NOT EXISTS copy_studio_jobs_claim
    ON public.copy_studio_jobs(agency_id, status, run_after);

-- Transactional claim: concurrent workers cannot claim the same live lease.
-- Expired leases are resumable from checkpoint; exhausted jobs become terminal.
CREATE OR REPLACE FUNCTION public.copy_studio_claim_job(
    p_agency_id uuid, p_worker_id text, p_lease_seconds integer DEFAULT 300)
RETURNS SETOF public.copy_studio_jobs LANGUAGE plpgsql
SET search_path = pg_catalog, public AS $$
DECLARE chosen uuid;
BEGIN
    IF length(trim(p_worker_id)) NOT BETWEEN 1 AND 200 THEN
        RAISE EXCEPTION 'A worker identity is required';
    END IF;
    UPDATE public.copy_studio_jobs j
    SET status = 'failed', worker_id = NULL, lease_expires_at = NULL,
        last_error = 'Worker lease expired; retry limit reached.', updated_at = now(), completed_at = now()
    WHERE j.agency_id = p_agency_id AND j.status = 'running'
      AND j.lease_expires_at <= now() AND j.attempts >= j.max_attempts;
    SELECT j.id INTO chosen FROM public.copy_studio_jobs j
    WHERE j.agency_id = p_agency_id AND j.attempts < j.max_attempts
      AND ((j.status = 'queued' AND j.run_after <= now())
           OR (j.status = 'running' AND j.lease_expires_at <= now()))
    ORDER BY j.run_after, j.created_at, j.id
    FOR UPDATE SKIP LOCKED LIMIT 1;
    IF chosen IS NOT NULL THEN
        RETURN QUERY UPDATE public.copy_studio_jobs j
        SET status = 'running', attempts = j.attempts + 1, worker_id = p_worker_id,
            lease_expires_at = now() + make_interval(secs => greatest(10, least(3600, p_lease_seconds))), updated_at = now()
        WHERE j.id = chosen AND j.agency_id = p_agency_id RETURNING j.*;
    END IF;
END;
$$;

CREATE OR REPLACE FUNCTION public.copy_studio_checkpoint_job(
    p_agency_id uuid, p_job_id uuid, p_worker_id text, p_checkpoint jsonb)
RETURNS boolean LANGUAGE plpgsql SET search_path = pg_catalog, public AS $$
BEGIN
    UPDATE public.copy_studio_jobs j SET checkpoint = p_checkpoint, updated_at = now()
    WHERE j.id = p_job_id AND j.agency_id = p_agency_id AND j.worker_id = p_worker_id
      AND j.status = 'running' AND j.lease_expires_at > now();
    RETURN FOUND;
END;
$$;

CREATE OR REPLACE FUNCTION public.copy_studio_heartbeat_job(
    p_agency_id uuid, p_job_id uuid, p_worker_id text, p_lease_seconds integer DEFAULT 300)
RETURNS boolean LANGUAGE plpgsql SET search_path = pg_catalog, public AS $$
BEGIN
    UPDATE public.copy_studio_jobs j
    SET lease_expires_at = now() + make_interval(secs => greatest(10, least(3600, p_lease_seconds))), updated_at = now()
    WHERE j.id = p_job_id AND j.agency_id = p_agency_id AND j.worker_id = p_worker_id
      AND j.status = 'running' AND j.lease_expires_at > now();
    RETURN FOUND;
END;
$$;

CREATE OR REPLACE FUNCTION public.copy_studio_complete_job(
    p_agency_id uuid, p_job_id uuid, p_worker_id text, p_result jsonb)
RETURNS boolean LANGUAGE plpgsql SET search_path = pg_catalog, public AS $$
BEGIN
    UPDATE public.copy_studio_jobs j SET status = 'succeeded', result = p_result, last_error = '',
        worker_id = NULL, lease_expires_at = NULL, updated_at = now(), completed_at = now()
    WHERE j.id = p_job_id AND j.agency_id = p_agency_id AND j.worker_id = p_worker_id
      AND j.status = 'running' AND j.lease_expires_at > now();
    RETURN FOUND;
END;
$$;

CREATE OR REPLACE FUNCTION public.copy_studio_fail_job(
    p_agency_id uuid, p_job_id uuid, p_worker_id text, p_message text, p_retry boolean DEFAULT true)
RETURNS boolean LANGUAGE plpgsql SET search_path = pg_catalog, public AS $$
BEGIN
    UPDATE public.copy_studio_jobs j
    SET status = CASE WHEN p_retry AND j.attempts < j.max_attempts THEN 'queued' ELSE 'failed' END,
        last_error = left(p_message, 1000), worker_id = NULL, lease_expires_at = NULL,
        run_after = now() + make_interval(secs => least(3600, (30 * power(2, j.attempts - 1))::integer)),
        updated_at = now(), completed_at = CASE WHEN p_retry AND j.attempts < j.max_attempts THEN NULL ELSE now() END
    WHERE j.id = p_job_id AND j.agency_id = p_agency_id AND j.worker_id = p_worker_id
      AND j.status = 'running' AND j.lease_expires_at > now();
    RETURN FOUND;
END;
$$;

ALTER TABLE public.copy_studio_records ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.copy_studio_reviews ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.copy_studio_jobs ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON public.copy_studio_records, public.copy_studio_reviews, public.copy_studio_jobs,
    public.copy_studio_record_status FROM PUBLIC, anon, authenticated, service_role;
-- Supabase can grant service_role ALL through default privileges. Reset those
-- object grants before narrowing them; additive GRANT alone leaves TRUNCATE,
-- UPDATE and DELETE available on the immutable corpus.
GRANT SELECT, INSERT ON public.copy_studio_records, public.copy_studio_reviews TO service_role;
GRANT SELECT ON public.copy_studio_record_status TO service_role;
GRANT SELECT, INSERT, UPDATE ON public.copy_studio_jobs TO service_role;

REVOKE ALL ON FUNCTION public.copy_studio_immutable() FROM PUBLIC, anon, authenticated, service_role;
REVOKE ALL ON FUNCTION public.copy_studio_claim_job(uuid,text,integer) FROM PUBLIC, anon, authenticated, service_role;
REVOKE ALL ON FUNCTION public.copy_studio_checkpoint_job(uuid,uuid,text,jsonb) FROM PUBLIC, anon, authenticated, service_role;
REVOKE ALL ON FUNCTION public.copy_studio_heartbeat_job(uuid,uuid,text,integer) FROM PUBLIC, anon, authenticated, service_role;
REVOKE ALL ON FUNCTION public.copy_studio_complete_job(uuid,uuid,text,jsonb) FROM PUBLIC, anon, authenticated, service_role;
REVOKE ALL ON FUNCTION public.copy_studio_fail_job(uuid,uuid,text,text,boolean) FROM PUBLIC, anon, authenticated, service_role;
GRANT EXECUTE ON FUNCTION public.copy_studio_claim_job(uuid,text,integer) TO service_role;
GRANT EXECUTE ON FUNCTION public.copy_studio_checkpoint_job(uuid,uuid,text,jsonb) TO service_role;
GRANT EXECUTE ON FUNCTION public.copy_studio_heartbeat_job(uuid,uuid,text,integer) TO service_role;
GRANT EXECUTE ON FUNCTION public.copy_studio_complete_job(uuid,uuid,text,jsonb) TO service_role;
GRANT EXECUTE ON FUNCTION public.copy_studio_fail_job(uuid,uuid,text,text,boolean) TO service_role;

NOTIFY pgrst, 'reload schema';
COMMIT;
