-- Run only in a disposable PostgreSQL 15+ database after ai_copy_studio.sql.
-- Requires the test agency IDs to be absent. All fixture writes are rolled back.
-- psql -v ON_ERROR_STOP=1 -f tests/sql/test_ai_copy_studio.sql
BEGIN;
INSERT INTO public.agencies(id) VALUES
    ('11111111-1111-1111-1111-111111111111'),
    ('22222222-2222-2222-2222-222222222222');

DO $$
DECLARE t text; f text;
BEGIN
    FOREACH t IN ARRAY ARRAY['copy_studio_records','copy_studio_reviews','copy_studio_jobs'] LOOP
        ASSERT (SELECT relrowsecurity FROM pg_class WHERE oid=('public.'||t)::regclass), 'RLS must be enabled';
        ASSERT NOT has_table_privilege('anon','public.'||t,'SELECT,INSERT,UPDATE,DELETE'), 'anon table access';
        ASSERT NOT has_table_privilege('authenticated','public.'||t,'SELECT,INSERT,UPDATE,DELETE'), 'authenticated table access';
        ASSERT has_table_privilege('service_role','public.'||t,'SELECT,INSERT'), 'server missing table access';
        ASSERT NOT has_table_privilege('service_role','public.'||t,'DELETE,TRUNCATE,REFERENCES,TRIGGER'), 'server inherited unsafe default grants';
        IF t <> 'copy_studio_jobs' THEN
            ASSERT NOT has_table_privilege('service_role','public.'||t,'UPDATE'), 'server can update immutable evidence';
        ELSE
            ASSERT has_table_privilege('service_role','public.'||t,'UPDATE'), 'server cannot update leased jobs';
        END IF;
    END LOOP;
    ASSERT NOT has_table_privilege('anon','public.copy_studio_record_status','SELECT'), 'anon view access';
    ASSERT NOT has_table_privilege('authenticated','public.copy_studio_record_status','SELECT'), 'authenticated view access';
    ASSERT has_table_privilege('service_role','public.copy_studio_record_status','SELECT'), 'server missing view access';
    ASSERT NOT has_table_privilege('service_role','public.copy_studio_record_status','INSERT,UPDATE,DELETE,TRUNCATE,REFERENCES,TRIGGER'), 'server inherited view mutation grants';
    ASSERT NOT has_function_privilege('service_role','public.copy_studio_immutable()','EXECUTE'), 'internal immutable helper exposed';
    FOREACH f IN ARRAY ARRAY[
        'copy_studio_claim_job(uuid,text,integer)',
        'copy_studio_checkpoint_job(uuid,uuid,text,jsonb)',
        'copy_studio_heartbeat_job(uuid,uuid,text,integer)',
        'copy_studio_complete_job(uuid,uuid,text,jsonb)',
        'copy_studio_fail_job(uuid,uuid,text,text,boolean)'
    ] LOOP
        ASSERT NOT has_function_privilege('anon','public.'||f,'EXECUTE'), 'anon RPC access';
        ASSERT NOT has_function_privilege('authenticated','public.'||f,'EXECUTE'), 'authenticated RPC access';
        ASSERT has_function_privilege('service_role','public.'||f,'EXECUTE'), 'server missing RPC access';
    END LOOP;
END;
$$;

SET LOCAL ROLE service_role;
INSERT INTO public.copy_studio_records
    (id,agency_id,client_id,kind,source_key,content,content_hash,version,created_by)
VALUES ('33333333-3333-3333-3333-333333333333','11111111-1111-1111-1111-111111111111','client-one','brief','fixture','{"offer":"Accounting"}',repeat('a',64),repeat('a',64),'author@example.test');
INSERT INTO public.copy_studio_reviews(id,agency_id,client_id,record_id,decision,reviewed_by)
VALUES ('44444444-4444-4444-4444-444444444444','11111111-1111-1111-1111-111111111111','client-one','33333333-3333-3333-3333-333333333333','approved','reviewer@example.test');
DO $$
BEGIN
    ASSERT (SELECT status FROM public.copy_studio_record_status WHERE id='33333333-3333-3333-3333-333333333333')='approved', 'review status view';
    ASSERT (SELECT review->>'reviewed_by' FROM public.copy_studio_record_status WHERE id='33333333-3333-3333-3333-333333333333')='reviewer@example.test', 'review attribution';
    BEGIN
        INSERT INTO public.copy_studio_reviews(id,agency_id,client_id,record_id,decision,reviewed_by)
        VALUES(gen_random_uuid(),'22222222-2222-2222-2222-222222222222','client-one','33333333-3333-3333-3333-333333333333','approved','other@example.test');
        RAISE EXCEPTION 'cross-agency review accepted';
    EXCEPTION WHEN foreign_key_violation THEN NULL;
    END;
    BEGIN
        INSERT INTO public.copy_studio_reviews(id,agency_id,client_id,record_id,decision,reviewed_by)
        VALUES(gen_random_uuid(),'11111111-1111-1111-1111-111111111111','client-two','33333333-3333-3333-3333-333333333333','approved','other@example.test');
        RAISE EXCEPTION 'cross-client review accepted';
    EXCEPTION WHEN foreign_key_violation THEN NULL;
    END;
    BEGIN
        INSERT INTO public.copy_studio_records(id,agency_id,client_id,kind,source_key,content,content_hash,version,created_by,parent_id)
        VALUES(gen_random_uuid(),'22222222-2222-2222-2222-222222222222','client-one','brief','bad-parent','{}',repeat('b',64),repeat('b',64),'author@example.test','33333333-3333-3333-3333-333333333333');
        RAISE EXCEPTION 'cross-agency parent accepted';
    EXCEPTION WHEN foreign_key_violation THEN NULL;
    END;
END;
$$;
RESET ROLE;
-- Immutability is enforced even for an owner, rather than only by missing grants.
DO $$
BEGIN
    BEGIN
        UPDATE public.copy_studio_records SET content='{}' WHERE id='33333333-3333-3333-3333-333333333333';
        RAISE EXCEPTION 'record update was accepted';
    EXCEPTION WHEN raise_exception THEN
        IF SQLERRM <> 'Copy Studio evidence and reviews are immutable' THEN RAISE; END IF;
    END;
    BEGIN
        DELETE FROM public.copy_studio_reviews WHERE id='44444444-4444-4444-4444-444444444444';
        RAISE EXCEPTION 'review deletion was accepted';
    EXCEPTION WHEN raise_exception THEN
        IF SQLERRM <> 'Copy Studio evidence and reviews are immutable' THEN RAISE; END IF;
    END;
END;
$$;

SET LOCAL ROLE service_role;
INSERT INTO public.copy_studio_jobs
    (id,agency_id,client_id,kind,payload,payload_hash,created_by,idempotency_key,max_attempts)
VALUES
    ('55555555-5555-5555-5555-555555555555','11111111-1111-1111-1111-111111111111','client-one','sync','{}',repeat('c',64),'author@example.test','test-sync',3);
DO $$
DECLARE claimed public.copy_studio_jobs;
BEGIN
    ASSERT (SELECT count(*) FROM public.copy_studio_claim_job('22222222-2222-2222-2222-222222222222','other-agency',300))=0, 'cross-agency job claim';
    SELECT * INTO claimed FROM public.copy_studio_claim_job('11111111-1111-1111-1111-111111111111','worker-one',300);
    ASSERT claimed.id='55555555-5555-5555-5555-555555555555' AND claimed.attempts=1, 'initial claim';
    ASSERT (SELECT count(*) FROM public.copy_studio_claim_job('11111111-1111-1111-1111-111111111111','worker-two',300))=0, 'live lease stolen';
    ASSERT NOT public.copy_studio_checkpoint_job('11111111-1111-1111-1111-111111111111',claimed.id,'wrong-worker','{}'), 'wrong worker checkpoint';
    ASSERT NOT public.copy_studio_complete_job('22222222-2222-2222-2222-222222222222',claimed.id,'worker-one','{}'), 'wrong agency complete';
    ASSERT public.copy_studio_checkpoint_job('11111111-1111-1111-1111-111111111111',claimed.id,'worker-one','{"page":7}'), 'checkpoint rejected';
    ASSERT public.copy_studio_heartbeat_job('11111111-1111-1111-1111-111111111111',claimed.id,'worker-one',600), 'heartbeat rejected';
    ASSERT public.copy_studio_fail_job('11111111-1111-1111-1111-111111111111',claimed.id,'worker-one','Provider unavailable',true), 'failure rejected';
    ASSERT (SELECT status FROM public.copy_studio_jobs WHERE id=claimed.id)='queued', 'retry did not queue';
    ASSERT (SELECT run_after FROM public.copy_studio_jobs WHERE id=claimed.id)>now(), 'no retry backoff';
    ASSERT (SELECT checkpoint FROM public.copy_studio_jobs WHERE id=claimed.id)='{"page":7}'::jsonb, 'checkpoint lost';
    UPDATE public.copy_studio_jobs SET run_after=now()-interval '1 minute' WHERE id=claimed.id;
    SELECT * INTO claimed FROM public.copy_studio_claim_job('11111111-1111-1111-1111-111111111111','worker-two',300);
    ASSERT claimed.attempts=2 AND claimed.checkpoint='{"page":7}'::jsonb, 'retry failed';
    UPDATE public.copy_studio_jobs SET lease_expires_at=now()-interval '1 minute' WHERE id=claimed.id;
    ASSERT NOT public.copy_studio_complete_job('11111111-1111-1111-1111-111111111111',claimed.id,'worker-two','{}'), 'expired lease committed';
    SELECT * INTO claimed FROM public.copy_studio_claim_job('11111111-1111-1111-1111-111111111111','worker-three',300);
    ASSERT claimed.attempts=3 AND claimed.checkpoint='{"page":7}'::jsonb, 'expired lease not resumed';
    ASSERT public.copy_studio_complete_job('11111111-1111-1111-1111-111111111111',claimed.id,'worker-three','{"records":12}'), 'completion failed';
    ASSERT (SELECT status FROM public.copy_studio_jobs WHERE id=claimed.id)='succeeded', 'result not completed';
    ASSERT (SELECT result FROM public.copy_studio_jobs WHERE id=claimed.id)='{"records":12}'::jsonb, 'result not saved';
    ASSERT NOT public.copy_studio_complete_job('11111111-1111-1111-1111-111111111111',claimed.id,'worker-three','{}'), 'completed job changed';
END;
$$;
RESET ROLE;
ROLLBACK;
SELECT 'PASS: Copy Studio PostgreSQL permissions, immutable history, tenant references, lease recovery and RPCs' AS result;
