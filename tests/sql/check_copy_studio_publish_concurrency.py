"""Exercise real two-session publishing races in a network-disabled test container.

Usage: python tests/sql/check_copy_studio_publish_concurrency.py --container copy-studio-publish-check
Requires an existing disposable postgres:17-alpine Docker container with network
mode 'none'. Creates/drops its own database; never connects to a remote database.
"""
import argparse
from concurrent.futures import ThreadPoolExecutor
import json
import os
from pathlib import Path
import subprocess
import time
import uuid


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--container', required=True)
    args = parser.parse_args()
    if not args.container.startswith('copy-studio-'):
        parser.error('Use an explicitly named copy-studio-* disposable test container.')
    env = {k: v for k, v in os.environ.items() if k not in {'DOCKER_HOST', 'DOCKER_CONTEXT', 'DOCKER_TLS', 'DOCKER_TLS_VERIFY', 'DOCKER_CERT_PATH'}}
    docker = ['docker', '--host=unix:///var/run/docker.sock']
    info = json.loads(subprocess.check_output(docker+['inspect', args.container], env=env, text=True))[0]
    if info['HostConfig']['NetworkMode'] != 'none' or not info['Config']['Image'].startswith('postgres:17'):
        parser.error('The fixture must be PostgreSQL 17 with networking disabled.')
    database = 'copy_studio_publish_test_'+uuid.uuid4().hex
    base = docker+['exec', '-i', args.container, 'psql', '-U', 'postgres', '-v', 'ON_ERROR_STOP=1', '-Atq']

    def sql(statement, db=database):
        result = subprocess.run(base+['-d', db], input=statement, text=True, env=env, capture_output=True, timeout=25)
        if result.returncode:
            raise AssertionError(result.stderr)
        return result.stdout.strip()

    def lit(value):
        return "'"+json.dumps(value, separators=(',', ':')).replace("'", "''")+"'::jsonb"

    root = Path(__file__).resolve().parents[2]
    agency = '11111111-1111-1111-1111-111111111111'
    client = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
    brief = '33333333-3333-3333-3333-333333333333'
    draft = '55555555-5555-5555-5555-555555555555'
    key = '__c__'+client+'__UK_Software'
    angles = [{'subjects': ['Subject'], 'steps': [{'kind': 'initial', 'body': 'Approved exact opening.'}, {'kind': 'follow_up', 'body': 'Approved follow-up.'}]}]
    content = {'channel': 'email', 'request': {'locale': 'en-GB'}, 'angles': angles,
               'studio_context': {'brief_id': brief, 'example_ids': []}, 'review': {'can_approve': True, 'issues': []},
               'final_check': {'complete': True, 'locale': 'en-GB', 'copy_hash': 'b'*64, 'checked_angles': angles,
                               'report': {'complete': True, 'locale': 'en-GB', 'copy_hash': 'b'*64, 'issues': []}}}
    selection = {'territory': 'UK', 'industry': 'Software', 'section': 'email', 'angle_index': 0, 'step_indices': [0]}
    sql('CREATE DATABASE '+database, db='postgres')
    try:
        sql("""DO $$ BEGIN
          IF NOT EXISTS(SELECT 1 FROM pg_roles WHERE rolname='anon') THEN CREATE ROLE anon NOLOGIN; END IF;
          IF NOT EXISTS(SELECT 1 FROM pg_roles WHERE rolname='authenticated') THEN CREATE ROLE authenticated NOLOGIN; END IF;
          IF NOT EXISTS(SELECT 1 FROM pg_roles WHERE rolname='service_role') THEN CREATE ROLE service_role NOLOGIN BYPASSRLS; END IF;
        END $$;
        CREATE TABLE agencies(id uuid PRIMARY KEY);
        CREATE TABLE clients(id uuid PRIMARY KEY,agency_id uuid,name text);
        CREATE TABLE copy_bank_templates(key text PRIMARY KEY,content jsonb NOT NULL);
        CREATE TABLE copy_bank_pending(key text PRIMARY KEY,content jsonb,status text);
        -- Match the inherited defaults observed in a hosted Supabase project.
        ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT ALL ON TABLES TO service_role;
        ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT ALL ON FUNCTIONS TO service_role;
        """)
        sql((root/'migrations/ai_copy_studio.sql').read_text())
        sql((root/'migrations/ai_copy_studio_publish.sql').read_text())
        sql((root/'tests/sql/test_ai_copy_studio.sql').read_text())
        sql((root/'tests/sql/test_ai_copy_studio_publish.sql').read_text())
        # Also cover repair of an already-applied schema that inherited broader
        # grants. Reapplying must narrow rights without erasing stored history.
        sql("""GRANT ALL ON copy_studio_records,copy_studio_reviews,copy_studio_jobs,
                  copy_studio_record_status,copy_studio_publications TO service_role;
               GRANT EXECUTE ON FUNCTION copy_studio_immutable(),
                  copy_studio_copy_bank_append(jsonb,jsonb,jsonb,jsonb) TO service_role;""")
        sql((root/'migrations/ai_copy_studio.sql').read_text())
        sql((root/'migrations/ai_copy_studio_publish.sql').read_text())
        sql((root/'tests/sql/test_ai_copy_studio.sql').read_text())
        sql((root/'tests/sql/test_ai_copy_studio_publish.sql').read_text())
        profiles = [{'client_id': client, 'type': 'client', 'name': 'Fixture', 'territories': ['UK', 'US', 'CA'], 'industries': ['Software']}]
        sql(f"INSERT INTO agencies VALUES('{agency}'); INSERT INTO clients VALUES('{client}','{agency}','Fixture');"
            f"INSERT INTO copy_bank_templates VALUES('__cb_profiles__',{lit(profiles)}),('{key}', '{{}}');")
        for identifier, kind, value in [(brief, 'brief', {}), (draft, 'draft', content)]:
            sql(f"INSERT INTO copy_studio_records(id,agency_id,client_id,kind,source_key,content,content_hash,version,created_by) "
                f"VALUES('{identifier}','{agency}','{client}','{kind}','{kind}',{lit(value)},repeat('a',64),repeat('a',64),'author@example.test');"
                f"INSERT INTO copy_studio_reviews(id,agency_id,client_id,record_id,decision,reviewed_by) "
                f"VALUES(gen_random_uuid(),'{agency}','{client}','{identifier}','approved','reviewer@example.test');")
        snapshot = json.loads(sql(f"SELECT jsonb_object_agg(id::text,jsonb_build_object('status',status,'review_id',review->>'id')) "
                                  f"FROM copy_studio_record_status WHERE agency_id='{agency}' AND client_id='{client}';"))

        def version(chosen):
            return json.loads(sql(f"SELECT copy_studio_copy_bank_destination('{agency}','{client}',{lit(chosen)});"))['version']

        def publish(chosen, expected, identity=None, hold=0):
            statement = (f"SELECT copy_studio_publish_copy_bank('{agency}','{client}','{draft}',{lit(chosen)},'{expected}',"
                         f"'reviewer@example.test','{identity or uuid.uuid4()}',{lit(snapshot)});")
            if hold:
                statement = 'BEGIN; '+statement+f' SELECT pg_sleep({float(hold)}); COMMIT;'
            return json.loads(sql(statement).splitlines()[0])

        def wait_for_lock(lock_key):
            # Observe the first transaction reaching its controlled sleep after
            # the write; no assumption about Docker process startup latency.
            for _ in range(100):
                if sql("SELECT count(*) FROM pg_stat_activity WHERE datname=current_database() AND wait_event='PgSleep';") != '0':
                    return
                time.sleep(.02)
            raise AssertionError('Fixture transaction did not acquire its lock.')

        with ThreadPoolExecutor(max_workers=2) as pool:
            expected = version(selection)
            identity = str(uuid.uuid4())
            first = pool.submit(publish, selection, expected, identity, 1)
            wait_for_lock(key)
            second = pool.submit(publish, selection, expected, identity)
            a, b = first.result(), second.result()
            assert a['published'] and b['already_published'] and a['publication_id'] == b['publication_id']
            assert sql(f"SELECT jsonb_array_length(content->'email'->'variations') FROM copy_bank_templates WHERE key='{key}';") == '1'

            # Two different selections from the same preview cannot silently
            # merge across an unseen change; the second must re-preview.
            email = dict(selection, territory='US')
            followup = dict(email, section='chaser', step_indices=[1])
            expected = version(email)
            first = pool.submit(publish, email, expected, None, 1)
            wait_for_lock('US')
            second = pool.submit(publish, followup, expected)
            assert first.result()['published']
            assert second.result()['code'] == 'conflict'

            # A legacy editor does not take the publisher advisory lock. Its row
            # lock must still force the publisher to compare the actual content.
            expected = version(selection)
            legacy = pool.submit(sql, f"BEGIN; UPDATE copy_bank_templates SET content=content||'{{\"legacy\":true}}' WHERE key='{key}'; SELECT pg_sleep(1); COMMIT;")
            wait_for_lock(key)
            attempt = pool.submit(publish, dict(selection, section='chaser', step_indices=[1]), expected)
            legacy.result()
            assert attempt.result()['code'] == 'conflict'

            # A concurrent legacy first insertion is protected by the unique key,
            # even though the absent row could not be locked by SELECT FOR UPDATE.
            new = dict(selection, territory='CA')
            expected = version(new)
            new_key = '__c__'+client+'__CA_Software'
            legacy = pool.submit(sql, f"BEGIN; INSERT INTO copy_bank_templates VALUES('{new_key}','{{\"legacy\":true}}'); SELECT pg_sleep(1); COMMIT;")
            wait_for_lock(new_key)
            attempt = pool.submit(publish, new, expected)
            legacy.result()
            assert attempt.result()['code'] == 'conflict'
            assert json.loads(sql(f"SELECT content FROM copy_bank_templates WHERE key='{new_key}';")) == {'legacy': True}
        assert sql('SELECT count(*) FROM copy_studio_publications;') == '2'
        print('PASS: Supabase default-grant narrowing on initial apply and repair reapply; private SQL assertions; concurrent duplicate, different-selection stale preview, legacy update and legacy first-insert races')
    finally:
        sql('DROP DATABASE '+database+' WITH (FORCE)', db='postgres')


if __name__ == '__main__':
    main()
