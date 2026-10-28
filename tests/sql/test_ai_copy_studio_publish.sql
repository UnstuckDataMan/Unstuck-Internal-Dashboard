-- Disposable PostgreSQL only, after both Copy Studio migrations. Requires
-- agencies, clients(id,agency_id,name), copy_bank_templates and copy_bank_pending.
-- No provider calls. Every fixture and publication below is rolled back.
BEGIN;
INSERT INTO public.agencies(id) VALUES ('11111111-1111-1111-1111-111111111111'), ('22222222-2222-2222-2222-222222222222');
INSERT INTO public.clients(id,agency_id,name) VALUES
 ('aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa','11111111-1111-1111-1111-111111111111','Fixture A'),
 ('bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb','22222222-2222-2222-2222-222222222222','Fixture B');
INSERT INTO public.copy_bank_templates(key,content) VALUES ('__cb_profiles__',
 '[{"client_id":"aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa","name":"Fixture A","type":"client","territories":["UK","US","CA"],"industries":["Software"],"industryLabels":{"Software":"Software & IT"}},
   {"client_id":"bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb","name":"Fixture B","type":"client","territories":["UK"],"industries":["Software"]}]');
INSERT INTO public.copy_bank_templates(key,content) VALUES ('__c__aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa__UK_Software',
 '{"email":{"subjects":["Existing subject"],"variations":[{"body":"Existing opening"}],"custom":"preserve"},"linkedin":{"steps":[{"body":"Untouched LinkedIn"}]},"other":{"keep":true}}');

-- Context snapshot deliberately covers only immutable source/draft records and
-- their latest review IDs. Cost/performance/job churn is not part of it.
CREATE FUNCTION pg_temp.publish_snapshot(agency uuid, client text, methodology text DEFAULT NULL) RETURNS jsonb
LANGUAGE sql AS $$
 SELECT coalesce(jsonb_object_agg(id::text,jsonb_build_object('status',status,'review_id',review->>'id')),'{}')
 FROM public.copy_studio_record_status WHERE agency_id=agency AND
 ((client_id=client AND kind IN ('draft','brief','example','transcript','sync_report'))
  OR (client_id='__agency__' AND kind='methodology' AND id::text=methodology));
$$;
CREATE FUNCTION pg_temp.add_record(identifier uuid, kind_value text, key_value text, value jsonb,
    client_value text DEFAULT 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa', parent uuid DEFAULT NULL) RETURNS void
LANGUAGE plpgsql AS $$
BEGIN
 INSERT INTO public.copy_studio_records(id,agency_id,client_id,kind,source_key,content,content_hash,version,created_by,parent_id,created_at)
 VALUES(identifier,'11111111-1111-1111-1111-111111111111',client_value,kind_value,key_value,value,repeat('a',64),repeat('a',64),'author@example.test',parent,clock_timestamp());
 INSERT INTO public.copy_studio_reviews(id,agency_id,client_id,record_id,decision,reviewed_by,created_at)
 VALUES(gen_random_uuid(),'11111111-1111-1111-1111-111111111111',client_value,identifier,'approved','reviewer@example.test',clock_timestamp());
END;
$$;
SELECT pg_temp.add_record('33333333-3333-3333-3333-333333333333','brief','client-brief','{"coverage":{"sources":[]}}');
SELECT pg_temp.add_record('44444444-4444-4444-4444-444444444444','example','example','{"source":"copy_bank"}');

DO $$
DECLARE agency uuid := '11111111-1111-1111-1111-111111111111'; client text := 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
 draft_id uuid := '55555555-5555-5555-5555-555555555555'; angles jsonb; value jsonb; selection jsonb; dest jsonb;
 result jsonb; repeated jsonb; snap jsonb; expected text; original jsonb; f text; bad jsonb; checked jsonb;
BEGIN
 ASSERT (SELECT relrowsecurity FROM pg_class WHERE oid='public.copy_studio_publications'::regclass), 'receipt RLS disabled';
 ASSERT NOT has_table_privilege('anon','public.copy_studio_publications','SELECT,INSERT,UPDATE,DELETE'), 'anon receipt access';
 ASSERT NOT has_table_privilege('authenticated','public.copy_studio_publications','SELECT,INSERT,UPDATE,DELETE'), 'authenticated receipt access';
 ASSERT has_table_privilege('service_role','public.copy_studio_publications','SELECT'), 'service missing receipt read';
 ASSERT NOT has_table_privilege('service_role','public.copy_studio_publications','INSERT,UPDATE,DELETE,TRUNCATE,REFERENCES,TRIGGER'), 'service can bypass publication transaction';
 FOREACH f IN ARRAY ARRAY['copy_studio_copy_bank_destinations(uuid,text)','copy_studio_copy_bank_destination(uuid,text,jsonb)',
                         'copy_studio_publish_copy_bank(uuid,text,uuid,jsonb,text,text,uuid,jsonb)'] LOOP
   ASSERT NOT has_function_privilege('anon','public.'||f,'EXECUTE'), 'anon publish RPC access';
   ASSERT NOT has_function_privilege('authenticated','public.'||f,'EXECUTE'), 'authenticated publish RPC access';
   ASSERT has_function_privilege('service_role','public.'||f,'EXECUTE'), 'service missing publish RPC access';
 END LOOP;
 ASSERT NOT has_function_privilege('service_role','public.copy_studio_copy_bank_append(jsonb,jsonb,jsonb,jsonb)','EXECUTE'), 'internal append helper exposed';
 angles := '[{"name":"Clarity","subjects":["Simpler reporting","Reporting question"],"steps":[{"kind":"initial","body":"Could simpler reporting help?"},{"kind":"follow_up","body":"Would a conversation help?"},{"kind":"closing","body":"Happy to leave this here."}]}]';
 value := jsonb_build_object('angles',angles,'channel','email','mode','assemble','request',jsonb_build_object('channel','email','locale','en-GB'),
   'studio_context',jsonb_build_object('brief_id','33333333-3333-3333-3333-333333333333','example_ids',jsonb_build_array('44444444-4444-4444-4444-444444444444')),
   'review',jsonb_build_object('can_approve',true,'issues','[]'::jsonb),
   'final_check',jsonb_build_object('complete',true,'checked_angles',angles,'locale','en-GB','copy_hash',repeat('b',64),
       'report',jsonb_build_object('complete',true,'locale','en-GB','copy_hash',repeat('b',64),'issues','[]'::jsonb)));
 PERFORM pg_temp.add_record(draft_id,'draft','generation-1',value);
 selection := '{"territory":"UK","industry":"Software","section":"email","angle_index":0,"step_indices":[0]}';
 snap := pg_temp.publish_snapshot(agency,client);
 dest := public.copy_studio_copy_bank_destination(agency,client,selection);
 ASSERT dest->>'ok'='true' AND dest->>'pending'='false', 'destination resolution';
 expected := dest->>'version'; original := dest->'content';
 -- Wrong agency/client and unknown combinations cannot fall back to a shared key.
 ASSERT public.copy_studio_copy_bank_destinations('22222222-2222-2222-2222-222222222222',client)->>'ok'='false', 'cross-agency profile read';
 ASSERT public.copy_studio_copy_bank_destination(agency,client,selection||'{"territory":"Other"}')->>'ok'='false', 'unmapped destination';
 result := public.copy_studio_publish_copy_bank(agency,'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb',draft_id,selection,expected,'reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'ok'='false', 'cross-client draft published';
 -- A changed destination is rejected without creating a receipt.
 UPDATE public.copy_bank_templates SET content=content||'{"concurrent":"edit"}' WHERE key=dest->>'key';
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,expected,'reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'code'='conflict', 'stale destination accepted';
 ASSERT (SELECT count(*) FROM public.copy_studio_publications)=0, 'stale preview wrote receipt';
 dest := public.copy_studio_copy_bank_destination(agency,client,selection); expected := dest->>'version';
 -- A pending legacy approval would overwrite an append: publishing is blocked.
 INSERT INTO public.copy_bank_pending(key,content,status) VALUES(dest->>'key','{}','pending');
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,expected,'reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'code'='conflict', 'pending approval accepted';
 DELETE FROM public.copy_bank_pending WHERE key=dest->>'key';
 -- Profile changes after preview invalidate the compare-and-append token.
 UPDATE public.clients SET name='Fixture A renamed' WHERE id::text=client;
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,expected,'reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'code'='conflict', 'changed canonical profile accepted';
 dest := public.copy_studio_copy_bank_destination(agency,client,selection); expected := dest->>'version';
 -- A context refresh queued after the HTTP check is still rejected atomically.
 INSERT INTO public.copy_studio_jobs(id,agency_id,client_id,kind,payload,payload_hash,created_by,idempotency_key)
 VALUES('99999999-9999-9999-9999-999999999999',agency,client,'sync','{"source":"context"}',repeat('c',64),'author@example.test','context-test');
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,expected,'reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'code'='conflict', 'active context refresh accepted';
 UPDATE public.copy_studio_jobs SET status='succeeded' WHERE id='99999999-9999-9999-9999-999999999999';
 -- Valid service call preserves every unrelated section and custom field.
 SET LOCAL ROLE service_role;
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,expected,'publisher@example.test','66666666-6666-6666-6666-666666666666',snap);
 ASSERT result->>'published'='true' AND result->>'already_published'='false', 'valid publication rejected';
 ASSERT result->'receipt'->>'approved_by'='reviewer@example.test', 'approval actor not derived from saved review';
 ASSERT result->'receipt'->>'published_by'='publisher@example.test', 'publisher not audited';
 repeated := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,expected,'publisher@example.test',gen_random_uuid(),snap);
 ASSERT repeated->>'already_published'='true' AND repeated->>'publication_id'=result->>'publication_id', 'duplicate appended twice';
 RESET ROLE;
 SELECT content INTO checked FROM public.copy_bank_templates WHERE key=dest->>'key';
 ASSERT checked->'linkedin'=original->'linkedin' AND checked->'other'=original->'other', 'unrelated sections changed';
 ASSERT checked->'email'->>'custom'='preserve' AND checked->>'concurrent'='edit', 'unrelated attributes changed';
 ASSERT jsonb_array_length(checked->'email'->'subjects')=3 AND jsonb_array_length(checked->'email'->'variations')=2, 'append counts incorrect';
 ASSERT checked->'email'->'variations'->1->>'body'='Could simpler reporting help?', 'sent text changed';
 ASSERT (SELECT count(*) FROM public.copy_studio_publications)=1, 'duplicate audit receipt';
 BEGIN
   UPDATE public.copy_studio_publications SET published_by='changed';
   RAISE EXCEPTION 'receipt mutable';
 EXCEPTION WHEN raise_exception THEN IF SQLERRM <> 'Copy Studio evidence and reviews are immutable' THEN RAISE; END IF; END;
 -- Changed reference approval after snapshot prevents append to another section.
 selection := selection||'{"section":"chaser","step_indices":[1,2]}';
 dest := public.copy_studio_copy_bank_destination(agency,client,selection);
 INSERT INTO public.copy_studio_reviews(id,agency_id,client_id,record_id,decision,reviewed_by,created_at)
 VALUES(gen_random_uuid(),agency,client,'44444444-4444-4444-4444-444444444444','rejected','reviewer@example.test',clock_timestamp());
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,dest->>'version','reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'code'='conflict', 'revoked reference raced publication';
 snap := pg_temp.publish_snapshot(agency,client);
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,dest->>'version','reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'ok'='false', 'revoked reference accepted with refreshed snapshot';
 INSERT INTO public.copy_studio_reviews(id,agency_id,client_id,record_id,decision,reviewed_by,created_at)
 VALUES(gen_random_uuid(),agency,client,'44444444-4444-4444-4444-444444444444','approved','reviewer@example.test',clock_timestamp());
 snap := pg_temp.publish_snapshot(agency,client);
 -- Reject cross-channel and duplicated step requests at the database boundary.
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection||'{"section":"linkedin"}',dest->>'version','reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'ok'='false', 'email published to linkedin';
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection||'{"step_indices":[1,1]}',dest->>'version','reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'ok'='false', 'duplicate steps accepted';
 -- New destination is atomically created with only its selected section.
 selection := selection||'{"territory":"US"}';
 dest := public.copy_studio_copy_bank_destination(agency,client,selection);
 ASSERT dest->>'exists'='false', 'fixture new destination already exists';
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,dest->>'version','reviewer@example.test',gen_random_uuid(),snap);
 ASSERT result->>'published'='true', 'new destination not created';
 SELECT content INTO checked FROM public.copy_bank_templates WHERE key=dest->>'key';
 ASSERT checked ? 'chaser' AND NOT checked ? 'email' AND jsonb_array_length(checked->'chaser'->'steps')=2, 'new section incorrect';
 -- An edited descendant makes its previously approved parent stale.
 PERFORM pg_temp.add_record('77777777-7777-7777-7777-777777777777','draft','generation-1',value||'{"human_edited":true}',client,draft_id);
 selection := selection||'{"territory":"CA"}'; dest := public.copy_studio_copy_bank_destination(agency,client,selection);
 result := public.copy_studio_publish_copy_bank(agency,client,draft_id,selection,dest->>'version','reviewer@example.test',gen_random_uuid(),pg_temp.publish_snapshot(agency,client));
 ASSERT result->>'code'='conflict', 'superseded draft published';
 -- Invalid final-check snapshot cannot publish even when a service supplied approval.
 bad := jsonb_set(value,'{final_check,checked_angles}','[]');
 PERFORM pg_temp.add_record('88888888-8888-8888-8888-888888888888','draft','bad-check',bad);
 result := public.copy_studio_publish_copy_bank(agency,client,'88888888-8888-8888-8888-888888888888',selection,dest->>'version','reviewer@example.test',gen_random_uuid(),pg_temp.publish_snapshot(agency,client));
 ASSERT result->>'ok'='false', 'stale final check accepted';
 -- Selected legacy array normalises without erasing other sections or item metadata.
 checked := public.copy_studio_copy_bank_append('{"email":[{"subject":"Legacy","body":"Kept body","custom":"kept"}],"linkedin":{"steps":[]}}',
     value,'{"section":"email","angle_index":0,"step_indices":[0]}','{"type":"client"}');
 ASSERT checked->>'ok'='true' AND checked->'content'->'email'->'variations'->0->>'custom'='kept', 'legacy item metadata lost';
 ASSERT checked->'content'->'linkedin'='{"steps":[]}'::jsonb, 'legacy conversion changed another section';
 -- Reserved namespace may never be reached through a crafted BizDev profile.
 UPDATE public.copy_bank_templates SET content='[{"client_id":"aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa","type":"bizdev","territories":["__c__orphan"],"industries":["Private"]}]' WHERE key='__cb_profiles__';
 result := public.copy_studio_copy_bank_destination(agency,client,'{"territory":"__c__orphan","industry":"Private"}');
 ASSERT result->>'ok'='false', 'BizDev escaped into reserved namespace';
END;
$$;
ROLLBACK;
SELECT 'PASS: private publishing RPC, exact draft/context approval, profile scoping, CAS, idempotency, immutable receipts and isolated append' AS result;
