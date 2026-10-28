-- Reviewed Copy Studio -> Copy Bank append. Apply after ai_copy_studio.sql.
-- Existing Copy Bank grants/RLS and editing routes are not changed.
-- Review in staging first. Re-runnable. Rollback: disable publishing routes and
-- revoke the publish RPC; retain immutable receipts and published copy for audit.
BEGIN;

CREATE TABLE IF NOT EXISTS public.copy_studio_publications (
    id uuid PRIMARY KEY,
    agency_id uuid NOT NULL REFERENCES public.agencies(id),
    client_id text NOT NULL,
    draft_id uuid NOT NULL,
    destination_key text NOT NULL,
    selection jsonb NOT NULL,
    receipt jsonb NOT NULL,
    published_by text NOT NULL,
    published_at timestamptz NOT NULL DEFAULT now(),
    FOREIGN KEY (agency_id, client_id, draft_id)
        REFERENCES public.copy_studio_records(agency_id, client_id, id),
    UNIQUE (agency_id, client_id, draft_id, destination_key, selection)
);
ALTER TABLE public.copy_studio_publications ENABLE ROW LEVEL SECURITY;
-- Reset Supabase's service_role default ALL grant too: receipts may only be
-- inserted by the transaction-owning SECURITY DEFINER publication function.
REVOKE ALL ON public.copy_studio_publications FROM PUBLIC, anon, authenticated, service_role;
GRANT SELECT ON public.copy_studio_publications TO service_role;
DROP TRIGGER IF EXISTS copy_studio_publications_immutable ON public.copy_studio_publications;
CREATE TRIGGER copy_studio_publications_immutable BEFORE UPDATE OR DELETE ON public.copy_studio_publications
    FOR EACH ROW EXECUTE FUNCTION public.copy_studio_immutable();

CREATE OR REPLACE FUNCTION public.copy_studio_copy_bank_destinations(p_agency_id uuid, p_client_id text)
RETURNS jsonb LANGUAGE plpgsql SECURITY DEFINER SET search_path = pg_catalog, public AS $$
DECLARE directory jsonb; profile jsonb; matches integer; industries jsonb;
BEGIN
    IF NOT EXISTS (SELECT 1 FROM public.clients WHERE id::text = p_client_id AND agency_id = p_agency_id) THEN
        RETURN jsonb_build_object('ok', false, 'message', 'The client does not belong to this agency.');
    END IF;
    SELECT content INTO directory FROM public.copy_bank_templates WHERE key = '__cb_profiles__';
    IF jsonb_typeof(directory) IS DISTINCT FROM 'array' THEN
        RETURN jsonb_build_object('ok', false, 'message', 'Configure this client in Copy Bank before publishing.');
    END IF;
    SELECT count(*) INTO matches FROM jsonb_array_elements(directory) p WHERE p->>'client_id' = p_client_id;
    IF matches <> 1 THEN
        RETURN jsonb_build_object('ok', false, 'message', 'Copy Bank needs one exact profile for this client.');
    END IF;
    SELECT p INTO profile FROM jsonb_array_elements(directory) p WHERE p->>'client_id' = p_client_id;
    IF profile->>'type' NOT IN ('client', 'bizdev') OR profile->>'type' IS NULL
       OR jsonb_typeof(profile->'territories') IS DISTINCT FROM 'array'
       OR jsonb_typeof(profile->'industries') IS DISTINCT FROM 'array'
       OR EXISTS (SELECT 1 FROM jsonb_array_elements(profile->'territories') x
                  WHERE jsonb_typeof(x) <> 'string' OR length(trim(x #>> '{}')) = 0)
       OR EXISTS (SELECT 1 FROM jsonb_array_elements(profile->'industries') x
                  WHERE jsonb_typeof(x) <> 'string' OR length(trim(x #>> '{}')) = 0) THEN
        RETURN jsonb_build_object('ok', false, 'message', 'The Copy Bank profile needs valid territory and industry choices.');
    END IF;
    SELECT coalesce(jsonb_agg(jsonb_build_object('value', x, 'label', coalesce(profile->'industryLabels'->>x, x))), '[]')
    INTO industries FROM jsonb_array_elements_text(profile->'industries') x;
    RETURN jsonb_build_object('ok', true, 'profile', profile,
        'territories', profile->'territories', 'industries', industries);
END;
$$;

CREATE OR REPLACE FUNCTION public.copy_studio_copy_bank_destination(
    p_agency_id uuid, p_client_id text, p_selection jsonb)
RETURNS jsonb LANGUAGE plpgsql SECURITY DEFINER SET search_path = pg_catalog, public AS $$
DECLARE result jsonb; profile jsonb; directory jsonb; client_row jsonb; destination text;
        existing jsonb; combinations integer; has_pending boolean;
BEGIN
    result := public.copy_studio_copy_bank_destinations(p_agency_id, p_client_id);
    IF result->>'ok' = 'false' THEN RETURN result; END IF;
    profile := result->'profile';
    IF NOT (profile->'territories' ? (p_selection->>'territory'))
       OR NOT (profile->'industries' ? (p_selection->>'industry'))
       OR p_selection->>'territory' IS NULL OR p_selection->>'industry' IS NULL THEN
        RETURN jsonb_build_object('ok', false, 'message', 'Choose a territory and industry from this client''s Copy Bank profile.');
    END IF;
    destination := CASE WHEN profile->>'type' = 'bizdev' THEN '' ELSE '__c__'||p_client_id||'__' END
                   ||(p_selection->>'territory')||'_'||(p_selection->>'industry');
    IF profile->>'type' = 'bizdev' AND left(destination, 2) = '__' THEN
        RETURN jsonb_build_object('ok', false, 'message', 'BizDev copy cannot use a reserved or client-scoped destination key.');
    END IF;
    SELECT content INTO directory FROM public.copy_bank_templates WHERE key = '__cb_profiles__';
    -- Underscores in legacy keys can make two distinct combinations collide.
    -- Shared BizDev namespaces additionally must have one unambiguous owner.
    SELECT count(*) INTO combinations FROM jsonb_array_elements(directory) p,
        LATERAL jsonb_array_elements_text(CASE WHEN jsonb_typeof(p->'territories') = 'array' THEN p->'territories' ELSE '[]' END) t,
        LATERAL jsonb_array_elements_text(CASE WHEN jsonb_typeof(p->'industries') = 'array' THEN p->'industries' ELSE '[]' END) i
    WHERE (CASE WHEN p->>'type' = 'bizdev' THEN '' ELSE '__c__'||(p->>'client_id')||'__' END)||t||'_'||i = destination;
    IF combinations <> 1 THEN
        RETURN jsonb_build_object('ok', false, 'message', 'This Copy Bank destination is ambiguous. Correct its profile mapping first.');
    END IF;
    SELECT content INTO existing FROM public.copy_bank_templates WHERE key = destination;
    IF existing IS NOT NULL AND jsonb_typeof(existing) <> 'object' THEN
        RETURN jsonb_build_object('ok', false, 'message', 'The Copy Bank destination has an unsupported format.');
    END IF;
    SELECT to_jsonb(c) INTO client_row FROM public.clients c WHERE id::text = p_client_id AND agency_id = p_agency_id;
    SELECT EXISTS (SELECT 1 FROM public.copy_bank_pending WHERE key = destination AND status = 'pending') INTO has_pending;
    RETURN jsonb_build_object('ok', true, 'key', destination, 'profile', profile,
        'content', coalesce(existing, '{}'::jsonb), 'exists', existing IS NOT NULL, 'pending', has_pending,
        'version', md5(coalesce(existing::text, 'missing')||chr(31)||directory::text||chr(31)||client_row::text));
END;
$$;

-- Narrow, deterministic helper. It derives append text from the immutable draft,
-- never from a browser-provided body. Only the selected section is normalised.
CREATE OR REPLACE FUNCTION public.copy_studio_copy_bank_append(
    p_content jsonb, p_draft jsonb, p_selection jsonb, p_profile jsonb)
RETURNS jsonb LANGUAGE plpgsql SET search_path = pg_catalog, public AS $$
DECLARE section text := p_selection->>'section'; channel text; angle jsonb; step jsonb;
        angle_index integer; indices jsonb; index_value jsonb; step_index integer; previous_index integer := -1;
        subjects jsonb := '[]'; steps jsonb := '[]'; target jsonb; variations jsonb; old_subjects jsonb;
        allowed text[]; item jsonb;
BEGIN
    channel := coalesce(p_draft->>'channel', p_draft->'request'->>'channel', p_draft->'studio_context'->'request'->>'channel');
    IF (channel = 'email' AND section NOT IN ('email', 'flyout', 'chaser'))
       OR (channel = 'linkedin' AND section <> 'linkedin') OR channel NOT IN ('email', 'linkedin')
       OR channel IS NULL OR section IS NULL
       OR (section = 'flyout' AND p_profile->>'type' IS DISTINCT FROM 'bizdev') THEN
        RETURN jsonb_build_object('ok', false, 'message', 'Draft channel and Copy Bank section do not match.');
    END IF;
    IF coalesce(p_selection->>'angle_index', '') !~ '^[0-9]{1,6}$'
       OR jsonb_typeof(p_selection->'angle_index') IS DISTINCT FROM 'number'
       OR jsonb_typeof(p_draft->'angles') IS DISTINCT FROM 'array' THEN
        RETURN jsonb_build_object('ok', false, 'message', 'Select one saved angle.');
    END IF;
    angle_index := (p_selection->>'angle_index')::integer;
    angle := p_draft->'angles'->angle_index;
    indices := p_selection->'step_indices';
    IF angle IS NULL OR jsonb_typeof(angle->'steps') IS DISTINCT FROM 'array'
       OR jsonb_typeof(indices) IS DISTINCT FROM 'array' OR indices = '[]'::jsonb THEN
        RETURN jsonb_build_object('ok', false, 'message', 'Select compatible saved message steps.');
    END IF;
    allowed := CASE WHEN section IN ('email', 'flyout') THEN ARRAY['initial']
                    WHEN section = 'chaser' THEN ARRAY['follow_up', 'closing']
                    ELSE ARRAY['connection', 'initial', 'follow_up', 'closing'] END;
    FOR index_value IN SELECT value FROM jsonb_array_elements(indices) LOOP
        IF jsonb_typeof(index_value) <> 'number' OR (index_value #>> '{}') !~ '^[0-9]{1,6}$' THEN
            RETURN jsonb_build_object('ok', false, 'message', 'Saved step indices must be integers.');
        END IF;
        step_index := (index_value #>> '{}')::integer;
        step := angle->'steps'->step_index;
        IF step_index <= previous_index OR step IS NULL OR NOT ((step->>'kind') = ANY(allowed))
           OR step->>'kind' IS NULL OR jsonb_typeof(step->'body') IS DISTINCT FROM 'string'
           OR length(trim(step->>'body')) = 0 THEN
            RETURN jsonb_build_object('ok', false, 'message', 'The selected steps are empty, duplicated or incompatible.');
        END IF;
        previous_index := step_index;
        steps := steps || jsonb_build_array(jsonb_build_object('kind', step->>'kind', 'body', step->>'body'));
    END LOOP;
    IF section IN ('email', 'flyout') THEN
        subjects := coalesce(angle->'subjects', '[]'::jsonb);
        IF jsonb_typeof(subjects) <> 'array' OR EXISTS (SELECT 1 FROM jsonb_array_elements(subjects) x
            WHERE jsonb_typeof(x) <> 'string' OR length(trim(x #>> '{}')) = 0) THEN
            RETURN jsonb_build_object('ok', false, 'message', 'Saved subject variants must contain text.');
        END IF;
    END IF;
    target := coalesce(p_content->section, '{}'::jsonb);
    IF jsonb_typeof(target) = 'array' THEN
        IF EXISTS (SELECT 1 FROM jsonb_array_elements(target) x WHERE jsonb_typeof(x) <> 'object'
              OR (x ? 'body' AND jsonb_typeof(x->'body') <> 'string')
              OR (x ? 'subject' AND jsonb_typeof(x->'subject') <> 'string')) THEN
            RETURN jsonb_build_object('ok', false, 'message', 'The existing destination section needs manual format review.');
        END IF;
        SELECT coalesce(jsonb_agg(x->'subject') FILTER (WHERE coalesce(x->>'subject', '') <> ''), '[]'),
               coalesce(jsonb_agg(x - 'subject'), '[]') INTO old_subjects, variations FROM jsonb_array_elements(target) x;
        target := CASE WHEN section IN ('email', 'flyout')
            THEN jsonb_build_object('subjects', old_subjects, 'variations', variations)
            ELSE jsonb_build_object('steps', variations) END;
    ELSIF jsonb_typeof(target) <> 'object' THEN
        RETURN jsonb_build_object('ok', false, 'message', 'The existing destination section needs manual format review.');
    END IF;
    IF section IN ('email', 'flyout') THEN
        IF jsonb_typeof(coalesce(target->'subjects', '[]')) <> 'array'
           OR jsonb_typeof(coalesce(target->'variations', '[]')) <> 'array' THEN
            RETURN jsonb_build_object('ok', false, 'message', 'The existing email section needs manual format review.');
        END IF;
        target := target || jsonb_build_object('subjects', coalesce(target->'subjects', '[]')||subjects,
            'variations', coalesce(target->'variations', '[]')||steps);
    ELSE
        IF jsonb_typeof(coalesce(target->'steps', '[]')) <> 'array' THEN
            RETURN jsonb_build_object('ok', false, 'message', 'The existing sequence section needs manual format review.');
        END IF;
        target := target || jsonb_build_object('steps', coalesce(target->'steps', '[]')||steps);
    END IF;
    RETURN jsonb_build_object('ok', true, 'content', p_content||jsonb_build_object(section, target),
        'appended_copy', jsonb_build_object('subjects', subjects, 'steps', steps));
END;
$$;

CREATE OR REPLACE FUNCTION public.copy_studio_publish_copy_bank(
    p_agency_id uuid, p_client_id text, p_draft_id uuid, p_selection jsonb,
    p_expected_version text, p_actor text, p_publication_id uuid, p_context_snapshot jsonb)
RETURNS jsonb LANGUAGE plpgsql SECURITY DEFINER SET search_path = pg_catalog, public AS $$
DECLARE draft public.copy_studio_record_status%ROWTYPE; latest_id uuid; check_result jsonb;
        destination jsonb; appended jsonb; actual_snapshot jsonb; receipt jsonb; previous_receipt jsonb;
        key_value text; updated_content jsonb; inserted_key text; final_check jsonb; normalized_selection jsonb;
        directory jsonb; client_row jsonb; after_version text; published_stamp timestamptz := clock_timestamp();
BEGIN
    IF p_actor IS NULL OR length(trim(p_actor)) NOT BETWEEN 1 AND 320 THEN
        RETURN jsonb_build_object('ok', false, 'message', 'An authenticated reviewer identity is required.');
    END IF;
    -- Short-lived table locks are deliberate: existing save/review routes insert
    -- directly, so an advisory lock alone cannot stop concurrent revocations or
    -- new source versions. No network/model call occurs inside this transaction.
    LOCK TABLE public.copy_studio_records, public.copy_studio_reviews, public.copy_studio_jobs IN SHARE MODE;
    PERFORM 1 FROM public.clients WHERE id::text = p_client_id AND agency_id = p_agency_id FOR SHARE;
    IF NOT FOUND THEN
        RETURN jsonb_build_object('ok', false, 'message', 'The client does not belong to this agency.');
    END IF;
    SELECT * INTO draft FROM public.copy_studio_record_status
        WHERE id = p_draft_id AND agency_id = p_agency_id AND client_id = p_client_id AND kind = 'draft';
    IF NOT FOUND OR draft.status <> 'approved' THEN
        RETURN jsonb_build_object('ok', false, 'message', 'Approve this exact saved draft before publishing.');
    END IF;
    SELECT id INTO latest_id FROM public.copy_studio_records
        WHERE agency_id = p_agency_id AND client_id = p_client_id AND kind = 'draft' AND source_key = draft.source_key
        ORDER BY created_at DESC, id DESC LIMIT 1;
    IF latest_id <> p_draft_id THEN
        RETURN jsonb_build_object('ok', false, 'code', 'conflict', 'message', 'A newer draft version exists. Review and preview that version.');
    END IF;
    final_check := draft.content->'final_check';
    IF final_check->'complete' IS DISTINCT FROM 'true'::jsonb
       OR final_check->'checked_angles' IS DISTINCT FROM draft.content->'angles'
       OR final_check->'report'->'complete' IS DISTINCT FROM 'true'::jsonb
       OR final_check->>'locale' NOT IN ('en-GB', 'en-US') OR final_check->>'locale' IS NULL
       OR final_check->>'locale' IS DISTINCT FROM final_check->'report'->>'locale'
       OR final_check->>'locale' IS DISTINCT FROM draft.content->'request'->>'locale'
       OR coalesce(final_check->>'copy_hash', '') !~ '^[a-f0-9]{64}$'
       OR final_check->>'copy_hash' IS DISTINCT FROM final_check->'report'->>'copy_hash'
       OR draft.content->'review'->'can_approve' IS DISTINCT FROM 'true'::jsonb
       OR jsonb_typeof(final_check->'report'->'issues') IS DISTINCT FROM 'array'
       OR EXISTS (SELECT 1 FROM jsonb_array_elements(final_check->'report'->'issues') x WHERE x->>'severity' IS DISTINCT FROM 'warning')
       OR EXISTS (SELECT 1 FROM jsonb_array_elements(coalesce(draft.content->'review'->'issues', '[]')) x WHERE x->>'severity' = 'error') THEN
        RETURN jsonb_build_object('ok', false, 'message', 'Complete the final checks on this exact saved draft before publishing.');
    END IF;
    SELECT coalesce(jsonb_object_agg(r.id::text, jsonb_build_object('status', r.status, 'review_id', r.review->>'id')), '{}')
    INTO actual_snapshot FROM public.copy_studio_record_status r
    WHERE r.agency_id = p_agency_id AND
        ((r.client_id = p_client_id AND r.kind IN ('draft', 'brief', 'example', 'transcript', 'sync_report'))
         OR (r.client_id = '__agency__' AND r.kind = 'methodology'
             AND r.id::text = draft.content->'studio_context'->>'methodology_id'));
    IF actual_snapshot IS DISTINCT FROM p_context_snapshot THEN
        RETURN jsonb_build_object('ok', false, 'code', 'conflict', 'message', 'Client evidence or approvals changed. Refresh and preview again.');
    END IF;
    IF EXISTS (SELECT 1 FROM public.copy_studio_jobs WHERE agency_id = p_agency_id AND client_id = p_client_id
               AND kind = 'sync' AND status IN ('queued', 'running') AND payload->>'source' IN ('all', 'context', 'onboarding')) THEN
        RETURN jsonb_build_object('ok', false, 'code', 'conflict', 'message', 'Client context is refreshing. Wait for it to finish and preview again.');
    END IF;
    -- Independently verify direct dependencies even for a service RPC caller.
    IF NOT EXISTS (SELECT 1 FROM public.copy_studio_record_status r WHERE r.agency_id = p_agency_id
        AND r.client_id = p_client_id AND r.kind = 'brief' AND r.status = 'approved'
        AND r.id::text = draft.content->'studio_context'->>'brief_id')
       OR EXISTS (SELECT 1 FROM jsonb_array_elements_text(coalesce(draft.content->'studio_context'->'example_ids', '[]')) identifier
          WHERE NOT EXISTS (SELECT 1 FROM public.copy_studio_record_status r WHERE r.agency_id = p_agency_id
            AND r.client_id = p_client_id AND r.kind = 'example' AND r.status = 'approved' AND r.id::text = identifier))
       OR (draft.content->'studio_context'->>'methodology_id' IS NOT NULL AND NOT EXISTS (
          SELECT 1 FROM public.copy_studio_record_status r WHERE r.agency_id = p_agency_id AND r.client_id = '__agency__'
            AND r.kind = 'methodology' AND r.status = 'approved' AND r.id::text = draft.content->'studio_context'->>'methodology_id')) THEN
        RETURN jsonb_build_object('ok', false, 'message', 'A required client brief, reference or methodology is no longer approved.');
    END IF;
    SELECT id INTO latest_id FROM public.copy_studio_record_status WHERE agency_id = p_agency_id
        AND client_id = p_client_id AND kind = 'brief' AND status = 'approved' ORDER BY created_at DESC, id DESC LIMIT 1;
    IF latest_id::text IS DISTINCT FROM draft.content->'studio_context'->>'brief_id' THEN
        RETURN jsonb_build_object('ok', false, 'message', 'The approved client brief has changed.');
    END IF;
    PERFORM 1 FROM public.copy_bank_templates WHERE key = '__cb_profiles__' FOR SHARE;
    -- Whole-content pending approvals would overwrite an append when approved.
    -- Block new pending submissions just for this short transaction.
    LOCK TABLE public.copy_bank_pending IN SHARE MODE;
    destination := public.copy_studio_copy_bank_destination(p_agency_id, p_client_id, p_selection);
    IF destination->>'ok' = 'false' THEN RETURN destination; END IF;
    key_value := destination->>'key';
    normalized_selection := jsonb_build_object('territory', p_selection->'territory', 'industry', p_selection->'industry',
        'section', p_selection->'section', 'angle_index', p_selection->'angle_index', 'step_indices', p_selection->'step_indices');
    SELECT p.receipt INTO previous_receipt FROM public.copy_studio_publications p
        WHERE p.agency_id = p_agency_id AND p.client_id = p_client_id AND p.draft_id = p_draft_id
          AND p.destination_key = key_value AND p.selection = normalized_selection;
    IF FOUND THEN
        RETURN previous_receipt||jsonb_build_object('already_published', true);
    END IF;
    -- Serialize same-destination bridge inserts, and also lock existing rows so
    -- legacy editors (which do not take advisory locks) participate in the CAS.
    PERFORM pg_advisory_xact_lock(hashtextextended('copy-studio-publish:'||key_value, 0));
    PERFORM 1 FROM public.copy_bank_templates WHERE key = key_value FOR UPDATE;
    -- Recheck receipt after waiting for a concurrent identical publication.
    SELECT p.receipt INTO previous_receipt FROM public.copy_studio_publications p
        WHERE p.agency_id = p_agency_id AND p.client_id = p_client_id AND p.draft_id = p_draft_id
          AND p.destination_key = key_value AND p.selection = normalized_selection;
    IF FOUND THEN RETURN previous_receipt||jsonb_build_object('already_published', true); END IF;
    destination := public.copy_studio_copy_bank_destination(p_agency_id, p_client_id, p_selection);
    IF destination->>'ok' = 'false' THEN RETURN destination; END IF;
    IF destination->>'version' IS DISTINCT FROM p_expected_version OR destination->>'pending' = 'true' THEN
        RETURN jsonb_build_object('ok', false, 'code', 'conflict', 'message', 'The destination or client profile changed, or has a pending approval. Preview again.');
    END IF;
    appended := public.copy_studio_copy_bank_append(destination->'content', draft.content, normalized_selection, destination->'profile');
    IF appended->>'ok' = 'false' THEN RETURN appended; END IF;
    updated_content := appended->'content';
    IF destination->>'exists' = 'true' THEN
        UPDATE public.copy_bank_templates SET content = updated_content WHERE key = key_value;
    ELSE
        INSERT INTO public.copy_bank_templates(key, content) VALUES (key_value, updated_content)
            ON CONFLICT (key) DO NOTHING RETURNING key INTO inserted_key;
        IF inserted_key IS NULL THEN
            RETURN jsonb_build_object('ok', false, 'code', 'conflict', 'message', 'Another editor created this destination. Preview it again.');
        END IF;
    END IF;
    SELECT content INTO directory FROM public.copy_bank_templates WHERE key = '__cb_profiles__';
    SELECT to_jsonb(c) INTO client_row FROM public.clients c WHERE id::text = p_client_id AND agency_id = p_agency_id;
    after_version := md5(updated_content::text||chr(31)||directory::text||chr(31)||client_row::text);
    receipt := jsonb_build_object('published', true, 'already_published', false,
        'publication_id', p_publication_id, 'draft_id', p_draft_id, 'destination_key', key_value,
        'section', p_selection->>'section', 'published_at', published_stamp, 'copy_bank_url', '/copy-bank',
        'receipt', jsonb_build_object('selection', normalized_selection, 'draft_hash', draft.content_hash,
            'checked_copy_hash', final_check->>'copy_hash', 'locale', final_check->>'locale',
            'approved_by', draft.review->>'reviewed_by', 'published_by', p_actor,
            'before_version', p_expected_version, 'after_version', after_version, 'appended_copy', appended->'appended_copy'));
    INSERT INTO public.copy_studio_publications(id, agency_id, client_id, draft_id, destination_key, selection, receipt, published_by, published_at)
        VALUES (p_publication_id, p_agency_id, p_client_id, p_draft_id, key_value, normalized_selection, receipt, p_actor, published_stamp);
    RETURN receipt;
END;
$$;

REVOKE ALL ON FUNCTION public.copy_studio_copy_bank_destinations(uuid,text) FROM PUBLIC, anon, authenticated, service_role;
REVOKE ALL ON FUNCTION public.copy_studio_copy_bank_destination(uuid,text,jsonb) FROM PUBLIC, anon, authenticated, service_role;
REVOKE ALL ON FUNCTION public.copy_studio_copy_bank_append(jsonb,jsonb,jsonb,jsonb) FROM PUBLIC, anon, authenticated, service_role;
REVOKE ALL ON FUNCTION public.copy_studio_publish_copy_bank(uuid,text,uuid,jsonb,text,text,uuid,jsonb) FROM PUBLIC, anon, authenticated, service_role;
GRANT EXECUTE ON FUNCTION public.copy_studio_copy_bank_destinations(uuid,text) TO service_role;
GRANT EXECUTE ON FUNCTION public.copy_studio_copy_bank_destination(uuid,text,jsonb) TO service_role;
GRANT EXECUTE ON FUNCTION public.copy_studio_publish_copy_bank(uuid,text,uuid,jsonb,text,text,uuid,jsonb) TO service_role;
NOTIFY pgrst, 'reload schema';
COMMIT;
