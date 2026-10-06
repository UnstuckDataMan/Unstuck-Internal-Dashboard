-- ============================================================================
-- Client Performance & Reports — copy recorded against a variant by hand
--
-- Run once in the Supabase SQL editor. Safe to re-run.
--
-- WHY THIS EXISTS
-- Smartlead's sequence-analytics endpoint reports a subject line per sequence
-- STEP, not per variant, and nothing at all for a manual campaign — there the
-- variant is a Copy Bank combination like "S2/B1", which names an index rather
-- than any text. So when a variant wins, we frequently know that it won and
-- not what it said.
--
-- That is the gap this fills: an account manager writes the subject and body
-- in once, against the variant, and both the internal panel and the client's
-- report can then show the message that worked rather than a letter.
--
-- WHY NOT STORE IT ON THE A/B RESULT
-- The result is computed live from Smartlead and from the campaign sheets and
-- is never stored. Copy is the opposite: entered once and true until someone
-- rewrites it. They have different lifetimes, so they are different rows.
--
-- WHY campaign_ref IS TEXT AND MAY BE EMPTY
-- A Smartlead variant belongs to one campaign; a manual Copy Bank combination
-- is reused across every campaign drawing the same copy. An empty campaign_ref
-- means "this variant, for this client, wherever it ran" — which is the honest
-- shape of a Copy Bank variant and avoids forcing the same text to be typed
-- once per campaign.
-- ============================================================================

BEGIN;

CREATE TABLE IF NOT EXISTS pulse_copy_variants (
    id           uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id    uuid NOT NULL REFERENCES agencies(id) ON DELETE CASCADE,
    client_id    uuid NOT NULL REFERENCES clients(id)  ON DELETE CASCADE,

    -- 'smartlead' or 'manual'. Which tool's idea of a variant this names.
    source_tool  text NOT NULL,
    -- '' for a variant that is not campaign-specific. See the note above.
    campaign_ref text NOT NULL DEFAULT '',
    -- 'A', 'B', 'S2/B1' — whatever that source calls the variant.
    variant_key  text NOT NULL,

    subject      text NOT NULL DEFAULT '',
    -- Sanitised HTML, cleaned by app/utils/pulse/richtext.py on the way in.
    -- This renders on a client-facing page, so it is never stored raw.
    body         text NOT NULL DEFAULT '',

    entered_by   text NOT NULL DEFAULT '',
    created_at   timestamptz NOT NULL DEFAULT now(),
    updated_at   timestamptz NOT NULL DEFAULT now()
);

-- One record per variant, so re-entering corrects it rather than leaving two
-- versions of the same message to choose between.
CREATE UNIQUE INDEX IF NOT EXISTS pulse_copy_variants_unique
    ON pulse_copy_variants (agency_id, client_id, source_tool, campaign_ref, variant_key);

CREATE INDEX IF NOT EXISTS pulse_copy_variants_client_idx
    ON pulse_copy_variants (agency_id, client_id);

-- A record with neither a subject nor a body says nothing; the panel treats
-- that as "no copy recorded" anyway, so it should not reach the table.
ALTER TABLE pulse_copy_variants DROP CONSTRAINT IF EXISTS pulse_copy_variants_not_blank;
ALTER TABLE pulse_copy_variants
    ADD CONSTRAINT pulse_copy_variants_not_blank
    CHECK (length(btrim(subject)) > 0 OR length(btrim(body)) > 0);

ALTER TABLE pulse_copy_variants DISABLE ROW LEVEL SECURITY;


DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'anon') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_copy_variants TO anon;
    END IF;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'authenticated') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_copy_variants TO authenticated;
    END IF;
END;
$$;

COMMIT;
