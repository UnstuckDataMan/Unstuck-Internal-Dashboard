-- Unique index on contacted_prospects so that
-- PostgREST's "resolution=ignore-duplicates" (ON CONFLICT DO NOTHING) works.
--
-- NULL campaign_name is treated as '' so two rows with the same
-- client_id + email and no campaign are considered duplicates.
-- (PostgreSQL treats NULLs as distinct in regular unique constraints,
--  which would silently allow unlimited re-uploads of the same email.)

CREATE UNIQUE INDEX IF NOT EXISTS uq_contacted_prospects_client_email_campaign
    ON contacted_prospects (client_id, email, COALESCE(campaign_name, ''));
