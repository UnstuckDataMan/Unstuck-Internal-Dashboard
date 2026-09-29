-- ============================================================================
-- Outbound Pulse — account manager notes on a client's report
--
-- Run once in the Supabase SQL editor. Safe to re-run.
--
-- WHY A TABLE AND NOT A COLUMN ON clients
-- A note is a dated entry with an author, and a client accumulates them over
-- reporting cycles. A single text column would overwrite last month's context
-- every time an account manager wrote this month's, and would not record who
-- said it. Keeping them as rows also means the client-facing report can show
-- the latest few while the internal view keeps the whole history.
--
-- VISIBILITY
-- show_on_report defaults to true, because the reason to write one is to say
-- something to the client. An internal-only note is the exception, so it is
-- the box you tick rather than the box you clear. Nothing here is ever shown
-- to a client without an account manager having written it for them.
-- ============================================================================

BEGIN;

CREATE TABLE IF NOT EXISTS pulse_client_notes (
    id              uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id       uuid NOT NULL REFERENCES agencies(id) ON DELETE CASCADE,
    client_id       uuid NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    body            text NOT NULL,
    author_email    text,
    author_name     text,
    show_on_report  boolean NOT NULL DEFAULT true,
    created_at      timestamptz NOT NULL DEFAULT now(),
    updated_at      timestamptz NOT NULL DEFAULT now()
);

-- A note with no text is not a note. The check is here rather than only in the
-- router because the portal renders this straight onto a client's page.
ALTER TABLE pulse_client_notes DROP CONSTRAINT IF EXISTS pulse_client_notes_body_not_blank;
ALTER TABLE pulse_client_notes
    ADD CONSTRAINT pulse_client_notes_body_not_blank
    CHECK (length(btrim(body)) > 0);

-- Both readers order by client then recency: the internal list for one client,
-- and the portal's "latest few that are visible".
CREATE INDEX IF NOT EXISTS pulse_client_notes_client_idx
    ON pulse_client_notes (agency_id, client_id, created_at DESC);


DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'anon') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_client_notes TO anon;
    END IF;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'authenticated') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_client_notes TO authenticated;
    END IF;
END;
$$;

COMMIT;
