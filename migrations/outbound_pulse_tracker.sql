-- ============================================================================
-- Client Performance & Reports — the per-client ICP Performance Tracker sheet
--
-- Run once in the Supabase SQL editor, after outbound_pulse_copy_variants.sql.
-- Safe to re-run.
--
-- WHY THIS EXISTS
-- Every client has a Google Sheet an account manager maintains by hand, whose
-- "All leads" tab grades each lead against that client's ICP: headcount, job
-- role, industry, location and a 1-10 rating. Nothing in this database knows
-- where that sheet is, so it has to be pointed at once, from the client page.
--
-- WHY NOT A COLUMN ON clients
-- `clients` is defined by dnc_schema.sql and shared with the DNC and Client
-- Profiles tools. Both of its readers (app/routers/profiles.py and
-- store.list_clients) ask for an optional column list and fall back to
-- "id,name" when that select fails. So a column added there is invisible until
-- both fallback chains are edited — and if one is missed the failure is
-- SILENT: the whole first select fails, taking the branding columns down with
-- it, and nothing errors. A link that quietly stops existing is worse than no
-- link. A separate table also gets a hard agency_id, which `clients` cannot
-- have because its rows predate agency scoping.
--
-- WHY IT IS MORE THAN A URL
-- The tab name is a column because a client who renames "All leads" should be
-- a settings change and not a deploy. The last read's outcome is a column so
-- the panel can say "this sheet stopped being readable on the 14th" instead of
-- failing afresh on every page load, and so a 403 can name the service account
-- the sheet needs sharing with.
--
-- WHY ONE ROW PER CLIENT, AND WHY UNLINKING DELETES IT
-- No row means "never linked", which is a different state from a linked sheet
-- that happens to be unreadable — the first wants a form, the second wants an
-- error. Blanking sheet_id would invent a third state nothing reads, so
-- unlinking removes the row. Same rule the per-user view preferences follow in
-- outbound_pulse_user_prefs.sql.
--
-- WHAT THIS TABLE DELIBERATELY DOES NOT HOLD
-- The leads themselves. The sheet is hand-maintained and is the truth; a copy
-- in here would be a second truth that drifts, and would need a sync job, a
-- staleness indicator and a reconciliation story for a few dozen rows per
-- client. Reports freeze their figures in pulse_reports.snapshot, which is the
-- mechanism that already exists for exactly this.
-- ============================================================================

BEGIN;

CREATE TABLE IF NOT EXISTS pulse_client_trackers (
    id              uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id       uuid NOT NULL REFERENCES agencies(id) ON DELETE CASCADE,
    client_id       uuid NOT NULL REFERENCES clients(id)  ON DELETE CASCADE,

    -- The spreadsheet key, extracted from whatever was pasted. Stored apart
    -- from the URL so that re-pasting the same sheet with a different #gid or
    -- query string is recognised as the same sheet.
    sheet_id        text NOT NULL,
    -- Exactly what was pasted, so "Open sheet" can link to it without us
    -- rebuilding a URL we were never given.
    sheet_url       text NOT NULL DEFAULT '',
    -- The data tab. Defaults to the name every tracker uses today.
    -- NEVER point this at "Raw Leads": that tab is locked, has a different
    -- schema, and is not ours to read.
    tab_title       text NOT NULL DEFAULT 'All leads',

    linked_by       text NOT NULL DEFAULT '',
    -- The outcome of the most recent read, for the panel's status line. Empty
    -- means the last read succeeded.
    last_error      text NOT NULL DEFAULT '',
    last_checked_at timestamptz,

    created_at      timestamptz NOT NULL DEFAULT now(),
    updated_at      timestamptz NOT NULL DEFAULT now()
);

-- One tracker per client, so re-linking corrects the sheet rather than leaving
-- two for the app to choose between. Leads with agency_id because this is the
-- on_conflict target.
CREATE UNIQUE INDEX IF NOT EXISTS pulse_client_trackers_unique
    ON pulse_client_trackers (agency_id, client_id);

-- A row naming no spreadsheet is the unlinked state, and the unlinked state is
-- the absence of a row. It must not be possible to write one.
ALTER TABLE pulse_client_trackers DROP CONSTRAINT IF EXISTS pulse_client_trackers_has_sheet;
ALTER TABLE pulse_client_trackers
    ADD CONSTRAINT pulse_client_trackers_has_sheet
    CHECK (length(btrim(sheet_id)) > 0 AND length(btrim(tab_title)) > 0);

ALTER TABLE pulse_client_trackers DISABLE ROW LEVEL SECURITY;

DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'anon') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_client_trackers TO anon;
    END IF;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'authenticated') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_client_trackers TO authenticated;
    END IF;
END;
$$;

COMMIT;
