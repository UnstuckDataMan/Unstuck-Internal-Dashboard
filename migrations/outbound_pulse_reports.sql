-- ============================================================================
-- Outbound Pulse — published client reports
--
-- Run once in the Supabase SQL editor, after outbound_pulse_unsubscribes.sql.
-- Safe to re-run.
--
-- WHAT CHANGES
-- The client portal stops being a live dashboard with range buttons and
-- becomes a series of reports an account manager publishes. One link per
-- client, for good; the link opens the newest published report and the client
-- pages back through earlier ones.
--
-- WHY THE NUMBERS ARE FROZEN
-- `snapshot` holds the figures as they stood when the report was published. It
-- is not a cache — it is the report. A published report that re-queried live
-- would change after the client read it, because a later sync can still add
-- events inside a closed period, and "the number moved since you sent it" is
-- exactly the conversation this module exists to avoid. Re-publishing is
-- deliberate and recomputes it.
--
-- WHY body IS HTML
-- The write-up is authored in a rich text editor and carries headings, bullets
-- and emphasis. It is sanitised on the way in by the router — never trust it
-- on the way out without that, since it renders on a client-facing page.
--
-- DRAFTS ARE INVISIBLE TO CLIENTS
-- Only status = 'published' is ever served to the portal. A draft is the
-- account manager's working copy, including its snapshot, which is taken fresh
-- on every publish rather than at creation.
-- ============================================================================

BEGIN;

CREATE TABLE IF NOT EXISTS pulse_reports (
    id           uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id    uuid NOT NULL REFERENCES agencies(id) ON DELETE CASCADE,
    client_id    uuid NOT NULL REFERENCES clients(id)  ON DELETE CASCADE,

    -- The period the report covers. Both inclusive, both dates, because a
    -- report is always whole days in UTC — the same clock the funnel uses.
    period_start date NOT NULL,
    period_end   date NOT NULL,

    -- Shown as the report's heading. Defaults to the period, but an account
    -- manager can name it ("September 2026", "Q3 review").
    title        text NOT NULL DEFAULT '',

    -- The write-up: analytics and recommendations, sanitised HTML.
    body         text NOT NULL DEFAULT '',

    -- The frozen figures. See above: this is the report, not a cache.
    snapshot     jsonb NOT NULL DEFAULT '{}'::jsonb,

    status       text NOT NULL DEFAULT 'draft',
    published_at timestamptz,
    created_by   text NOT NULL DEFAULT '',
    created_at   timestamptz NOT NULL DEFAULT now(),
    updated_at   timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE pulse_reports DROP CONSTRAINT IF EXISTS pulse_reports_status_check;
ALTER TABLE pulse_reports
    ADD CONSTRAINT pulse_reports_status_check
    CHECK (status IN ('draft', 'published'));

-- A period that ends before it starts would render as an empty report with a
-- nonsense heading, and the range picker cannot be trusted to prevent it.
ALTER TABLE pulse_reports DROP CONSTRAINT IF EXISTS pulse_reports_period_check;
ALTER TABLE pulse_reports
    ADD CONSTRAINT pulse_reports_period_check
    CHECK (period_end >= period_start);

-- The portal's only query: this client's published reports, newest first. The
-- arrows page through exactly this ordering.
CREATE INDEX IF NOT EXISTS pulse_reports_client_period_idx
    ON pulse_reports (agency_id, client_id, period_end DESC);

-- One report per client per period. Publishing September twice should correct
-- the September report, not leave the client two of them to choose between.
CREATE UNIQUE INDEX IF NOT EXISTS pulse_reports_client_period_unique
    ON pulse_reports (agency_id, client_id, period_start, period_end);


DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'anon') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_reports TO anon;
    END IF;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'authenticated') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_reports TO authenticated;
    END IF;
END;
$$;

COMMIT;
