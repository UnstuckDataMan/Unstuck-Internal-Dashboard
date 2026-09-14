-- ============================================================================
-- Outbound Pulse — funnel rollup table
--
-- Run once in the Supabase SQL editor, after outbound_pulse_schema.sql.
-- Safe to re-run: it rebuilds the rollup from the event table every time, so
-- re-running is also the repair procedure if counts are ever suspected wrong.
--
-- WHY
-- pulse_funnel_daily was a plain view: every dashboard request re-aggregated
-- the ENTIRE pulse_campaign_events table, then filtered the result. That cost
-- grows with every event ever synced, and Supabase cancels anon-role queries
-- after 3 seconds — surfacing as HTTP 500 on the Outbound Pulse page.
--
-- WHAT
-- A small rollup table (one row per campaign / event type / UTC day) kept
-- current by a trigger at insert time. pulse_funnel_daily is recreated over it
-- with the same columns and types, so no application query changes. Reads now
-- scan rollup rows for the requested date range instead of raw events.
--
-- CORRECTNESS
-- Connectors insert with ON CONFLICT (agency_id, dedupe_key) DO NOTHING. A
-- statement trigger's transition table contains only rows actually inserted,
-- so re-synced duplicates never reach the rollup — counts stay exact without
-- the trigger needing to know about deduplication.
--
-- client_id and channel are read from pulse_campaigns at query time rather
-- than stored in the rollup. Remapping a campaign therefore moves its whole
-- history to the new client instantly, with no rollup rewrite.
-- ============================================================================

BEGIN;

-- Hold writes to the event table for the length of this transaction.
-- Without it, a sync inserting during the rebuild below is either counted twice
-- (trigger fires AND the rebuild sees the row) or not at all. SHARE ROW
-- EXCLUSIVE blocks INSERT but still allows the dashboard to read. A sync that
-- arrives meanwhile waits; if it times out instead, it logs a partial run and
-- re-syncs next hour, which is harmless because ingest is idempotent.
LOCK TABLE pulse_campaign_events IN SHARE ROW EXCLUSIVE MODE;


-- ── Rollup table ────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS pulse_funnel_rollup (
    agency_id   UUID   NOT NULL REFERENCES agencies(id),
    campaign_id UUID   NOT NULL REFERENCES pulse_campaigns(id) ON DELETE CASCADE,
    event_type  TEXT   NOT NULL,
    day         DATE   NOT NULL,
    events      BIGINT NOT NULL DEFAULT 0,
    PRIMARY KEY (agency_id, campaign_id, event_type, day)
);

-- Serves the dashboard's agency + date-range reads.
CREATE INDEX IF NOT EXISTS pulse_funnel_rollup_day_idx
    ON pulse_funnel_rollup (agency_id, day);

ALTER TABLE pulse_funnel_rollup DISABLE ROW LEVEL SECURITY;


-- ── Maintenance trigger ─────────────────────────────────────────────────────
-- Statement-level with a transition table: one grouped upsert per insert
-- statement (the connectors insert 500 rows at a time), not one per row.
--
-- SECURITY DEFINER so the rollup write does not depend on the inserting role
-- holding write privileges on the rollup table. EXECUTE is revoked below, so
-- this cannot be invoked other than by the trigger.
CREATE OR REPLACE FUNCTION pulse_rollup_events() RETURNS trigger
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
    INSERT INTO pulse_funnel_rollup (agency_id, campaign_id, event_type, day, events)
    SELECT agency_id,
           campaign_id,
           event_type,
           (occurred_at AT TIME ZONE 'UTC')::date,
           COUNT(*)
      FROM new_rows
     GROUP BY 1, 2, 3, 4
    ON CONFLICT (agency_id, campaign_id, event_type, day)
    DO UPDATE SET events = pulse_funnel_rollup.events + EXCLUDED.events;
    RETURN NULL;
END;
$$;

REVOKE ALL ON FUNCTION pulse_rollup_events() FROM PUBLIC;

DROP TRIGGER IF EXISTS pulse_campaign_events_rollup ON pulse_campaign_events;
CREATE TRIGGER pulse_campaign_events_rollup
    AFTER INSERT ON pulse_campaign_events
    REFERENCING NEW TABLE AS new_rows
    FOR EACH STATEMENT
    EXECUTE FUNCTION pulse_rollup_events();


-- ── Rebuild from the event table ────────────────────────────────────────────
-- TRUNCATE then full rebuild, rather than an incremental backfill, so this file
-- is both the first-time setup and the repair tool. The day expression must
-- match the trigger's exactly, or rebuilt and live rows would bucket
-- differently around midnight.
TRUNCATE pulse_funnel_rollup;

INSERT INTO pulse_funnel_rollup (agency_id, campaign_id, event_type, day, events)
SELECT agency_id,
       campaign_id,
       event_type,
       (occurred_at AT TIME ZONE 'UTC')::date,
       COUNT(*)
  FROM pulse_campaign_events
 GROUP BY 1, 2, 3, 4;


-- ── Replace the view over the rollup ────────────────────────────────────────
-- Same column names, order and types as before (events stays BIGINT), so
-- every existing PostgREST query against pulse_funnel_daily keeps working.
-- DROP + CREATE rather than CREATE OR REPLACE, which refuses any column change.
DROP VIEW IF EXISTS pulse_funnel_daily;

CREATE VIEW pulse_funnel_daily AS
SELECT r.agency_id,
       c.client_id,
       r.campaign_id,
       c.channel,
       c.source_tool,
       r.event_type,
       r.day,
       r.events
  FROM pulse_funnel_rollup r
  JOIN pulse_campaigns c ON c.id = r.campaign_id;


-- ── Grants ──────────────────────────────────────────────────────────────────
-- Dropping the view discards its privileges. Supabase's default privileges
-- normally re-grant a new relation, but the dashboard reads this view with the
-- anon key, so grant explicitly rather than rely on that. Guarded so the file
-- also runs on a plain Postgres without Supabase's roles.
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'anon') THEN
        GRANT SELECT ON pulse_funnel_daily TO anon;
        GRANT SELECT ON pulse_funnel_rollup TO anon;
    END IF;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'authenticated') THEN
        GRANT SELECT ON pulse_funnel_daily TO authenticated;
        GRANT SELECT ON pulse_funnel_rollup TO authenticated;
    END IF;
END;
$$;

COMMIT;
