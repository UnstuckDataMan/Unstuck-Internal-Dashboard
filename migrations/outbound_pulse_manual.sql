-- ============================================================================
-- Outbound Pulse — manual campaign stats from the DNC & Merger tool
--
-- Run once in the Supabase SQL editor, after outbound_pulse_outcomes.sql.
-- Safe to re-run. Set the editor's row limit to "No limit" first.
--
-- WHY A VIEW, NOT A SYNC
-- Manual campaign data already lives in this database — contacted_prospects
-- for sends and dnc_entries for outcomes, both written by the mail-merge
-- auto-sync. Copying it into pulse_campaign_events would duplicate it and add
-- a third connector that could drift from the tool that owns the data. This
-- reads it in place, so the Pulse numbers can never disagree with the DNC &
-- Merger tool about what was synced.
--
-- WHAT MAPS TO WHAT
--   sent            ← contacted_prospects.contacted_at (+ chaser sends)
--   lead            ← dnc_entries where reason = 'lead'
--   positive_reply  ← dnc_entries where reason = 'interested'
--
-- `lead` is a stage of its own. A manual campaign records a lead without
-- saying whether it was an information or a meeting request, so folding it
-- into either would invent a distinction the source does not make. It counts
-- towards Leads alongside those two.
--
-- REPLIES ARE NOT HERE. The sheet's "Reply" status writes no row — only a
-- running total on the campaign — so there is no date to filter it by. Manual
-- replies stay visible in the DNC & Merger tool; Pulse does not show a manual
-- reply count it cannot place in time.
--
-- DATES. A lead is dated by when the hourly sync first recorded it, which is
-- within an hour of the team setting the status in the sheet. Leads are stored
-- deduplicated by domain per client, so a domain that becomes a lead twice is
-- counted once, and these totals read slightly lower than the per-campaign
-- lead_count the DNC & Merger tool shows. That is a property of the source.
-- ============================================================================

BEGIN;

-- Serves the lead/interested lookups by client and date below.
CREATE INDEX IF NOT EXISTS dnc_entries_reason_date_idx
    ON dnc_entries (client_id, reason, created_at);


-- ── Manual daily stats ──────────────────────────────────────────────────────
-- campaign_id is NULL: dnc_entries records a client, not a campaign, so a
-- manual lead cannot be attributed to one. Sends could be, but are left NULL
-- too so a manual row is never half-attributed.
--
-- Built by a DO block because chaser_contacted_at was added to
-- contacted_prospects outside the migrations folder and may not exist on every
-- database. Where it is missing, chaser sends are simply not counted.
DO $$
DECLARE
    has_chaser boolean;
    chaser_sql text := '';
BEGIN
    SELECT EXISTS (
        SELECT 1 FROM information_schema.columns
         WHERE table_name = 'contacted_prospects' AND column_name = 'chaser_contacted_at'
    ) INTO has_chaser;

    IF has_chaser THEN
        chaser_sql := $q$
        UNION ALL
        SELECT c.agency_id, cp.client_id, NULL::uuid, 'email', 'manual', 'sent',
               cp.chaser_contacted_at, COUNT(*)::bigint
          FROM contacted_prospects cp
          JOIN clients c ON c.id = cp.client_id
         WHERE cp.chaser_contacted_at IS NOT NULL AND c.agency_id IS NOT NULL
         GROUP BY 1, 2, 3, 4, 5, 6, 7
        $q$;
    END IF;

    EXECUTE $q$
    CREATE OR REPLACE VIEW pulse_manual_daily
        (agency_id, client_id, campaign_id, channel, source_tool, event_type, day, events)
    AS
    -- Initial sends
    SELECT c.agency_id, cp.client_id, NULL::uuid, 'email', 'manual', 'sent',
           cp.contacted_at, COUNT(*)::bigint
      FROM contacted_prospects cp
      JOIN clients c ON c.id = cp.client_id
     WHERE c.agency_id IS NOT NULL
     GROUP BY 1, 2, 3, 4, 5, 6, 7
    $q$ || chaser_sql || $q$
    UNION ALL
    -- Leads and interested, dated by when the sync recorded them
    SELECT c.agency_id, d.client_id, NULL::uuid, 'email', 'manual',
           CASE d.reason WHEN 'lead' THEN 'lead' ELSE 'positive_reply' END,
           (d.created_at AT TIME ZONE 'UTC')::date, COUNT(*)::bigint
      FROM dnc_entries d
      JOIN clients c ON c.id = d.client_id
     WHERE d.reason IN ('lead', 'interested') AND c.agency_id IS NOT NULL
     GROUP BY 1, 2, 3, 4, 5, 6, 7
    $q$;
END;
$$;


-- ── Recreate the funnel view with manual included ───────────────────────────
-- Same column names, order and types as before.
DROP VIEW IF EXISTS pulse_funnel_daily;

CREATE VIEW pulse_funnel_daily AS
SELECT r.agency_id, c.client_id, r.campaign_id, c.channel, c.source_tool,
       r.event_type, r.day, r.events
  FROM pulse_funnel_rollup r
  JOIN pulse_campaigns c ON c.id = r.campaign_id
 WHERE r.event_type IN ('sent', 'opened', 'replied')
UNION ALL
SELECT o.agency_id, c.client_id, o.campaign_id, c.channel, c.source_tool,
       o.stage, o.day, COUNT(*)::bigint
  FROM pulse_lead_outcomes o
  JOIN pulse_campaigns c ON c.id = o.campaign_id
 WHERE o.stage IS NOT NULL
 GROUP BY 1, 2, 3, 4, 5, 6, 7
UNION ALL
SELECT agency_id, client_id, campaign_id, channel, source_tool,
       event_type, day, events
  FROM pulse_manual_daily;


DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'anon') THEN
        GRANT SELECT ON pulse_funnel_daily, pulse_manual_daily TO anon;
    END IF;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'authenticated') THEN
        GRANT SELECT ON pulse_funnel_daily, pulse_manual_daily TO authenticated;
    END IF;
END;
$$;

COMMIT;
