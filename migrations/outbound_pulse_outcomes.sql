-- ============================================================================
-- Outbound Pulse — lead outcomes by CURRENT Smartlead category
--
-- Run once in the Supabase SQL editor, after outbound_pulse_rollup.sql.
-- Safe to re-run. Set the editor's row limit to "No limit" first.
--
-- WHY
-- Reply outcomes were stored as append-only events keeping the FIRST stage a
-- lead ever reached. That worked while the stages were nested (a meeting
-- request also counted as interested). The team's categories are mutually
-- exclusive, and they change: a lead marked Interested is later re-marked
-- Meeting Request. Append-only events cannot move a lead out of a bucket, so
-- Interested would keep counting every lead that progressed past it.
--
-- WHAT
-- pulse_lead_outcomes holds one row per replying lead per campaign with that
-- lead's current stage. The connectors upsert it on every sync, so a
-- re-categorised lead moves bucket, and one re-marked Not Interested drops out
-- (stage becomes NULL). Nothing is ever deleted to achieve this.
--
--   stage = positive_reply       Interested
--           information_request  Information Request   ─┐ leads
--           meeting_booked       Meeting Request       ─┘
--           NULL                 replied, not positive
--
-- pulse_funnel_daily is recreated with the same columns: sent/opened/replied
-- still come from the event rollup; the three outcome stages now come from
-- this table. Legacy positive_reply / meeting_booked EVENTS stay in the event
-- table untouched but are no longer read.
-- ============================================================================

BEGIN;

-- ── Outcomes table ──────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS pulse_lead_outcomes (
    agency_id   UUID        NOT NULL REFERENCES agencies(id),
    campaign_id UUID        NOT NULL REFERENCES pulse_campaigns(id) ON DELETE CASCADE,
    lead_key    TEXT        NOT NULL,
    stage       TEXT,                           -- NULL = replied, not positive
    category    TEXT        NOT NULL DEFAULT '', -- source category, for audit
    day         DATE        NOT NULL,           -- UTC day of the reply
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (agency_id, campaign_id, lead_key)
);

-- Serves the dashboard's agency + date-range reads. Partial: non-positive rows
-- are never counted, so they need not be indexed.
CREATE INDEX IF NOT EXISTS pulse_lead_outcomes_day_idx
    ON pulse_lead_outcomes (agency_id, day) WHERE stage IS NOT NULL;

ALTER TABLE pulse_lead_outcomes DISABLE ROW LEVEL SECURITY;


-- ── One-time backfill from history ──────────────────────────────────────────
-- Without this, every outcome would read zero until the rolling sync revisited
-- each campaign — up to ~40 hours for inactive ones. Stored events keep the
-- Smartlead lead_category in raw_payload, so the latest known category per
-- lead can be classified now. The next sync of each campaign then overwrites
-- these rows with the live category, so any imperfection here is temporary.
--
-- The classifier is created and dropped inside this transaction, so it never
-- becomes visible outside it. Its term lists MUST match classify_reply() in
-- app/utils/pulse/normalize.py — a test compares them.
CREATE FUNCTION pulse_classify_backfill(category text) RETURNS text
LANGUAGE sql IMMUTABLE AS $fn$
    SELECT CASE
        WHEN k = '' THEN NULL
        WHEN k = ANY (meeting)     THEN 'meeting_booked'
        WHEN k = ANY (information) THEN 'information_request'
        WHEN k = ANY (interested)  THEN 'positive_reply'
        WHEN k = ANY (negative)    THEN NULL
        -- Substring fallback, negative FIRST: when in doubt, under-report.
        WHEN EXISTS (SELECT 1 FROM unnest(negative)    t WHERE strpos(k, t) > 0) THEN NULL
        WHEN EXISTS (SELECT 1 FROM unnest(meeting)     t WHERE strpos(k, t) > 0) THEN 'meeting_booked'
        WHEN EXISTS (SELECT 1 FROM unnest(information) t WHERE strpos(k, t) > 0) THEN 'information_request'
        WHEN EXISTS (SELECT 1 FROM unnest(interested)  t WHERE strpos(k, t) > 0) THEN 'positive_reply'
        ELSE NULL
    END
    FROM (SELECT lower(btrim(coalesce(category, ''), E' \t\r\n')) AS k) key,
         (SELECT
            ARRAY['meeting request', 'meeting requested', 'meeting booked',
                  'meeting completed', 'meeting scheduled', 'booked',
                  'demo booked', 'call booked']                          AS meeting,
            ARRAY['information request', 'info request',
                  'more information', 'more info']                       AS information,
            ARRAY['interested', 'positive', 'positive reply',
                  'warm', 'hot lead']                                    AS interested,
            ARRAY['not interested', 'no longer interested', 'uninterested',
                  'not a fit', 'negative', 'do not contact', 'unsubscribed',
                  'out of office', 'wrong person', 'bounced', 'spam']    AS negative
         ) terms
$fn$;

WITH replied AS (
    -- Every lead that replied, with the UTC day of its first reply.
    SELECT agency_id, campaign_id, lead_key,
           (MIN(occurred_at) AT TIME ZONE 'UTC')::date AS day
      FROM pulse_campaign_events
     WHERE event_type = 'replied' AND lead_key <> ''
     GROUP BY 1, 2, 3
),
latest AS (
    -- The most recently recorded category per lead.
    SELECT DISTINCT ON (agency_id, campaign_id, lead_key)
           agency_id, campaign_id, lead_key,
           btrim(coalesce(raw_payload -> 'row' ->> 'lead_category',
                          raw_payload -> 'row' ->> 'category', '')) AS category
      FROM pulse_campaign_events
     WHERE lead_key <> ''
       AND btrim(coalesce(raw_payload -> 'row' ->> 'lead_category',
                          raw_payload -> 'row' ->> 'category', '')) <> ''
     ORDER BY agency_id, campaign_id, lead_key, created_at DESC
),
legacy_meeting AS (
    -- A lead with a legacy meeting event and no category at all (a tracked
    -- meeting from a Meet Alfred import) is still meeting-tier.
    SELECT DISTINCT agency_id, campaign_id, lead_key
      FROM pulse_campaign_events
     WHERE event_type = 'meeting_booked'
)
INSERT INTO pulse_lead_outcomes (agency_id, campaign_id, lead_key, stage, category, day)
SELECT r.agency_id,
       r.campaign_id,
       r.lead_key,
       CASE
           WHEN l.category IS NOT NULL THEN pulse_classify_backfill(l.category)
           WHEN m.lead_key IS NOT NULL THEN 'meeting_booked'
           ELSE NULL
       END,
       coalesce(l.category, ''),
       r.day
  FROM replied r
  LEFT JOIN latest l         USING (agency_id, campaign_id, lead_key)
  LEFT JOIN legacy_meeting m USING (agency_id, campaign_id, lead_key)
-- DO NOTHING, not an update: on a re-run, rows a live sync already wrote hold
-- the current category and must not be overwritten with an older one.
ON CONFLICT (agency_id, campaign_id, lead_key) DO NOTHING;

DROP FUNCTION pulse_classify_backfill(text);


-- ── Recreate the view ───────────────────────────────────────────────────────
-- Same column names, order and types as before (events stays BIGINT).
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
  JOIN pulse_campaigns c ON c.id = r.campaign_id
 -- Outcome stages are read from pulse_lead_outcomes below. Legacy outcome
 -- events stay in the rollup but are excluded here, or they'd count twice.
 WHERE r.event_type IN ('sent', 'opened', 'replied')
UNION ALL
SELECT o.agency_id,
       c.client_id,
       o.campaign_id,
       c.channel,
       c.source_tool,
       o.stage,
       o.day,
       COUNT(*)::bigint
  FROM pulse_lead_outcomes o
  JOIN pulse_campaigns c ON c.id = o.campaign_id
 WHERE o.stage IS NOT NULL
 GROUP BY 1, 2, 3, 4, 5, 6, 7;


-- ── Grants ──────────────────────────────────────────────────────────────────
-- The dashboard reads the view and the sync upserts outcomes with the anon
-- key. An upsert needs SELECT, INSERT and UPDATE. Explicit rather than relying
-- on Supabase's default privileges, and guarded so plain Postgres runs it too.
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'anon') THEN
        GRANT SELECT ON pulse_funnel_daily TO anon;
        GRANT SELECT, INSERT, UPDATE ON pulse_lead_outcomes TO anon;
    END IF;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'authenticated') THEN
        GRANT SELECT ON pulse_funnel_daily TO authenticated;
        GRANT SELECT, INSERT, UPDATE ON pulse_lead_outcomes TO authenticated;
    END IF;
END;
$$;

COMMIT;
