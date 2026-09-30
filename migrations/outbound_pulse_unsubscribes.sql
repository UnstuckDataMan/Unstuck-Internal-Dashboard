-- ============================================================================
-- Outbound Pulse — unsubscribes as a counted stage
--
-- Run once in the Supabase SQL editor, after outbound_pulse_manual.sql.
-- Safe to re-run. Set the editor's row limit to "No limit" first.
--
-- WHY A COLUMN AND NOT A STAGE VALUE
-- pulse_lead_outcomes holds one row per lead with a single `stage`, because
-- the reply categories are mutually exclusive — a lead is Interested or a
-- Meeting Request, never both. An unsubscribe is not exclusive of them: a lead
-- can be marked Interested in March and unsubscribe in April, and both remain
-- true. Storing it as another `stage` value would force the sync to throw one
-- fact away to record the other.
--
-- So it is its own nullable date column. Non-null means "this lead
-- unsubscribed, on this day", and the funnel view counts those rows
-- independently of whatever `stage` the same row carries.
--
-- DATING. Smartlead's statistics rows carry `is_unsubscribed` as a flag with
-- no timestamp of its own. The connector dates it by the lead's reply, and
-- where there is no reply, by the last send to that lead. Both are the closest
-- dated fact available; an unsubscribe cannot precede the send that prompted
-- it, so the figure is right for any range wider than the sequence itself.
--
-- MANUAL. dnc_entries already records opt_out with a created_at. The manual
-- view has been excluding it — it counted only 'lead' and 'interested'. That
-- exclusion is what this migration removes.
-- ============================================================================

BEGIN;

-- ── Smartlead / Meet Alfred: a date on the outcome row ──────────────────────
ALTER TABLE pulse_lead_outcomes
    ADD COLUMN IF NOT EXISTS unsub_day DATE;

-- Partial: only the rows that unsubscribed are ever read through this.
CREATE INDEX IF NOT EXISTS pulse_lead_outcomes_unsub_idx
    ON pulse_lead_outcomes (agency_id, unsub_day) WHERE unsub_day IS NOT NULL;


-- ── Backfill from stored events ────────────────────────────────────────────
-- Events keep `is_unsubscribed` in raw_payload, so the flag is already on
-- disk for every campaign synced so far. Without this, unsubscribes would read
-- zero until the rolling backfill revisited each campaign — up to ~40 hours.
--
-- Dated by the event's own day, which for the reply event is the reply and for
-- a send is the send: the same rule the connector uses going forward.
UPDATE pulse_lead_outcomes o
   SET unsub_day = src.day
  FROM (
    SELECT e.agency_id, e.campaign_id, e.lead_key, MAX(e.occurred_at::date) AS day
      FROM pulse_campaign_events e
     WHERE e.raw_payload->>'is_unsubscribed' IN ('true', 't', '1')
     GROUP BY 1, 2, 3
  ) src
 WHERE o.agency_id   = src.agency_id
   AND o.campaign_id = src.campaign_id
   AND o.lead_key    = src.lead_key
   AND o.unsub_day IS NULL;

-- Leads who unsubscribed without ever replying have no outcome row yet.
-- stage stays NULL, so they are not counted as a reply category — only as an
-- unsubscribe.
INSERT INTO pulse_lead_outcomes (agency_id, campaign_id, lead_key, stage,
                                 category, day, unsub_day, updated_at)
SELECT src.agency_id, src.campaign_id, src.lead_key, NULL, '', src.day,
       src.day, now()
  FROM (
    SELECT e.agency_id, e.campaign_id, e.lead_key, MAX(e.occurred_at::date) AS day
      FROM pulse_campaign_events e
     WHERE e.raw_payload->>'is_unsubscribed' IN ('true', 't', '1')
     GROUP BY 1, 2, 3
  ) src
 WHERE NOT EXISTS (
        SELECT 1 FROM pulse_lead_outcomes o
         WHERE o.agency_id   = src.agency_id
           AND o.campaign_id = src.campaign_id
           AND o.lead_key    = src.lead_key)
ON CONFLICT (agency_id, campaign_id, lead_key) DO NOTHING;


-- ── Manual: stop excluding opt_out ─────────────────────────────────────────
-- Same shape as outbound_pulse_manual.sql, with the third reason added. The
-- DO block is kept because chaser_contacted_at exists on some databases and
-- not others.
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
    SELECT c.agency_id, cp.client_id, NULL::uuid, 'email', 'manual', 'sent',
           cp.contacted_at, COUNT(*)::bigint
      FROM contacted_prospects cp
      JOIN clients c ON c.id = cp.client_id
     WHERE c.agency_id IS NOT NULL
     GROUP BY 1, 2, 3, 4, 5, 6, 7
    $q$ || chaser_sql || $q$
    UNION ALL
    SELECT c.agency_id, d.client_id, NULL::uuid, 'email', 'manual',
           CASE d.reason
               WHEN 'lead'     THEN 'lead'
               WHEN 'opt_out'  THEN 'unsubscribed'
               ELSE 'positive_reply'
           END,
           (d.created_at AT TIME ZONE 'UTC')::date, COUNT(*)::bigint
      FROM dnc_entries d
      JOIN clients c ON c.id = d.client_id
     WHERE d.reason IN ('lead', 'interested', 'opt_out') AND c.agency_id IS NOT NULL
     GROUP BY 1, 2, 3, 4, 5, 6, 7
    $q$;
END;
$$;


-- ── Recreate the funnel view with unsubscribes ─────────────────────────────
-- Same column names, order and types as before, plus a fourth branch.
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
-- Counted off unsub_day, independently of the row's stage, so a lead who was
-- marked Interested and then unsubscribed is counted once in each.
SELECT o.agency_id, c.client_id, o.campaign_id, c.channel, c.source_tool,
       'unsubscribed', o.unsub_day, COUNT(*)::bigint
  FROM pulse_lead_outcomes o
  JOIN pulse_campaigns c ON c.id = o.campaign_id
 WHERE o.unsub_day IS NOT NULL
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
