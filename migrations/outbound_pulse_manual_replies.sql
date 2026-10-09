-- ============================================================================
-- Client Performance & Reports — a classified manual lead has replied
--
-- Run once in the Supabase SQL editor, after outbound_pulse_hand_entry.sql.
-- Safe to re-run. Set the editor's row limit to "No limit" first.
--
-- WHY
-- pulse_manual_daily has never emitted a `replied` row, so every DNC & Merger
-- campaign showed a reply count of zero and a reply rate of nothing. The
-- client page even carried a note explaining it. But a prospect does not get
-- marked Lead, Interested or Unsubscribe unless they answered: the
-- classification IS the evidence of a reply. Counting the outcome while
-- refusing to count the reply that produced it made manual campaigns look
-- like they generated leads out of silence, and made the combined reply rate
-- an average over a denominator that included manual sends and a numerator
-- that did not.
--
-- WHAT COUNTS
-- The three classifications that write a dated row: 'lead', 'interested' and
-- 'opt_out'. Each already becomes an outcome stage; each now also becomes a
-- reply, on the same day, from the same row.
--
-- WHAT DOES NOT, AND WHY
-- The sheet's fourth status, a plain "Reply", writes nothing to this database
-- at all — app/utils/auto_sync.py only bumps campaigns.reply_count, which is
-- a running total with no date. A daily funnel cannot take an undated total
-- without inventing a day for it, and dnc_entries is not available as a home
-- for it either: that table is the suppression list and the scrub matches on
-- email with no filter on reason, so a row written there would quietly stop
-- that prospect ever being contacted again. A reply is usually the start of a
-- conversation, not the end of one. Plain replies therefore stay uncounted,
-- and manual reply figures read low wherever the team uses that status.
--
-- DOUBLE COUNTING
-- dnc_entries carries a unique index on (client_id, lower(email)), so a
-- prospect holds at most one classification ever — there is no second row
-- to count, on that day or any other. COUNT(DISTINCT email) is therefore
-- belt and braces rather than the thing holding this together, and it is
-- kept so that relaxing that index later cannot quietly inflate every
-- reply rate.
--
-- The one case it cannot see: a 'lead' is stored by the DNC tool as a
-- DOMAIN rather than an address, so a lead and an interested reply from
-- the same company are two different values and count as two replies.
-- That is a property of how that tool records leads, and fixing it
-- belongs there.
-- ============================================================================

BEGIN;

DO $$
DECLARE
    has_chaser boolean;
    chaser_sql text := '';
BEGIN
    SELECT EXISTS (
        SELECT 1 FROM information_schema.columns
         WHERE table_name = 'contacted_prospects'
           AND column_name = 'chaser_contacted_at'
    ) INTO has_chaser;

    IF has_chaser THEN
        chaser_sql := $c$
    UNION ALL
    SELECT c.agency_id, cp.client_id, NULL::uuid, 'email', 'manual', 'sent',
           cp.chaser_contacted_at, COUNT(*)::bigint
      FROM contacted_prospects cp
      JOIN clients c ON c.id = cp.client_id
     WHERE cp.chaser_contacted_at IS NOT NULL AND c.agency_id IS NOT NULL
     GROUP BY 1, 2, 3, 4, 5, 6, 7
        $c$;
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
    -- The outcome itself, unchanged.
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
    UNION ALL
    -- NEW: the reply that produced it. Nobody is classified without answering.
    -- DISTINCT so one prospect classified twice on a day is one reply.
    SELECT c.agency_id, d.client_id, NULL::uuid, 'email', 'manual', 'replied',
           (d.created_at AT TIME ZONE 'UTC')::date, COUNT(DISTINCT d.email)::bigint
      FROM dnc_entries d
      JOIN clients c ON c.id = d.client_id
     WHERE d.reason IN ('lead', 'interested', 'opt_out') AND c.agency_id IS NOT NULL
     GROUP BY 1, 2, 3, 4, 5, 6, 7
    $q$;
END;
$$;


-- ── Recreate the funnel view ────────────────────────────────────────────────
-- Unchanged apart from being rebuilt on the new pulse_manual_daily. All five
-- branches are listed because CREATE OR REPLACE VIEW cannot change a view's
-- column list and the funnel has to be dropped and rebuilt either way.
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
SELECT o.agency_id, c.client_id, o.campaign_id, c.channel, c.source_tool,
       'unsubscribed', o.unsub_day, COUNT(*)::bigint
  FROM pulse_lead_outcomes o
  JOIN pulse_campaigns c ON c.id = o.campaign_id
 WHERE o.unsub_day IS NOT NULL
 GROUP BY 1, 2, 3, 4, 5, 6, 7
UNION ALL
SELECT agency_id, client_id, campaign_id, channel, source_tool,
       event_type, day, events
  FROM pulse_manual_daily
UNION ALL
SELECT agency_id, client_id, campaign_id, channel, source_tool,
       event_type, day, events
  FROM pulse_hand_entry_daily;


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
