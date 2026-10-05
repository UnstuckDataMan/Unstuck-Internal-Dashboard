-- ============================================================================
-- Client Performance & Reports — figures typed in by an account manager
--
-- Run once in the Supabase SQL editor, after outbound_pulse_unsubscribes.sql.
-- Safe to re-run. Set the editor's row limit to "No limit" first.
--
-- THIS IS NOT THE 'manual' SOURCE
-- source_tool 'manual' already exists and means something else: campaigns run
-- through the DNC & Merger tool, whose numbers this database already holds and
-- pulse_manual_daily reads in place. Nobody types those. This is the opposite
-- — a person reading a dashboard the app cannot reach (Meet Alfred, whose API
-- access is still unconfirmed) and keying the totals in.
--
-- Different provenance, so: a different source_tool ('hand_entry'), a
-- different view, and a differently labelled tab. Nothing is called just
-- "Manual" any more — that word is what made the two confusable. The existing
-- key stays 'manual' because it is written into every report snapshot already
-- on disk; only its label changed, to "DNC & Merger".
--
-- WHY A TABLE AND NOT A VIEW
-- pulse_manual_daily can be a view because its data already exists. This data
-- exists nowhere until somebody types it. It is deliberately NOT written into
-- pulse_campaign_events either: those are append-only per-lead events with a
-- dedupe key, and these are aggregate monthly totals with no lead identity.
-- Forcing them in would mean inventing leads that were never observed.
--
-- DATING — WHY THE FIRST OF THE MONTH
-- pulse_funnel_daily is daily and this figure covers a month. It is written on
-- ONE day and never spread across the month: dividing a monthly total by 30
-- would invent a daily series nobody measured, and every number in it would be
-- false.
--
-- The first of the month rather than the last because:
--   * it is exactly where the trend strip's month bucket starts, so the figure
--     lands in the right column at any bucket width;
--   * a figure entered for the current month shows up in "This month"
--     immediately, instead of only once the month has closed;
--   * it is derivable from the period alone, so re-saving can never move it.
--
-- The cost, which the entry form states plainly: a range that cuts through a
-- month either contains the whole figure or none of it. A monthly number
-- cannot be truthfully sliced.
-- ============================================================================

BEGIN;

CREATE TABLE IF NOT EXISTS pulse_hand_entries (
    id           uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id    uuid NOT NULL REFERENCES agencies(id) ON DELETE CASCADE,
    client_id    uuid NOT NULL REFERENCES clients(id)  ON DELETE CASCADE,

    -- Per entry, not per source: unlike Smartlead and Meet Alfred, the same
    -- person keys LinkedIn figures for one client and email for another.
    channel      text NOT NULL CHECK (channel IN ('email', 'linkedin')),

    -- The month covered, stored as its first day. This column IS the day the
    -- funnel reads; see the dating note above.
    period_month date NOT NULL,

    -- One column per funnel stage this source can report. `opened` carries
    -- LinkedIn connection acceptances as well as email opens: they are the
    -- same stage, and two columns feeding one stage would let a single entry
    -- count both.
    sent         bigint NOT NULL DEFAULT 0 CHECK (sent         >= 0),
    opened       bigint NOT NULL DEFAULT 0 CHECK (opened       >= 0),
    replied      bigint NOT NULL DEFAULT 0 CHECK (replied      >= 0),
    meetings     bigint NOT NULL DEFAULT 0 CHECK (meetings     >= 0),
    bounced      bigint NOT NULL DEFAULT 0 CHECK (bounced      >= 0),
    unsubscribed bigint NOT NULL DEFAULT 0 CHECK (unsubscribed >= 0),

    -- Where the numbers came from ("Meet Alfred dashboard, 3 Oct"). A typed
    -- figure with no provenance cannot be checked against anything later.
    source_note  text NOT NULL DEFAULT '',
    entered_by   text NOT NULL DEFAULT '',
    created_at   timestamptz NOT NULL DEFAULT now(),
    updated_at   timestamptz NOT NULL DEFAULT now()
);

-- A month is its first day, always. Without this the same month could be
-- stored under 30 different dates and the unique index below would not hold.
ALTER TABLE pulse_hand_entries DROP CONSTRAINT IF EXISTS pulse_hand_entries_month_start;
ALTER TABLE pulse_hand_entries
    ADD CONSTRAINT pulse_hand_entries_month_start
    CHECK (period_month = date_trunc('month', period_month)::date);

-- The upsert target. Re-entering a month corrects it instead of leaving two
-- contradictory totals for the same period.
CREATE UNIQUE INDEX IF NOT EXISTS pulse_hand_entries_unique
    ON pulse_hand_entries (agency_id, client_id, channel, period_month);

CREATE INDEX IF NOT EXISTS pulse_hand_entries_month_idx
    ON pulse_hand_entries (agency_id, period_month);

ALTER TABLE pulse_hand_entries DISABLE ROW LEVEL SECURITY;


-- ── One funnel row per non-zero figure ──────────────────────────────────────
-- Same column names, order and types as every other branch of the union, so
-- nothing downstream needs to know this source exists.
--
-- Zero figures produce no row, so a client who has never had anything entered
-- does not get an "Entered by hand" tab lit up with zeros.
CREATE OR REPLACE VIEW pulse_hand_entry_daily
    (agency_id, client_id, campaign_id, channel, source_tool, event_type, day, events)
AS
SELECT e.agency_id,
       e.client_id,
       -- NULL: these are client-level totals, not a campaign's. Attributing
       -- them to one would be a guess, and the campaign table would then show
       -- figures nobody can trace back to a source.
       NULL::uuid,
       e.channel,
       'hand_entry'::text,
       v.event_type,
       e.period_month,
       v.events
  FROM pulse_hand_entries e
  CROSS JOIN LATERAL (VALUES
        ('sent'::text,           e.sent),
        ('opened'::text,         e.opened),
        ('replied'::text,        e.replied),
        -- The existing stage key, labelled "Meeting requests" in the UI. The
        -- entry form says "Meeting requests / meetings booked" so the number
        -- typed means the number printed.
        ('meeting_booked'::text, e.meetings),
        ('bounced'::text,        e.bounced),
        ('unsubscribed'::text,   e.unsubscribed)
  ) AS v(event_type, events)
 WHERE v.events > 0;


-- ── Recreate the funnel view with hand entries ──────────────────────────────
-- The four existing branches are unchanged. pulse_manual_daily is deliberately
-- NOT touched: this migration adds a source, it does not alter the DNC one.
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
        GRANT SELECT ON pulse_funnel_daily, pulse_hand_entry_daily TO anon;
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_hand_entries TO anon;
    END IF;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'authenticated') THEN
        GRANT SELECT ON pulse_funnel_daily, pulse_hand_entry_daily TO authenticated;
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_hand_entries TO authenticated;
    END IF;
END;
$$;

COMMIT;
