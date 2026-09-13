-- ============================================================================
-- Outbound Pulse — normalized outbound reporting (Smartlead + Meet Alfred)
--
-- Run once in the Supabase SQL editor.  Every statement is additive and
-- idempotent (matching the house pattern — there is no migration runner), so
-- re-running is safe and the app keeps working before it is applied.
--
-- NAMING: the dashboard already has a `campaigns` table (mail-merge campaigns
-- linked to Google Sheets, see campaigns_schema.sql).  That is a different
-- concept from an outbound-tool campaign, so every table here is prefixed
-- `pulse_` rather than colliding with it.  The `clients` table IS reused — it
-- is already the single source of truth for a client profile.
--
-- TENANCY: every table carries agency_id.  There will only ever be one agency
-- row for now; the column exists so spinning this out later is a migration
-- rather than a rebuild.  All application queries filter on it.
-- ============================================================================


-- ── agencies ────────────────────────────────────────────────────────────────
-- One row for now.  Modeled properly so agency_id is a real FK everywhere.
CREATE TABLE IF NOT EXISTS agencies (
    id         UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    name       TEXT        UNIQUE NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

INSERT INTO agencies (name) VALUES ('Unstuck')
ON CONFLICT (name) DO NOTHING;


-- ── clients.agency_id ───────────────────────────────────────────────────────
-- The existing clients table becomes agency-scoped.  Backfilled to the single
-- agency row; the app also backfills on demand so a client created by another
-- tool (DNC, Client Profiles) does not silently fall out of Pulse's scope.
ALTER TABLE clients ADD COLUMN IF NOT EXISTS agency_id UUID REFERENCES agencies(id);

UPDATE clients
   SET agency_id = (SELECT id FROM agencies WHERE name = 'Unstuck')
 WHERE agency_id IS NULL;

CREATE INDEX IF NOT EXISTS clients_agency_id_idx ON clients (agency_id);


-- ── pulse_campaigns ─────────────────────────────────────────────────────────
-- One row per campaign in an external tool, mapped to one of our clients.
CREATE TABLE IF NOT EXISTS pulse_campaigns (
    id                   UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id            UUID        NOT NULL REFERENCES agencies(id),
    client_id            UUID        REFERENCES clients(id) ON DELETE SET NULL,
    channel              TEXT        NOT NULL,        -- 'email' | 'linkedin'
    source_tool          TEXT        NOT NULL,        -- 'smartlead' | 'meet_alfred'
    external_campaign_id TEXT        NOT NULL,
    name                 TEXT        NOT NULL DEFAULT '',
    status               TEXT        NOT NULL DEFAULT 'unknown',
    last_synced_at       TIMESTAMPTZ,
    raw_payload          JSONB       NOT NULL DEFAULT '{}'::jsonb,
    created_at           TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at           TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Upsert target for the connectors: one row per external campaign per agency.
CREATE UNIQUE INDEX IF NOT EXISTS pulse_campaigns_external_unique
    ON pulse_campaigns (agency_id, source_tool, external_campaign_id);

CREATE INDEX IF NOT EXISTS pulse_campaigns_client_idx
    ON pulse_campaigns (agency_id, client_id);


-- ── pulse_campaign_events ───────────────────────────────────────────────────
-- Append-only, normalized across both connectors.  The funnel view never needs
-- to know which tool a data point came from.
--
-- COUNTING CONTRACT (enforced at ingest — see app/utils/pulse/normalize.py):
--   * 'sent' events are per send.  A 4-step sequence to one lead = 4 rows.
--   * 'opened' | 'replied' | 'positive_reply' | 'meeting_booked' are per lead
--     per campaign — at most ONE row each, because their dedupe_key omits the
--     timestamp.  A lead who opens six times still produces one 'opened' row.
--
-- That makes the funnel a plain COUNT(*) per event_type: unique-lead semantics
-- for the qualifying stages, true volume for sends.  No COUNT(DISTINCT) is
-- needed, which matters because PostgREST cannot express one.
CREATE TABLE IF NOT EXISTS pulse_campaign_events (
    id          UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id   UUID        NOT NULL REFERENCES agencies(id),
    campaign_id UUID        NOT NULL REFERENCES pulse_campaigns(id) ON DELETE CASCADE,
    client_id   UUID        REFERENCES clients(id) ON DELETE SET NULL,
    event_type  TEXT        NOT NULL,   -- sent|opened|replied|positive_reply|meeting_booked
    occurred_at TIMESTAMPTZ NOT NULL,
    lead_key    TEXT        NOT NULL DEFAULT '',   -- lowercased email / LinkedIn profile URL
    dedupe_key  TEXT        NOT NULL,              -- see counting contract above
    raw_payload JSONB       NOT NULL DEFAULT '{}'::jsonb,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Makes re-syncing idempotent: the connectors POST with
-- Prefer: resolution=ignore-duplicates and on_conflict=agency_id,dedupe_key.
CREATE UNIQUE INDEX IF NOT EXISTS pulse_campaign_events_dedupe_unique
    ON pulse_campaign_events (agency_id, dedupe_key);

CREATE INDEX IF NOT EXISTS pulse_campaign_events_campaign_idx
    ON pulse_campaign_events (agency_id, campaign_id, event_type, occurred_at);

CREATE INDEX IF NOT EXISTS pulse_campaign_events_client_idx
    ON pulse_campaign_events (agency_id, client_id, occurred_at);


-- ── pulse_funnel_daily (view) ───────────────────────────────────────────────
-- Pre-aggregated funnel counts, read over PostgREST by the dashboard.  Reading
-- the raw event rows into Python would move tens of thousands of rows per page
-- load; this collapses them to one row per campaign/type/day server-side.
--
-- SUM(events) over a date range is the correct funnel figure because of the
-- ingest dedupe contract documented on pulse_campaign_events above.
CREATE OR REPLACE VIEW pulse_funnel_daily AS
SELECT
    e.agency_id,
    e.client_id,
    e.campaign_id,
    c.channel,
    c.source_tool,
    e.event_type,
    (e.occurred_at AT TIME ZONE 'UTC')::date AS day,
    COUNT(*)                                 AS events
FROM pulse_campaign_events e
JOIN pulse_campaigns c ON c.id = e.campaign_id
GROUP BY 1, 2, 3, 4, 5, 6, 7;


-- ── pulse_sync_logs ─────────────────────────────────────────────────────────
-- One row per connector run, so breakage is visible internally before a client
-- notices a stale dashboard.
CREATE TABLE IF NOT EXISTS pulse_sync_logs (
    id                UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id         UUID        NOT NULL REFERENCES agencies(id),
    source_tool       TEXT        NOT NULL,
    run_at            TIMESTAMPTZ NOT NULL DEFAULT now(),
    finished_at       TIMESTAMPTZ,
    status            TEXT        NOT NULL DEFAULT 'running',  -- running|ok|partial|error
    campaigns_synced  INT         NOT NULL DEFAULT 0,
    events_inserted   INT         NOT NULL DEFAULT 0,
    duration_s        NUMERIC,
    triggered_by      TEXT        NOT NULL DEFAULT 'schedule',  -- schedule|manual:<email>
    error_message     TEXT
);

CREATE INDEX IF NOT EXISTS pulse_sync_logs_recent_idx
    ON pulse_sync_logs (agency_id, source_tool, run_at DESC);


-- ── pulse_client_access ─────────────────────────────────────────────────────
-- Magic-link access to the read-only client portal.  Only a SHA-256 hash of the
-- token is stored, so a leaked database dump does not hand over live portal
-- links.  The plaintext token is shown to staff exactly once, at creation.
CREATE TABLE IF NOT EXISTS pulse_client_access (
    id           UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id    UUID        NOT NULL REFERENCES agencies(id),
    client_id    UUID        NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    token_hash   TEXT        NOT NULL UNIQUE,
    label        TEXT        NOT NULL DEFAULT '',   -- who it was issued to
    created_by   TEXT        NOT NULL DEFAULT '',
    created_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    expires_at   TIMESTAMPTZ,
    revoked_at   TIMESTAMPTZ,
    last_used_at TIMESTAMPTZ,
    view_count   INT         NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS pulse_client_access_client_idx
    ON pulse_client_access (agency_id, client_id);


-- ── pulse_portal_visits ─────────────────────────────────────────────────────
-- Append-only visit log.  Exists to answer success-criteria question #2 —
-- "did clients engage with the portal, or ignore it?" — which cannot be
-- answered retrospectively if we do not record it from day one.
CREATE TABLE IF NOT EXISTS pulse_portal_visits (
    id         UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id  UUID        NOT NULL REFERENCES agencies(id),
    client_id  UUID        NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    access_id  UUID        REFERENCES pulse_client_access(id) ON DELETE SET NULL,
    viewed_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    user_agent TEXT
);

CREATE INDEX IF NOT EXISTS pulse_portal_visits_client_idx
    ON pulse_portal_visits (agency_id, client_id, viewed_at DESC);


-- ── Row Level Security ──────────────────────────────────────────────────────
-- Disabled to match every other table in this project: the app is server-side
-- and reaches Supabase with the anon key, never from a browser.  The portal is
-- served by our own FastAPI route after token verification — the client's
-- browser never holds a Supabase key.
ALTER TABLE agencies              DISABLE ROW LEVEL SECURITY;
ALTER TABLE pulse_campaigns       DISABLE ROW LEVEL SECURITY;
ALTER TABLE pulse_campaign_events DISABLE ROW LEVEL SECURITY;
ALTER TABLE pulse_sync_logs       DISABLE ROW LEVEL SECURITY;
ALTER TABLE pulse_client_access   DISABLE ROW LEVEL SECURITY;
ALTER TABLE pulse_portal_visits   DISABLE ROW LEVEL SECURITY;
