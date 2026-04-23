-- ============================================================
-- Campaign tracking table
-- Run this once in the Supabase SQL editor.
-- ============================================================

CREATE TABLE IF NOT EXISTS campaigns (
    id                  UUID        DEFAULT gen_random_uuid() PRIMARY KEY,
    created_at          TIMESTAMPTZ DEFAULT now(),
    campaign_name       TEXT        NOT NULL,
    sender_profile_name TEXT,
    client_id           TEXT,
    client_name         TEXT,
    sheet_id            TEXT,
    sheet_url           TEXT,
    total_prospects     INT         DEFAULT 0,
    sent_count          INT         DEFAULT 0
);

-- Optional: index for fast client-based lookups
CREATE INDEX IF NOT EXISTS campaigns_client_id_idx ON campaigns (client_id);
CREATE INDEX IF NOT EXISTS campaigns_created_at_idx ON campaigns (created_at DESC);
