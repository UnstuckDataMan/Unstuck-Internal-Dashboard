-- Recently Contacted Filtering Feature — Supabase Schema Migration
-- Run this in the Supabase SQL editor (Dashboard → SQL Editor → New query)
-- Safe to run on an existing project — does not touch clients, dnc_entries, or scrub_logs rows.

-- ── client_industries ─────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS client_industries (
    id         UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    client_id  UUID NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    name       TEXT NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Prevent duplicate industry names per client (case-insensitive)
CREATE UNIQUE INDEX IF NOT EXISTS client_industries_client_name_unique
    ON client_industries (client_id, LOWER(name));

-- ── contacted_prospects ───────────────────────────────────────────────────────
-- emails are always stored lowercase (enforced in application code)
-- No unique constraint on email alone — same prospect can be contacted on
-- different dates or in different campaigns. Unique on full context.
CREATE TABLE IF NOT EXISTS contacted_prospects (
    id                 UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    client_id          UUID NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    client_industry_id UUID NOT NULL REFERENCES client_industries(id) ON DELETE CASCADE,
    email              TEXT NOT NULL,
    contacted_at       DATE NOT NULL,
    campaign_name      TEXT,
    source             TEXT NOT NULL DEFAULT 'manual',
    created_at         TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Composite unique: same email can only appear once per client+industry+date
CREATE UNIQUE INDEX IF NOT EXISTS contacted_prospects_unique_idx
    ON contacted_prospects (client_id, client_industry_id, email, contacted_at);

-- Fast lookup index for scrub queries
CREATE INDEX IF NOT EXISTS contacted_prospects_scrub_idx
    ON contacted_prospects (client_id, client_industry_id, email, contacted_at);

-- Date range index for management UI filtering
CREATE INDEX IF NOT EXISTS contacted_prospects_date_idx
    ON contacted_prospects (client_id, contacted_at);

-- ── Extend scrub_logs with new audit columns ──────────────────────────────────
ALTER TABLE scrub_logs ADD COLUMN IF NOT EXISTS contacted_removed_count INTEGER DEFAULT 0;
ALTER TABLE scrub_logs ADD COLUMN IF NOT EXISTS lookback_days INTEGER;
ALTER TABLE scrub_logs ADD COLUMN IF NOT EXISTS industry_filter TEXT;

-- ── Row Level Security (disabled — internal tool, matches existing tables) ────
ALTER TABLE client_industries   DISABLE ROW LEVEL SECURITY;
ALTER TABLE contacted_prospects DISABLE ROW LEVEL SECURITY;
