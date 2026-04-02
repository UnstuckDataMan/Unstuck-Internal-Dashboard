-- DNC List Scrubbing Tool — Supabase Schema Migration
-- Run this in the Supabase SQL editor (Dashboard → SQL Editor → New query)

-- ── clients ───────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS clients (
    id         UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    name       TEXT UNIQUE NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ── dnc_entries ───────────────────────────────────────────────────────────────
-- emails are always stored lowercase (enforced in application code before insert)
CREATE TABLE IF NOT EXISTS dnc_entries (
    id         UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    client_id  UUID NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    email      TEXT NOT NULL,
    reason     TEXT NOT NULL DEFAULT 'manual',
    added_by   TEXT NOT NULL DEFAULT 'system',
    notes      TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Prevents duplicate emails per client (case-insensitive via functional index)
CREATE UNIQUE INDEX IF NOT EXISTS dnc_entries_client_email_unique
    ON dnc_entries (client_id, LOWER(email));

-- Fast lookup by client + email for scrubbing
CREATE INDEX IF NOT EXISTS dnc_entries_client_id_idx
    ON dnc_entries (client_id);

-- ── scrub_logs ────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS scrub_logs (
    id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    client_id       UUID NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    uploaded_count  INT NOT NULL,
    removed_count   INT NOT NULL,
    remaining_count INT NOT NULL,
    performed_by    TEXT NOT NULL DEFAULT 'unknown',
    created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ── Row Level Security ────────────────────────────────────────────────────────
-- This is an internal tool accessed via the anon key.
-- RLS is disabled so the anon key has full read/write access.
-- If you enable Supabase Auth later, switch to authenticated-user policies.
ALTER TABLE clients    DISABLE ROW LEVEL SECURITY;
ALTER TABLE dnc_entries DISABLE ROW LEVEL SECURITY;
ALTER TABLE scrub_logs  DISABLE ROW LEVEL SECURITY;

-- ── Sample client (optional — delete before production) ───────────────────────
-- INSERT INTO clients (name) VALUES ('Test Client');
