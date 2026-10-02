-- Credit Control — Supabase Schema Migration
-- Run this in the Supabase SQL editor (Dashboard → SQL Editor → New query)
--
-- Deliberately a single JSONB blob, not a normalised relational schema: the
-- app's JS (churn grace-period logic, due-date month assignment, commission
-- splits, retention math — see app/templates/credit_control/app.html) keeps
-- reading/writing the exact same {settings, clients, invoices} shape it
-- always has, whether that's coming from the old local data.json or this
-- table. Re-modelling 872 invoices and 170 clients into normalised tables
-- would mean re-deriving every one of those calculations against a new data
-- shape — exactly the risk this migration is designed to avoid.

-- ── credit_control_store ────────────────────────────────────────────────────
-- Single-row table (id is pinned to 1) holding the live data blob.
CREATE TABLE IF NOT EXISTS credit_control_store (
    id         INT PRIMARY KEY DEFAULT 1,
    data       JSONB NOT NULL,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_by TEXT,
    CONSTRAINT credit_control_store_singleton CHECK (id = 1)
);

-- ── credit_control_backups ──────────────────────────────────────────────────
-- Append-only history, written before every save — the Supabase equivalent
-- of serve.ps1's timestamped backups/*.json files (last 30 kept there).
CREATE TABLE IF NOT EXISTS credit_control_backups (
    id       BIGSERIAL PRIMARY KEY,
    data     JSONB NOT NULL,
    saved_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    saved_by TEXT
);

CREATE INDEX IF NOT EXISTS credit_control_backups_saved_at_idx
    ON credit_control_backups (saved_at DESC);

-- ── Row Level Security ──────────────────────────────────────────────────────
-- Unlike the dashboard's other tables, these hold financial data (invoices,
-- commissions, client billing), so they're closed to the public anon key:
-- RLS on, no policies. Only the server-side service_role key (env var
-- SUPABASE_SERVICE_ROLE_KEY) can read/write them. (Choosing "Run and enable
-- RLS" in the Supabase SQL editor does the same thing as these two lines.)
ALTER TABLE credit_control_store   ENABLE ROW LEVEL SECURITY;
ALTER TABLE credit_control_backups ENABLE ROW LEVEL SECURITY;

-- No seed row here on purpose — the one-time import script
-- (scripts/import_credit_control_data.py) inserts the real data.json content
-- as the single credit_control_store row. Running this migration alone
-- leaves the table empty; the app's GET /api/credit-control/data will 404
-- until that import has run.
