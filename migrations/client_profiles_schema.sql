-- Central Client Profiles — extend the existing `clients` table into the single
-- source of truth for a client profile. The Client Profiles manager
-- (app/routers/profiles.py) does CRUD on these columns and mirrors the relevant
-- slices out to the legacy stores (sender_profiles, copy_bank_templates'
-- __cb_profiles__ row, and campaigns.client_name).
--
-- There is no migration runner — run this once in the Supabase SQL editor.
-- Every statement is additive and idempotent, so re-running is safe and the
-- app keeps working before it is applied (the router falls back to selecting
-- just id,name when these columns are absent).

ALTER TABLE clients ADD COLUMN IF NOT EXISTS type            text        DEFAULT 'client';
ALTER TABLE clients ADD COLUMN IF NOT EXISTS territories     text[]      DEFAULT '{}';
ALTER TABLE clients ADD COLUMN IF NOT EXISTS industries      text[]      DEFAULT '{}';
ALTER TABLE clients ADD COLUMN IF NOT EXISTS industry_labels jsonb       DEFAULT '{}';   -- { "<key>": "<display label>" }
ALTER TABLE clients ADD COLUMN IF NOT EXISTS senders         text[]      DEFAULT '{}';   -- sender display names (Copy Bank)
ALTER TABLE clients ADD COLUMN IF NOT EXISTS sender_emails   text[]      DEFAULT '{}';   -- sending inboxes (Mail Merge)
ALTER TABLE clients ADD COLUMN IF NOT EXISTS color           text;                       -- card colour, e.g. #7c3aed
ALTER TABLE clients ADD COLUMN IF NOT EXISTS emoji           text;                       -- card emoji, e.g. 🏢
ALTER TABLE clients ADD COLUMN IF NOT EXISTS booking_link    text;                       -- calendar / meeting URL
ALTER TABLE clients ADD COLUMN IF NOT EXISTS notes           text;
ALTER TABLE clients ADD COLUMN IF NOT EXISTS active          boolean     DEFAULT true;
ALTER TABLE clients ADD COLUMN IF NOT EXISTS updated_at      timestamptz DEFAULT now();
