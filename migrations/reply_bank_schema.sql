-- Reply Bank: shared template/snippet/options storage
-- Run once in the Supabase SQL editor

CREATE TABLE IF NOT EXISTS reply_bank_templates (
    key     TEXT PRIMARY KEY,   -- e.g. "US_Standard_Email_PR"
    content TEXT                -- NULL means "not available" for this combination
);

CREATE TABLE IF NOT EXISTS reply_bank_snippets (
    key     TEXT PRIMARY KEY,   -- e.g. "leoLink", "usPricing", "payPerLead"
    content TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS reply_bank_options (
    key    TEXT PRIMARY KEY,    -- "industries", "territories", or "senders"
    values TEXT[] NOT NULL
);

ALTER TABLE reply_bank_templates DISABLE ROW LEVEL SECURITY;
ALTER TABLE reply_bank_snippets  DISABLE ROW LEVEL SECURITY;
ALTER TABLE reply_bank_options   DISABLE ROW LEVEL SECURITY;
