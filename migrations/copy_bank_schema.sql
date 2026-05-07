-- Copy Bank: shared outreach copy storage
-- Run once in the Supabase SQL editor

CREATE TABLE IF NOT EXISTS copy_bank_templates (
    key     TEXT PRIMARY KEY,   -- e.g. "US_PR", "UK_Marketing"
    content JSONB NOT NULL DEFAULT '{
        "email":    [{"subject":"","body":""}],
        "flyout":   [{"body":""}],
        "linkedin": [{"body":""}]
    }'::jsonb
    -- content.email    = array of {subject, body}  — multiple variations
    -- content.flyout   = array of {body}            — multiple variations
    -- content.linkedin = array of {body}            — ordered sequence steps
);

ALTER TABLE copy_bank_templates DISABLE ROW LEVEL SECURITY;


-- ── Copy Bank: pending approval drafts ────────────────────────────────────────
-- Holds draft content submitted for review before going live.
-- One row per key; re-requesting overwrites the existing pending draft.

CREATE TABLE IF NOT EXISTS copy_bank_pending (
    id          UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    key         TEXT        NOT NULL UNIQUE,   -- matches copy_bank_templates key format
    client_name TEXT        NOT NULL,
    territory   TEXT        NOT NULL,
    industry    TEXT        NOT NULL,
    content     JSONB       NOT NULL,
    status      TEXT        NOT NULL DEFAULT 'pending', -- 'pending' | 'approved'
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE copy_bank_pending DISABLE ROW LEVEL SECURITY;


-- ── Copy Bank: approval audit log ─────────────────────────────────────────────
-- Immutable append-only log. One row written per approval event.

CREATE TABLE IF NOT EXISTS copy_approval_logs (
    id               UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    key              TEXT        NOT NULL,
    client_name      TEXT        NOT NULL,
    territory        TEXT        NOT NULL,
    industry         TEXT        NOT NULL,
    content_snapshot JSONB       NOT NULL,
    approved_by      TEXT        NOT NULL,  -- 'Ollie' | 'Chris' | 'Leo'
    requested_at     TIMESTAMPTZ NOT NULL,
    approved_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE copy_approval_logs DISABLE ROW LEVEL SECURITY;
