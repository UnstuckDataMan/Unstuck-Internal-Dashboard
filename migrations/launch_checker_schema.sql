-- Launch Checker records table
-- Run this once in the Supabase SQL Editor

CREATE TABLE IF NOT EXISTS launch_checker_records (
    id                   TEXT        PRIMARY KEY,
    launch_name          TEXT        NOT NULL,
    user_name            TEXT        NOT NULL,
    role                 TEXT        NOT NULL,
    client_name          TEXT        NOT NULL,
    client_type          TEXT        NOT NULL,
    channels             JSONB       NOT NULL DEFAULT '[]',
    active_checks        JSONB       NOT NULL DEFAULT '[]',
    completed_sub_steps  JSONB       NOT NULL DEFAULT '[]',
    status               TEXT        NOT NULL DEFAULT 'in_progress',
    timestamp_started    TIMESTAMPTZ NOT NULL,
    timestamp_updated    TIMESTAMPTZ NOT NULL,
    timestamp_completed  TIMESTAMPTZ
);

-- Index for the 2-month auto-deletion query
CREATE INDEX IF NOT EXISTS idx_launch_checker_started
    ON launch_checker_records (timestamp_started);
