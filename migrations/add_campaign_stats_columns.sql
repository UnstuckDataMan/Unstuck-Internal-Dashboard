-- Add per-campaign lead/reply/interested/unsubscribe counters.
-- Run once in the Supabase SQL Editor.

ALTER TABLE campaigns
  ADD COLUMN IF NOT EXISTS lead_count        INT DEFAULT 0,
  ADD COLUMN IF NOT EXISTS reply_count       INT DEFAULT 0,
  ADD COLUMN IF NOT EXISTS interested_count  INT DEFAULT 0,
  ADD COLUMN IF NOT EXISTS unsubscribe_count INT DEFAULT 0;
