-- Run in Supabase SQL Editor
ALTER TABLE campaigns
  ADD COLUMN IF NOT EXISTS completed_at timestamptz;
