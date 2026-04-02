-- Add location filtering to contacted_prospects
-- Run this in the Supabase SQL editor (Dashboard → SQL Editor → New query)

ALTER TABLE contacted_prospects ADD COLUMN IF NOT EXISTS location TEXT;

-- Index to speed up location-based scrub queries and table filtering
CREATE INDEX IF NOT EXISTS contacted_prospects_location_idx
    ON contacted_prospects (client_id, location);
