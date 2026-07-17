-- ============================================================
-- Copy Bank → campaign linkage (A/B winner tracing)
-- Records which Copy Bank copy a merge pulled, so Copy Bank can
-- surface the A/B winner for campaigns that used a given copy.
-- Run once in the Supabase SQL editor.
-- ============================================================

ALTER TABLE campaigns ADD COLUMN IF NOT EXISTS copy_territory TEXT;
ALTER TABLE campaigns ADD COLUMN IF NOT EXISTS copy_industry  TEXT;

-- Fast lookup of campaigns by the copy they used
CREATE INDEX IF NOT EXISTS campaigns_copy_source_idx
    ON campaigns (client_id, copy_territory, copy_industry);
