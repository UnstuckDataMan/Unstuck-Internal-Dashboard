-- ============================================================================
-- Client Performance & Reports — "show only these" as well as "hide these"
--
-- Run once in the Supabase SQL editor, after outbound_pulse_user_prefs.sql.
-- Safe to re-run.
--
-- WHY A MODE RATHER THAN A SECOND LIST
-- Hiding three clients and showing only three are the same selection applied
-- two ways. A second `included_client_ids` column would let both be non-empty
-- at once, and nothing in the data would say which one won — so the question
-- "what am I looking at?" would have no answer in the row.
--
-- One list plus a mode has exactly one reading. The column is renamed from
-- excluded_client_ids to client_ids because it no longer always means
-- excluded, and a column whose name contradicts half its rows is worse than a
-- rename.
--
-- THE THREE STATES STILL HOLD
--   no row                    → never chosen; the house default applies
--   row, exclude, {}          → show everything
--   row, exclude, {a,b}       → hide a and b
--   row, only,    {a,b}       → show only a and b
--
-- "only" with an empty list would be an overview with nothing in it, which
-- reads as broken rather than as a filter. The check below makes that state
-- unstorable, and the panel refuses it with a reason before it gets here.
-- ============================================================================

BEGIN;

ALTER TABLE pulse_user_prefs
    ADD COLUMN IF NOT EXISTS mode text NOT NULL DEFAULT 'exclude';

-- Renamed only if it has not been already, so the migration can be re-run.
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM information_schema.columns
                WHERE table_name = 'pulse_user_prefs'
                  AND column_name = 'excluded_client_ids') THEN
        ALTER TABLE pulse_user_prefs
            RENAME COLUMN excluded_client_ids TO client_ids;
    END IF;
END;
$$;

ALTER TABLE pulse_user_prefs DROP CONSTRAINT IF EXISTS pulse_user_prefs_mode_check;
ALTER TABLE pulse_user_prefs
    ADD CONSTRAINT pulse_user_prefs_mode_check
    CHECK (mode IN ('exclude', 'only'));

-- An empty "show only" selects nothing at all. Unstorable rather than merely
-- discouraged: a blank overview with no explanation is the worst outcome here.
ALTER TABLE pulse_user_prefs DROP CONSTRAINT IF EXISTS pulse_user_prefs_only_not_empty;
ALTER TABLE pulse_user_prefs
    ADD CONSTRAINT pulse_user_prefs_only_not_empty
    CHECK (mode <> 'only' OR cardinality(client_ids) > 0);

COMMIT;
