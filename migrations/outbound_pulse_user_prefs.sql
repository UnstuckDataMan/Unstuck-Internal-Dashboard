-- ============================================================================
-- Client Performance & Reports — per-user overview preferences
--
-- Run once in the Supabase SQL editor. Safe to re-run.
--
-- WHY A ROW'S EXISTENCE IS THE SIGNAL
-- Three states have to be told apart, and two of them look identical if the
-- list is the only thing stored:
--
--   no row       → this person has never chosen. Apply the house default
--                  (Unstuck - Business Development is hidden).
--   row, {}      → this person deliberately cleared their exclusions and wants
--                  to see every client, including that one.
--   row, {a,b}   → hide exactly a and b.
--
-- A column that is NOT NULL DEFAULT '{}' cannot express the difference between
-- the first and the second, which would make "show me everything" unsayable:
-- every save of an empty list would read back as "has not chosen yet" and
-- silently re-hide the client the user had just un-hidden.
--
-- So the row is created on the first save, and Reset DELETEs it. The delete is
-- what returns someone to the default, and it is a different action from
-- saving nothing.
--
-- WHY NO FOREIGN KEY TO app_users
-- The dashboard runs with AUTH_DISABLED in the whole test suite and as the
-- production rollback lever. The user is then 'dev@local', who has no
-- app_users row. A foreign key would turn saving a preference into an error in
-- exactly the mode that exists so the app keeps working. A preference left
-- behind by a departed colleague is a dead row, not a bug.
--
-- WHY NO FOREIGN KEY ON THE CLIENT IDS EITHER
-- Postgres cannot foreign-key the elements of an array. A deleted client
-- leaves a stale id behind, which matches nothing and changes no filter. The
-- alternative — a row per (user, client) — would need a second table just to
-- record "has chosen", which is the problem this design exists to avoid.
-- ============================================================================

BEGIN;

CREATE TABLE IF NOT EXISTS pulse_user_prefs (
    agency_id           uuid NOT NULL REFERENCES agencies(id) ON DELETE CASCADE,
    -- Always lowercased, the same convention as app_users.email.
    user_email          text NOT NULL,
    excluded_client_ids uuid[] NOT NULL DEFAULT '{}',
    updated_at          timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY (agency_id, user_email)
);

-- Enforced here rather than trusted to every caller: a mixed-case email would
-- create a second, invisible preference row for the same person.
ALTER TABLE pulse_user_prefs DROP CONSTRAINT IF EXISTS pulse_user_prefs_email_lower;
ALTER TABLE pulse_user_prefs
    ADD CONSTRAINT pulse_user_prefs_email_lower
    CHECK (user_email = lower(user_email) AND length(btrim(user_email)) > 0);

ALTER TABLE pulse_user_prefs DISABLE ROW LEVEL SECURITY;


DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'anon') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_user_prefs TO anon;
    END IF;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'authenticated') THEN
        GRANT SELECT, INSERT, UPDATE, DELETE ON pulse_user_prefs TO authenticated;
    END IF;
END;
$$;

COMMIT;
