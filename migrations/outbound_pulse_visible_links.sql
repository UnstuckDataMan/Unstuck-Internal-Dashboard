-- ============================================================================
-- Client Performance & Reports — keep the portal link so it can be re-read
--
-- Run once in the Supabase SQL editor. Safe to re-run.
--
-- WHAT THIS CHANGES, AND WHAT IT COSTS
-- Until now only the SHA-256 of a portal token was stored, so the link was
-- shown exactly once and could never be recovered — only revoked and
-- re-issued. That was a deliberate trade: a database dump could not be turned
-- into a set of live client report links.
--
-- The team needs to open a client's report themselves, from the client's
-- profile, which means the link has to be readable. So the token is now stored
-- alongside its hash.
--
-- Be clear about what that actually trades. The internal dashboard sits behind
-- the organisation's Google auth, which is why showing the link in the UI is
-- safe. It is not why STORING it is safe: anyone with read access to this
-- table — a dump, a backup, a leaked service key — now holds working links to
-- every client's report. The mitigations are that the links are read-only, can
-- be revoked individually, and expire after 180 days.
--
-- token_hash stays and remains the lookup key. Resolving a link still hashes
-- the incoming token and matches on that, so the plaintext column is never
-- part of authentication — it exists only to be displayed.
--
-- EXISTING LINKS CANNOT BE RECOVERED. Rows created before this ran hold only a
-- hash, and a hash is not reversible. Those show as unavailable in the UI, with
-- re-issuing as the way to get a readable one.
-- ============================================================================

BEGIN;

ALTER TABLE pulse_client_access
    ADD COLUMN IF NOT EXISTS token TEXT;

COMMENT ON COLUMN pulse_client_access.token IS
    'Plaintext portal token, so the link can be shown on the client profile. '
    'NULL for links issued before this column existed — a hash cannot be '
    'reversed, so those can only be re-issued. Never used for lookup; '
    'token_hash remains the key.';

COMMIT;
