-- ============================================================================
-- Client Performance & Reports — Read AI meetings, their AI summaries, and
-- per-client onboarding notes
--
-- Run once in the Supabase SQL editor, after pulse_growth_plans.sql.
-- Safe to re-run.
--
-- WHAT ARRIVES
-- Read AI posts a meeting report to /webhooks/read-ai when a meeting ends
-- (a workspace webhook, so every team member's meetings come through). Each
-- report becomes one pulse_meetings row, matched to a client by the
-- attendees' email domains against clients.domains. A meeting that matches no
-- client, or more than one, waits with client_id NULL until an account
-- manager assigns it.
--
-- WHY THE TRANSCRIPT IS STORED
-- Every point in the AI summary links back to the passage it came from, and
-- the passage is shown in the dashboard. Fetching it from Read AI on demand
-- would need an API this account may not have; the webhook already sends it.
--
-- WHY service_role ONLY
-- Transcripts are client conversations. Like pulse_growth_plans, these
-- tables are closed to the anon key and read only by the server with
-- SUPABASE_SERVICE_ROLE_KEY.
--
-- clients.domains
-- The client's own email and web domains (acme.com), edited on Client
-- Profiles. The app reads it with a select fallback, so a project that has
-- not run this file keeps working; meetings simply match nobody.
-- ============================================================================

BEGIN;

ALTER TABLE public.clients
    ADD COLUMN IF NOT EXISTS domains text[] NOT NULL DEFAULT '{}';

CREATE TABLE IF NOT EXISTS public.pulse_meetings (
    id                   uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    agency_id            uuid NOT NULL REFERENCES public.agencies(id) ON DELETE CASCADE,
    client_id            uuid REFERENCES public.clients(id) ON DELETE SET NULL,
    session_id           text NOT NULL,
    request_id           text NOT NULL DEFAULT '',
    title                text NOT NULL DEFAULT '',
    start_time           timestamptz,
    end_time             timestamptz,
    owner_email          text NOT NULL DEFAULT '',
    participants         jsonb NOT NULL DEFAULT '[]'::jsonb,
    report_url           text NOT NULL DEFAULT '',
    read_summary         text NOT NULL DEFAULT '',
    action_items         jsonb NOT NULL DEFAULT '[]'::jsonb,
    -- Normalised speaker blocks: [{"i", "speaker", "start_ms", "end_ms", "text"}]
    transcript           jsonb NOT NULL DEFAULT '[]'::jsonb,
    match_status         text NOT NULL DEFAULT 'unmatched'
                         CHECK (match_status IN ('matched', 'ambiguous', 'unmatched', 'manual', 'ignored')),
    candidate_client_ids uuid[] NOT NULL DEFAULT '{}',
    is_onboarding        boolean NOT NULL DEFAULT false,
    -- {"summary", "positives": [{"point", "block", "quote"}], "negatives": [...]}
    analysis             jsonb NOT NULL DEFAULT '{}'::jsonb CHECK (jsonb_typeof(analysis) = 'object'),
    analysis_status      text NOT NULL DEFAULT 'pending'
                         CHECK (analysis_status IN ('pending', 'running', 'done', 'failed', 'skipped')),
    analysis_error       text NOT NULL DEFAULT '',
    analysed_at          timestamptz,
    received_at          timestamptz NOT NULL DEFAULT now(),
    updated_at           timestamptz NOT NULL DEFAULT now(),
    UNIQUE (agency_id, session_id)
);

CREATE INDEX IF NOT EXISTS pulse_meetings_client_idx
    ON public.pulse_meetings (agency_id, client_id, start_time DESC);
CREATE INDEX IF NOT EXISTS pulse_meetings_status_idx
    ON public.pulse_meetings (agency_id, match_status);

-- One AI overview per client, rebuilt from the per-meeting analyses.
-- meeting_ids records which meetings it was built from, so the panel can tell
-- when a newer meeting has not been folded in yet.
CREATE TABLE IF NOT EXISTS public.pulse_meeting_insights (
    agency_id    uuid NOT NULL REFERENCES public.agencies(id) ON DELETE CASCADE,
    client_id    uuid NOT NULL REFERENCES public.clients(id) ON DELETE CASCADE,
    -- {"summary", "successes": [{"point", "refs": [...]}], "pain_points": [...]}
    overview     jsonb NOT NULL DEFAULT '{}'::jsonb CHECK (jsonb_typeof(overview) = 'object'),
    status       text NOT NULL DEFAULT 'pending'
                 CHECK (status IN ('pending', 'running', 'done', 'failed')),
    error        text NOT NULL DEFAULT '',
    meeting_ids  uuid[] NOT NULL DEFAULT '{}',
    generated_at timestamptz,
    updated_at   timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY (agency_id, client_id)
);

-- One set of onboarding notes per client. The AI draft fills an untouched
-- row directly; once a person has edited the notes (edited_by set), a new
-- draft waits in pending_draft and never overwrites their work.
CREATE TABLE IF NOT EXISTS public.pulse_onboarding_notes (
    agency_id          uuid NOT NULL REFERENCES public.agencies(id) ON DELETE CASCADE,
    client_id          uuid NOT NULL REFERENCES public.clients(id) ON DELETE CASCADE,
    body               text NOT NULL DEFAULT '',
    source_meeting_id  uuid REFERENCES public.pulse_meetings(id) ON DELETE SET NULL,
    ai_generated_at    timestamptz,
    edited_by          text NOT NULL DEFAULT '',
    edited_at          timestamptz,
    pending_draft      text NOT NULL DEFAULT '',
    pending_meeting_id uuid REFERENCES public.pulse_meetings(id) ON DELETE SET NULL,
    version            integer NOT NULL DEFAULT 1 CHECK (version > 0),
    updated_at         timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY (agency_id, client_id)
);

ALTER TABLE public.pulse_meetings         ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pulse_meeting_insights ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pulse_onboarding_notes ENABLE ROW LEVEL SECURITY;

REVOKE ALL ON public.pulse_meetings         FROM PUBLIC, anon, authenticated;
REVOKE ALL ON public.pulse_meeting_insights FROM PUBLIC, anon, authenticated;
REVOKE ALL ON public.pulse_onboarding_notes FROM PUBLIC, anon, authenticated;

GRANT SELECT, INSERT, UPDATE, DELETE ON public.pulse_meetings         TO service_role;
GRANT SELECT, INSERT, UPDATE, DELETE ON public.pulse_meeting_insights TO service_role;
GRANT SELECT, INSERT, UPDATE, DELETE ON public.pulse_onboarding_notes TO service_role;

COMMIT;
