-- Optional growth planner. Server access only, scoped to agency + client.
-- Reports carry only the inclusion flag; the explicitly live planner is separate.
BEGIN;
CREATE TABLE IF NOT EXISTS public.pulse_growth_plans (
    agency_id uuid NOT NULL REFERENCES public.agencies(id) ON DELETE CASCADE,
    client_id uuid NOT NULL REFERENCES public.clients(id) ON DELETE CASCADE,
    enabled boolean NOT NULL DEFAULT false,
    include_in_reports boolean NOT NULL DEFAULT false,
    approved_by text NOT NULL DEFAULT '',
    data jsonb NOT NULL DEFAULT '{}'::jsonb CHECK (jsonb_typeof(data) = 'object'),
    version integer NOT NULL DEFAULT 1 CHECK (version > 0),
    updated_at timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY (agency_id, client_id),
    CHECK (NOT include_in_reports OR enabled)
);
ALTER TABLE public.pulse_growth_plans ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON public.pulse_growth_plans FROM PUBLIC, anon, authenticated;
GRANT SELECT, INSERT, UPDATE, DELETE ON public.pulse_growth_plans TO service_role;
CREATE INDEX IF NOT EXISTS pulse_growth_plans_client_idx ON public.pulse_growth_plans(client_id);
COMMIT;
