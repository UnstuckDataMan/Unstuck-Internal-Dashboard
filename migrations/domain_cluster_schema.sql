-- Domain Cluster Builder. Apply once in Supabase SQL Editor before enabling storage.
-- Mirrors the repository's existing manual, idempotent SQL migration workflow.
-- Server service key only; Google dashboard sessions are enforced by FastAPI.
BEGIN;
CREATE TABLE IF NOT EXISTS public.domain_cluster_portfolios (
    id uuid PRIMARY KEY,
    workspace_id text NOT NULL,
    owner_id text NOT NULL,
    name text NOT NULL CHECK (char_length(name) BETWEEN 1 AND 100),
    revision integer NOT NULL DEFAULT 1 CHECK (revision > 0),
    state jsonb NOT NULL DEFAULT '{}'::jsonb CHECK (jsonb_typeof(state) = 'object'),
    updated_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS domain_cluster_owner_updated_idx
    ON public.domain_cluster_portfolios (workspace_id, owner_id, updated_at DESC);
ALTER TABLE public.domain_cluster_portfolios ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON public.domain_cluster_portfolios FROM PUBLIC, anon, authenticated;
GRANT SELECT, INSERT, UPDATE ON public.domain_cluster_portfolios TO service_role;
-- No browser RLS policies: only the server service role may access this table.
NOTIFY pgrst, 'reload schema';
COMMIT;
