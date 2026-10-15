-- Shared workspace tracker. Existing private domain portfolios are retained.
BEGIN;
CREATE TABLE IF NOT EXISTS public.gmail_setup_boards (
    workspace_id text PRIMARY KEY,
    revision integer NOT NULL DEFAULT 1 CHECK (revision > 0),
    state jsonb NOT NULL CHECK (jsonb_typeof(state) = 'object'),
    updated_at timestamptz NOT NULL DEFAULT now()
);
ALTER TABLE public.gmail_setup_boards ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON public.gmail_setup_boards FROM PUBLIC, anon, authenticated;
GRANT SELECT, INSERT, UPDATE ON public.gmail_setup_boards TO service_role;
NOTIFY pgrst, 'reload schema';
COMMIT;
