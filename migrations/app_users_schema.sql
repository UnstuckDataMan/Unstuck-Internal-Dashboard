-- ── app_users ────────────────────────────────────────────────────────────────
-- Access-control allowlist for the internal dashboard.
--
-- Only emails present here (and active = true) may sign in via Google SSO.
-- `role` sets the default tool bundle (see ROLE_TOOLS in app/auth.py); the two
-- override arrays let an admin grant or revoke individual tools per person:
--
--     effective tools = ROLE_TOOLS[role]  ∪ tools_add  −  tools_remove
--
-- Tool keys (must match TOOLS in app/auth.py):
--   dnc, mail_merge, copy_bank, reply_bank, bd_targeting, campaigns,
--   launch_checker, targeting_checker, gender, city, admin_users
--
-- Roles: admin | reviewer | sdr | viewer

create table if not exists app_users (
    email        text primary key,                       -- always stored lowercased
    name         text,
    role         text        not null default 'sdr',
    tools_add    text[]      not null default '{}',       -- per-user grants on top of role
    tools_remove text[]      not null default '{}',       -- per-user revokes from role
    active       boolean     not null default true,
    created_at   timestamptz not null default now(),
    last_login   timestamptz
);

-- Seed the owner as admin so the dashboard is reachable immediately after
-- auth is switched on. Add the rest of the team via the /admin/users UI.
insert into app_users (email, name, role)
values ('chris@unstuck-agency.com', 'Chris', 'admin')
on conflict (email) do nothing;
