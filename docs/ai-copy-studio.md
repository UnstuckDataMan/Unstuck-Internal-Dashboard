# Unstuck AI Copy Studio

Copy Studio is an internal drafting and review workflow at `/copy-studio`, linked
from **Campaign Management** in the existing dashboard. An explicit reviewer action can append checked,
approved copy to a selected Copy Bank area. It does not send campaigns,
mutate Smartlead/Sheets, or connect to Meet Alfred. Production migrations are applied;
the web/worker release still requires deployment verification.

## What is implemented

Select an authorised client from the same canonical directory as Client Profiles.
The workspace loads their saved domain and profile link, then queues their existing
Copy Bank and Read AI context automatically (without calling a writing model).
Missing or hour-old Smartlead/manual campaign snapshots are refreshed as well;
opening the client enrols them in the worker's hourly source refreshes. Calls already
assigned to the client are included across all meeting types; previously unassigned
calls require a unique participant-domain match across the complete agency directory.
Missing domains point back to Client Profiles rather than inferring a website from
sending inboxes. Imports cover stored webhook transcripts, not unverified full Read AI
account history.

Review sources, confirm full-call coverage,
extract and approve a cited brief, choose an audience/channel and generate copy.
Approved-template assembly preserves reviewed wording and works without a model.
New drafts propose distinct angles with email subject variants, initial messages,
follow-ups and a closing message, or LinkedIn connection/initial/follow-up/closing
steps. Limits are configurable. Citations appear in the review panel, outside the
outreach. Edits save new immutable versions; AI draft edits need a fresh semantic
check before human approval. Nothing sends automatically.

UK/US English is suggested from the approved, cited call facts. An explicit language
preference takes priority; geography is only a tentative hint, and conflicting
preferences require a manual choice. The user can override the suggestion.
**Check spelling & final review** reviews the exact saved subjects/messages,
returns proposed corrections without rewriting them, and creates a new pending
version. Human edits invalidate that check and require another review. These
checks use the configured provider and are included in costs.

After approving the checked version, **Send to Copy Bank** lets an admin/reviewer
select the client's territory, industry, compatible section and angle. The preview
shows exactly what will be appended. Confirmation rechecks the source version,
approval and destination snapshot; a conflicting edit requires a fresh preview.
Existing subjects, variations and unrelated sections are preserved. Retrying the
same publication returns its receipt instead of duplicating content. Editing the
draft after approval blocks publication of that superseded version.

The separate **Brief the writer** field supplies creative direction to drafting
and rechecks. It cannot approve new facts, permission or claims, or override client
constraints. Switching clients clears creative inputs. A new conversation manifest
excludes stale assignments from brief extraction; changed or incomplete context
blocks generation/approval until refreshed and reviewed. Existing client Copy Bank
copy is available for review and, once approved, informs generation. A current
source manifest excludes removed or externally changed entries; reviewed edits
retain their original import lineage rather than silently losing eligibility.

LinkedIn uses approved LinkedIn Copy Bank examples plus the selected client's
approved facts. It has no performance-proven badge or Meet Alfred metrics. The
future Meet Alfred boundary is explicitly disabled even if an old Pulse setting
exists. Email references can use reviewed reusable versions of exact sent copy;
results remain attached to their original version, not to the rewritten template.

Smartlead and manual campaign results reuse Pulse's comparison rules. They remain
separate measures, with denominators, periods, attribution and freshness. A
positive-response leader is not statistical significance. Losing arms stay in the
corpus. Recipient-specific sent text stays private; preparing a reusable reference
requires explicit sanitisation and approval. See [source contracts](ai-copy-studio-sources.md).

Client facts, methodology, references, sent versions, performance, reviews, usage
and job checkpoints use separate record kinds in private, immutable storage.
Imports are pending; they never overwrite human-approved facts. Source versions
and brief edits remain visible. To resolve contradictory onboarding statements,
exclude the outdated fact or correct it against the source. A resolution note
does not make two conflicting statements safe to use.

## Configuration

Use the existing Python dependencies. Web and background worker use the same
database and configuration, entered through server/environment secret settings.
Never place credentials in Git, case exports, browser JavaScript or command lines.

| Setting | Purpose |
| --- | --- |
| `SUPABASE_URL` | Existing project; prefer staging during validation |
| `SUPABASE_SERVICE_ROLE_KEY` | Server-only access to the new private tables and existing source records |
| `COPY_STUDIO_AGENCY_ID` | Explicit existing agency UUID; `PULSE_AGENCY_ID` is accepted as a fallback |
| `SMARTLEAD_API_KEY` | Reuse the dashboard reporting key for GET-only source access |
| `GOOGLE_SHEETS_SA_JSON` | Existing base64 service-account credential; share manual campaign sheets as Viewer |
| `COPY_STUDIO_PROVIDER` | `openai` or `anthropic`; this setup selects `anthropic` at the user's request |
| `COPY_STUDIO_MODEL` | Explicit company-approved model ID; no untested studio model default |
| `COPY_STUDIO_OPENAI_API_KEY` | Company OpenAI API credential; existing `OPENAI_API_KEY` is also accepted |
| `COPY_STUDIO_ANTHROPIC_API_KEY` | Optional company Claude credential; existing `ANTHROPIC_API_KEY` is also accepted |
| `COPY_STUDIO_PRICING_JSON` | Verified USD rates per million tokens, keyed by `provider:model` |
| `COPY_STUDIO_EMBEDDED_WORKER` | Set exactly `1` to run the Studio worker alongside the web process using `scripts/run_dashboard.py`; otherwise web only |

The existing `OPENAI_MODEL`/`ANTHROPIC_MODEL` can supply the selected model if the
studio-specific setting is absent. Studio settings do not modify other tools'
model settings. Only one provider is necessary; choose between providers using
the same evidence-backed briefs and blind human reviews, not a brand preference.
See the [evaluation workflow](copy-studio-evaluation.md).

This environment selects `claude-sonnet-5-5`, verified in the authenticated models
listing on 28 October 2026. Standard synchronous global inference rates verified
against [Anthropic's pricing documentation](https://platform.claude.com/docs/en/about-claude/pricing)
are USD 2 per million input tokens, USD 10 output and USD 0.10 cache reads. The
bundled rate card is dated 28 October 2026 and used for that exact model when no
explicit override is present. An optional override is
`{"anthropic:claude-sonnet-5-5":{"input":2,"output":10,"cached_input":0.1}}`.
Invalid configuration and unsupported models remain unpriced; historical records
are never repriced.
This adapter does not request batch processing, regional inference or explicit
cache creation. Cache-write rates are intentionally absent: unexpected cache
writes remain unpriced until their duration-specific charges are configured.

Read AI source sync reads full stored `pulse_meetings` client transcripts from
the existing webhook pipeline. Calls missing from that pipeline need ingestion
there first. Stored transcript completeness is distinguished from confirmed
full-call coverage. A reviewer can compare the original call and confirm coverage,
creating an immutable source revision; extract the brief again from that revision.
The linked Notion SOPs/training history have **not** been ingested in this build.
The seeded methodology is labelled a research summary for review, not an imported
or approved corpus. Fixed implementation guardrails remain in force.

## Cost reporting

The compact cost bar shows **today, this calendar month and all time**, using
`Europe/London` boundaries, in USD. It covers this tool's recorded extraction,
writing, independent reviews, rechecks and opted-in evaluation calls. It does not
invent costs for activity before tracking began or for other dashboard tools.
Initial approved-template assembly makes no model calls. Final proofreading and
semantic checks of edited assemblies use the model and incur tracked charges.

Every HTTP attempt first creates an unknown usage record. Returned provider token
usage finalises that attempt before output parsing, so charged malformed or
incomplete output is still counted. Retries are distinct attempts. A crash leaves
an unknown attempt; absent usage or pricing means an incomplete total, never zero.
Database failure before recording an attempt prevents the model call. Failure
after an attempt stops the job for review rather than silently repeating charges.

Configure prices from the current provider pricing/account terms. The shape is:

```text
{"openai:<model-id>": {"input": <USD-per-million>, "output": <USD-per-million>,
                        "cached_input": <USD-per-million>},
 "anthropic:<model-id>": {"input": <USD-per-million>, "output": <USD-per-million>,
                          "cached_input": <USD-per-million>,
                          "cache_creation_input": <USD-per-million>}}
```

These are placeholders, not quoted model prices. Cache rates are required only
when that usage occurs. Each call saves its rate snapshot; changing future rates
does not rewrite prior costs. Totals are estimated API charges, not invoices or
hosting/tax totals. Incomplete historical entries require reconciliation against
provider records; setting a rate later does not backfill an unknown charge.

## Migration and access checks

1. Back up the target staging database and confirm its existing `agencies` table
   and explicit client-to-agency mappings. Rows with null agency ownership are not
   silently admitted. Validate changes in staging before an authorised production release.
2. Review and apply [`migrations/ai_copy_studio.sql`](../migrations/ai_copy_studio.sql),
   then [`migrations/ai_copy_studio_publish.sql`](../migrations/ai_copy_studio_publish.sql)
   to staging using the established SQL/migration workflow. The publishing bridge
   requires the existing `clients`, `copy_bank_templates` and `copy_bank_pending`
   tables. PostgreSQL 15+ is
   required for the security-invoker view. The migration is transactional and
   idempotent; it creates only `copy_studio_*` objects.
3. Verify anonymous/authenticated callers cannot read or write these tables/view
   or invoke their RPCs. The server service role has scoped application access.
   New tables do not inherit the broad public access on older dashboard tables.
4. Run the offline tests and disposable-database SQL assertions below, then check
   the actual staging PostgREST routes/RPC schema with the server credential.
5. Sign in again to receive the new `copy_studio` tool grant. Default access is
   internal admin/reviewer/SDR; viewer/client roles cannot enter. An explicit
   `client_ids` assignment restricts the client list and every operation.
   Approval is admin/reviewer only, with identity taken from the signed session.

For the authorised 28 October 2026 pilot, both migrations have now been applied to
production and their access controls verified. The new RLS tables, views and RPCs
are private to the server service role. The deployed web service's Supabase and
Anthropic credentials also authenticated successfully (HTTP 200). These checks
confirm database/provider access, not a completed application rollout.

The old Copy Bank approve endpoint also now requires admin/reviewer and ignores
submitted approver names. Its existing browser/anonymous-table exposure and
nontransactional legacy approval workflow are **not** repaired by this isolated
change. Review inherited access before widening client access; this implementation
does not enable an external client portal.

## Durable worker and historical backfill

The web process does not run model jobs in in-process background threads. The
pilot can use the **existing Render web service**, without creating or paying for
an additional service:

1. Keep its existing runtime, build command, plan and other environment settings.
2. Set `COPY_STUDIO_EMBEDDED_WORKER=1` in that service's environment.
3. Set its start command to `python scripts/run_dashboard.py` and deploy the approved commit.

The launcher runs the dashboard and one separate Studio worker process with the
same private settings. The worker uses durable database leases and hourly source
scheduling. Without the opt-in flag, the launcher runs only the web process.
Worker failures retry with backoff, then enter a five-minute cooldown while the
web process stays running; recovery is attempted afterwards. Web-process failure
still shuts down both child process groups. Normal shutdown allows 20 seconds
before terminating remaining children.
Check both process startup logs and a completed Studio job after deployment;
`/healthz` alone does not prove that the worker is processing jobs.

A **dedicated background worker** remains an optional way to isolate CPU/memory.
Disable `COPY_STUDIO_EMBEDDED_WORKER` on the web service before enabling a dedicated
worker, so both deployments do not schedule the same source refreshes. Use the
same approved commit, dependencies, private configuration and this command:

```bash
.venv/bin/python -m app.utils.copy_studio.worker --schedule
```

For a single diagnostic/local job:

```bash
.venv/bin/python -m app.utils.copy_studio.worker --once
```

On Render, its optional dedicated-worker start command is
`python -m app.utils.copy_studio.worker --schedule`. This is an alternative to the
shared-service pilot, not a deployment requirement.

Production memory usage has not been measured: the Render metrics request returned
HTTP 403. A local fixture worker peaked at approximately 122 MiB; this is not a
production capacity test and does not include combined web/worker load. Monitor
the existing instance during the pilot before increasing concurrent usage.

Selecting a client or choosing **Sync sources** enrolls it for hourly refreshes. A sync
walks historical campaign pages and checkpoints each step. Retries/restarts resume
the saved page/cursor. Repeated imports deduplicate immutable source content.
Subsequent refreshes revisit history because old reply statuses can change; they
do not assume provider events are append-only. Provider rate limits use bounded
backoff. Multiple workers claim leased jobs atomically; interrupted jobs become
claimable after lease expiry. The scheduler enqueues idempotently per client,
source and hour. Model extraction/generation is explicit, never an hourly expense.

Partial source failures retain successful imports and report missing coverage.
`succeeded` means the job finished its attempt; inspect `result.complete` and
source reports. Unavailable sheets do not block other campaigns. Job state,
attempts, checkpoints, errors and last-success timestamps are durable. There is no
implicit CSV upload requirement or dummy-live-data fallback.

## Development and verification

An explicit `COPY_STUDIO_SQLITE_PATH` enables local development storage (file
permissions 0600). It is never selected because Supabase fails, and must not be set
on production. Local fixtures must be clearly synthetic, outside live source data.

```bash
.venv/bin/pytest -q
node tests/test_copy_studio_ui.cjs
```

For a **disposable** PostgreSQL database with Supabase-style roles and a minimal
`agencies` fixture table, apply the migration and run:

```bash
psql -v ON_ERROR_STOP=1 -f migrations/ai_copy_studio.sql
psql -v ON_ERROR_STOP=1 -f tests/sql/test_ai_copy_studio.sql
```

The SQL assertions use reserved fixture agency IDs and roll back their writes;
they are for an isolated database, not production. The migration was applied twice
and tested in isolated PostgreSQL 17, including RLS privileges, immutable history,
cross-client references, retries and concurrent job claims.

The local browser check used a real local HTTP API, private SQLite storage and a
separate worker to complete keyless assembly, editing, approval and sanitised
reference promotion. Source/model contract tests use fixtures. These checks do
not prove current live provider contracts or subjective copy quality.

Final regression validation: 1,834 Python tests passed, along with the Copy Studio,
benchmark, ROI audit and growth-math JavaScript checks. The broad historical
browser-script glob is not a passing release check: several scripts require
separate fixture servers on ports 8765/8766, and some older mocked Gmail browser
checks report console assertions. Those unrelated scripts were not changed.
The Copy Studio browser workflow was checked separately against its real local
HTTP API and worker. New publishing panels were additionally checked at desktop
and mobile sizes with synthetic gateway responses; no external copy was written.

The publishing migration was applied twice in a network-disabled PostgreSQL 17
container. `tests/sql/test_ai_copy_studio_publish.sql` validates private RPC/table
permissions, version/approval/context checks, append preservation and immutable
receipts. `tests/sql/check_copy_studio_publish_concurrency.py --container <name>`
exercises actual concurrent sessions: duplicate publications, competing selections,
a legacy editor's update, and a legacy editor's first insert. All four races passed.
These scripts are only for a disposable fixture database containing the existing
table shapes, never production. The test container and volumes were removed.

Live checks on 28 October 2026 authenticated Supabase, Smartlead (1,140 campaigns),
Google Sheets and Anthropic's model list. Client Profiles/stored Read AI schemas
were checked read-only. A synthetic Sonnet 5.5 response then verified native JSON
output and token usage with no client data sent. The first two requests exposed
Sonnet 5.5's unsupported forced-tool option; the adapter now uses documented
`output_config.format` and retains local schema validation. The setup ledger keeps
those rejected attempts' absent usage unknown instead of claiming they were free.
Those initial source probes did not write client content. During the later authorised
release preparation, both production migrations were applied and private access
controls verified; the earlier missing-table blocker is resolved. Web-service
Supabase and Anthropic credential checks returned HTTP 200. Application deployment
verification, production memory measurements, human copy-quality ratings and
real-client model pilots remain pending. Full Read AI account history has not been
established by the stored-webhook checks.

## Release gates and rollback

Before broad client access: verify PostgREST permissions and client mapping; run
bounded read-only Smartlead/Sheets/onboarding probes; review actual source coverage
and exact historical attribution; verify current provider contracts and price
settings; run paid pilot calls; collect blind human ratings on 5–10 varied clients
and work toward 50–100 held-out real briefs. No unsupported claims/cross-client
leakage, enforced exclusions, valid channel variables and visible completeness
are gates for the evaluated set. Demonstrate improved reviewer preference/editing
effort before broad access. Automated checks alone cannot satisfy those gates.

The user has explicitly authorised this internal pilot release. Deploy the approved
commit with either the shared-service worker opt-in or the optional dedicated
worker, then verify the live page and job processing. The remaining human quality
evaluation is required before broad client access.

Rollback: for the shared-service pilot, unset `COPY_STUDIO_EMBEDDED_WORKER` and
redeploy to keep the dashboard running without Studio job execution. Stop a
dedicated worker if that option was used. Remove the tool grant or revert this
feature's application commit as needed, and leave the private tables intact for audit and
cost reconciliation. The migrations do not alter existing source tables or their
grants; the publishing RPC writes only after an explicit reviewer action. To undo
a published append, use its immutable receipt and review the current Copy Bank
entry before editing; do not restore an old whole-entry snapshot over later work.
The legacy Copy Bank editor can still overwrite an entry in a later independent
edit; the new bridge guarantees preservation and concurrency checks during its
own transaction, not a rewrite of that older editing workflow.
Do not drop the version/usage history as a rollback shortcut. Paused jobs resume
from durable state when an approved worker is restarted.
