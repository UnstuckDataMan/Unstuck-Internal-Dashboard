# Unstuck AI Copy Studio

Copy Studio is an internal drafting and review workflow at `/copy-studio`, linked
from the existing dashboard. It does not send campaigns, publish Copy Bank
content, mutate Smartlead/Sheets, or connect to Meet Alfred. This feature branch
has not been released to production.

## What is implemented

Select an authorised client, import existing sources, confirm onboarding coverage,
extract and approve a cited brief, choose an audience/channel and generate copy.
Approved-template assembly preserves reviewed wording and works without a model.
New drafts propose distinct angles with email subject variants, initial messages,
follow-ups and a closing message, or LinkedIn connection/initial/follow-up/closing
steps. Limits are configurable. Citations appear in the review panel, outside the
outreach. Edits save new immutable versions; AI draft edits need a fresh semantic
check before human approval. Nothing sends automatically.

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
| `COPY_STUDIO_PROVIDER` | `openai` (default selection) or `anthropic`; not a claim of writing quality |
| `COPY_STUDIO_MODEL` | Explicit company-approved model ID; no untested studio model default |
| `COPY_STUDIO_OPENAI_API_KEY` | Company OpenAI API credential; existing `OPENAI_API_KEY` is also accepted |
| `COPY_STUDIO_ANTHROPIC_API_KEY` | Optional company Claude credential; existing `ANTHROPIC_API_KEY` is also accepted |
| `COPY_STUDIO_PRICING_JSON` | Verified USD rates per million tokens, keyed by `provider:model` |

The existing `OPENAI_MODEL`/`ANTHROPIC_MODEL` can supply the selected model if the
studio-specific setting is absent. Studio settings do not modify other tools'
model settings. Only one provider is necessary; choose between providers using
the same evidence-backed briefs and blind human reviews, not a brand preference.
See the [evaluation workflow](copy-studio-evaluation.md).

Read AI source sync reads full stored `pulse_meetings` onboarding transcripts from
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
Assembly makes no model calls and therefore has no model API charge.

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

## Staging migration and access checks

1. Back up the target staging database and confirm its existing `agencies` table
   and explicit client-to-agency mappings. Rows with null agency ownership are not
   silently admitted. Do not run the migration against production at this stage.
2. Review and apply [`migrations/ai_copy_studio.sql`](../migrations/ai_copy_studio.sql)
   to staging using the established SQL/migration workflow. PostgreSQL 15+ is
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

The old Copy Bank approve endpoint also now requires admin/reviewer and ignores
submitted approver names. Its existing browser/anonymous-table exposure and
nontransactional legacy approval workflow are **not** repaired by this isolated
change. Review inherited access before widening client access; this implementation
does not enable an external client portal.

## Durable worker and historical backfill

The web process does not run model jobs in in-process background threads. Start a
separate worker with the same private database and server settings:

```bash
.venv/bin/python -m app.utils.copy_studio.worker --schedule
```

For a single diagnostic/local job:

```bash
.venv/bin/python -m app.utils.copy_studio.worker --once
```

On Render, after release approval, create a **Background Worker** using the same
approved commit/repository, build command `pip install -r requirements.txt`, and
start command `python -m app.utils.copy_studio.worker --schedule`. Give it the same
configuration as the web service. This is a deployment step, not performed by the
feature branch or by `render.yaml` in this change.

The first **Sync sources** for a client enrolls it for hourly refreshes. A sync
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

Final offline validation: 1,696 Python tests passed, along with the Copy Studio,
benchmark, ROI audit and growth-math JavaScript checks. The broad historical
browser-script glob is not a passing release check: several scripts require
separate fixture servers on ports 8765/8766, and some older mocked Gmail browser
checks report console assertions. Those unrelated scripts were not changed.
The Copy Studio browser workflow was checked separately against its real local
HTTP API and worker. Live sources, live model calls and human copy-quality ratings
remain pending; no paid model calls were made during implementation.

## Release gates and rollback

Before production: verify staging PostgREST permissions and client mapping; run
bounded read-only Smartlead/Sheets/onboarding probes; review actual source coverage
and exact historical attribution; verify current provider contracts and price
settings; run paid pilot calls; collect blind human ratings on 5–10 varied clients
and work toward 50–100 held-out real briefs. No unsupported claims/cross-client
leakage, enforced exclusions, valid channel variables and visible completeness
are gates for the evaluated set. Demonstrate improved reviewer preference/editing
effort before broad access. Automated checks alone cannot satisfy those gates.

After explicit release approval, apply the reviewed migration and deploy the
approved commit plus worker. Do not treat earlier benchmark deployment approval
as approval for this feature.

Rollback: stop the Copy Studio worker first, remove the tool grant or revert this
feature's application commit, and leave the private tables intact for audit and
cost reconciliation. Existing tools' databases are unchanged by the migration.
Do not drop the version/usage history as a rollback shortcut. Paused jobs resume
from durable state when an approved worker is restarted.
