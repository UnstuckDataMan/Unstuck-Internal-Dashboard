# Unstuck Internal Dashboard

Existing FastAPI/Jinja dashboard deployed to Render. Run with Python 3.11+:

```sh
python -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

Configure the established Google sign-in and integrations using `.env.example` and
Render environment variables; the app reads its environment, not a committed `.env`.
`/healthz` is the existing public health endpoint. `render.yaml` retains the existing
build (`pip install -r requirements.txt`) and start commands. This feature introduces
no new service, frontend build step or database provider.

## Gmail Set / account setup tracker

Open **Operations → Gmail Set** (`/gmail-set`). The main tool tracks the shared
company setup journey: **Not started → Purchase made → Setup complete**.
Create a client first, click its card, then add accounts by entering only the
sending email. The server derives the domain automatically. The
client list shows counts instead of every domain; unfinished domains appear above
completed ones in each client’s domain list. Each domain has a stage dropdown,
and the add/edit form also lets you choose its stage. Existing flat records are
grouped automatically without changing their sheet rows or purchase history. It starts
with the instruction to use the GPT domain clustering plugin in Notion; clustering
is no longer the main workflow.

- The tracker seeds its member list from active dashboard users in the signed-in
  workspace. Members can be added, edited or removed within this tool without
  changing dashboard login access. Open work must be reassigned before removal.
- Record account manager, setup owner, client/Biz Dev section, root domain,
  sending email, order and purchase device. Purchase records snapshot the actual
  device and asset label; removing/transferring equipment preserves that history.
- Add actual inventory in batches of 0–10 phones and 0–10 PCs per person, with
  labels Phone 1–10 and PC 1–10 plus stable physical asset IDs. The unallocated pool
  supports stock awaiting assignment. No imaginary devices are seeded.
- Unused, Used and Blocked states are separate from allocation. **Block QR code**
  and **Block Payment** are explicit block reasons. All blocked equipment is red
  and excluded from purchases; previously used devices never count as new stock.
- Both the tool and the Operations card alert when the company has fewer than
  20 unused devices (phones + PCs combined). Counts refresh after every change
  and while visible. No external notifications or purchases are sent.
- Selecting a later stage records purchase/device history as needed; choosing
  Setup complete also publishes the sending email. Earlier purchase/Live history
  cannot be erased by moving a stage backwards. Completion is saved before attempting Google sync. The correct authorised tab
  is selected, columns are resolved by headings (including the moved BD Live
  column), existing exact emails/row markers are reconciled, and a native row is
  inserted immediately after the last actual email/domain entry (ignoring the
  blank/formatted tail), atomically with Live checked and checkbox validation. Readback verifies
  the email and Live value. No recovery code or password columns are read/copied.
- Failed sync remains visibly retryable. Workspace-wide conditional revisions
  and a persistent 180-second write lease prevent overlapping tool writes. A
  retry reconciles committed external writes after a timeout or process restart.
- The strategy-sheet service account requires **Editor** access. The tool can
  check Drive `capabilities.canEdit` without writing a fake setup row.
- Apply `migrations/gmail_setup_schema.sql` to the existing Supabase project. RLS
  is enabled with no browser policies/grants; only the server service role can
  access the workspace-scoped board. Google dashboard sessions, tool access and
  CSRF protect every tracker mutation. State is shared within the workspace.

The previous domain-planning APIs and private saved portfolios remain intact.
The old planner UI is available at `/domain-cluster-builder/legacy` for historical
records. `/domain-cluster-builder` now opens the account setup tracker.

## Legacy Domain Cluster Builder

The legacy planner is at `/domain-cluster-builder/legacy`. Staff and reviewer roles
receive the `gmail_set` tool by default; viewers need an explicit admin grant. Existing sessions
must sign out and back in to receive the new tool permission.

The default workflow reads the current Live checkboxes from **both** strategy tabs
for every submitted Analyse, Reanalyse or Generate request. It never substitutes a
cached inventory if Google fails. The interface uses the dashboard's dark surfaces,
purple accents and typography, with three steps: scan live domains, choose ten,
then set up the new batch. Manual paste / TXT / CSV appears when selected; cluster
editing, DNS checks, exports and settings are expandable. Planning selections saves
any unsaved inventory first, then rechecks live capacity before creating the batch.
Save portfolio remains available for keeping an analysis without a new batch.

### Storage setup

1. Apply `migrations/domain_cluster_schema.sql` in the existing Supabase SQL Editor.
   This repository uses manual, additive SQL migrations. The script is idempotent.
2. Keep `SUPABASE_URL` and the server-only `SUPABASE_SERVICE_ROLE_KEY` configured
   in Render. The anon key intentionally cannot access portfolio records.
3. Retain `SESSION_SECRET`, `APP_BASE_URL`, `GOOGLE_CLIENT_ID`,
   `GOOGLE_CLIENT_SECRET` and `ALLOWED_EMAIL_DOMAIN` for existing Google sign-in.
4. Deploy the changed repository to the existing Render service.

Portfolios are private to their signed-in owner and scoped to their email-domain
workspace. IDs are random UUIDs; server requests always filter both workspace and
owner, including revision-checked writes. Browser roles have no database grants or
RLS policies. Each atomic JSONB document contains inventory, metadata, settings,
manual overrides, planned selections, ten reversible save checkpoints and the last
30 summary analysis snapshots. The document structure follows the existing Credit
Control persistence pattern. No cross-user sharing or workspace membership editor
is introduced. `/api/domain-cluster/*` also checks sign-in and tool access when
`AUTH_DISABLED` is set; it never makes saved portfolios public.

### Fresh strategy-sheet integration

The fixed authorised spreadsheet is:
https://docs.google.com/spreadsheets/d/1PN5Lx-5c4nbouDl-Apw9YJgrYTuoDsBKqzk5Hh6Z0po/edit

- Client emails: `1) Emails - entry`, sheet ID `0`.
- Biz Dev emails: `1) BD Emails - entry`, sheet ID `838806676`.
- Reuses `GOOGLE_SHEETS_SA_JSON`, the existing base64 service-account credential.
- The user has shared the sheet as Viewer with
  `sheets-writer-unstuck-dashboar@unstuck-dnc-and-merger.iam.gserviceaccount.com`.
- An independent **read-only** credential scope reads metadata and a bounded A:G
  `values:batchGet`, with unformatted booleans. No recovery codes/addresses from
  later columns are fetched. The existing Sheets quota counter/limiter is reused.
- Only boolean `true` / text `TRUE` in the Live column includes a row. False, blank
  and unrecognised values are excluded. At risk / Needs deleting / Being replaced
  are separate flags and never silently override a checked Live value.
- Root domains are deduplicated across people and both tabs. Client/Biz Dev sources
  and supplied Client/Biz Dev lead names remain filterable; mailbox local parts are
  discarded before responses and storage. Same-root mailboxes consume one slot.
- Newly Live sheet domains become Active in the working preview. Previously
  imported sheet domains no longer Live become Paused. Manual-only records and
  confirmed cluster overrides remain intact. No source sheet is modified.
- Generate reads the sheet again immediately before classification and capacity
  validation. Saving planned domains also checks for inventory changes since the
  last saved analysis and rejects stale plans. Google failures or invalid Live
  domains prevent live-mode generation; no cached success or partial inventory is
  treated as current. A zero-domain scan never clears a portfolio automatically.
- Sheet validation shows the affected tab, row and cell, correction guidance and
  a direct Google Sheets cell link in Analyse, blocked Generate/Planned saves and
  Check setup. A fresh successful scan clears the issues; no cell contents or
  mailbox addresses are included in the feedback. Missing headings point to the
  tab's header area; access/network failures report the connection problem.
- Each tab is limited to 10,000 allocated rows; exceeding that fails explicitly.
  Portfolios support 5,000 distinct domains including historical records.

### Optional Porkbun availability

Add **PORKBUN_API_KEY** and **PORKBUN_SECRET_API_KEY** to Render environment variables.
Availability checks need read permission. URL forwarding and DMARC setup need both read and DNS permissions. Configure these at https://porkbun.com/account/api. Never paste keys into
portfolio settings, browser code, Git, or the strategy sheet. No purchase permission
is needed. Credentials remain on Render. Use Check setup in Gmail Set to verify live access; a read-only check also logs sanitized readiness flags after each production boot.

- **Check on Porkbun** checks an existing ten-name batch.
- **Only return names verified available on Porkbun** searches up to 30 candidates
  per request and returns exactly ten only when all ten passed the API and capacity
  checks. A shortfall is explicit and no partial batch is approved.
- Uses `POST https://api.porkbun.com/api/json/v3/domain/checkDomain` with at most
  ten names per call. Understands `domains`, `unresolved` and `invalid` results.
- Unknown/missing answers, provider errors and mock responses are never labelled
  available. Confirmed results show check time, premium status and price when
  supplied; availability is not a reservation. Cached provider checks retain their
  original timestamps and expire after five minutes. Inventory scans are **never**
  cached. Retry-After, account and per-user request limits are enforced per process.
- Credentials never follow redirects. No registration, purchase or payment operation is implemented.
  Forwarding and DMARC creation use the separate current-batch setup workflow below. The Porkbun link opens checkout/search for the
  user to review separately.

Official contract consulted: https://porkbun.com/api/json/v3/spec.

### Current-batch website forwarding and DMARC

After saving selected recommendations as Planned, the server records that exact new
batch. Register the domains on Porkbun, enable API access, then use **Set up this
new batch**. Enter the destination URL, select the domains, preview and apply. The
browser submits one domain at a time and displays progress and individual receipts.
There is no registrar-wide bulk action, background purchase or write to older/live
portfolio domains. The feature requires a dashboard administrator because the
registrar credential is shared. Earlier batches and manually imported inventories
cannot be substituted for the current saved batch.

- URL forwarding defaults to HTTP 301, preserving paths, at the root and `www`.
  302/307, path handling and `www` are configurable. Wildcard forwarding is off.
- DMARC creates one TXT record named `_dmarc`, TTL **600**, with this exact pattern:
  `v=DMARC1; p=none; pct=100; rua=mailto:me@{domain}; ruf=mailto:me@{domain}; fo=1`.
- **Report mailbox name** defaults to `me`. A user may supply a known name for the
  batch or an individual domain (e.g. `richard`). Both report destinations are
  then `richard@that-domain`. Names are never guessed from an owner display name.
  The recipient mailbox must exist; DNS setup does not provision an email inbox.
- The preview verifies credentials/read/DNS scopes, account ownership, Active
  registrar status, per-domain API access and Porkbun nameserver delegation.
  It inspects existing forwarding and DMARC records. Existing differing records
  (including DMARC CNAMEs) block that domain; no record is deleted or overwritten.
- A persisted preview expires in 30 minutes. Apply rechecks ownership, delegation
  and the record fingerprint. A revision-checked database claim precedes external
  writes; concurrent changes fail before another write. A running claim blocks
  ordinary portfolio edits for up to five minutes after an interrupted request.
- Identical existing settings are skipped. Stable idempotency keys and API readback
  prevent duplicate records on retries. A timeout/partial success is reported as
  **needs review**; create a new preview to reconcile it. Completed receipts remain
  with the portfolio and are not rolled back by inventory Undo.
- Only missing `/domain/addUrlForward/{domain}` and `/dns/create/{domain}` entries
  are created. SPF, DKIM, MX, other domains and nameserver delegation are untouched
  by this code. Porkbun manages the web-forwarding records needed for its redirect
  service. Conflicting configurations must be reviewed separately.
- “Verified” means Porkbun's API readback matches; public DNS/redirect propagation
  and mailbox delivery can still take time. A `p=none` policy collects reports and
  does not enforce quarantine/rejection.

### Analysis policy and scope

Rules `2026.10.1` use the bundled `tldextract` Public Suffix List, IDNA normalization,
maximum-coverage word segmentation, synonym families, conservative spelling
similarity, weighted intent evidence, overlap tags and explicit low-confidence
review. No external LLM or API key is required. Identical input/settings produce
stable results; generation accepts a deterministic batch index for new batches.

Default thresholds: 1–6 Healthy, 7–8 Watch, 9 Near Limit, 10 Full, >10 Over Limit.
Capacity can be tightened to 1–10 alongside the warning thresholds. Existing full
or over-limit clusters remain visible. Unclassified generic names stay in a review
queue, not an invented semantic family. Suggestions check both primary assignments
and cumulative natural-family overlap across Active, Planned and accepted proposals,
even after manual cluster moves. Recommendations never reuse any existing root,
including Paused and Retired ones. New names use preserved natural component words,
letters only and `.com`. Additional Strategy, Research, Client Success and Talent
families represent distinct commercial intent. They do not split a full family.

The deterministic classifier is a transparent heuristic, not universal semantic
understanding; ambiguous names need human review. Neither naming similarity nor a
shared MX/NS provider establishes inbox-provider association, reputation isolation
or guaranteed delivery. AI classification is intentionally not connected.

The interface includes reviewable add/replace imports, manual reassignment/rename/
merge/review/restore-auto controls, status/owner/platform/notes, filters, paginated
large tranches, an optional scroll/zoom cluster map (20 nodes per tranche), one-click
copy, TXT/CSV/JSON exports and printing. Save explicitly persists working edits;
replace pauses missing Active records, and Undo restores the previous saved state.
Proposed batches are transient; selected Planned records and their proposals are
stored when saved. UI map and history bounds are deliberate, visible limits.

Optional DNS analysis checks up to twenty saved Active roots for MX, SPF, DMARC,
NS and a known DKIM selector. Five workers, two-second per-query limits, ten-minute
cache and per-user limits bound lookups. This is a separate read-only diagnostic;
no port probing or redirect crawling is performed. DNS depends on runtime egress.

## Verification

Install test dependencies separately from production:

```sh
pip install pytest==8.4.2 httpx
python -m pytest -q
node --check app/static/domain-cluster.js
```

Browser journey (requires Playwright and Chromium):

```sh
# Terminal 1: local synthetic data only, no production credentials or writes.
uvicorn --app-dir tests domain_cluster_preview:app --host 127.0.0.1 --port 8765 --lifespan off
# Terminal 2:
node tests/test_domain_cluster_browser.cjs
```

Set `CHROMIUM_PATH` if Chromium is elsewhere. Browser screenshots are temporary
files under `work/`. The harness is never imported by `app.main`; its fake sign-in,
portfolio storage and strategy rows only exist in the local test process.

Tests cover parsing/PSL/IDNs, semantic synonyms and generic non-matches, 13-member
warnings without splitting, cumulative and overlap capacity, strict candidate rules,
manual override persistence, private owner isolation, concurrency conflicts, CSRF,
replace/undo, source checkboxes and fresh reads, honest Porkbun availability, exact DMARC construction, selected-batch limits,
DNS permissions, conflict detection, safe retries and setup receipt persistence.
External Google service-account, live Supabase migration and live Porkbun checks
require the existing production bindings and cannot be claimed verified locally.
