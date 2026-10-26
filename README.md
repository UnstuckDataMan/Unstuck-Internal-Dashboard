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

## Gmail Setup Tool / account setup tracker

Open **Operations → Gmail Setup Tool** (`/gmail-set`). The main tool tracks the shared
company setup journey: **Not started → Started → Purchase made → Setup complete**.
Create a client first, click its card, then paste up to 500 domains or emails.
Client details also store the sender's full name, client website for forwarding,
and account manager's email. These can be edited later. A website entered without
a scheme is saved with `https://`; saved websites prefill DNS setup confirmations.
An optional sender name (for example, `ashley`) populates `ashley@domain.com`
for bare domains; explicit pasted email addresses are preserved. This mailbox
name is separate from the sender's full name.
One account manager applies to the client. Optionally choose a default setup
person, then change setup person, sending email, device and stage inline for
each domain. Bare domains can be planned without an email or setup person;
completion requires both. Invalid paste lines are identified and duplicates
are skipped. The
client list shows counts instead of every domain; clients with no setups left to
do appear in a separate **Completed** section. Unfinished domains appear above
completed ones in each client’s domain list. Each domain has a stage dropdown,
and the add/edit form also lets you choose its stage. Existing flat records are
grouped automatically without changing their sheet rows or purchase history. It starts
with a prominent instruction for SDRs and Scale Clients to use the GPT domain
clustering plugin in Notion. This step is optional for normal clients; clustering
is no longer the main workflow.

Gmail Setup Tool checks for a new release every 30 seconds and when a tab becomes visible.
Idle tabs reload the current release; open forms, unsaved edits and active saves
delay that reload. The Refresh button checks the release as well as shared data.
Render's commit ID identifies releases; local previews use the asset hash.
If an open tab has an older CSRF token after sign-in, only the server's explicit
pre-mutation session rejection triggers a current-token fetch and one retry.
Ordinary permission failures and uncertain writes are never replayed this way.
The Operations card also uses an asset content hash and no-store responses,
refreshes every 30 seconds and on returning to the page, and clears a stale
figure if the latest summary cannot be read. Browser Back restoration rechecks
the tracker version without reloading over drafts.

- The tracker seeds its member list from active dashboard users in the signed-in
  workspace. Members can be added, edited or removed within this tool without
  changing dashboard login access. Open work must be reassigned before removal.
- Record account manager, setup owner, client/Biz Dev section, root domain,
  sending email. Completion order is automatic per client: the first completion
  gets 1, the next 2, and so on. The number stays fixed through reopening and
  deletion; manual order inputs are ignored.
  Device allocation is optional, even with zero inventory;
  accounts without one show **Device not added yet** and an inline device selector.
  A missing device can be recorded after completion.
  Purchase records snapshot the actual
  device and asset label; removing/transferring equipment preserves that history.
- Clients can be renamed and deleted from their detail page; every domain has
  a **Remove** action, including completed work. Its warning directs deleted Gmail
  subscriptions to **Gmail deleted** in the Stage dropdown instead. Cancel is
  focused first; **Remove from tracker** must be confirmed explicitly. Removing
  a domain hides its tracker entry and preserves sheet rows, completion counters and physical device
  history. Historical sheet-entry references remain available.
- Each domain has a **Notes** box and **Save notes** button. Notes remain saved
  when a domain is completed, burned or reopened.
- **Add client** accepts an optional password and **Card last 4 digits**. The
  password is encrypted in the same atomic save that creates the client. The
  card field stores exactly four digits, including leading zeros; use **Edit
  client** to change or clear it. The last four digits appear in the client's
  key information at the top. Existing clients can leave both fields unset.
- Each client's details include an optional password input with **Save**,
  **Show / Hide** and **Remove**. Passwords are masked by default; saved values
  are encrypted server-side and omitted from shared board responses. A current
  signed-in user with Gmail Setup Tool access must explicitly reveal a password.
  Encryption uses `GMAIL_SETUP_PASSWORD_KEY`, falling back to the existing
  `SESSION_SECRET`; keep that key stable or existing saved passwords cannot be
  recovered. Password edits are kept in memory while saving other tracker work;
  they are never stored in browser storage.
- The optional **Master domain** dropdown groups domains within a single client.
  Choose **Make master domain** for up to five masters per client, then select
  **Under [domain]** on related domains. A crown identifies each master; its
  solid left-hand colour strip is shared by linked domains. Stage backgrounds
  remain visible. Choose **No master domain** to undo a selection. Demoting,
  reassigning or deleting a master ungroups its linked domains; other masters
  retain their colours. Grouping is shared tracking metadata and leaves setup
  progress, device capacity, DNS and forwarding unchanged. New sheet backup rows
  label selected masters and secondary domains beside the sending email.
- Add actual inventory in batches of 0–10 phones and 0–10 PCs per person, with
  labels Phone 1–10 and PC 1–10 plus stable physical asset IDs. The unallocated pool
  supports stock awaiting assignment. No imaginary devices are seeded.
  When adding inventory, choose Unused, Used once or Used twice (or a block
  reason). Previous uses count towards the device's setup limit immediately.
- Setup stages include **Failed at QR** and **Failed at payment**. Selecting either
  blocks the linked device in red with the matching reason.
  Retry with a replacement device or leave it unallocated; the failed
  asset stays in account history.
- Domain stages include **Started**, which tracks active work without consuming a
  device use. Started and Purchase made cards have a
  subtle amber tint; Setup complete is green; Burned and QR/payment failures are
  red. Button styles remain unchanged. Started stays in setup-demand calculations
  and has its own overview count and stage filter.
- **Gmail deleted** requires a **Yes, proceed** confirmation that domain auto-renew
  has been stopped. This records the team's manual check; it does not change
  Porkbun auto-renew or delete a Google account. Deleted domains appear below
  Burned, retain notes/master labels/history and remain replacement work against
  the client's target. They are excluded from unfinished lists, bulk completion
  and DNS/DMARC setup. Reopen with **Not started** if needed. Completed domains
  are listed under **Live domains (complete)**.
- Unused, Used and Blocked states are separate from allocation. **Block QR code**
  and **Block Payment** are explicit block reasons. All blocked equipment is red
  and excluded from purchases; previously used devices never count as new stock.
- Device cards and their detail dialog show linked domains, client names,
  stages and reset history. Domain links jump to the correct client and row;
  deleted records remain visible as history.
- Setup limits are shared workspace settings: defaults are **2 per phone** and
  **1 per laptop/PC**. Change them in **Setup limits & order calculation**, along
  with the type of new equipment to order (phones by default). Limits accept 1–100.
  Green means zero uses; amber means used with capacity left; red means at/above
  the saved limit or blocked. Raising/lowering limits updates every card, domain
  picker, allocation check and consumption check without erasing physical history.
  Manually blocked equipment remains blocked even if limits are raised.
- Each client has a saved **Domain target** (0–5000): the number of completed
  domains needed for that client. Set it manually before buying domains, or let
  pasted domains establish it. Partial imports retain a larger manual target;
  adding domains grows the target only when the active, nonburned domain count
  exceeds it. Duplicate-only and failed imports do not increase the target.
  Remaining work is **target − completed domains**. For a target of 10, three
  complete and two burned leaves **7 to do**, including **2 replacements**.
  Burned or deleted domains retain their place in the target until replaced or
  the target is explicitly reduced. Replacement additions fill those places
  without inflating the target; targets cannot be reduced below the number of
  current nonburned domains. Client counts, Completed grouping and device orders
  all use this same remaining-work calculation. Missing domains create demand,
  not synthetic email rows or DNS targets.
  Older clients receive a target once from the greater of their old planned
  count and recorded domain history, including burned/deleted rows. This target
  can then be edited in the client details.
- Both Gmail Setup Tool and Operations compare **setup slots**: required setups are
  **all clients' remaining work + 10 contingency setups**, not ten physical devices.
  Remaining capacity sums max(limit − recorded uses, 0) on active, unblocked
  phones and PCs. Purchase-stage jobs already paid for on a recorded device also
  count as covered; jobs reset to Not started require fresh capacity. Completed
  domains reduce demand; burned/deleted domains still need replacements against
  the client's target. New orders are
  ceil(max(0, required − available setup slots) / setups per new device).
  The card shows phone/PC capacity separately, demand, setup shortfall and whole
  devices to order. For 19 to do + 10 contingency, 22 phone + 2 PC slots give
  24/29 coverage: 5 slots short, requiring **3 phones** at two setups each.
  The same saved limits control ordering, device colours and allocations.
  No external orders or notifications are sent.
- Every team member with Gmail Setup Tool access can use the domain-level
  **One-click setup** button and client-wide **Set up all domains**. This uses
  the same tool permission as Master Domains, stages, notes and client passwords;
  no administrator role is required. After the
  purchase warning, enter the client website URL and select **Proceed**. It removes
  only reviewed Porkbun forwarding MX, default Porkbun SPF and ACME challenge TXT
  records; adds/verifies `t` CNAME → `open.sleadtrack.com`, the domain's sending-email
  DMARC policy, root MX `smtp.google.com` (priority 1) and root SPF TXT
  `v=spf1 include:_spf.google.com ~all`, all with TTL 600. Existing Google mail
  TTL/priority is corrected without duplicates; custom mail policies require review.
  It removes the reviewed root forwarding before adding the client's URL (302,
  wildcard enabled, no appended path); unrelated subdomain forwards are retained.
  The URL is remembered for the client's next domain.
  **Cancel** makes no external changes. Conflicting mail, tracking or DMARC records
  require review; forwarding changed since preview requires fresh confirmation.
  Exact deletion snapshots, expiring confirmations, shared write claims,
  readback and idempotent retries protect unrelated records and domains.
- **Set up all domains** in the client header runs the
  same full setup for every actual, active domain in that client. One purchase
  confirmation and website URL cover the batch. Domains run sequentially, with a
  progress bar and saved Ready / Needs attention / Waiting results. Burned and
  deleted domains are excluded. One domain's failure does not stop the others;
  a provider rate limit pauses the batch. Keep the page open, or use **Pause after
  this domain** and **Resume setup** after a refresh. Each waiting domain receives
  a fresh DNS preview; interrupted writes are reconciled before retrying.
- DMARC is included in both the individual and client-wide DNS setup actions.
  Both rua/ruf use the domain's sending email, with
  `v=DMARC1; p=none; pct=100; rua=mailto:<email>; ruf=mailto:<email>; fo=1`,
  TXT `_dmarc` and TTL 600. Missing sending emails need attention before DNS setup.
  Existing Porkbun credentials remain server-side. Each domain checks ownership,
  API access and Porkbun nameservers; missing records are created and read back.
  Matching records are verified without another write; conflicting records are
  reported without duplication or overwrite. Persistent CAS claims serialize DNS
  writes; per-domain results survive refresh and safe retries.
- Selecting a later stage records purchase/device history as needed; choosing
  **Setup complete** saves completion in Gmail Setup Tool and queues its sheet backup. **Mark all as complete**,
  beside **Set up all domains**, completes all unfinished domains for that client
  after confirmation; burned and deleted domains are excluded. Each domain needs
  a sending email and active setup person before completion.
  Select **Not started** to reopen a completed or failed domain and change its
  setup person or device. Resetting a failed attempt removes that attempt’s use and block from its
  device: it returns to green/Unused if no other use remains, or orange with
  one other use. Independent manual blocks and other failed domains still apply.
  Older detached resets are reconciled on read. Failure history, original
  completion order and existing sheet rows stay recorded. Re-completing
  the same domain keeps its completion order rather than counting another setup.
- **Google Sheets backups run in the background.** Completing a domain adds its
  sending email below the last populated entry in `1) Emails - entry` for clients
  or `1) BD Emails - entry` for Biz Dev, with **Live** checked. Existing matching
  rows are reused, avoiding duplicate entries on retries. **Gmail deleted** only
  unticks **Live** on its existing row; it never removes the entry or creates a
  missing deleted row. Domain cards show backup progress, row links and actionable
  errors with **Retry backup**. **Back up to Sheets** at client level includes
  historical completed and Gmail-deleted domains from the manual-only period.
  New rows use the actual sending email followed by **(Master)** or
  **(Secondary: masterdomain.com is Master)** when that grouping is selected.
  Ungrouped domains stay plain; existing email-cell labels are not rewritten.
- Workspace-wide conditional revisions protect shared tracker updates. Narrow
  stage, assignment and notes saves can retry after an unrelated teammate edit;
  a conflicting change to the same fields shows the latest saved values instead
  of overwriting them. Background refreshes do not replace active edits or
  responses from a newer save. A persistent 180-second write lease prevents
  overlapping DNS writes. A separate background sheet lease preserves team edits
  while sheet requests run; newer Live intentions are reconciled after an older
  in-flight request. Boot diagnostics check tracker storage and sheet readiness.
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
- Gmail Setup Tool backups require Editor access for
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
is needed. Credentials remain on Render. Use **Check connections** in the legacy
planner to verify access.

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
