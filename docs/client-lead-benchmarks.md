# Client Lead Benchmarks (methodology v3)

`/client-benchmarks` uses the existing `client_benchmarks` permission and displays
Average Monthly Leads by Target Company Size. This replaces the earlier
lead-producing-month calculation on the live endpoint. Legacy upload endpoints
remain backward compatible and are not used by this page.

## Sources and access

Lead Database: `1mhJ0lytLPmfkhMstOy2aF6FMjZPoWalxIigfK18ZqsY`.

- Form Responses, gid 1768027115: sole source of individual leads.
- Monthly leads, gid 830915082: reconciliation, never additional lead rows.
- Lead Data Averages Backend, gid 786537287: estimated X:AE targeting weights.
- Lead Data Averages, gid 65292198: source identification only; values are not reused.
- Optional Benchmark campaign months tab: explicit dated activity and allocations.
- ERI performance tracker: reference-only access check. ERI data never enters the
  benchmark; automatic cross-tracker reconciliation is currently unavailable
  because the tracker schema has not been verified.

Tab IDs are preferred; an exact tab-title fallback is reported in methodology.
Missing required tabs fail the sync rather than silently substituting data.

Reuse `GOOGLE_SHEETS_SA_JSON` (base64 JSON) without changing its value. The reader
creates an independent read-only session, shares the existing Sheets read-rate
limiter, and does not write to Google or alter other integrations. The existing
service account needs Viewer access to Lead Database and optionally the ERI
tracker. Do not publish sheets or upload credentials into Git/chat.

## Live synchronisation

A dedicated daemon worker syncs 10 seconds after startup, then every 30 minutes
when credentials are configured. The first page request can initiate a sync.
Concurrent requests share a lock and cache; a Refresh click does not bypass the
30-minute interval. An open browser checks for new snapshots every five minutes.
This is scheduled polling, not real-time streaming.

A successful calculation is atomically saved with mode 0600 to
`CLIENT_BENCHMARK_CACHE_PATH`, default `/tmp/unstuck-client-benchmarks-v3.json`.
Memory and disk snapshots survive temporary sync failures, visibly marked stale.
Numbers older than seven days are withheld from the headline. For retention
across Render redeploys, set this optional variable to a path on an **existing
persistent disk**. The default /tmp cache is ephemeral across redeploys. Do not
put this audit-containing file under static files or inside a public repository.
There are no schema migrations or new application dependencies.

A source/parser failure never overwrites the successful snapshot. A first-sync
failure shows an error instead of fabricated results. Read-only API responses
are no-store. Freshness and calculation timestamps are displayed separately.

## Definition and arithmetic

Qualifying categories are Lead and Interested. This intentionally differs from
the legacy upload parser. Exclude ERI and Unstuck aliases, explicit test/demo/sample
client labels, invalid dates and future records. Deduplicate exact rows and
repeated (client, valid email) identities, retaining the earliest qualifying
observation. Without a stable email identifier only exact deduplication is safe;
counts of these rows are reported as a quality limitation. Company names alone
are not unique person identifiers. Unknown company sizes remain a reported
quality exclusion and do not acquire guessed segment assignments.

Reporting window: six completed calendar months for all three display segments.
The incomplete current month never enters numerator or denominator. Sparse
samples stay marked insufficient rather than switching windows. Data before
January 2025 is not included.

Each client-month is the client's complete campaign portfolio. An allocation
vector represents all eight source headcount bands; normalise its entire sum
before grouping. Prospect-count vectors must contain deduplicated prospects
across campaigns. Multiple campaign schedules must be consolidated into one
client/month row, never independently counted as full client-months.

Segments: Small businesses 1–50 (bands 1–10 + 11–50), Mid-market 51–500
(51–200 + 201–500), Enterprise 501+ (501–1,000 + 1,001–5,000 +
5,001–10,000 + 10,001+). Exactly 500 belongs only to Mid-market. The six-band
research comparison still reports 501–1,000 and 1,001+ separately.

For each segment, pool qualifying leads from eligible client-months and divide
by the sum of their allocation fractions. Zero-lead exposure counts. Leads in
bands with no assigned exposure are flagged and omitted from the numerator.
Never infer an allocation from lead distribution, sum sub-band rates, average
client rates equally, or fit the result to preliminary sales estimates.

## Maintainable monthly controls

Administrators can download headers at
`/api/client-benchmarks/control-template` and create a tab called
**Benchmark campaign months** in Lead Database. Header order:

Client | Month | Activity status | Allocation source | 1–10 | 11–50 | 51–200 |
201–500 | 501–1,000 | 1,001–5,000 | 5,001–10,000 | 10,001+ | Evidence

- Client: exact normalised lead-sheet name. No fuzzy company-family joining.
- Month: YYYY-MM; one row per client/month. All full-month zeros must be included.
- Activity status: full, partial, inactive or unknown. Partial months are excluded
  consistently, not rounded up. Evidence should refer to launch, pause and
  termination records or a reviewed campaign log.
- Allocation source: outreach (preferred actual unique-prospect counts), approved
  (dated approved weights), estimated (documented estimates), or unknown.
- Eight band values: finite nonnegative numbers with a positive total. All eight
  are normalised together. Uniform percent strings are accepted; mixed percent
  strings and numeric fractions are rejected as ambiguous.
- Evidence: required source/reference, not just an undocumented guess.

A reviewed row overrides fallback inference for that month, including explicit
partial/inactive/unknown rows. Dates and changing weights are represented by
separate monthly rows. Duplicate client/months and malformed controls fail the
sync, retaining the last successful snapshot. The approver should select actual
outreach over approved allocations over estimates when composing a row.

Without a dated row, use backend X:AE weights only when their eight headers are
verified. Those weights are estimates, not proof of historical targeting.
Fallback activity is limited to complete months bounded by the first and last
qualifying lead observations; first/last partial observation months are excluded
and intervening quiet months count. This can include unrecorded pauses, so it is
explicitly labelled an observed-span estimate, not verified live activity.
Unknown activity or allocation is excluded and counted, never invented. No
ongoing Gmail access is required. Existing lifetime invoice counts cannot supply
month-level launch/termination dates and are not joined automatically.

## Confidence, reconciliation and audit

At least three clients and six equivalent months are required for a displayed
figure. Sparse segments say Insufficient data. Estimated activity/allocations
are labelled Estimated inputs. One client supplying over half the leads triggers
a dominance note, not deletion of that client's results. Leave-one-client-out
rates are available in Methodology.

Display policy uses pooled rates, never a forced target. Small businesses and
Mid-market display the adjacent whole numbers bracketing the pooled average
(e.g. 4.19 → 4–5); an exact whole number is shown once. Enterprise displays the
nearest approximate whole number (e.g. 2.56 → ~3). Positive rates below one show
<1. These are rounded estimates, not observed outcome ranges or confidence
intervals. Compatible cached snapshots are reformatted on read, preserving
last-good data while preventing an obsolete display string from persisting.

The branded main page shows only the three company-size tiers, six-month period,
estimated-input/sample qualifications, freshness and a short indicative-results
note. The in-depth Methodology section is removed. Detailed calculations,
reconciliation and source warnings remain accessible through the protected audit
API for internal review.

The public API omits client identities, evidence and per-client reconciliation.
The admin-only `/api/client-benchmarks/audit` returns each client/month's activity,
allocation, included/recorded leads and exclusion reasons, plus Monthly leads
comparison rows. Duplicate Monthly leads totals are omitted rather than summed;
report definition differences can cause discrepancies. The internal audit includes sample sizes, verified/estimated exposure, missing
data, sync status, and the six preliminary research comparisons. Preliminary numbers never enter calculations.

## Validation and limits

Tests cover exclusions, Lead + Interested deduplication, all size boundaries,
zero-lead exposure, pooled weights, changing month controls, partial months,
six-month windows, combined enterprise exposure, thresholds, dominance, source IDs, source changes,
allocation updates, stale disk recovery, privacy and admin access. Run
`.venv/bin/python -m pytest -q` and
`node --check app/static/client-benchmarks.js`.

Offline workbook reconciliation and browser checks were performed separately;
private workbooks and audit exports stay outside Git. Live Google totals and ERI
reconciliation cannot be verified from a workspace without the runtime Google
credential. Source identity/freshness and any live access errors are exposed on
the deployed page. Snapshot validation is not a claim of live verification.
