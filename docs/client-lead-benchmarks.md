# Client Lead Benchmarks

Open Reporting & Analytics → Client Lead Benchmarks. Access uses the new
`client_benchmarks` tool permission. Existing tool permissions are unchanged.

## Isolation

All calculations, data reads, assets and endpoints are specific to this tool.
Shared integration changes only register its router, permission and home card.
There are no database migrations, scheduled jobs, new dependencies or changes to
other tools' calculations, Google configuration or deployment settings.

The Google reader uses a separate session with the Sheets read-only OAuth scope.
It reads the existing `GOOGLE_SHEETS_SA_JSON` base64 service-account binding without
modifying it or the shared Google module's client/session caches. Leave that
Render setting unchanged. Share the lead spreadsheet with its configured account
as Viewer. Reads are cached for 60 seconds and occur only when requested, not in
background polling. A connection failure returns a tool-specific error without
substituting invented results. Google access must be verified after deployment;
it cannot be verified locally without the runtime credential.

The configured sheet/tab is in `app/utils/client_benchmarks.py`. The live tab
expects Client, Date, Lead, Job role, Company size, Company Industry, Channel,
Country. CSV export is a fallback only for sources that are already public.
Private client sheets should remain private. Required network destinations are
docs.google.com, sheets.googleapis.com and oauth2.googleapis.com.

## Excel and measurement

Uploads are processed in memory for the request and are not saved. Limits are
20 MB compressed, 200 MB expanded and 100,000 records. Formulas are read as cached
values, never executed. Prefer the Form Responses raw tab over Responses (clean)
when both exist; raw data requires Timestamp, Client, Lead company size and Lead
category. Only Lead-category records count. Exact duplicate rows are removed;
quality exclusions are shown. Unknown sizes cannot contribute to size metrics.

Lead-only input computes observed leads per account per selected calendar month.
All selected months count for accounts represented by leads in the segment.
It does not infer active campaigns from lead-producing months.

Optional `Delivery months` tab headers: Client | Lead sheet client | Delivery months.
The alias maps billing names to lead names; month counts must be positive whole
numbers. Only matched accounts contribute. Lifetime counts are used with all
available lead history, so date filters are disabled and coverage differences are
explicitly disclosed. Each matched client's full tenure counts once per observed
segment. Targeting duration is not allocated by lead share. Live lead-only reads
do not automatically fetch invoice counts or preserve an Excel delivery roster.

Explicit campaign activity takes precedence: Campaign ID | Client | Month |
Company size | Active | Leads. Include active zero-lead months and allocate every
lead to one size band. Month is YYYY-MM and campaign ownership must be unique.
The downloadable template contains instructions and no fabricated records.

## Presentation

The headline shows the overall metric with sample size, denominator, median and
coverage limitations. Size-specific averages are in an expandable detail table;
they use different client samples and are not additive.

Prominent segment cards show selected highest recorded single account/calendar-
month counts, not averages or expected volumes. Group bands are combined first;
ties choose the most recent month then stable account order. Each card discloses
the date, candidate sample and one-account/one-month basis. Unstuck's own
prospecting is explicitly identified and other account names are anonymised.
Missing results show a dash rather than a fabricated example.

Delivery-month examples use the last 12 completed months and matched roster
accounts plus Unstuck, separately from the headline's all-history scope. Other
modes use the selected reporting period. Examples may come from different months
and must not be added together. Own prospecting without a supplied delivery count
can appear as an example without entering the billing-denominator average.

## Validation

Run `.venv/bin/python -m pytest -q` and
`node --check app/static/client-benchmarks.js`. Tests cover access controls,
parsing, source errors, private session isolation, read-only requests, caching,
zero months, denominator arithmetic, selected examples and Excel limits.
Browser checks cover real uploads, mobile width, PDF output and source switching.
Client workbooks, private audit results and credentials are kept outside Git.
