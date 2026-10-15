# Average Client Data

Open Reporting & Analytics → Average Client Data. Access uses the new
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

The configured sheet/tab is in `app/utils/client_benchmarks.py`. The live reader prefers the Form Responses tab by name and parses Timestamp,
Client, Lead company size and Lead category, including Google long-form dates.
Only Lead-category records count. If that tab is absent, the configured tab
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

The page automatically reads Google Sheets and displays only four target-size
cards (Micro, SME, Medium, Enterprise), a refresh control and a short period/basis
note. Cards use `groups[].average`, never selected maximum results. The page and
home card are named Average Client Data; existing routes and permissions stay
compatible. Excel controls, examples, overall metric, detailed tables, trends and
PDF controls are removed from this page. Upload APIs remain available for existing
integrations but are not invoked by this UI.

The live lead-only view uses all available dated history through today, from the
earliest valid record to the current month, with no 36-month history cap. Future
records are excluded. The current partial month counts once and is explicitly
labelled. The live view divides each segment’s leads by distinct client/month pairs with
at least one lead in that segment. A client/month spanning multiple bands in a
group counts once. Cards say “Average leads in months with leads”. This is a
conditional average, not an average across all active delivery months. Existing
Excel/calendar calculations retain their prior denominator.
Explicit API date filters retain their existing completed-month validation. The note identifies Unstuck when included and makes
clear that active delivery months are not verified. Lead-producing months cannot
serve as a substitute for a complete activity roster. A missing segment displays
a dash, not zero. Failures hide old results; refresh retries the live source.

## Validation

Run `.venv/bin/python -m pytest -q` and
`node --check app/static/client-benchmarks.js`. Tests cover access controls,
parsing, source errors, private session isolation, read-only requests, caching,
zero months, denominator arithmetic, selected examples and Excel limits.
Browser checks cover automatic Google loading, actual averages rather than peak examples, no-data results, mobile width, source failures and recovery.
Client workbooks, private audit results and credentials are kept outside Git.
