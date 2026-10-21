# Client CRO Insights

The CRO Insights tab sits between Lead quality and Reports. It is separate from agency CRO Insights.

## Setup and publishing

Apply `migrations/pulse_growth_plans.sql` to the dashboard database before deploying. The app requires its existing Supabase service-role key. The new table has RLS enabled, no anonymous/authenticated grants, and server-only access scoped to agency and client. No existing report rows are migrated.

Link the client's Performance Tracker in Lead quality. Set a tracking start date matching the entered conversions, then refresh generated leads. Only dated tracker rows in that window count. The planner's window is independent of other dashboard date filters.

An account manager (reviewer role) or admin records at least one actual conversion and confirms both activation questions. Activation enables inclusion; publishing or re-publishing a report adds its planner flag. Unchecking inclusion or deactivating removes portal access, including older reports. Existing reports without the flag stay unchanged. Staff can save a draft before activation.

## Inputs and forecasts

Business figures, annual churn, individual conversion dates/values, 12/24-month revenue goals, and monthly lead and annual sales targets are editable in place. Changing the sales target calculates the required whole number of leads; editing leads recalculates expected sales. The chosen target basis is saved: annual sales stay fixed when editing sales, while lead-driven plans derive annual sales from the monthly lead target. Annual sales can be entered and saved before lead or conversion data is available. Save planner persists edits; other open editors must reload after a version conflict.

Only conversions within the saved lead snapshot window are used. Later conversions are identified and included after an AM refreshes leads. The server ignores submitted lead counts. A failed refresh never substitutes invented leads.

Observed rate = conversions / generated leads. Leads per sale = generated leads / conversions. Below 3%, only leads per sale is displayed. At 3% and above, the percentage also appears. Calculations always use the actual rate.

Monthly retention = `(1 - annual_churn / 100) ** (1 / 12)`. Forecasts use expected average sales, allowing fractional averages without promising fractional actual sales. Monthly mode retains the prior month's revenue then adds new sales at their average monthly value. Project mode counts new sales once; current annual baseline is spread over months and reduced by retention. The year-two project result sums months 13–24. Current revenue already includes historical conversions, so those are never added again. Suggested lead targets use this same model and are capped at one million leads/month.

Clients with a valid, unrevoked report link can edit the included live planner. Campaign report snapshots remain immutable. Clients cannot activate, change the measurement window, change inclusion or supply lead volumes. Both staff and client writes require a signed, route- and user-bound CSRF token and current record version.

## Verification

- `.venv/bin/python -m pytest tests/test_pulse_growth.py tests/test_outbound_pulse.py tests/test_pulse_tracker.py -q`
- `node tests/test_pulse_growth_math.cjs`
- `node tests/test_pulse_growth_browser.cjs` (Playwright and Chromium required; `CHROMIUM_PATH` overrides `/usr/bin/chromium`)

The CSRF token does not expire during editing and is independent of session-cookie updates from concurrent requests. While the page is visible and recently edited, a five-minute keepalive renews the existing authenticated session. Saves refresh the token without overwriting inputs or record versions. Real sign-outs or expired access links still require authentication; failed saves preserve the form.
