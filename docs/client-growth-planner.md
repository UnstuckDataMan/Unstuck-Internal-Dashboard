# Reporting CRO Insights

This optional reporting section is separate from the agency's main CRO tool. It appears at the bottom of all published reports for a client once an AM activates it and includes it in reports. Existing report links continue to work.

## Simple inputs

Current monthly business revenue, average monthly revenue per customer (or average one-off project value), annual revenue churn, and optional Year 1–5 goals. Goals are annualised revenue at that point: monthly revenue × 12, not cumulative cash collected. Blank years remain null. The graph ends at the last entered year, defaulting to two years if no goals exist.

One row represents one Unstuck deal, with an optional name, month won and value. The count is automatic. Lead measurement settings are under an expandable section. AMs refresh the existing linked Performance Tracker; client users cannot replace server-owned lead volumes.

## Forecast and analysis

Leads per deal uses only conversions inside the saved lead window. Forecast sales pace is matched conversions divided by calendar months in that window (partial months prorated). The model compounds annual churn monthly and adds expected new customer revenue using the user's average value. Retainers recur; one-off projects count once. Current revenue already includes historical deals, so those deals are never added again.

A single graph shows annualised revenue at that pace and the supplied goals. Without matched lead/deal data it shows an explicitly labelled churn-only baseline, not an invented growth forecast. Selected-goal analysis calculates the monthly and total deals needed to close the gap and translates these into leads using observed leads per deal. It explicitly frames this as an illustration if Unstuck alone filled the gap. A permanent note says Unstuck should work alongside existing activity and other sources of growth. The tool never displays conversion-rate percentages.

The new plan stores schema_version=2, average_value, and five nullable goals in the existing JSON data column. Previous monthly goals are converted to annualised figures on load; previous project current revenue is converted from annual to monthly. Original data is not rewritten until saved. Activation, CSRF protection, optimistic concurrency, client scoping, and session renewal remain intact. Main CRO Insights code is unchanged.

## Verification

- `.venv/bin/python -m pytest tests/test_pulse_growth.py tests/test_outbound_pulse.py tests/test_pulse_tracker.py -q`
- `node tests/test_pulse_growth_math.cjs`
- `node tests/test_pulse_growth_browser.cjs` (Playwright and Chromium required)
