"""The Performance Tracker parser.

No HTTP and no fakes: every rule in tracker.py takes a list of dicts, so these
tests are the specification for what a tracker sheet means.

The fixture below is the live Dovetail sheet in miniature — same headers, same
shapes, same awkward bits: a client-named date column, a lead with a blank Lead
cell, "-" for ungraded, a blank channel, and a few thousand rows of filled-down
formula underneath. Those awkward bits are the point; the clean rows are easy.
"""
from __future__ import annotations

from datetime import date

import pytest

from app.utils.pulse import tracker


HEADERS = ["Date/Time", "Lead", "Channel", "Headcount", "Job role",
           "Industry", "Location", "ICP rating", "Category", "Sale stage",
           "Additional notes"]


# The live sheets open with a banner naming the client and put the column
# names on row 2, so every fixture here is a grid with that shape.
BANNER = ["Dovetail"] + [""] * (len(HEADERS) - 1)


def _row(day, lead="acme.com", channel="Smartlead", headcount="11-50",
         role="Head of Growth", industry="Marketing", location="United Kingdom",
         icp="8", category="Lead", stage="", notes=""):
    return [day, lead, channel, headcount, role, industry,
            location, icp, category, stage, notes]


def _grid(rows, headers=None, banner=True):
    out = [list(BANNER)] if banner else []
    out.append(list(headers if headers is not None else HEADERS))
    return out + [list(r) for r in rows]


def _sheet():
    """Eight real leads, then the filled-down tail the live sheet carries."""
    rows = [
        _row("9/23/2025", "boxtoboxfilms.com", "Smartlead", "51-200",
             "Head of People & Operations UK", "Media Production", icp="8"),
        # A real lead with no Lead cell. Filtering on Lead loses it, and the
        # client's own total then disagrees with ours by one.
        _row("10/7/2025", lead="", channel="Smartlead", headcount="11-50",
             role="Managing Director", industry="Advertising Services", icp="-"),
        _row("10/21/2025", "northpr.co.uk", "LinkedIn", "1-10",
             "Founding Partner", "PR", icp="10"),
        _row("11/4/2025", "globex.io", "LinkedIn", "201-500",
             "VP of Engineering", "SaaS", icp="-"),
        _row("1/5/2026", "initech.com", "Manual", "501-1000",
             "Chief Marketing Officer", "Marketing", icp="9"),
        _row("2/18/2026", "hooli.com", "", "11-50",
             "Growth Manager", "Marketing", icp="4"),
        _row("3/2/2026", "umbrella.org", "Smartlead", "",
             "Procurement Lead", "Charity & non-profit", icp="-",
             category="Interested"),
        _row("3/30/2026", "stark.com", "Smartlead", "51-200",
             "Head of Growth", "Media Production", icp="7"),
    ]
    # The tail: ~2,500 of these sit under the real rows in the live sheet,
    # carrying nothing but a formula result.
    rows += [_row("", lead="", channel="", headcount="", role="", industry="",
                  location="", icp="-", category="") for _ in range(40)]
    return _grid(rows)


@pytest.fixture
def parsed():
    rows, meta = tracker.parse_rows(_sheet())
    return rows, meta


# ── What counts as a lead ────────────────────────────────────────────────────

def test_a_row_is_a_lead_when_its_date_parses(parsed):
    """Not when the Lead cell is filled, and not when any cell is: the live
    sheet is ~2,500 rows of filled-down formula under a few dozen real ones,
    and one of the real ones has no Lead cell at all."""
    rows, meta = parsed
    assert len(rows) == 8
    assert meta["skipped"] == 40
    assert any(r["lead"] == "" for r in rows), "the blank-Lead row is a real lead"


def test_the_undated_tail_is_not_dated_today(parsed):
    rows, _ = parsed
    assert all(r["day"] <= "2026-03-30" for r in rows)


# ── The client-named date column ─────────────────────────────────────────────

@pytest.mark.parametrize("header", [
    "Dovetail Date/Time", "Acme Date/Time", "Globex Corp Date/Time",
    "Date/Time", "Date", "  dovetail   date/time  ",
])
def test_the_date_column_is_found_whatever_the_client_is_called(header):
    columns, warnings = tracker.resolve_columns([header] + HEADERS[1:])
    assert columns["date"] == header
    assert not warnings


def test_a_column_merely_mentioning_a_date_is_not_the_date_column():
    """"Date added" is a different thing. Reporting no date column beats
    scoping every client report by the wrong one."""
    headers = ["Notes", "Date added", "Lead", "ICP rating"]
    columns, warnings = tracker.resolve_columns(headers)
    assert "date" not in columns
    assert any("No column is named like a date" in w or "No date column" in w
               for w in warnings)


def test_a_known_column_is_never_taken_as_the_date_column():
    """Column 0 is the fallback, but only if nothing else claims it."""
    columns, _ = tracker.resolve_columns(["Lead", "Industry", "ICP rating"])
    assert "date" not in columns


def test_the_first_column_is_the_fallback_and_says_so():
    columns, warnings = tracker.resolve_columns(["When", "Lead", "ICP rating"])
    assert columns["date"] == "When"
    assert any("being\nused as one" in w.replace("“", "").replace("”", "")
               or "used as one" in w for w in warnings)


def test_headers_match_whatever_the_spacing_and_case():
    columns, _ = tracker.resolve_columns(
        ["Date", "  ICP  Rating ", "Job  Role", "COMPANY SIZE"])
    assert columns["icp"] == "  ICP  Rating "
    assert columns["role"] == "Job  Role"
    assert columns["headcount"] == "COMPANY SIZE"


def test_a_missing_column_disables_its_breakdown_rather_than_guessing():
    drop = HEADERS.index("Industry")
    headers = [h for i, h in enumerate(HEADERS) if i != drop]
    rows, meta = tracker.parse_rows(_grid(
        [[c for i, c in enumerate(r) if i != drop] for r in _sheet()[2:]],
        headers))
    assert all(r["industry"] == tracker.UNSPECIFIED for r in rows)
    assert len(rows) == 8, "losing one column must not lose the leads"


def test_without_a_date_column_nothing_is_returned():
    """A period figure built from undated rows would be fiction."""
    rows, meta = tracker.parse_rows(
        _grid([["a.com", "8"]], ["Lead", "ICP rating"]))
    assert rows == []
    assert meta["has_dates"] is False


# ── Dates ────────────────────────────────────────────────────────────────────

def test_us_dates_are_read_month_first(parsed):
    rows, meta = parsed
    assert meta["date_order"] == "mdy"
    assert rows[-1]["day"] == "2025-09-23"          # 9/23/2025


def test_a_day_first_column_is_detected_across_the_whole_sheet():
    """"23/9/2025" can only be day-first, and that settles "1/5/2026" too.
    Deciding per row would read one column as a mixture of both."""
    parsed, meta = tracker.parse_rows(
        _grid([_row("23/9/2025"), _row("1/5/2026")]))
    assert meta["date_order"] == "dmy"
    assert sorted(r["day"] for r in parsed) == ["2025-09-23", "2026-05-01"]


def test_an_entirely_ambiguous_column_falls_back_to_month_first():
    parsed, meta = tracker.parse_rows(_grid([_row("1/5/2026")], HEADERS))
    assert meta["date_order"] == "mdy"
    assert parsed[0]["day"] == "2026-01-05"


def test_an_iso_date_cell_is_understood():
    parsed, _ = tracker.parse_rows(_grid([_row("2026-04-09")], HEADERS))
    assert parsed[0]["day"] == "2026-04-09"


def test_an_impossible_date_is_not_a_lead():
    parsed, meta = tracker.parse_rows(_grid([_row("13/45/2026")], HEADERS))
    assert parsed == []
    assert meta["skipped"] == 1


# ── ICP ──────────────────────────────────────────────────────────────────────

def test_a_dash_is_ungraded_not_zero():
    for blank in ("-", "–", "—", "", "n/a", "TBC", "?"):
        assert tracker.parse_icp(blank) is None, blank


def test_the_average_ignores_ungraded_leads(parsed):
    """The sheet's own average does, which is the only way the two agree."""
    rows, _ = parsed
    summary = tracker.summarise(rows)
    assert summary["total"] == 8
    assert summary["rated"] == 5                     # 8, 10, 9, 4, 7
    assert summary["unrated"] == 3
    assert summary["average"] == 7.6                 # 38 / 5, not 38 / 8


def test_an_ungraded_lead_never_drags_the_average_down(parsed):
    rows, _ = parsed
    graded_only = [r for r in rows if r["icp"] is not None]
    assert tracker.summarise(rows)["average"] == \
        tracker.summarise(graded_only)["average"]


def test_the_average_is_absent_rather_than_zero_when_nothing_is_graded():
    """0.00 average ICP is the single worst thing this feature could print."""
    rows, _ = tracker.parse_rows(_grid([_row("1/5/2026", icp="-")], HEADERS))
    summary = tracker.summarise(rows)
    assert summary["average"] is None
    assert summary["median"] is None
    assert summary["total"] == 1


def test_an_out_of_range_grade_is_ungraded():
    assert tracker.parse_icp("12") is None
    assert tracker.parse_icp("0") is None
    assert tracker.parse_icp("abc") is None
    assert tracker.parse_icp("8.4") == 8


def test_ungraded_is_a_visible_band_and_the_bands_total_everything(parsed):
    rows, _ = parsed
    bands = tracker.summarise(rows)["bands"]
    unrated = next(b for b in bands if b["key"] == "unrated")
    assert unrated["count"] == 3
    assert sum(b["count"] for b in bands) == len(rows)
    assert round(sum(b["pct"] for b in bands)) == 100


# ── Headcount ────────────────────────────────────────────────────────────────

def test_headcount_sorts_by_lower_bound_not_lexically():
    """Sorted as text, 1001-5000 lands second and the chart reads as nonsense."""
    bands = ["501-1000", "11-50", "10001+", "1-10", "1001-5000", "51-200"]
    assert sorted(bands, key=tracker.headcount_order) == [
        "1-10", "11-50", "51-200", "501-1000", "1001-5000", "10001+"]
    assert sorted(bands) != sorted(bands, key=tracker.headcount_order)


def test_a_blank_headcount_is_unspecified_and_still_counted(parsed):
    rows, _ = parsed
    headcount = tracker.summarise(rows)["headcount"]
    assert sum(e["count"] for e in headcount) == len(rows)
    assert any(e["label"] == tracker.UNSPECIFIED for e in headcount)


def test_a_band_google_turned_into_a_date_is_surfaced_not_dropped():
    """Typing 1-10 into Sheets can produce 1/10/2025. The fix is in the sheet,
    so it has to be visible rather than quietly binned."""
    rows, meta = tracker.parse_rows(_grid([_row("1/5/2026", headcount="1/10/2025")],
                                    HEADERS))
    assert rows[0]["headcount"] == tracker.UNRECOGNISED
    assert meta["unrecognised_bands"] == 1


def test_unspecified_and_unrecognised_bands_sort_last():
    order = tracker.headcount_order
    assert order("10001+") < order(tracker.UNSPECIFIED) < order(tracker.UNRECOGNISED)


# ── Seniority ────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("title,expected", [
    ("Managing Director", "director"),          # not "manager"
    ("Founding Partner", "founder"),            # not "partner"
    ("VP of Engineering", "vp"),                # not "ic" via engineer
    ("Chief Marketing Officer", "c_level"),
    ("Head of People & Operations", "head"),
    ("Growth Manager", "manager"),
    ("Global Client Partner", "partner"),
    ("Marketing Executive", "ic"),
    ("Chief Bottle Washer", "c_level"),
    ("Wrangler of Vibes", "unclassified"),
    ("", "unclassified"),
])
def test_seniority_buckets_in_order(title, expected):
    assert tracker.seniority(title) == expected


def test_an_unknown_title_is_not_quietly_called_an_individual_contributor():
    """A wrong IC count on a client's report is worse than an honest gap."""
    assert tracker.seniority("Wrangler of Vibes") == "unclassified"


def test_trailing_geography_is_stripped_so_one_role_is_one_row():
    assert tracker.clean_role("Head of People & Operations UK") == \
        "Head of People & Operations"
    assert tracker.clean_role("Head of Growth (EMEA)") == "Head of Growth"
    rows, _ = tracker.parse_rows(_grid([_row("1/5/2026", role="Head of Growth UK"),
         _row("1/6/2026", role="Head of Growth")], HEADERS))
    roles = tracker.summarise(rows)["roles"]
    assert roles == [{"label": "Head of Growth", "count": 2}]


def test_seniority_is_ordered_most_senior_first(parsed):
    rows, _ = parsed
    keys = [e["key"] for e in tracker.summarise(rows)["seniority"]]
    assert keys == [k for k in tracker.SENIORITY_ORDER if k in keys]


# ── Channel ──────────────────────────────────────────────────────────────────

def test_sheet_channels_map_onto_the_apps_own_names():
    assert tracker.clean_channel("Smartlead")[0] == "smartlead"
    assert tracker.clean_channel("LinkedIn")[0] == "meet_alfred"
    assert tracker.clean_channel("Meet Alfred")[0] == "meet_alfred"
    assert tracker.clean_channel("Manual") == ("manual", "DNC & Merger")


def test_a_blank_channel_is_its_own_bucket_so_the_split_adds_up(parsed):
    rows, _ = parsed
    channels = tracker.summarise(rows)["channel"]
    assert sum(e["count"] for e in channels) == len(rows)
    assert any(e["label"] == tracker.NOT_RECORDED for e in channels)


# ── Top-N ────────────────────────────────────────────────────────────────────

def test_a_long_industry_tail_collapses_into_other():
    rows, _ = tracker.parse_rows(_grid([_row(f"1/{d}/2026", industry=f"Industry {d}") for d in range(1, 13)],
        HEADERS))
    industries = tracker.summarise(rows)["industry"]
    other = industries[-1]
    assert other["label"] == "Other"
    assert other["distinct"] == 4                    # 12 - top 8
    assert sum(e["count"] for e in industries) == 12


def test_a_tail_of_one_is_named_rather_than_called_other():
    """"Other: 1" hides a name for no gain."""
    rows, _ = tracker.parse_rows(_grid([_row(f"1/{d}/2026", industry=f"Industry {d}") for d in range(1, 10)],
        HEADERS))
    labels = [e["label"] for e in tracker.summarise(rows)["industry"]]
    assert "Other" not in labels
    assert len(labels) == 9


# ── Periods ──────────────────────────────────────────────────────────────────

def test_a_period_takes_only_the_leads_inside_it(parsed):
    rows, _ = parsed
    march = tracker.in_period(rows, date(2026, 3, 1), date(2026, 3, 31))
    assert sorted(r["day"] for r in march) == ["2026-03-02", "2026-03-30"]


def test_a_period_is_inclusive_at_both_ends(parsed):
    rows, _ = parsed
    assert len(tracker.in_period(rows, date(2026, 3, 2), date(2026, 3, 30))) == 2


def test_a_period_average_is_computed_over_that_period_only(parsed):
    rows, _ = parsed
    march = tracker.in_period(rows, date(2026, 3, 1), date(2026, 3, 31))
    summary = tracker.summarise(march)
    assert summary["total"] == 2
    assert summary["rated"] == 1                     # the other is "-"
    assert summary["average"] == 7.0


def test_a_period_with_nothing_graded_has_no_average(parsed):
    rows, _ = parsed
    november = tracker.in_period(rows, date(2025, 11, 1), date(2025, 11, 30))
    summary = tracker.summarise(november)
    assert summary["total"] == 1
    assert summary["average"] is None


# ── The shape handed on ──────────────────────────────────────────────────────

def test_every_value_is_json_serialisable(parsed):
    """The same rows go into a report snapshot, where a date object silently
    failed to encode and made Publish do nothing at all."""
    import json

    rows, _ = parsed
    json.dumps({"rows": rows, "summary": tracker.summarise(rows)})


def test_a_row_carries_its_own_derived_fields(parsed):
    """The browser filters these rows, so every rule has to be decided here —
    otherwise the same rule exists twice and they drift."""
    rows, _ = parsed
    row = next(r for r in rows if r["lead"] == "northpr.co.uk")
    assert row["seniority"] == "founder"
    assert row["icp_band"] == "9_10"
    assert row["headcount_order"] == 1.0
    assert row["channel_label"] == "LinkedIn"


# ── Finding the header row ───────────────────────────────────────────────────
#
# This is the bug that reached production. A real tracker opens with a banner
# naming the client -- "Dovetail" alone in A1, the rest of the row blank -- and
# puts the column names on row 2. Taking row 1 as the header gave a sheet with
# one usable column and no ICP rating, and the panel reported that as a warning
# rather than a failure, so it looked linked and was empty.
#
# It was not caught earlier because the sheet was profiled through Google's
# gviz CSV endpoint, which INFERS headers and silently concatenates a two-row
# header into one label ("Dovetail" + "Date/Time" -> "Dovetail Date/Time").
# The Sheets API returns the raw grid. These tests use raw grids.

def test_a_banner_row_above_the_headers_is_not_the_header():
    index, headers = tracker.find_header(_sheet())
    assert index == 1, "row 2 holds the column names"
    assert headers[0] == "Date/Time"
    assert "ICP rating" in headers


def test_the_real_sheets_shape_resolves_every_column():
    _rows, meta = tracker.parse_rows(_sheet())
    assert meta["header_row"] == 2
    assert not meta["warnings"], meta["warnings"]
    assert set(meta["columns"]) >= {"date", "icp", "headcount", "role",
                                    "industry", "location", "category"}


def test_a_sheet_with_no_banner_still_works():
    """The banner is a convention, not a rule."""
    grid = _grid([_row("1/5/2026")], banner=False)
    index, headers = tracker.find_header(grid)
    assert index == 0
    assert headers[0] == "Date/Time"
    rows, meta = tracker.parse_rows(grid)
    assert len(rows) == 1 and not meta["warnings"]


def test_several_banner_rows_are_skipped():
    grid = [["Dovetail x Unstuck"] + [""] * 10,
            [""] * 11,
            ["Updated weekly"] + [""] * 10] + _grid([_row("1/5/2026")], banner=False)
    rows, meta = tracker.parse_rows(grid)
    assert meta["header_row"] == 4
    assert len(rows) == 1


def test_the_header_is_the_row_naming_the_most_columns():
    """Scored, not hardcoded to row 2 -- so a sheet that puts its headers
    anywhere near the top still reads."""
    assert tracker.header_score(["Dovetail", "", ""]) == 0
    assert tracker.header_score(list(HEADERS)) >= 10
    assert tracker.header_score(["9/23/2025", "acme.com", "Smartlead"]) <= 1


def test_a_grid_with_no_recognisable_header_falls_back_to_the_first_row():
    """And then says what it could not find, by name, rather than guessing."""
    grid = [["alpha", "beta"], ["1", "2"]]
    index, headers = tracker.find_header(grid)
    assert index == 0 and headers == ["alpha", "beta"]
    _rows, meta = tracker.parse_rows(grid)
    assert any("ICP" in w for w in meta["warnings"])


def test_a_repeated_column_name_is_read_once():
    grid = [list(HEADERS) + ["Industry"],
            _row("1/5/2026") + ["ignored"]]
    rows, _meta = tracker.parse_rows(grid)
    assert rows[0]["industry"] == "Marketing", "the first Industry column wins"


def test_data_above_the_header_row_is_not_counted_as_leads():
    """A banner that happens to hold a date must not become a lead."""
    grid = [["9/23/2025"] + [""] * 10] + _grid([_row("1/5/2026")], banner=False)
    rows, meta = tracker.parse_rows(grid)
    assert len(rows) == 1
    assert rows[0]["day"] == "2026-01-05"
