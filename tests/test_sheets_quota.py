"""Offline tests for the Google Sheets quota discipline.

The hourly campaign sync failed intermittently with 429s because every per-sheet
helper opened the spreadsheet for itself, and gspread's `Spreadsheet.sheet1` is
a real spreadsheets.get request — not a free attribute lookup.  One campaign
therefore spent 3–7 reads against the 60-reads/min-per-user quota while the sync
loop's pacing was tuned for 1–2, and the metadata fetch was the one call with no
429 retry around it.

These tests pin the read budget per campaign, the retry, and the shared limiter.
"""
from __future__ import annotations

import pytest
from gspread.exceptions import APIError

from conftest import FakeGC, FakeResponse, FakeSpreadsheet, FakeWorksheet

import app.utils.auto_sync as auto_sync
import app.utils.google_sheets as gs


SHEET_HEADERS = ["Send Status", "Recipient Email", "Lead Status",
                 "Sent Date", "Chaser Sent?", "Chaser Date"]

# Row 1 is chased with no Chaser Date, row 2 is sent with no Sent Date, so a
# sync of this sheet exercises BOTH date back-fills — the pair that used to cost
# two metadata fetches plus a full re-read of the sheet between them.
SHEET_ROWS = [
    ["TRUE", "a@acme.com", "Lead", "2026-06-01", "TRUE",  ""],
    ["TRUE", "b@corp.com", "",     "",           "FALSE", ""],
]


def _api_error(status: int) -> APIError:
    """Build a gspread APIError carrying an HTTP status, as the API would."""
    payload = {"error": {"code": status, "message": "quota exceeded",
                         "status": "RESOURCE_EXHAUSTED"}}
    return APIError(FakeResponse(status, payload))


@pytest.fixture
def sheet(monkeypatch):
    """Wire one fake spreadsheet in behind gs._client(); returns the fake."""
    ws = FakeWorksheet(SHEET_HEADERS, [list(r) for r in SHEET_ROWS])
    sh = FakeSpreadsheet(ws, tabs={"Stats": ws})
    monkeypatch.setattr(gs, "_client", lambda: FakeGC({"sheet-1": sh}))
    return sh


@pytest.fixture
def sb_ready(fake_sb):
    """Supabase routes for a sync that should run all the way through."""
    fake_sb.route("GET",   "contacted_prospects", lambda c: FakeResponse(200, []))
    fake_sb.route("POST",  "contacted_prospects", lambda c: FakeResponse(201))
    fake_sb.route("POST",  "dnc_entries",         lambda c: FakeResponse(201))
    fake_sb.route("PATCH", "campaigns",           lambda c: FakeResponse(204))
    return fake_sb


# ── Read budget ───────────────────────────────────────────────────────────────

def test_full_campaign_sync_spends_two_sheet_reads(sheet, sb_ready):
    """A cold sync of one campaign = one metadata fetch + one values fetch.

    Regression: this used to be four metadata fetches (read_leads,
    read_sent_with_dates, write_sent_dates, write_chaser_dates each opened the
    sheet) plus two values fetches, because the Sent Date write dropped the
    records cache and the Chaser Date write immediately re-read the whole sheet.
    Six reads per campaign against a 60/min budget is what exhausted the quota.
    """
    result = auto_sync.sync_campaign_core(
        "camp-1", "sheet-1", "client-1", "June Campaign",
    )

    assert result["error"] == ""
    assert sheet.metadata_fetches == 1, \
        "the worksheet handle must be fetched once, not once per helper"
    assert gs.sheet_reads_count() == 2, \
        "one metadata fetch + one values fetch is the whole read cost"


def test_worksheet_handle_is_shared_by_every_helper(sheet):
    """Independent read helpers must reuse the cached handle, not re-open."""
    gs.read_leads("sheet-1")
    gs.read_sent_with_dates("sheet-1")
    gs.read_ab_stats("sheet-1")

    assert sheet.metadata_fetches == 1


def test_second_date_backfill_reuses_the_cached_rows(sheet, sb_ready):
    """Writing Sent Date must not force Chaser Date to re-read the sheet.

    The write patches the just-written values into the cached records instead of
    dropping the cache, so the rows stay both warm and correct.
    """
    auto_sync.sync_campaign_core("camp-1", "sheet-1", "client-1", "June")

    values_reads = gs.sheet_reads_count() - sheet.metadata_fetches
    assert values_reads == 1, "the sheet's rows must be fetched exactly once"

    # ...and the cache now reflects what was written, not the pre-write blanks.
    cached = gs._get_all_records(sheet.sheet1, "sheet-1",
                                 value_render_option="UNFORMATTED_VALUE")
    assert cached[0]["Chaser Date"], "written Chaser Date must be in the cache"
    assert cached[1]["Sent Date"],   "written Sent Date must be in the cache"


def test_header_change_drops_both_caches(monkeypatch, sb_ready):
    """Appending a date column invalidates the records AND worksheet caches.

    The handle carries the grid width used to decide whether the grid must grow
    first, so a stale one could reintroduce the "exceeds grid limits" failure.
    """
    # No Sent Date / Chaser Date columns → the write has to append one.
    ws = FakeWorksheet(["Send Status", "Recipient Email", "Lead Status"],
                       [["TRUE", "a@acme.com", ""]])
    sh = FakeSpreadsheet(ws)
    monkeypatch.setattr(gs, "_client", lambda: FakeGC({"sheet-1": sh}))

    gs.write_sent_dates("sheet-1", [(2, "2026-06-01")])

    assert ws.added_cols, "a tight grid must be widened before the append"
    assert not gs._ws_cache, "worksheet cache must drop after a header change"
    assert not gs._records_cache, "records cache must drop after a header change"


# ── Retry ─────────────────────────────────────────────────────────────────────

def test_metadata_fetch_retries_on_429_instead_of_failing_the_campaign(
    monkeypatch, sb_ready
):
    """A 429 on the metadata fetch must back off and retry, not fail the sync.

    This was the actual failure in the logs: `sh.sheet1` sat outside the retry
    wrapper, so a quota burst propagated straight out of read_leads and the
    campaign was recorded as failed for the hour.
    """
    monkeypatch.setattr(gs.time, "sleep", lambda s: None)   # no real backoff
    ws = FakeWorksheet(SHEET_HEADERS, [list(r) for r in SHEET_ROWS])

    class Flaky(FakeSpreadsheet):
        fails = 2

        @property
        def sheet1(self):
            self.metadata_fetches += 1
            if self.metadata_fetches <= self.fails:
                raise _api_error(429)
            return self._ws

    # One instance, reused across _client() calls — a fresh object per call
    # would reset the failure counter and never let the retry succeed.
    flaky = Flaky(ws)
    monkeypatch.setattr(gs, "_client", lambda: FakeGC({"sheet-1": flaky}))

    result = auto_sync.sync_campaign_core("camp-1", "sheet-1", "client-1", "June")

    assert result["error"] == "", "a transient 429 must not fail the campaign"


def test_non_retryable_api_error_still_raises(monkeypatch, sheet):
    """403/404 are permanent — retrying them just burns quota."""
    monkeypatch.setattr(gs.time, "sleep", lambda s: None)

    def boom(*_a, **_kw):
        raise _api_error(403)

    monkeypatch.setattr(sheet.sheet1, "get_all_values", boom)

    with pytest.raises(APIError):
        gs.read_leads("sheet-1")


# ── Shared quota limiter ──────────────────────────────────────────────────────

def test_rate_limiter_allows_the_budget_then_waits_out_the_window(monkeypatch):
    clock = {"t": 1000.0}
    slept: list[float] = []

    def fake_sleep(seconds):
        slept.append(seconds)
        clock["t"] += seconds

    monkeypatch.setattr(gs.time, "monotonic", lambda: clock["t"])
    monkeypatch.setattr(gs.time, "sleep", fake_sleep)

    limiter = gs._RateLimiter(3)
    for _ in range(3):
        limiter.acquire()
    assert slept == [], "requests within the budget must never block"

    limiter.acquire()   # 4th in the same minute — must wait the window out
    assert slept and sum(slept) >= 60.0

    # Once the window has rolled, the budget is available again immediately.
    slept.clear()
    clock["t"] += 60.0
    limiter.acquire()
    assert slept == []


def test_reads_and_writes_draw_on_separate_budgets(sheet, sb_ready):
    """Sheets meters reads and writes separately, so the limiters must too."""
    assert gs._read_limiter is not gs._write_limiter

    auto_sync.sync_campaign_core("camp-1", "sheet-1", "client-1", "June")

    # Both back-fills are writes; neither may be charged to the read counter.
    assert len(gs._write_limiter._hits) == 2, "two date back-fills = two writes"
    assert gs.sheet_reads_count() == 2, "writes must not inflate the read count"


# ── read_stats_cells never breaks a sync ──────────────────────────────────────

def test_read_stats_cells_returns_empty_on_failure(monkeypatch, sheet):
    """Regression: the failure path raised NameError instead of returning {}.

    google_sheets.py referenced `logger` without ever importing logging, so the
    one branch that used it — read_stats_cells' except handler — blew up with
    NameError.  auto_sync calls read_stats_cells unguarded, so a rate-limited
    Stats read took the whole campaign's sync down with it.
    """
    monkeypatch.setattr(gs.time, "sleep", lambda s: None)
    sheet.sheet1.batch_get_result = _api_error(429)

    assert gs.read_stats_cells("sheet-1") == {}


def test_missing_stats_tab_is_silent(monkeypatch):
    """Sheets created before the Stats tab existed must skip, not raise."""
    ws = FakeWorksheet(SHEET_HEADERS, [])
    monkeypatch.setattr(gs, "_client",
                        lambda: FakeGC({"sheet-1": FakeSpreadsheet(ws)}))  # no tabs

    assert gs.read_stats_cells("sheet-1") == {}
