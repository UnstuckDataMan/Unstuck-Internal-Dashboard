"""Offline tests for Google Sheet marking and conditional-format rules.

gspread and the AuthorizedSession are replaced with in-memory fakes; the
batchUpdate payloads are inspected directly.
"""
from __future__ import annotations

import pytest

from conftest import (
    FakeGC, FakeResponse, FakeSession, FakeSpreadsheet, FakeWorksheet, cf_requests,
)

import app.utils.google_sheets as gs


SHEET_HEADERS = ["Send Status", "Sender Account", "Recipient Email", "Lead Status"]
LS_LETTER = "D"   # Lead Status = 4th column


def make_ws(emails):
    rows = [["", f"s{i % 2}@unstuck.com", e, ""] for i, e in enumerate(emails)]
    return FakeWorksheet(SHEET_HEADERS, rows)


@pytest.fixture
def marked_sheet(monkeypatch):
    """Wire a fake sheet into mark_email_in_sheet; returns (ws, session)."""
    ws = make_ws(["a@acme.com", "b@other.com", "a@acme.com"])
    session = FakeSession()
    monkeypatch.setattr(gs, "_client", lambda: FakeGC({"sheet-1": FakeSpreadsheet(ws)}))
    monkeypatch.setattr(gs, "_authed_session", lambda: session)
    return ws, session


# ── mark_email_in_sheet: exact matching + reason mapping ──────────────────────

def test_mark_updates_all_exact_matches_only(marked_sheet):
    ws, _ = marked_sheet
    n = gs.mark_email_in_sheet("sheet-1", "a@acme.com", reason="lead")
    assert n == 2
    updates = ws.batch_updates[0]
    # Rows 2 and 4 hold a@acme.com; b@other.com (row 3) untouched
    assert {u["range"] for u in updates} == {f"{LS_LETTER}2", f"{LS_LETTER}4"}
    assert all(u["values"] == [["Lead"]] for u in updates)


@pytest.mark.parametrize("reason,expected", [
    ("lead", "Lead"), ("interested", "Interested"), ("reply", "Reply"),
    ("opt_out", "Unsubscribe"), ("manual", "Unsubscribe"),
    ("unknown_reason", "Unsubscribe"),   # safe default
])
def test_reason_maps_to_lead_status(marked_sheet, reason, expected):
    ws, _ = marked_sheet
    gs.mark_email_in_sheet("sheet-1", "a@acme.com", reason=reason)
    assert ws.batch_updates[0][0]["values"] == [[expected]]


def test_mark_is_case_insensitive(marked_sheet):
    ws, _ = marked_sheet
    assert gs.mark_email_in_sheet("sheet-1", "  A@ACME.COM ", reason="lead") == 2


def test_bare_domain_marks_nothing(monkeypatch):
    # Must return 0 BEFORE any sheet access — marking a whole domain would
    # inflate lead_count stats (CF rules grey the other rows instead).
    monkeypatch.setattr(gs, "_client",
                        lambda: (_ for _ in ()).throw(AssertionError("must not open sheet")))
    assert gs.mark_email_in_sheet("sheet-1", "acme.com", reason="lead") == 0


def test_no_match_returns_zero_and_adds_no_cf(marked_sheet):
    ws, session = marked_sheet
    assert gs.mark_email_in_sheet("sheet-1", "nobody@nowhere.com") == 0
    assert ws.batch_updates == []
    assert session.posts == []


def test_sheet_without_expected_columns_returns_zero(monkeypatch):
    ws = FakeWorksheet(["Email", "Status"], [["a@b.com", ""]])
    monkeypatch.setattr(gs, "_client", lambda: FakeGC({"s": FakeSpreadsheet(ws)}))
    assert gs.mark_email_in_sheet("s", "a@b.com") == 0


# ── CF rules added on mark ────────────────────────────────────────────────────

def lead_grey_formula():
    return (f'=OR(${LS_LETTER}2="Lead",${LS_LETTER}2="Reply",'
            f'${LS_LETTER}2="Interested",${LS_LETTER}2="Unsubscribe")')


def sender_stripe_formula():
    return '=AND($B2<>"",$B2<>$B1)'   # Sender Account = column B


def test_mark_lead_adds_row_grey_domain_grey_and_stripe(marked_sheet, monkeypatch):
    ws, session = marked_sheet
    monkeypatch.setattr(gs, "_existing_cf_formulas", lambda sid, wid: set())
    gs.mark_email_in_sheet("sheet-1", "a@acme.com", reason="lead")

    reqs = cf_requests(session)
    formulas = [r["addConditionalFormatRule"]["rule"]["booleanRule"]["condition"]
                ["values"][0]["userEnteredValue"] for r in reqs]
    assert lead_grey_formula() in formulas
    assert sender_stripe_formula() in formulas
    assert any("COUNTIFS" in f for f in formulas), "domain-grey rule missing"
    assert all(r["addConditionalFormatRule"]["index"] == 0 for r in reqs)


def test_mark_non_lead_skips_domain_grey(marked_sheet, monkeypatch):
    _, session = marked_sheet
    monkeypatch.setattr(gs, "_existing_cf_formulas", lambda sid, wid: set())
    gs.mark_email_in_sheet("sheet-1", "a@acme.com", reason="opt_out")
    formulas = [r["addConditionalFormatRule"]["rule"]["booleanRule"]["condition"]
                ["values"][0]["userEnteredValue"] for r in cf_requests(session)]
    assert not any("COUNTIFS" in f for f in formulas), \
        "domain-grey is Lead-only — unsubscribes must not grey the whole domain"


def test_repeat_marks_do_not_duplicate_cf_rules(marked_sheet, monkeypatch):
    """The June 2026 fix: identical rules must not pile up on every mark."""
    _, session = marked_sheet
    existing = {lead_grey_formula(), sender_stripe_formula()}
    monkeypatch.setattr(gs, "_existing_cf_formulas", lambda sid, wid: set(existing))
    gs.mark_email_in_sheet("sheet-1", "a@acme.com", reason="opt_out")
    assert cf_requests(session) == [], \
        "rules already on the sheet must not be re-added"


def test_cf_lookup_failure_falls_back_to_always_add(marked_sheet, monkeypatch):
    """None (lookup failed) → old always-add behaviour: missing rule is worse
    than a duplicate."""
    _, session = marked_sheet
    monkeypatch.setattr(gs, "_existing_cf_formulas", lambda sid, wid: None)
    gs.mark_email_in_sheet("sheet-1", "a@acme.com", reason="opt_out")
    assert len(cf_requests(session)) == 2   # row-grey + sender stripe


def test_lead_grey_ranges_exclude_lead_status_column(marked_sheet, monkeypatch):
    """The coloured LS cell must never be covered by the grey row fill."""
    _, session = marked_sheet
    monkeypatch.setattr(gs, "_existing_cf_formulas", lambda sid, wid: set())
    gs.mark_email_in_sheet("sheet-1", "a@acme.com", reason="opt_out")
    grey = [r for r in cf_requests(session)
            if r["addConditionalFormatRule"]["rule"]["booleanRule"]["condition"]
               ["values"][0]["userEnteredValue"] == lead_grey_formula()]
    ranges = grey[0]["addConditionalFormatRule"]["rule"]["ranges"]
    ls0 = 3   # Lead Status 0-based index
    covered = set()
    for rng in ranges:
        covered.update(range(rng["startColumnIndex"], rng["endColumnIndex"]))
    assert ls0 not in covered
    assert covered == {0, 1, 2}   # all other columns are greyed


# ── _existing_cf_formulas parsing ─────────────────────────────────────────────

def test_existing_cf_formulas_filters_by_worksheet(monkeypatch):
    payload = {"sheets": [
        {"properties": {"sheetId": 777},
         "conditionalFormats": [
             {"booleanRule": {"condition": {"type": "CUSTOM_FORMULA",
                                            "values": [{"userEnteredValue": "=A1"}]}}},
             {"booleanRule": {"condition": {"type": "TEXT_EQ",
                                            "values": [{"userEnteredValue": "Lead"}]}}},
         ]},
        {"properties": {"sheetId": 999},
         "conditionalFormats": [
             {"booleanRule": {"condition": {"type": "CUSTOM_FORMULA",
                                            "values": [{"userEnteredValue": "=B1"}]}}},
         ]},
    ]}
    monkeypatch.setattr(gs, "_authed_session", lambda: FakeSession(get_json=payload))
    out = gs._existing_cf_formulas("sheet-1", 777)
    assert out == {"=A1"}      # other worksheet + non-custom-formula rules excluded


def test_existing_cf_formulas_returns_none_on_http_error(monkeypatch):
    class FailingSession(FakeSession):
        def get(self, url, params=None, **kw):
            return FakeResponse(403, [], text="forbidden")
    monkeypatch.setattr(gs, "_authed_session", lambda: FailingSession())
    assert gs._existing_cf_formulas("sheet-1", 777) is None


# ── Full-sheet formatting (creation-time CF rules) ────────────────────────────

OUTREACH_HEADERS = ["Send Status", "Send Time", "Sender Account", "First Name",
                    "Recipient Email", "Subject Line", "Email Body", "A/B Variant",
                    "Chaser Sent?", "Lead Status", "Notes", "__divider__", "Company"]


def outreach_rows():
    rows = [OUTREACH_HEADERS]
    rows += [["", "09:00", "s1@x.com", "A", "a@acme.com", "S", "B", "S1/B1",
              "", "", "", "", "ACME"],
             ["", "09:10", "s2@x.com", "B", "b@beta.com", "S", "B", "S2/B1",
              "", "", "", "", "Beta"],
             ["No More Emails For Today."] + [""] * (len(OUTREACH_HEADERS) - 1)]
    return rows


def test_sheet_formatting_requests(monkeypatch):
    session = FakeSession()
    gs._apply_sheet_formatting(session, "sheet-1", 777, outreach_rows())
    reqs = cf_requests(session)
    by_kind: dict[str, list] = {}
    for r in reqs:
        for k in r:
            by_kind.setdefault(k, []).append(r[k])

    # Separator row: merged + styled
    assert len(by_kind.get("mergeCells", [])) == 1
    merge = by_kind["mergeCells"][0]["range"]
    assert merge["startRowIndex"] == 3 and merge["endColumnIndex"] == len(OUTREACH_HEADERS)

    # Checkboxes for Send Status AND Chaser Sent?; dropdown with all 4 statuses
    validations = by_kind.get("setDataValidation", [])
    bool_cols = {v["range"]["startColumnIndex"] for v in validations
                 if v["rule"]["condition"]["type"] == "BOOLEAN"}
    assert bool_cols == {0, 8}, "Send Status (A) and Chaser Sent? (I) need checkboxes"
    dropdowns = [v for v in validations
                 if v["rule"]["condition"]["type"] == "ONE_OF_LIST"]
    values = {x["userEnteredValue"] for x in dropdowns[0]["rule"]["condition"]["values"]}
    assert values == {"Lead", "Reply", "Interested", "Unsubscribe"}

    # CF rules: sent-blue + domain-grey + 4 status colours + row-grey + stripe
    cf = by_kind.get("addConditionalFormatRule", [])
    assert len(cf) == 8
    assert all(r["index"] == 0 for r in cf)
    # Last-added rule wins priority — must be the sender stripe
    last = cf[-1]["rule"]
    stripe_col = OUTREACH_HEADERS.index("Sender Account")
    assert last["ranges"][0]["startColumnIndex"] == stripe_col

    # Divider column: narrow width + grey fills
    dims = by_kind.get("updateDimensionProperties", [])
    div_col = OUTREACH_HEADERS.index("__divider__")
    assert any(d["range"]["startIndex"] == div_col and
               d["properties"]["pixelSize"] == 20 for d in dims)


# ── create_outreach_sheet helper-column injection ─────────────────────────────

def test_create_outreach_sheet_adds_tracking_columns(monkeypatch, tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Outreach List"
    ws.append(["Send Status", "Recipient Email", "Chaser Body", "Lead Status",
               "Sender Account"])
    ws.append(["", "a@b.com", "bump", "", "s1@x.com"])
    path = tmp_path / "m.xlsx"
    wb.save(str(path))

    fake_ws = FakeWorksheet([], [])
    sh = FakeSpreadsheet(fake_ws)
    monkeypatch.setattr(gs, "_create_spreadsheet", lambda title: ("NEW123", "http://sheet"))
    monkeypatch.setattr(gs, "_client", lambda: FakeGC({"NEW123": sh}))
    session = FakeSession()
    monkeypatch.setattr(gs, "_authed_session", lambda: session)

    out = gs.create_outreach_sheet("Client – June – 2026-06-12", str(path))
    assert out == {"sheet_id": "NEW123", "sheet_url": "http://sheet",
                   "title": "Client – June – 2026-06-12"}

    written = [u for u in fake_ws.updates if u[0] == "update"][0]
    headers = written[2][0]
    assert headers[-1] == "Sent Date", "Sent Date appended for date tracking"
    cb = headers.index("Chaser Body")
    assert headers[cb + 1] == "Chaser Sent?" and headers[cb + 2] == "Chaser Date"
    # Every data row padded to the same width
    assert all(len(r) == len(headers) for r in written[2])
    # Shared as anyone-with-link editor + formatting batch sent
    assert sh.shares and sh.shares[0][1] == {"perm_type": "anyone", "role": "writer"}
    assert session.posts, "formatting batchUpdate must run"


# ── Cell-value parsing + cache behaviour ──────────────────────────────────────

def test_is_sent_accepts_all_representations():
    assert gs._is_sent(True) and gs._is_sent("TRUE") and gs._is_sent(" true ")
    assert gs._is_sent("Sent") and gs._is_sent(1)
    assert not gs._is_sent(False) and not gs._is_sent("") and not gs._is_sent("FALSE")
    assert not gs._is_sent("Reply") and not gs._is_sent(0)


def test_to_iso_date_normalises_serials_and_strings():
    assert gs._to_iso_date("2026-06-12") == "2026-06-12"
    assert gs._to_iso_date("2026-06-12T10:00:00") == "2026-06-12"
    assert gs._to_iso_date(46100) == "2026-03-19"      # Sheets serial number
    assert gs._to_iso_date(True) == ""                 # bool is not a date
    assert gs._to_iso_date(None) == "" and gs._to_iso_date("junk") == ""
    assert gs._to_iso_date(99999999) == ""             # implausible serial


def test_records_cache_hits_and_invalidation(monkeypatch):
    ws = make_ws(["a@acme.com"])
    calls = {"n": 0}
    real = ws.get_all_values
    def counting(**kw):
        calls["n"] += 1
        return real(**kw)
    ws.get_all_values = counting

    gs._records_cache.clear()
    r1 = gs._get_all_records(ws, "cache-test", value_render_option="UNFORMATTED_VALUE")
    r2 = gs._get_all_records(ws, "cache-test", value_render_option="UNFORMATTED_VALUE")
    assert r1 == r2 and calls["n"] == 1, "second read within TTL must be cached"

    gs._invalidate_sheet_cache("cache-test")
    gs._get_all_records(ws, "cache-test", value_render_option="UNFORMATTED_VALUE")
    assert calls["n"] == 2, "write invalidation must force a fresh read"
