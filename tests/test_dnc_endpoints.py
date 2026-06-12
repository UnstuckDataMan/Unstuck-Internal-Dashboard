"""Offline endpoint tests for the DNC tool (FastAPI TestClient + FakeSupabase)."""
from __future__ import annotations

import re
from datetime import date

from conftest import FakeResponse, in_list_values, param_values

import app.routers.dnc as dnc_mod
import app.utils.google_sheets as gs


# ── helpers ────────────────────────────────────────────────────────────────────

def stat_map(html: str) -> dict[str, str]:
    """Extract {label: number} pairs from the scrub-result stat cards."""
    return {lbl: num for num, lbl in re.findall(
        r'<span class="stat-num">([\d,]+)</span>\s*<span class="stat-lbl">([^<]+)</span>',
        html,
    )}


class _SyncThread:
    """threading.Thread stand-in that runs the target synchronously on start()."""
    def __init__(self, target=None, args=(), kwargs=None, daemon=None):
        self._target = target
        self._args   = args
        self._kwargs = kwargs or {}

    def start(self):
        self._target(*self._args, **self._kwargs)


class _ShimThreading:
    Thread = _SyncThread


DNC_TABLE_ROW = [{"id": "e1", "email": "x@y.com", "reason": "manual",
                  "added_by": "dashboard_user", "notes": "",
                  "created_at": "2026-06-12T10:00:00"}]


def table_response(rows):
    return FakeResponse(200, rows, headers={"Content-Range": f"0-0/{len(rows)}"})


# ── Scrub flow ────────────────────────────────────────────────────────────────

SCRUB_CSV = (
    "Email Address,Name\n"
    "A@X.COM,Alice\n"            # on DNC (exact email, case-normalised)
    "b@y.com,Bob\n"              # recently contacted
    "C@z.com,Carol\n"            # clean
    "C@z.com,Carol2\n"           # duplicate clean row
    "d@dnc-domain.com,Dave\n"    # domain-level DNC block
)


def wire_scrub_fakes(fake_sb):
    dnc_hits = {"a@x.com", "dnc-domain.com"}

    def dnc_handler(call):
        wanted = in_list_values(param_values(call, "email")[0])
        return FakeResponse(200, [{"email": v} for v in wanted if v in dnc_hits])

    def contacted_get(call):
        if any(v.startswith("gte.") for v in param_values(call, "contacted_at")):
            return FakeResponse(200, [{"email": "b@y.com"}])   # recent-contact match
        return FakeResponse(200, [])                            # existence pre-check

    fake_sb.route("GET", "dnc_entries", dnc_handler)
    fake_sb.route("GET", "contacted_prospects", contacted_get)
    fake_sb.route("POST", "contacted_prospects", lambda c: FakeResponse(201))
    fake_sb.route("POST", "scrub_logs", lambda c: FakeResponse(201))


def run_scrub(client):
    return client.post(
        "/api/dnc/scrub",
        files={"file": ("prospects.csv", SCRUB_CSV.encode(), "text/csv")},
        data={"client_id": "client-1", "remove_contacted": "on",
              "lookback_days_raw": "30", "save_contacted": "on",
              "campaign_name": "June Camp"},
    )


def test_scrub_counts_and_removal_reasons(fake_sb, client):
    wire_scrub_fakes(fake_sb)
    resp = run_scrub(client)
    assert resp.status_code == 200
    stats = stat_map(resp.text)
    assert stats["Uploaded"] == "5"
    assert stats["Remaining"] == "2"          # both Carol rows survive
    assert stats["Saved to Contacted"] == "1"  # deduped before saving


def test_scrub_save_posts_only_new_deduped_emails(fake_sb, client):
    wire_scrub_fakes(fake_sb)
    run_scrub(client)
    saves = fake_sb.calls_to("POST", "contacted_prospects")
    assert len(saves) == 1
    rows = saves[0]["json"]
    assert [r["email"] for r in rows] == ["c@z.com"]   # 2 file rows → 1 insert
    assert rows[0]["source"] == "scrub_upload"
    assert rows[0]["campaign_name"] == "June Camp"
    assert rows[0]["contacted_at"] == date.today().isoformat()


def test_scrub_skips_emails_already_in_history(fake_sb, client):
    wire_scrub_fakes(fake_sb)

    def contacted_get(call):
        if any(v.startswith("gte.") for v in param_values(call, "contacted_at")):
            return FakeResponse(200, [{"email": "b@y.com"}])
        return FakeResponse(200, [{"email": "c@z.com"}])   # already in history

    fake_sb.route("GET", "contacted_prospects", contacted_get)
    resp = run_scrub(client)
    assert stat_map(resp.text)["Saved to Contacted"] == "0"
    assert fake_sb.calls_to("POST", "contacted_prospects") == [], \
        "no new emails → no insert at all (and an accurate zero, not a fake count)"


def test_scrub_downloads_and_send_to_merge(fake_sb, client):
    wire_scrub_fakes(fake_sb)
    html = run_scrub(client).text

    clean_token   = re.search(r"/api/dnc/download/([0-9a-f-]+)", html).group(1)
    removed_token = re.search(r"/api/dnc/download-removed/([0-9a-f-]+)", html).group(1)

    clean = client.get(f"/api/dnc/download/{clean_token}")
    assert clean.status_code == 200
    assert clean.text.count("C@z.com") == 2
    for gone in ("A@X.COM", "b@y.com", "d@dnc-domain.com"):
        assert gone not in clean.text

    removed = client.get(f"/api/dnc/download-removed/{removed_token}")
    assert removed.status_code == 200
    assert "removal_reason" in removed.text
    assert removed.text.count("DNC") == 2
    assert removed.text.count("Recently Contacted") == 1

    # Scrub → Mail Merge handoff (CSV is converted to xlsx and parsed)
    merge = client.get(f"/api/dnc/send-to-merge/{clean_token}")
    assert merge.status_code == 200
    data = merge.json()
    assert data["total_rows"] == 2
    assert data["headers"] == ["Email Address", "Name"]

    assert client.get("/api/dnc/download/not-a-token").status_code == 404


def test_scrub_rejects_ambiguous_email_columns(fake_sb, client):
    csv = "Email,Work Email\na@b.com,c@d.com\n"
    resp = client.post(
        "/api/dnc/scrub",
        files={"file": ("two.csv", csv.encode(), "text/csv")},
        data={"client_id": "client-1"},
    )
    assert "Multiple possible email columns" in resp.text


# ── Sync from Google Sheet (Manage DNC tab) ───────────────────────────────────

SHEET_LEADS = [
    {"email": "lead1@acme.com", "status": "Lead"},
    {"email": "lead2@acme.com", "status": "Lead"},       # same domain
    {"email": "int@corp.com",   "status": "Interested"},
    {"email": "unsub@corp.com", "status": "Unsubscribe"},
    {"email": "reply@corp.com", "status": "Reply"},      # must NOT be blocked
]


def test_sync_from_sheet_maps_statuses_like_auto_sync(fake_sb, client, monkeypatch):
    monkeypatch.setattr(gs, "read_leads", lambda sid: list(SHEET_LEADS))
    fake_sb.route("POST", "dnc_entries", lambda c: FakeResponse(201))

    resp = client.post("/api/dnc/sync-from-sheet",
                       json={"sheet_id": "sheet-1", "client_id": "client-1"})
    assert resp.status_code == 200
    assert resp.json() == {"leads_added": 2, "interested_added": 1,
                           "unsubscribes_added": 1}

    post = fake_sb.calls_to("POST", "dnc_entries")[0]
    assert param_values(post, "on_conflict") == ["client_id,email"]
    assert "ignore-duplicates" in post["headers"]["Prefer"]
    rows = {r["email"]: r["reason"] for r in post["json"]}
    assert rows == {"acme.com": "lead",            # domain block, deduped
                    "int@corp.com": "interested",
                    "unsub@corp.com": "opt_out"}
    assert "reply@corp.com" not in rows, "a reply is NOT an opt-out"


def test_sync_from_sheet_repeat_run_succeeds_via_fallback(fake_sb, client, monkeypatch):
    """Everything already on the DNC list: chunk 409s, per-row 409s → success."""
    monkeypatch.setattr(gs, "read_leads", lambda sid: list(SHEET_LEADS))
    fake_sb.route("POST", "dnc_entries", lambda c: FakeResponse(409, [], text="conflict"))

    resp = client.post("/api/dnc/sync-from-sheet",
                       json={"sheet_id": "sheet-1", "client_id": "client-1"})
    assert resp.status_code == 200, "re-syncing an already-synced sheet must not error"
    assert resp.json()["leads_added"] == 2
    posts = fake_sb.calls_to("POST", "dnc_entries")
    assert len(posts) == 1 + 3   # failed chunk + one retry per unique row


def test_sync_from_sheet_error_is_valid_json_even_with_quotes(client, monkeypatch):
    def boom(sid):
        raise RuntimeError('bad "quoted" failure')
    monkeypatch.setattr(gs, "read_leads", boom)
    resp = client.post("/api/dnc/sync-from-sheet",
                       json={"sheet_id": "sheet-1", "client_id": "client-1"})
    assert resp.status_code == 500
    assert 'bad "quoted" failure' in resp.json()["error"]   # .json() must not raise


def test_sync_from_sheet_no_actionable_rows(fake_sb, client, monkeypatch):
    monkeypatch.setattr(gs, "read_leads",
                        lambda sid: [{"email": "r@x.com", "status": "Reply"}])
    resp = client.post("/api/dnc/sync-from-sheet",
                       json={"sheet_id": "sheet-1", "client_id": "client-1"})
    assert resp.json()["message"].startswith("No Lead")
    assert fake_sb.calls_to("POST", "dnc_entries") == []


def test_sync_from_sheet_requires_args(client):
    resp = client.post("/api/dnc/sync-from-sheet", json={"sheet_id": "", "client_id": ""})
    assert resp.status_code == 400
    assert "error" in resp.json()


# ── Manual DNC entry → background sheet marking ───────────────────────────────

def wire_add_entry(fake_sb, monkeypatch, campaign_sheets):
    fake_sb.route("POST", "dnc_entries", lambda c: FakeResponse(201))
    fake_sb.route("POST", "contacted_prospects", lambda c: FakeResponse(201))
    fake_sb.route("GET", "campaigns",
                  lambda c: FakeResponse(200, [{"sheet_id": s} for s in campaign_sheets]))
    fake_sb.route("GET", "dnc_entries", lambda c: table_response(DNC_TABLE_ROW))

    marked: list[tuple] = []
    monkeypatch.setattr(gs, "mark_email_in_sheet",
                        lambda sid, value, reason: marked.append((sid, value, reason)) or 1)
    monkeypatch.setattr(dnc_mod, "threading", _ShimThreading)   # run thread inline
    return marked


def test_manual_dnc_marks_all_active_sheets(fake_sb, client, monkeypatch):
    marked = wire_add_entry(fake_sb, monkeypatch, ["sheet-A", "sheet-B"])

    resp = client.post("/api/dnc/entries", data={
        "client_id": "client-1", "email": "Person@Acme.com", "reason": "manual",
    })
    assert resp.status_code == 200

    post = fake_sb.calls_to("POST", "dnc_entries")[0]
    assert post["json"]["email"] == "person@acme.com"   # manual keeps full email

    assert marked == [("sheet-A", "person@acme.com", "manual"),
                      ("sheet-B", "person@acme.com", "manual")]

    # Full-email entries are auto-logged to contacted history
    logged = fake_sb.calls_to("POST", "contacted_prospects")[0]["json"]
    assert logged["email"] == "person@acme.com" and logged["source"] == "dnc_manual"

    # Active-sheet lookup must use the NULL-safe completed filter
    camp_call = fake_sb.calls_to("GET", "campaigns")[0]
    assert "not.is.true" in param_values(camp_call, "completed")


def test_manual_lead_stores_domain_but_marks_original_email(fake_sb, client, monkeypatch):
    marked = wire_add_entry(fake_sb, monkeypatch, ["sheet-A"])

    # Domain mode: the JS strips the email to a domain and passes the original
    resp = client.post("/api/dnc/entries", data={
        "client_id": "client-1", "email": "acme.com", "reason": "lead",
        "original_email": "Person@Acme.com",
    })
    assert resp.status_code == 200

    post = fake_sb.calls_to("POST", "dnc_entries")[0]
    assert post["json"]["email"] == "acme.com"          # domain-level block stored
    assert marked == [("sheet-A", "person@acme.com", "lead")], \
        "the specific prospect row must be marked via the pre-strip email"

    # Domain entries are NOT auto-logged to contacted history
    assert fake_sb.calls_to("POST", "contacted_prospects") == []


def test_manual_lead_with_full_email_blocks_domain(fake_sb, client, monkeypatch):
    marked = wire_add_entry(fake_sb, monkeypatch, ["sheet-A"])
    client.post("/api/dnc/entries", data={
        "client_id": "client-1", "email": "person@acme.com", "reason": "lead",
    })
    post = fake_sb.calls_to("POST", "dnc_entries")[0]
    assert post["json"]["email"] == "acme.com"
    assert marked[0][1] == "person@acme.com"


def test_manual_dnc_already_exists_notification(fake_sb, client, monkeypatch):
    wire_add_entry(fake_sb, monkeypatch, [])
    fake_sb.route("POST", "dnc_entries", lambda c: FakeResponse(409, [], text="dup"))
    fake_sb.route("GET", "dnc_entries", lambda c: FakeResponse(200, [
        {"email": "acme.com", "reason": "lead", "created_at": "2026-01-15T09:00:00"},
    ]))
    resp = client.post("/api/dnc/entries", data={
        "client_id": "client-1", "email": "person@acme.com", "reason": "manual",
    })
    assert "already on the DNC list" in resp.text
    assert "Domain-level block" in resp.text       # tells the user the REAL level
    assert resp.headers.get("HX-Retarget") == "#add-entry-error"


def test_manual_dnc_validation_errors(fake_sb, client):
    bad_domain = client.post("/api/dnc/entries", data={
        "client_id": "c1", "email": "not-a-domain", "reason": "manual"})
    assert "valid domain" in bad_domain.text
    double_at = client.post("/api/dnc/entries", data={
        "client_id": "c1", "email": "a@@b.com", "reason": "manual"})
    assert "valid email" in double_at.text


# ── Bulk import + contacted upload (targeted duplicate pre-check) ─────────────

def test_bulk_import_skips_existing_via_targeted_query(fake_sb, client):
    def dnc_get(call):
        if param_values(call, "email") and "in.(" in param_values(call, "email")[0]:
            wanted = in_list_values(param_values(call, "email")[0])
            assert set(wanted) == {"new@x.com", "old@x.com", "spam.com"}, \
                "pre-check must query ONLY the uploaded values"
            return FakeResponse(200, [{"email": "old@x.com"}])
        return table_response(DNC_TABLE_ROW)   # final table refresh

    fake_sb.route("GET", "dnc_entries", dnc_get)
    fake_sb.route("POST", "dnc_entries", lambda c: FakeResponse(201))

    csv = "email\nnew@x.com\nOLD@X.COM\nnew@x.com\nspam.com\n"
    resp = client.post(
        "/api/dnc/entries/bulk",
        files={"file": ("bulk.csv", csv.encode(), "text/csv")},
        data={"client_id": "client-1", "reason": "bulk_import"},
    )
    assert resp.status_code == 200
    inserted = fake_sb.calls_to("POST", "dnc_entries")[0]["json"]
    assert {r["email"] for r in inserted} == {"new@x.com", "spam.com"}, \
        "existing entry skipped, file duplicates collapsed, bare domains accepted"


def test_contacted_upload_reports_added_and_skipped(fake_sb, client):
    def contacted_get(call):
        sel = param_values(call, "select")
        if sel and "count()" in sel[0]:
            return FakeResponse(200, [{"campaign_name": "Camp A", "count": 3}])
        return FakeResponse(200, [{"email": "old@x.com"}])     # existence pre-check

    fake_sb.route("GET", "contacted_prospects", contacted_get)
    fake_sb.route("POST", "contacted_prospects", lambda c: FakeResponse(201))

    csv = "email\nnew1@x.com\nold@x.com\nnew2@x.com\n"
    resp = client.post(
        "/api/dnc/contacted/upload",
        files={"file": ("c.csv", csv.encode(), "text/csv")},
        data={"client_id": "client-1", "contacted_at": "2026-06-10",
              "campaign_name": "Camp A"},
    )
    assert resp.status_code == 200
    assert "2 emails added" in resp.text and "1 already existed" in resp.text
    rows = fake_sb.calls_to("POST", "contacted_prospects")[0]["json"]
    assert {r["email"] for r in rows} == {"new1@x.com", "new2@x.com"}
    assert all(r["contacted_at"] == "2026-06-10" and r["source"] == "csv_upload"
               for r in rows)

    # Existence pre-check is scoped to the campaign
    pre = [c for c in fake_sb.calls_to("GET", "contacted_prospects")
           if param_values(c, "email")][0]
    assert "eq.Camp A" in param_values(pre, "campaign_name")


# ── Campaign grouping: aggregate fast-path + paging fallback ──────────────────

def test_contacted_campaigns_uses_single_aggregate_request(fake_sb, client):
    fake_sb.route("GET", "contacted_prospects", lambda c: FakeResponse(200, [
        {"campaign_name": "Big",   "count": 120},
        {"campaign_name": "Small", "count": 3},
        {"campaign_name": None,    "count": 7},
    ]))
    resp = client.get("/api/dnc/contacted/campaigns", params={"client_id": "c1"})
    assert resp.status_code == 200
    assert "Big" in resp.text and "Small" in resp.text
    gets = fake_sb.calls_to("GET", "contacted_prospects")
    assert len(gets) == 1, "aggregate path must be a single request"
    assert "count()" in param_values(gets[0], "select")[0]


def test_contacted_campaigns_falls_back_when_aggregates_disabled(fake_sb, client):
    def handler(call):
        if "count()" in param_values(call, "select")[0]:
            return FakeResponse(400, [], text="aggregates disabled")
        return FakeResponse(200, [{"campaign_name": "Solo"},
                                  {"campaign_name": None}])
    fake_sb.route("GET", "contacted_prospects", handler)
    resp = client.get("/api/dnc/contacted/campaigns", params={"client_id": "c1"})
    assert resp.status_code == 200
    assert "Solo" in resp.text


# ── Delete failures are surfaced, not swallowed ───────────────────────────────

def test_delete_contacted_surfaces_supabase_failure(fake_sb, client):
    fake_sb.route("DELETE", "contacted_prospects",
                  lambda c: FakeResponse(500, [], text="boom"))
    resp = client.delete("/api/dnc/contacted/entry-1", params={"client_id": "c1"})
    assert resp.status_code == 200          # HTMX still swaps the partial
    assert "Delete failed" in resp.text


def test_campaign_sheet_options_escape_html(fake_sb, client):
    fake_sb.route("GET", "campaigns", lambda c: FakeResponse(200, [
        {"campaign_name": 'Evil <script>"x"&', "sheet_id": "sheet-1"},
    ]))
    resp = client.get("/api/dnc/campaign-sheet-options", params={"client_id": "c1"})
    assert "<script>" not in resp.text
    assert "&lt;script&gt;" in resp.text
