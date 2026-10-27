"""Read AI meetings: webhook, matching, AI grounding, onboarding notes, panels."""
from __future__ import annotations

import base64
import hashlib
import hmac
import json

import pytest

from tests.conftest import FakeResponse, param_values

from app import auth
from app.routers import pulse_meetings as routes
from app.utils import llm
from app.utils import supabase
from app.utils.pulse import meeting_insights as mi
from app.utils.pulse import meetings_store as ms
from app.utils.pulse import readai, store

AGENCY = "11111111-1111-1111-1111-111111111111"
CLIENT = "33333333-3333-3333-3333-333333333333"
OTHER = "44444444-4444-4444-4444-444444444444"
MEETING = "55555555-5555-5555-5555-555555555555"
SECRET = base64.b64encode(b"read-ai-signing-key").decode()

TRANSCRIPT = [
    {"i": 0, "speaker": "Sam (Unstuck)", "start_ms": 1000, "end_ms": 5000,
     "text": "Thanks for joining, how have the campaigns been landing for you?"},
    {"i": 1, "speaker": "Jo (Acme)", "start_ms": 6000, "end_ms": 20000,
     "text": "Honestly the meetings from the fintech list have been brilliant, three of them converted."},
    {"i": 2, "speaker": "Jo (Acme)", "start_ms": 21000, "end_ms": 40000,
     "text": "The healthcare leads were mostly too junior, we need heads of department not coordinators."},
]


@pytest.fixture(autouse=True)
def _setup(monkeypatch):
    store.reset_agency_cache()
    monkeypatch.setenv("PULSE_AGENCY_ID", AGENCY)
    monkeypatch.setattr(supabase, "SUPABASE_SERVICE_ROLE_KEY", "test-service-key")
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    monkeypatch.delenv("READ_AI_WEBHOOK_SECRET", raising=False)
    monkeypatch.setenv("AGENCY_EMAIL_DOMAINS", "unstuck-agency.com")
    started = []
    monkeypatch.setattr(mi, "run_async", lambda target, *a, **kw: started.append((target.__name__, a, kw)))
    mi._drafting.clear()
    yield started
    store.reset_agency_cache()


def _sign(raw: bytes, secret: str = SECRET) -> str:
    return hmac.new(base64.b64decode(secret), raw, hashlib.sha256).hexdigest()


def _payload(**over):
    body = {
        "session_id": "sess-1", "request_id": "req-1", "trigger": "meeting_end",
        "title": "Acme onboarding call", "start_time": "2026-10-12T10:00:00Z",
        "end_time": 1760266800000,
        "owner": {"name": "Sam", "email": "sam@unstuck-agency.com"},
        "participants": [{"name": "Sam", "email": "sam@unstuck-agency.com"},
                         {"name": "Jo", "email": "jo@Acme.com"}],
        "summary": "Kick-off with Acme.", "action_items": [{"text": "Send ICP doc"}],
        "report_url": "https://app.read.ai/analytics/meetings/abc",
        "transcript": {"speakers": [{"name": "Sam"}], "speaker_blocks": [
            {"start_time": 1000, "end_time": 5000, "speaker": {"name": "Sam"}, "words": "Hello there."},
            {"start_time": 6000, "end_time": 9000, "speaker": {"name": "Jo"}, "words": "  "},
            {"start_time": 9000, "end_time": 12000, "speaker": "Jo", "words": "We sell to CFOs."},
        ]},
    }
    body.update(over)
    return body


# ── Signature ─────────────────────────────────────────────────────────────────

def test_signature_accepts_the_right_key_and_rejects_others():
    raw = b'{"a":1}'
    assert readai.verify_signature(raw, _sign(raw), SECRET)
    assert readai.verify_signature(raw, "sha256=" + _sign(raw).upper(), SECRET)
    assert not readai.verify_signature(raw, _sign(b'{"a":2}'), SECRET)
    assert not readai.verify_signature(raw, "", SECRET)
    assert not readai.verify_signature(raw, _sign(raw), "")
    other = base64.b64encode(b"another-key").decode()
    assert not readai.verify_signature(raw, _sign(raw, other), SECRET)


def test_webhook_is_unavailable_without_a_secret(client):
    r = client.post("/webhooks/read-ai", content=b"{}")
    assert r.status_code == 503


def test_webhook_rejects_a_bad_signature(client, monkeypatch):
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    raw = json.dumps(_payload()).encode()
    r = client.post("/webhooks/read-ai", content=raw, headers={"X-Read-Signature": "0" * 64})
    assert r.status_code == 401


def test_webhook_is_public_but_api_routes_still_need_the_tool(client, monkeypatch):
    monkeypatch.delenv("AUTH_DISABLED", raising=False)
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    monkeypatch.setattr(mi, "ingest", lambda parsed: {"result": "matched"})
    raw = json.dumps(_payload()).encode()
    r = client.post("/webhooks/read-ai", content=raw, headers={"X-Read-Signature": _sign(raw)})
    assert r.status_code == 200 and r.json() == {"result": "matched"}

    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/read-ai", follow_redirects=False)
    assert r.status_code == 401
    user = {"email": "v@unstuck-agency.com", "role": "viewer", "tools": ["campaigns"]}
    monkeypatch.setattr(auth, "get_session_user", lambda request: user)
    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/read-ai", follow_redirects=False)
    assert r.status_code == 403


def test_webhook_ignores_meeting_start(client, monkeypatch):
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    monkeypatch.setattr(mi, "ingest", lambda parsed: pytest.fail("should not ingest"))
    raw = json.dumps(_payload(trigger="meeting_start")).encode()
    r = client.post("/webhooks/read-ai", content=raw, headers={"X-Read-Signature": _sign(raw)})
    assert r.status_code == 200 and r.json()["result"] == "ignored"


def test_webhook_answers_400_to_a_body_without_a_session(client, monkeypatch):
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    raw = json.dumps({"trigger": "meeting_end"}).encode()
    r = client.post("/webhooks/read-ai", content=raw, headers={"X-Read-Signature": _sign(raw)})
    assert r.status_code == 400


# ── Parsing and matching ──────────────────────────────────────────────────────

def test_parse_payload_normalises_people_times_and_transcript():
    p = readai.parse_payload(_payload())
    assert p["participants"][1] == {"name": "Jo", "email": "jo@acme.com"}
    assert p["owner_email"] == "sam@unstuck-agency.com"
    assert p["start_time"].startswith("2026-10-12T10:00:00")
    assert p["end_time"].startswith("2025-10-12")  # Unix milliseconds
    assert p["action_items"] == ["Send ICP doc"]
    # The blank block is dropped and indices stay contiguous.
    assert [(b["i"], b["speaker"], b["text"]) for b in p["transcript"]] == [
        (0, "Sam", "Hello there."), (1, "Jo", "We sell to CFOs.")]


@pytest.mark.parametrize("raw,expected", [
    ("https://www.Acme.com/about", "acme.com"), ("jo@acme.co.uk", "acme.co.uk"),
    ("acme.com", "acme.com"), ("not a domain", ""), ("", ""),
])
def test_normalise_domain(raw, expected):
    assert readai.normalise_domain(raw) == expected


def test_match_client_single_ambiguous_and_ignored_domains():
    clients = [{"id": CLIENT, "domains": ["acme.com"]},
               {"id": OTHER, "domains": ["globex.com", "gmail.com"]}]
    people = lambda *emails: [{"name": "", "email": e} for e in emails]
    assert readai.match_client(people("jo@eu.acme.com", "sam@unstuck-agency.com"), clients) == ("matched", [CLIENT])
    assert readai.match_client(people("jo@acme.com", "al@globex.com"), clients)[0] == "ambiguous"
    # A personal address never identifies a client, even one that lists it.
    assert readai.match_client(people("someone@gmail.com"), clients) == ("unmatched", [])
    assert readai.match_client(people("sam@unstuck-agency.com"), clients) == ("unmatched", [])


@pytest.mark.parametrize("title,expected", [
    ("Acme onboarding", True), ("Kick-off: Acme", True), ("Acme kickoff call", True),
    ("Acme on-boarding", True), ("Weekly check-in", False), ("Boarding pass", False),
])
def test_looks_like_onboarding(title, expected):
    assert readai.looks_like_onboarding(title) is expected


# ── Ingest ────────────────────────────────────────────────────────────────────

def _clients_route(fake_sb, rows=None):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(200, rows if rows is not None else [
        {"id": CLIENT, "name": "Acme", "domains": ["acme.com"]}]))


def test_ingest_files_a_matched_onboarding_meeting_and_starts_the_work(fake_sb, _setup):
    _clients_route(fake_sb)
    fake_sb.route("POST", "pulse_meetings", lambda call: FakeResponse(201, [{"id": MEETING, **call["json"]}]))
    out = mi.ingest(readai.parse_payload(_payload()))
    assert out == {"result": "matched"}
    row = fake_sb.calls_to("POST", "pulse_meetings")[0]["json"]
    assert row["client_id"] == CLIENT and row["agency_id"] == AGENCY
    assert row["is_onboarding"] is True and row["match_status"] == "matched"
    assert fake_sb.calls_to("POST", "pulse_meetings")[0]["headers"]["apikey"] == "test-service-key"
    assert _setup == [("process_meeting", (MEETING,), {"draft_notes": True})]


def test_ingest_drops_a_repeat_delivery(fake_sb, _setup):
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "request_id": "req-1", "client_id": CLIENT, "transcript": []}]))
    assert mi.ingest(readai.parse_payload(_payload())) == {"result": "duplicate"}
    assert not fake_sb.calls_to("POST", "pulse_meetings")
    assert not fake_sb.calls_to("PATCH", "pulse_meetings")
    assert _setup == []


def test_ingest_refreshes_a_repushed_meeting_but_keeps_its_client(fake_sb, _setup):
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "request_id": "old", "client_id": CLIENT, "is_onboarding": False,
         "transcript": []}]))
    assert mi.ingest(readai.parse_payload(_payload())) == {"result": "updated"}
    patch = fake_sb.calls_to("PATCH", "pulse_meetings")[0]["json"]
    assert "client_id" not in patch and patch["analysis_status"] == "pending"
    assert _setup[0][0] == "process_meeting"


def test_an_internal_meeting_is_ignored_not_sent_to_the_inbox(fake_sb, _setup):
    _clients_route(fake_sb)
    fake_sb.route("POST", "pulse_meetings", lambda call: FakeResponse(201, [{"id": MEETING}]))
    payload = _payload(participants=[{"name": "Sam", "email": "sam@unstuck-agency.com"}])
    assert mi.ingest(readai.parse_payload(payload)) == {"result": "ignored"}
    assert fake_sb.calls_to("POST", "pulse_meetings")[0]["json"]["client_id"] is None
    assert _setup == []


def test_a_missing_domains_column_still_keeps_the_meeting(fake_sb, _setup):
    fake_sb.route("GET", "clients", lambda call: FakeResponse(400, {
        "code": "42703", "message": "column clients.domains does not exist"}))
    fake_sb.route("POST", "pulse_meetings", lambda call: FakeResponse(201, [{"id": MEETING}]))
    assert mi.ingest(readai.parse_payload(_payload())) == {"result": "unmatched"}


# ── AI grounding ──────────────────────────────────────────────────────────────

def test_locate_quote_tolerates_a_neighbouring_block_and_rejects_inventions():
    assert mi.locate_quote(TRANSCRIPT, 1, "three of them converted") == 1
    assert mi.locate_quote(TRANSCRIPT, 1, "we need heads of department") == 2
    assert mi.locate_quote(TRANSCRIPT, 0, "the CEO loved every single lead") is None
    assert mi.locate_quote(TRANSCRIPT, 0, "ok") is None


def test_analyse_meeting_keeps_only_grounded_points(fake_sb, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "title": "Acme QBR", "transcript": TRANSCRIPT}]))
    monkeypatch.setattr(llm, "complete_json", lambda *a, **kw: {
        "about_campaigns": True, "summary": "Mixed.",
        "positives": [{"point": "Fintech meetings converted.", "block": 1, "quote": "three of them converted"}],
        "negatives": [{"point": "Healthcare leads too junior.", "block": 2, "quote": "mostly too junior"},
                      {"point": "Invented.", "block": 2, "quote": "they want to cancel the contract"}],
    })
    analysis = mi.analyse_meeting(MEETING)
    assert [p["point"] for p in analysis["positives"]] == ["Fintech meetings converted."]
    assert [p["point"] for p in analysis["negatives"]] == ["Healthcare leads too junior."]
    final = fake_sb.calls_to("PATCH", "pulse_meetings")[-1]["json"]
    assert final["analysis_status"] == "done"


def test_analyse_meeting_records_a_failure_without_openai(fake_sb):
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "transcript": TRANSCRIPT}]))
    assert mi.analyse_meeting(MEETING) is None
    assert fake_sb.calls_to("PATCH", "pulse_meetings")[-1]["json"]["analysis_status"] == "failed"


def test_overview_cites_only_findings_that_exist(fake_sb, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "title": "QBR", "start_time": "2026-10-01", "analysis_status": "done",
         "analysis": {"positives": [{"point": "Fintech converted", "block": 1, "quote": "three of them converted"}],
                      "negatives": [{"point": "Too junior", "block": 2, "quote": "mostly too junior"}]}}]))
    monkeypatch.setattr(llm, "complete_json", lambda *a, **kw: {
        "summary": "Going well with one problem.",
        "successes": [{"point": "Fintech works", "sources": ["M1.P1", "M9.P9"]}],
        "pain_points": [{"point": "Seniority", "sources": ["M1.N1"]},
                        {"point": "Made up", "sources": ["M7.N2"]}],
    })
    overview = mi.rebuild_overview(CLIENT)
    assert overview["successes"] == [{"point": "Fintech works", "refs": [
        {"meeting_id": MEETING, "block": 1, "quote": "three of them converted"}]}]
    assert [p["point"] for p in overview["pain_points"]] == ["Seniority"]
    saved = fake_sb.calls_to("POST", "pulse_meeting_insights")[-1]
    assert saved["json"]["meeting_ids"] == [MEETING]
    assert param_values(saved, "on_conflict") == ["agency_id,client_id"]


def test_notes_html_is_escaped_and_ordered():
    html = mi.notes_html([
        {"heading": "Goals and KPIs", "bullets": ["10 meetings a month"]},
        {"heading": "Offer and services", "bullets": ["<script>x</script>Payroll software"]},
        {"heading": "Made up", "bullets": ["dropped"]},
    ])
    assert html.index("Offer and services") < html.index("Goals and KPIs")
    assert "<script>" not in html and "dropped" not in html


# ── Onboarding notes ──────────────────────────────────────────────────────────

def test_a_draft_fills_untouched_notes(fake_sb):
    fake_sb.route("GET", "pulse_onboarding_notes", lambda call: FakeResponse(200, []))
    fake_sb.route("POST", "pulse_onboarding_notes", lambda call: FakeResponse(201, [call["json"]]))
    assert mi.file_draft(CLIENT, MEETING, "<p>Draft</p>") == "filled"
    body = fake_sb.calls_to("POST", "pulse_onboarding_notes")[0]["json"]
    assert body["body"] == "<p>Draft</p>" and body["source_meeting_id"] == MEETING


def test_a_draft_never_overwrites_notes_a_person_edited(fake_sb):
    fake_sb.route("GET", "pulse_onboarding_notes", lambda call: FakeResponse(200, [
        {"body": "<p>Mine</p>", "edited_by": "Sam", "version": 3}]))
    fake_sb.route("PATCH", "pulse_onboarding_notes", lambda call: FakeResponse(200, [call["json"]]))
    assert mi.file_draft(CLIENT, MEETING, "<p>Draft</p>") == "pending"
    call = fake_sb.calls_to("PATCH", "pulse_onboarding_notes")[0]
    assert "body" not in call["json"] and call["json"]["pending_draft"] == "<p>Draft</p>"
    assert param_values(call, "version") == ["eq.3"]


def test_saving_notes_sanitises_and_records_the_editor(client, fake_sb):
    fake_sb.route("GET", "pulse_onboarding_notes", lambda call: FakeResponse(200, [
        {"body": "", "version": 2}]))
    fake_sb.route("PATCH", "pulse_onboarding_notes", lambda call: FakeResponse(200, [call["json"]]))
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/onboarding-notes",
                    data={"body": "<p>ICP: CFOs</p><script>alert(1)</script>", "version": "2"})
    assert r.status_code == 200 and "Notes saved." in r.text
    sent = fake_sb.calls_to("PATCH", "pulse_onboarding_notes")[0]["json"]
    assert sent["body"] == "<p>ICP: CFOs</p>" and sent["edited_by"] and sent["version"] == 3


def test_a_stale_save_keeps_the_typed_text_and_says_why(client, fake_sb):
    fake_sb.route("GET", "pulse_onboarding_notes", lambda call: FakeResponse(200, [
        {"body": "<p>Theirs</p>", "version": 5, "edited_by": "Alex"}]))
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/onboarding-notes",
                    data={"body": "<p>My careful notes</p>", "version": "4"})
    assert r.status_code == 200
    assert "Someone else changed these notes" in r.text
    assert "&lt;p&gt;My careful notes&lt;/p&gt;" in r.text  # back in the editor's textarea
    assert not fake_sb.calls_to("PATCH", "pulse_onboarding_notes")


def test_applying_a_pending_draft_replaces_the_body(client, fake_sb):
    fake_sb.route("GET", "pulse_onboarding_notes", lambda call: FakeResponse(200, [
        {"body": "<p>Mine</p>", "version": 1, "edited_by": "Sam",
         "pending_draft": "<p>New</p>", "pending_meeting_id": MEETING}]))
    fake_sb.route("PATCH", "pulse_onboarding_notes", lambda call: FakeResponse(200, [call["json"]]))
    r = client.post(f"/api/outbound-pulse/clients/{CLIENT}/onboarding-notes/draft/apply",
                    data={"version": "1"})
    assert "Notes replaced with the new draft." in r.text
    sent = fake_sb.calls_to("PATCH", "pulse_onboarding_notes")[0]["json"]
    assert sent["body"] == "<p>New</p>" and sent["pending_draft"] == ""


# ── Panels ────────────────────────────────────────────────────────────────────

def test_panels_explain_a_missing_migration(client, fake_sb):
    missing = lambda call: FakeResponse(404, {"code": "PGRST205", "message": "Could not find the table"})
    fake_sb.route("GET", "pulse_onboarding_notes", missing)
    fake_sb.route("GET", "pulse_meetings", missing)
    for path in ("onboarding-notes", "read-ai"):
        r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/{path}")
        assert "outbound_pulse_readai.sql" in r.text


def test_read_ai_panel_says_when_read_ai_is_not_connected(client, fake_sb):
    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/read-ai")
    assert "Read AI isn't connected" in r.text


def test_read_ai_panel_points_to_client_domains_when_there_are_none(client, fake_sb, monkeypatch):
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    _clients_route(fake_sb, [{"id": CLIENT, "name": "Acme", "domains": []}])
    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/read-ai")
    assert "No Read AI meetings for this client yet" in r.text and "Add client domains" in r.text


def test_read_ai_panel_renders_the_overview_with_links_to_passages(client, fake_sb, monkeypatch):
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    _clients_route(fake_sb)
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "client_id": CLIENT, "title": "Acme QBR", "start_time": "2026-10-12T10:00:00+00:00",
         "report_url": "javascript:alert(1)", "is_onboarding": True, "analysis_status": "done",
         "analysis": {"summary": "Mixed.", "positives": [], "negatives": [
             {"point": "Healthcare leads too junior.", "block": 2, "quote": "mostly too junior"}]}}]))
    fake_sb.route("GET", "pulse_meeting_insights", lambda call: FakeResponse(200, [
        {"status": "done", "meeting_ids": [MEETING], "generated_at": "2026-10-13T09:00:00+00:00",
         "overview": {"summary": "Going well.", "successes": [],
                      "pain_points": [{"point": "Seniority is off.", "refs": [
                          {"meeting_id": MEETING, "block": 2, "quote": "mostly too junior"}]}]}}]))
    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/read-ai")
    assert "Seniority is off." in r.text and "Pain points" in r.text
    assert f"/api/outbound-pulse/read-ai/meetings/{MEETING}?block=2" in r.text
    assert "Acme QBR · 12 Oct" in r.text and "Updated 13 Oct 2026" in r.text
    assert "javascript:" not in r.text
    assert "A newer meeting isn't in this overview" not in r.text


def test_transcript_highlights_the_cited_block(client, fake_sb):
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "title": "Acme QBR", "transcript": TRANSCRIPT}]))
    r = client.get(f"/api/outbound-pulse/read-ai/meetings/{MEETING}?block=2")
    hit = r.text.split('id="tr-b2"', 1)[1].split(">", 1)[0]
    assert "data-tr-hit" in hit and r.text.count("data-tr-hit") == 1
    assert "0:20" in r.text  # block 2 starts 20s after the first


def test_inbox_is_empty_markup_when_nothing_waits(client, fake_sb):
    r = client.get("/api/outbound-pulse/read-ai/inbox")
    assert r.status_code == 200 and r.text == ""


def test_assigning_from_the_inbox_files_the_meeting_and_analyses_it(client, fake_sb, monkeypatch, _setup):
    monkeypatch.setattr(store, "list_clients", lambda active_only=False: [{"id": CLIENT, "name": "Acme"}])
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "client_id": None, "title": "Acme kickoff", "analysis_status": "pending"}]
        if "id" in dict(call["params"]) else []))
    r = client.post(f"/api/outbound-pulse/read-ai/meetings/{MEETING}/assign",
                    data={"client_id": CLIENT, "view": "inbox"})
    assert r.status_code == 200 and "Assigned" in r.text
    patch = fake_sb.calls_to("PATCH", "pulse_meetings")[0]["json"]
    assert patch["client_id"] == CLIENT and patch["match_status"] == "manual"
    assert patch["is_onboarding"] is True
    assert _setup[0] == ("process_meeting", (MEETING,), {"draft_notes": True})


def test_client_page_has_both_panels(client, fake_sb, monkeypatch):
    from tests.test_outbound_pulse import _detail_routes
    _detail_routes(fake_sb)
    r = client.get(f"/outbound-pulse/clients/{CLIENT}")
    assert r.status_code == 200
    assert "Onboarding notes" in r.text and "Read AI summary" in r.text
    assert f"/api/outbound-pulse/clients/{CLIENT}/onboarding-notes" in r.text
    assert 'id="ra-transcript"' in r.text


# ── Client Profiles: domains ──────────────────────────────────────────────────

_PROFILE_FORM = {"name": "Acme", "type": "client", "sender_emails": "robyn@acme-mail.com",
                 "domains": "https://www.Acme.com, jo@acme.co.uk, acme.com", "active": "true"}


def test_profile_save_normalises_domains(client, fake_sb):
    fake_sb.route("PATCH", "clients", lambda call: FakeResponse(204))
    r = client.patch(f"/api/client-profiles/{CLIENT}", data=_PROFILE_FORM)
    assert r.status_code == 200
    body = fake_sb.calls_to("PATCH", "clients")[0]["json"]
    assert body["domains"] == ["acme.com", "acme.co.uk"]


def test_profile_save_without_the_domains_column_keeps_every_other_field(client, fake_sb):
    def patch(call):
        if "domains" in call["json"]:
            return FakeResponse(400, {"code": "PGRST204", "message": "Could not find the 'domains' column"})
        return FakeResponse(204)
    fake_sb.route("PATCH", "clients", patch)
    client.patch(f"/api/client-profiles/{CLIENT}", data=_PROFILE_FORM)
    calls = fake_sb.calls_to("PATCH", "clients")
    assert len(calls) == 2
    assert calls[1]["json"]["sender_emails"] == ["robyn@acme-mail.com"]
    assert "domains" not in calls[1]["json"]


# ── Visibility: nothing a webhook delivers may vanish ─────────────────────────

def test_a_meeting_without_attendee_emails_waits_with_title_suggestions(fake_sb, _setup):
    _clients_route(fake_sb, [{"id": CLIENT, "name": "Insight Associates", "domains": ["insightassociates.co.uk"]},
                             {"id": OTHER, "name": "Acme", "domains": []}])
    fake_sb.route("POST", "pulse_meetings", lambda call: FakeResponse(201, [{"id": MEETING}]))
    payload = _payload(title="Insight Associates - kick off",
                       participants=[{"name": "Sam", "email": "sam@unstuck-agency.com"},
                                     {"name": "Priya Shah", "email": None}])
    assert mi.ingest(readai.parse_payload(payload)) == {"result": "unmatched"}
    row = fake_sb.calls_to("POST", "pulse_meetings")[0]["json"]
    assert row["client_id"] is None and row["candidate_client_ids"] == [CLIENT]


def test_is_internal_needs_every_attendee_to_be_ours():
    ours = {"name": "Sam", "email": "sam@unstuck-agency.com"}
    assert readai.is_internal([ours])
    assert not readai.is_internal([ours, {"name": "Priya", "email": ""}])
    assert not readai.is_internal([ours, {"name": "Jo", "email": "jo@gmail.com"}])
    assert not readai.is_internal([])


def test_title_candidates_match_whole_names_only():
    clients = [{"id": "a", "name": "Insight Associates"}, {"id": "b", "name": "ERI"},
               {"id": "c", "name": "Hopscotch"}]
    assert readai.title_candidates("Insight Associates / Unstuck kick-off", clients) == ["a"]
    assert readai.title_candidates("Merit review", clients) == []
    assert readai.title_candidates("ERI weekly", clients) == ["b"]


def test_inbox_says_when_nothing_has_arrived(client, fake_sb, monkeypatch):
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    r = client.get("/api/outbound-pulse/read-ai/inbox")
    assert "No meetings received from Read AI yet" in r.text


def test_inbox_lists_recent_meetings_including_ignored_ones(client, fake_sb, monkeypatch):
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    monkeypatch.setattr(store, "list_clients", lambda active_only=False: [{"id": CLIENT, "name": "Acme"}])
    rows = [
        {"id": MEETING, "client_id": None, "title": "Team stand-up", "match_status": "ignored",
         "received_at": "2026-10-27T10:00:00+00:00",
         "participants": [{"name": "Sam", "email": "sam@unstuck-agency.com"}]},
        {"id": OTHER, "client_id": CLIENT, "title": "Acme QBR", "match_status": "matched",
         "received_at": "2026-10-26T10:00:00+00:00",
         "participants": [{"name": "Jo", "email": "jo@acme.com"}]},
    ]
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(
        200, [] if "client_id" in dict(call["params"]) else rows))
    r = client.get("/api/outbound-pulse/read-ai/inbox")
    assert "Recently received (2)" in r.text
    assert "Team stand-up" in r.text and "Ignored" in r.text
    assert f"/api/outbound-pulse/read-ai/meetings/{MEETING}/assign" in r.text  # can be rescued
    assert f'href="/outbound-pulse/clients/{CLIENT}#overview"' in r.text and "Acme QBR" in r.text
    assert "only Unstuck or personal email addresses" in r.text


def test_inbox_shows_a_failure_instead_of_nothing(client, fake_sb, monkeypatch):
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(500, {"code": "XX000", "message": "boom"}))
    r = client.get("/api/outbound-pulse/read-ai/inbox")
    assert "Couldn't load Read AI meetings" in r.text


# ── OpenAI limits ─────────────────────────────────────────────────────────────

def _openai(monkeypatch, responses):
    """Serve `responses` in order to llm's POSTs; record sleeps instead of sleeping."""
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    calls, sleeps = [], []
    def post(url, **kw):
        calls.append(kw["json"])
        return responses.pop(0)
    monkeypatch.setattr(llm.requests, "post", post)
    monkeypatch.setattr(llm.time, "sleep", sleeps.append)
    return calls, sleeps


def _ok(content):
    return FakeResponse(200, {"choices": [{"message": {"content": json.dumps(content)},
                                           "finish_reason": "stop"}]})


def test_a_rate_limit_is_waited_out_then_retried(monkeypatch):
    limited = FakeResponse(429, {"error": {"code": "rate_limit_exceeded", "message":
        "Rate limit reached for gpt-4o on tokens per min (TPM): Limit 30000, Used 29000, "
        "Requested 5000. Please try again in 6.5s."}}, headers={"x-ratelimit-reset-tokens": "4s"})
    calls, sleeps = _openai(monkeypatch, [limited, _ok({"a": 1})])
    assert llm.complete_json("s", "u", schema={}, schema_name="x") == {"a": 1}
    assert len(calls) == 2 and sleeps == [7.5]


def test_an_account_without_credit_says_so_and_does_not_retry(monkeypatch):
    broke = FakeResponse(429, {"error": {"code": "insufficient_quota", "message":
        "You exceeded your current quota, please check your plan and billing details."}})
    calls, sleeps = _openai(monkeypatch, [broke])
    with pytest.raises(llm.OutOfCredit, match="out of credit"):
        llm.complete_json("s", "u", schema={}, schema_name="x")
    assert len(calls) == 1 and sleeps == []


def test_a_request_too_large_for_the_account_fails_fast(monkeypatch):
    big = FakeResponse(429, {"error": {"code": "rate_limit_exceeded", "message":
        "Request too large for gpt-4o on tokens per min (TPM): Limit 30000, Requested 41000."}})
    calls, _ = _openai(monkeypatch, [big])
    with pytest.raises(llm.RequestTooLarge):
        llm.complete_json("s", "u", schema={}, schema_name="x")
    assert len(calls) == 1


@pytest.mark.parametrize("raw,seconds", [("6s", 6), ("1m20.5s", 80.5), ("250ms", 0.25), ("", None)])
def test_reset_header_durations(raw, seconds):
    assert llm._duration_s(raw) == seconds


def test_analysis_sends_less_transcript_when_the_request_is_too_large(fake_sb, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    long_blocks = [{"i": i, "speaker": "Jo", "start_ms": i, "end_ms": i,
                    "text": f"Block {i}: " + "the fintech meetings went well " * 40} for i in range(120)]
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "title": "QBR", "transcript": long_blocks}]))
    sizes = []
    def complete(system, user, **kw):
        sizes.append(len(user))
        if len(sizes) == 1:
            raise llm.RequestTooLarge("too big")
        return {"about_campaigns": True, "summary": "ok", "positives": [], "negatives": []}
    monkeypatch.setattr(llm, "complete_json", complete)
    assert mi.analyse_meeting(MEETING)["truncated"] is True
    assert sizes[1] < sizes[0] <= mi.TRANSCRIPT_CHAR_BUDGET + 1000


def test_a_failed_analysis_is_shown_as_a_failure_not_as_no_findings(client, fake_sb, monkeypatch):
    monkeypatch.setenv("READ_AI_WEBHOOK_SECRET", SECRET)
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    _clients_route(fake_sb)
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "client_id": CLIENT, "title": "Onboarding", "analysis_status": "failed",
         "analysis_error": "OpenAI says this account is out of credit.", "analysis": {}}]))
    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/read-ai")
    assert "Couldn't analyse 1 meeting" in r.text and "out of credit" in r.text
    assert "No campaign findings yet" not in r.text


def test_a_failed_notes_draft_is_explained(client, fake_sb, monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    fake_sb.route("GET", "pulse_meetings", lambda call: FakeResponse(200, [
        {"id": MEETING, "client_id": CLIENT, "title": "Kick-off", "transcript": TRANSCRIPT,
         "is_onboarding": True}]))
    monkeypatch.setattr(llm, "complete_json", lambda *a, **kw: (_ for _ in ()).throw(
        llm.OutOfCredit("OpenAI says this account is out of credit.")))
    assert mi.draft_onboarding_notes(MEETING) == ""
    r = client.get(f"/api/outbound-pulse/clients/{CLIENT}/onboarding-notes")
    assert "Couldn't draft the onboarding notes" in r.text and "out of credit" in r.text
    mi._draft_errors.clear()


def test_refresh_retries_failures_and_drafts_missing_notes(fake_sb, monkeypatch):
    done = []
    monkeypatch.setattr(ms, "list_client_meetings", lambda client_id: [
        {"id": MEETING, "analysis_status": "failed", "is_onboarding": True}])
    monkeypatch.setattr(ms, "get_notes", lambda client_id: None)
    monkeypatch.setattr(mi, "analyse_meeting", lambda mid: done.append(("analyse", mid)))
    monkeypatch.setattr(mi, "draft_onboarding_notes", lambda mid: done.append(("draft", mid)) or "filled")
    monkeypatch.setattr(mi, "rebuild_overview", lambda cid: done.append(("overview", cid)))
    mi.refresh_client(CLIENT)
    assert done == [("analyse", MEETING), ("draft", MEETING), ("overview", CLIENT)]
