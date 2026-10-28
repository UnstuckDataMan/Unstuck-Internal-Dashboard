from copy import deepcopy

import pytest
import requests

from app.utils.copy_studio import integrations, sources
from app.utils import google_sheets


@pytest.fixture
def scoped(monkeypatch):
    monkeypatch.setenv("COPY_STUDIO_AGENCY_ID", "agency-1")
    monkeypatch.setattr(sources, "_verify_client", lambda client_id: None)
    monkeypatch.setattr(sources, "_verify_campaign", lambda client_id, source, campaign: None)
    return {"id": "campaign-1", "client_id": "client-1", "agency_id": "agency-1",
            "external_campaign_id": "123", "sheet_id": "sheet-1", "name": "Client campaign",
            "status": "ACTIVE", "source_created_at": "2026-01-01"}


def sent_row(number, *, variant="101", body="A useful message", sequence="10", step=1):
    return {"stats_id": number, "sequence_number": step, "seq_variant_id": variant,
            "email_campaign_seq_id": sequence, "sent_time": "2026-04-01T10:00:00Z",
            "email_subject": "Hello", "email_message": body,
            "lead_email": "private@example.invalid", "lead_name": "Private Recipient"}


def analytics(sent=2, variant="101", sequence="10"):
    return {"data": {"sequences": [{"id": sequence, "seq_number": 1, "variants": [
        {"id": variant, "variant_label": "A", "stats": {"sent_count": sent, "positive_reply_count": 1, "reply_count": 1}}
    ]}]}}


def test_no_network_for_disabled_meet_alfred(monkeypatch):
    monkeypatch.setenv("MEET_ALFRED_API_KEY", "set-but-must-be-ignored")
    monkeypatch.setattr(requests, "get", lambda *a, **k: pytest.fail("Unexpected network"))
    provider = integrations.MeetAlfredProvider()
    assert provider.capabilities()["status"] == "disabled"
    assert not provider.capabilities()["read_performance"]
    with pytest.raises(integrations.IntegrationUnavailable, match="disabled"):
        provider.sync()


def test_capabilities_do_not_claim_configuration_is_a_successful_sync(monkeypatch):
    monkeypatch.setenv("SMARTLEAD_API_KEY", "present")
    monkeypatch.setenv("SUPABASE_URL", "https://database.example.invalid")
    monkeypatch.setenv("SUPABASE_SERVICE_ROLE_KEY", "present")
    monkeypatch.setenv("COPY_STUDIO_AGENCY_ID", "agency-1")
    statuses = {row["id"]: row for row in integrations.capabilities()}
    assert statuses["smartlead"]["status"] == "configured_not_verified"
    assert not statuses["smartlead"]["live_verified"]
    assert statuses["smartlead"]["read_performance"]
    assert not statuses["copy_bank"]["read_performance"]
    assert all(not row["send"] for row in statuses.values())


def test_provider_key_without_scoped_source_database_is_not_ready(monkeypatch):
    monkeypatch.setenv("SMARTLEAD_API_KEY", "present")
    monkeypatch.setenv("GOOGLE_SHEETS_SA_JSON", "present")
    monkeypatch.delenv("SUPABASE_SERVICE_ROLE_KEY", raising=False)
    statuses = {row["id"]: row for row in integrations.capabilities()}
    assert not statuses["smartlead"]["configured"]
    assert not statuses["manual"]["configured"]
    assert not statuses["copy_bank"]["configured"]


def test_copy_bank_strict_client_channel_and_review_contract(scoped, monkeypatch):
    rows = [
        {"key": "__c__client-1__UK_Marketing", "content": {
            "email": {"subjects": ["Subject"], "variations": [{"body": "Email message"}]},
            "linkedin": {"steps": [{"body": "Connection message"}, {"body": ""}]},
            "chaser": {"steps": [{"body": "Short follow-up"}]}}},
        {"key": "__c__client-2__UK_Marketing", "content": {"linkedin": {"steps": [{"body": "Other client secret"}]}}},
        {"key": "UK_Marketing", "content": {"email": {"variations": [{"body": "Biz Dev fallback must not leak"}]}}},
        {"key": "__cb_profiles__", "content": [{"client_id": "client-2"}]},
    ]
    monkeypatch.setattr(sources, "_sb_get", lambda table, params: rows)
    result = sources.sync_copy_bank("client-1")
    assert result["complete"] and len(result["items"]) == 3
    assert all(item["content"]["approval_required"] for item in result["items"])
    linkedin = next(item["content"] for item in result["items"] if item["content"]["channel"] == "linkedin")
    assert linkedin["source"] == "copy_bank"
    assert not linkedin["performance_proven"]
    assert linkedin["steps"] == [{"kind": "initial", "body": "Connection message"}]
    assert "Other client secret" not in str(result)


def test_onboarding_imports_entire_tail_and_reports_unknown_upstream_coverage(scoped, monkeypatch):
    blocks = [{"i": i, "text": "Long transcript section " * 100, "start_ms": i*1000} for i in range(300)]
    blocks[-1]["text"] = "Critical exclusion in the last sentence."
    rows = [{"id": "call-1", "client_id": "client-1", "agency_id": "agency-1", "transcript": blocks},
            {"id": "other", "client_id": "client-2", "agency_id": "agency-1", "transcript": blocks}]
    monkeypatch.setattr(sources, "_sb_get", lambda table, params: rows)
    result = sources.sync_onboarding("client-1")
    assert len(result["items"]) == 1
    content = result["items"][0]["content"]
    assert len(content["transcript"]) == 300
    assert content["transcript"][-1]["text"].startswith("Critical exclusion")
    assert not content["truncated"]
    assert content["stored_payload_complete"] is True
    assert content["source_complete"] is False


def test_copy_bank_email_alternatives_are_not_concatenated_into_a_sequence(scoped, monkeypatch):
    monkeypatch.setattr(sources, "_sb_get", lambda table, params: [{
        "key": "__c__client-1__UK_Consulting", "content": {
            "email": {"subjects": ["A", "B"], "variations": [{"body": "First angle"}, {"body": ""}, {"body": "Different angle"}]},
            "linkedin": {"steps": [{"type": "connection_request", "body": "Connect"}, {"type": "initial_message", "body": "Initial"}, {"type": "follow-up", "body": "Follow up"}, {"type": "closing_message", "body": "Close"}]}
        }}])
    result = sources.sync_copy_bank("client-1")
    email = [item for item in result["items"] if item["content"]["channel"] == "email"]
    assert len(email) == 2
    assert [item["content"]["steps"] for item in email] == [[{"kind": "initial", "body": "First angle"}], [{"kind": "initial", "body": "Different angle"}]]
    assert [item["content"]["source_position"] for item in email] == [1, 3]
    assert email[0]["source_key"] != email[1]["source_key"]
    linkedin = next(item for item in result["items"] if item["content"]["channel"] == "linkedin")
    assert [step["kind"] for step in linkedin["content"]["steps"]] == ["connection", "initial", "follow_up", "closing"]


def test_copy_bank_context_manifest_matches_persisted_versions_without_auto_approval(scoped, monkeypatch, tmp_path):
    from app.utils.copy_studio.store import SQLiteStore
    rows = [{"key": "__c__client-1__UK_Management_Consulting", "content": {
        "email": {"subjects": ["Question"], "variations": [{"body": "Original client copy"}]}}}]
    monkeypatch.setattr(sources, "_sb_get", lambda table, params: rows)
    store = SQLiteStore(str(tmp_path / "sources.db"), agency_id="agency-1")
    first = sources.sync_copy_bank("client-1")
    item = first["items"][0]
    content = item["content"]
    assert content["audience"] == {"geography": "UK", "sector": "Management_Consulting"}
    assert content["step_kinds"] == ["initial"]
    assert content["context_provenance"] == "copy_bank_key"
    saved = store.save(item["kind"], item["client_id"], item["source_key"], content, "importer")
    assert saved["status"] != "approved"
    assert first["coverage"]["imported_versions"] == [{"source_key": saved["source_key"], "content_hash": saved["content_hash"]}]
    assert first["coverage"]["source_manifest_authoritative"]
    assert sources.sync_copy_bank("client-1")["coverage"]["imported_versions"] == first["coverage"]["imported_versions"]
    rows[0]["content"]["email"]["variations"][0]["body"] = "Changed client copy"
    changed = sources.sync_copy_bank("client-1")
    assert changed["coverage"]["imported_source_keys"] == first["coverage"]["imported_source_keys"]
    assert changed["coverage"]["imported_versions"] != first["coverage"]["imported_versions"]


def test_copy_bank_empty_manifest_records_removed_examples(scoped, monkeypatch):
    monkeypatch.setattr(sources, "_sb_get", lambda *args: [])
    result = sources.sync_copy_bank("client-1")
    assert result["complete"]
    assert result["coverage"]["source_manifest_authoritative"]
    assert result["coverage"]["imported_versions"] == []


@pytest.mark.parametrize("failure", ["repeated", "unavailable"])
def test_copy_bank_incomplete_page_keeps_pending_items_without_authoritative_manifest(scoped, monkeypatch, failure):
    monkeypatch.setattr(sources, "PAGE_SIZE", 1)
    rows = [{"key": "__c__client-1__UK_Software", "content": {"email": {"variations": [{"body": "Client wording"}]}}}]
    def read(table, params):
        if failure == "unavailable" and int(params["offset"]):
            raise integrations.IntegrationUnavailable("Unavailable")
        return rows
    monkeypatch.setattr(sources, "_sb_get", read)
    result = sources.sync_copy_bank("client-1")
    assert len(result["items"]) == 1
    assert not result["complete"] and result["warnings"]
    assert not result["coverage"]["source_manifest_authoritative"]


def test_sent_copy_context_uses_explicit_mapping_not_campaign_name(scoped, monkeypatch):
    monkeypatch.setattr(sources, "_read_manual_records", lambda sheet: [{
        "Recipient Email": "person@example.invalid", "A/B Variant": "S1/B1", "Send Status": "Sent",
        "Lead Status": "Lead", "Subject Line": "Hello", "Email Body": "A relevant opening"}])
    campaign = {**scoped, "name": "US Healthcare", "copy_territory": "UK", "copy_industry": "Software"}
    result = sources.sync_campaign("client-1", "manual", campaign)
    assert all(item["content"]["audience"] == {"geography": "UK", "sector": "Software"} for item in result["items"])
    missing = sources.sync_campaign("client-1", "manual", scoped)
    assert all(item["content"]["audience"] == {} for item in missing["items"])
    assert all(not item["content"]["context_complete"] for item in missing["items"])


def test_manual_current_statuses_and_actual_copy_are_campaign_scoped(scoped, monkeypatch):
    rows = [
        {"Recipient Email": "one@example.invalid", "A/B Variant": "S1/B1", "Send Status": "Sent", "Lead Status": "Lead", "Subject Line": "Hello", "Email Body": "Old wording"},
        {"Recipient Email": "two@example.invalid", "A/B Variant": "S1/B1", "Send Status": "Sent", "Lead Status": "Meeting Request", "Subject Line": "Hello", "Email Body": "Changed wording"},
        {"Recipient Email": "three@example.invalid", "A/B Variant": "S2/B2", "Send Status": "", "Lead Status": "", "Subject Line": "Hello", "Email Body": "Unsent wording"},
    ]
    monkeypatch.setattr(sources, "_read_manual_records", lambda sheet_id: rows)
    result = sources.sync_campaign("client-1", "manual", scoped)
    copies = [item for item in result["items"] if item["kind"] == "sent_copy"]
    assert len(copies) == 2
    assert all(not item["content"]["reusable"] for item in copies)
    report = next(item["content"] for item in result["items"] if item["kind"] == "performance")
    first = next(v for v in report["variants"] if v["variant"] == "S1/B1")
    assert first["sent"] == 2 and first["positive"] == 1
    assert first["attribution"] == "unresolved"
    assert first["meeting_requests"] == 1 and first["meeting_booked"] == 0
    assert report["period"]["kind"] == "all_time_status_snapshot"
    second_campaign = sources.sync_campaign("client-1", "manual", {**scoped, "id": "campaign-2"})
    assert not {i["source_key"] for i in copies} & {i["source_key"] for i in second_campaign["items"]}
    assert "one@example.invalid" not in str(result)


def test_manual_missing_copy_bank_mapping_does_not_block_stats(scoped, monkeypatch):
    monkeypatch.setattr(sources, "_read_manual_records", lambda sheet_id: [{
        "Recipient Email": "person@example.invalid", "A/B Variant": "S1/B1", "Send Status": "Sent", "Lead Status": "Interested"}])
    result = sources.sync_campaign("client-1", "manual", scoped)
    report = result["items"][0]["content"]
    assert report["variants"][0]["positive"] == 1
    assert report["variants"][0]["attribution"] == "unresolved"


def test_manual_legacy_allocation_is_explicit(scoped, monkeypatch):
    monkeypatch.setattr(sources, "_read_manual_records", lambda sheet_id: [{
        "Recipient Email": "person@example.invalid", "A/B Variant": "S1/B1", "Lead Status": "Interested"}])
    result = sources.sync_campaign("client-1", "manual", scoped)
    row = result["items"][0]["content"]["variants"][0]
    assert row["sent"] == 1 and row["denominator"] == "allocated_legacy"
    assert not row["reference_eligible"]


def test_manual_sheet_failure_has_actionable_safe_message(scoped, monkeypatch):
    def fail(*args):
        raise requests.RequestException("token secret sensitive-url")
    monkeypatch.setattr(sources, "_read_manual_records", fail)
    with pytest.raises(integrations.IntegrationUnavailable, match="Viewer") as exc:
        sources.sync_campaign("client-1", "manual", scoped)
    assert "secret" not in str(exc.value)


def test_wrong_client_campaign_or_cursor_cannot_be_synchronised(scoped):
    with pytest.raises(ValueError, match="ownership"):
        sources.sync_campaign("client-2", "smartlead", scoped)
    with pytest.raises(ValueError, match="cursor"):
        sources.sync_campaign("client-1", "smartlead", scoped, cursor={"client_id": "client-2", "campaign_id": "123"})


def test_smartlead_resumes_history_deduplicates_copy_and_joins_stable_ids(scoped, monkeypatch):
    monkeypatch.setattr(sources, "PAGE_SIZE", 2)
    calls = []
    def read(path, params=None):
        calls.append((path, params))
        if path.endswith("/sequences"):
            return [{"id": 10, "seq_number": 1, "sequence_variants": [{"id": 101, "variant_label": "A", "subject": "New subject", "email_body": "Today's edited template"}]}]
        if path.endswith("/statistics"):
            return [sent_row(1), sent_row(2)] if params["offset"] == 0 else []
        return analytics()
    monkeypatch.setattr(sources, "_smartlead_get", read)
    first = sources.sync_campaign("client-1", "smartlead", scoped, max_pages=1)
    assert first["complete"] is False and first["next_cursor"]["offset"] == 2
    assert len([i for i in first["items"] if i["kind"] == "sent_copy"]) == 1
    assert not any(i["kind"] == "performance" for i in first["items"])
    assert "A useful message" not in str(first["next_cursor"])
    assert "private@example.invalid" not in str(first)
    second = sources.sync_campaign("client-1", "smartlead", scoped, cursor=first["next_cursor"], max_pages=1)
    assert second["complete"]
    report = second["items"][0]["content"]
    assert report["variants"][0]["attribution"] == "exact"
    assert report["variants"][0]["variant_id"] == "101"
    assert report["variants"][0]["observed_sent_rows"] == 2
    assert report["period"] == {"kind": "provider_window", "from": "2026-01-01", "to": sources.date.today().isoformat()}


def test_smartlead_current_sequence_default_is_preserved_without_inventing_variant_id():
    records = sources._templates("client-1", "123", [
        {"id": 10, "seq_number": 1, "subject": "Subject", "email_body": "Initial", "sequence_variants": [
            {"id": 101, "subject": "Subject", "email_body": "Initial", "variant_label": "A"}]},
        {"id": 20, "seq_number": 2, "subject": "Follow up", "email_body": "Default follow up", "sequence_variants": []},
    ])
    assert len(records) == 2  # duplicate default/variant wording is not repeated
    default = next(item for item in records if item["content"].get("template_kind") == "sequence_default")
    assert default["content"]["sequence_id"] == "20"
    assert default["content"]["variant_id"] == ""
    assert default["content"]["steps"] == [{"kind": "follow_up", "body": "Default follow up"}]
    assert default["content"]["attribution"] == "current_template_only"
    assert not default["content"]["reusable"] and not default["content"]["performance_proven"]


def test_smartlead_baseline_analytics_without_variant_id_remains_unresolved(scoped, monkeypatch):
    def read(path, params=None):
        if path.endswith("/sequences"):
            return []
        if path.endswith("/statistics"):
            return {"data": [sent_row(1)], "limit": 100, "offset": 0, "total_stats": 1}
        result = analytics(sent=1, variant=None)
        result["data"]["sequences"][0]["variants"][0]["is_baseline"] = True
        return result
    monkeypatch.setattr(sources, "_smartlead_get", read)
    result = sources.sync_campaign("client-1", "smartlead", scoped)
    report = next(item["content"] for item in result["items"] if item["kind"] == "performance")
    assert report["variants"][0]["attribution"] == "unresolved"
    assert report["variants"][0]["copy_source_keys"] == []
    assert not report["variants"][0]["reference_eligible"]


def test_smartlead_mismatched_id_does_not_join_by_label(scoped, monkeypatch):
    def read(path, params=None):
        if path.endswith("/sequences"):
            return []
        if path.endswith("/statistics"):
            return [sent_row(1, variant="999")]
        return analytics(sent=1, variant="101")
    monkeypatch.setattr(sources, "_smartlead_get", read)
    result = sources.sync_campaign("client-1", "smartlead", scoped)
    report = next(item["content"] for item in result["items"] if item["kind"] == "performance")
    assert report["variants"][0]["attribution"] == "unresolved"
    assert report["variants"][0]["copy_source_keys"] == []


def test_smartlead_changed_sent_text_is_retained_without_false_version_attribution(scoped, monkeypatch):
    def read(path, params=None):
        if path.endswith("/sequences"):
            return []
        if path.endswith("/statistics"):
            return [sent_row(1, body="Old"), sent_row(2, body="New")]
        return analytics()
    monkeypatch.setattr(sources, "_smartlead_get", read)
    result = sources.sync_campaign("client-1", "smartlead", scoped)
    assert len([i for i in result["items"] if i["kind"] == "sent_copy"]) == 2
    report = next(item["content"] for item in result["items"] if item["kind"] == "performance")
    assert report["variants"][0]["attribution"] == "unresolved"
    assert len(report["variants"][0]["observed_versions"]) == 2


def test_same_step_variant_text_in_different_sequence_ids_stays_distinct(scoped, monkeypatch):
    def read(path, params=None):
        if path.endswith("/sequences"):
            return []
        if path.endswith("/statistics"):
            return [sent_row(1, sequence="10"), sent_row(2, sequence="20")]
        return {"data": {"sequences": [analytics(sent=1, sequence="10")["data"]["sequences"][0],
                                        analytics(sent=1, sequence="20")["data"]["sequences"][0]]}}
    monkeypatch.setattr(sources, "_smartlead_get", read)
    result = sources.sync_campaign("client-1", "smartlead", scoped)
    copies = [item for item in result["items"] if item["kind"] == "sent_copy"]
    reports = [item for item in result["items"] if item["kind"] == "performance"]
    assert len({item["source_key"] for item in copies}) == 2
    assert len({item["source_key"] for item in reports}) == 2
    for report in reports:
        content = report["content"]
        variant = content["variants"][0]
        assert variant["attribution"] == "exact"
        assert f":sequence:{content['sequence_id']}:" in variant["copy_source_keys"][0]


def test_statistics_window_uses_utc_dates(scoped, monkeypatch):
    row = sent_row(1)
    row["sent_time"] = "2026-04-01T23:15:00-08:00"  # April 2 in provider's UTC window
    def read(path, params=None):
        if path.endswith("/sequences"):
            return []
        if path.endswith("/statistics"):
            return [row]
        return analytics(sent=1)
    monkeypatch.setattr(sources, "_smartlead_get", read)
    result = sources.sync_campaign("client-1", "smartlead", {**scoped, "source_created_at": "2026-04-02"})
    assert result["coverage"]["rows_seen"] == 1


def test_manual_unsent_outcomes_do_not_get_attributed_to_verified_sent_copy(scoped, monkeypatch):
    rows = []
    for variant, leads in (("S1/B1", 6), ("S2/B2", 1)):
        for i in range(30):
            rows.append({"Recipient Email": f"{variant}-{i}@example.invalid", "A/B Variant": variant,
                         "Send Status": "Sent", "Lead Status": "Lead" if i < leads else "",
                         "Subject Line": "Hello", "Email Body": f"Copy for {variant}"})
    rows.append({"Recipient Email": "unsent@example.invalid", "A/B Variant": "S1/B1", "Send Status": "",
                 "Lead Status": "Lead", "Subject Line": "Hello", "Email Body": "Never sent"})
    monkeypatch.setattr(sources, "_read_manual_records", lambda sheet_id: rows)
    result = sources.sync_campaign("client-1", "manual", scoped)
    report = next(item["content"] for item in result["items"] if item["kind"] == "performance")
    leader = report["variants"][0]
    assert leader["positive"] == 7  # preserve reporting's native aggregation
    assert leader["sent"] == 30 and leader["unverified_outcome_rows"] == 1
    assert leader["attribution"] == "unresolved" and not leader["reference_eligible"]
    assert any("without verified sends" in note for note in leader["limitations"])
    sent = next(item["content"] for item in result["items"] if item["kind"] == "sent_copy" and item["content"]["variant_id"] == leader["variant_id"])
    assert sent["sequence_id"] == leader["sequence_id"] == "manual"


def test_smartlead_duplicate_statistics_do_not_inflate_sends(scoped, monkeypatch):
    def read(path, params=None):
        if path.endswith("/sequences"):
            return []
        if path.endswith("/statistics"):
            return [sent_row(1), sent_row(1)]
        return analytics(sent=1)
    monkeypatch.setattr(sources, "_smartlead_get", read)
    result = sources.sync_campaign("client-1", "smartlead", scoped)
    assert result["coverage"]["rows_seen"] == 1


def test_smartlead_repeated_page_is_partial_not_a_success(scoped, monkeypatch):
    monkeypatch.setattr(sources, "PAGE_SIZE", 1)
    monkeypatch.setattr(sources, "_smartlead_get", lambda path, params=None: [] if path.endswith("/sequences") else [sent_row(1)])
    result = sources.sync_campaign("client-1", "smartlead", scoped, max_pages=3)
    assert not result["complete"]
    assert result["warnings"] and result["next_cursor"] is None
    assert not any(item["kind"] == "performance" for item in result["items"])


def test_smartlead_history_does_not_start_at_dashboard_import_time(scoped, monkeypatch):
    campaign = {**scoped, "source_created_at": None, "created_at": "2026-10-28"}
    monkeypatch.setattr(sources, "_smartlead_get", lambda path, params=None: analytics(sent=1) if path.endswith("sequence-analytics") else [sent_row(1)] if path.endswith("statistics") else [])
    result = sources.sync_campaign("client-1", "smartlead", campaign)
    assert result["coverage"]["rows_seen"] == 1
    assert result["coverage"]["requested_period"]["from"] == "2000-01-01"


def test_smartlead_get_retries_rate_limit_and_never_logs_credential(monkeypatch):
    monkeypatch.setattr(sources.smartlead, "api_key", lambda: "private-key")
    class Response:
        def __init__(self, status):
            self.status_code = status
            self.headers = {"Retry-After": "2"}
        def json(self):
            return []
    responses = iter([Response(429), Response(503), Response(200)])
    monkeypatch.setattr(requests, "get", lambda *a, **k: next(responses))
    sleeps = []
    assert sources._smartlead_get("/campaigns/123/sequences", sleep=sleeps.append) == []
    assert sleeps == [2, 2]
    monkeypatch.setattr(requests, "get", lambda *a, **k: Response(403))
    with pytest.raises(integrations.IntegrationUnavailable) as exc:
        sources._smartlead_get("/campaigns/123/sequences")
    assert "private-key" not in str(exc.value)


def test_client_mapping_is_verified_without_null_agency_fallback(monkeypatch):
    monkeypatch.setenv("COPY_STUDIO_AGENCY_ID", "agency-1")
    captured = []
    def read(table, params):
        captured.append(params)
        return [{"id": "client-1", "agency_id": None}]
    monkeypatch.setattr(sources, "_sb_get", read)
    with pytest.raises(integrations.IntegrationUnavailable, match="not mapped"):
        sources._verify_client("client-1")
    assert captured[0]["agency_id"] == "eq.agency-1"
    assert "or" not in captured[0]


def test_manual_campaigns_inherit_verified_client_agency(scoped, monkeypatch):
    captured = []
    def read(table, params):
        captured.append((table, params))
        return [{"id": "manual-1", "client_id": "client-1"}, {"id": "wrong", "client_id": "client-2"}]
    monkeypatch.setattr(sources, "_sb_get", read)
    result = sources.list_campaigns("client-1", "manual")
    assert len(result["campaigns"]) == 1
    assert result["campaigns"][0]["agency_id"] == "agency-1"
    assert captured[0][0] == "campaigns"
    assert "agency_id" not in captured[0][1]


def test_manual_record_reader_requests_readonly_scopes(monkeypatch):
    from google.oauth2.service_account import Credentials
    from google.auth.transport.requests import AuthorizedSession
    import google.auth.transport.requests as transport
    captured = {}
    def credentials(info, scopes):
        captured["scopes"] = scopes
        return object()
    monkeypatch.setattr(Credentials, "from_service_account_info", credentials)
    monkeypatch.setattr(google_sheets, "_decode_sa_json", lambda: {})
    monkeypatch.setattr(google_sheets._read_limiter, "acquire", lambda: None)
    class Response:
        def __init__(self, data): self.data = data
        def json(self): return self.data
        def raise_for_status(self): pass
    class Session:
        def __init__(self, creds): pass
        def get(self, url, **kwargs):
            if "/values/" in url:
                return Response({"values": [["Recipient Email", "A/B Variant", "Send Status", "Lead Status", "Email Body"], ["x@example.invalid", "S1/B1", "Sent", "Lead", "Private copy"]]})
            return Response({"sheets": [{"properties": {"title": "Outreach", "index": 0}}]})
        def close(self): captured["closed"] = True
    monkeypatch.setattr(transport, "AuthorizedSession", Session)
    rows = sources._read_manual_records("sheet-1")
    assert rows[0]["Email Body"] == "Private copy"
    assert captured["scopes"] == ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    assert captured["closed"]


def test_changed_campaign_mapping_cannot_resume_old_client_data(monkeypatch):
    monkeypatch.setenv("COPY_STUDIO_AGENCY_ID", "agency-1")
    monkeypatch.setattr(sources, "_sb_get", lambda table, params: [{"id": "campaign-1", "client_id": "client-2", "agency_id": "agency-1", "external_campaign_id": "123"}])
    with pytest.raises(integrations.IntegrationUnavailable, match="mapping changed"):
        sources._verify_campaign("client-1", "smartlead", {"id": "campaign-1", "external_campaign_id": "123"})


def test_campaign_reconciliation_retains_missing_ids_without_other_client_names(scoped, monkeypatch):
    monkeypatch.setattr(sources, "_smartlead_get", lambda path: [{"id": 123}, {"id": 456}, {"id": 999}])
    monkeypatch.setattr(sources, "_sb_get", lambda table, params: [
        {"agency_id": "agency-1", "client_id": "client-1", "external_campaign_id": "123"},
        {"agency_id": "agency-1", "client_id": "client-1", "external_campaign_id": "deleted"},
        {"agency_id": "agency-1", "client_id": "client-2", "external_campaign_id": "456"},
        {"agency_id": "agency-1", "client_id": None, "external_campaign_id": "999"},
    ])
    result = sources.reconcile_smartlead("client-1")
    assert result["provider_campaign_count"] == 3
    assert result["agency_mapped_campaign_count"] == 3
    assert result["client_mappings_missing_from_provider"] == ["deleted"]
    assert result["provider_without_agency_mapping_count"] == 1
    assert "client-2" not in str(result)


def profile(identifier="client-1", domain="example.com", **extra):
    return {"id": identifier, "agency_id": "agency-1", "name": identifier,
            "domains": [domain], "updated_at": "2026-10-28T00:00:00Z", **extra}


def call(identifier, client="client-1", email="person@example.com", **extra):
    return {"id": identifier, "client_id": client, "agency_id": "agency-1", "match_status": "matched" if client else "unmatched",
            "participants": [{"email": email}], "transcript": [{"i": 0, "text": "Verified source conversation.", "start_ms": 1000}],
            "is_onboarding": False, **extra}


def source_database(monkeypatch, profiles, meetings):
    queries = []
    def read(table, params):
        queries.append((table, params))
        if table == "clients":
            rows = profiles
        else:
            assert "is_onboarding" not in params  # all client meetings, not just onboarding
            client_filter = params.get("client_id")
            rows = [row for row in meetings if (not row.get("client_id") if client_filter == "is.null"
                    else str(row.get("client_id")) == client_filter.removeprefix("eq."))]
        offset, limit = int(params.get("offset", 0)), int(params.get("limit", 100))
        return deepcopy(rows[offset:offset+limit])
    monkeypatch.setattr(sources, "_sb_get", read)
    return queries


def test_canonical_profile_domain_identifies_website_without_sending_inbox_guess():
    metadata = sources.profile_metadata(profile(domains=["https://www.Example.com/about", "mail.example.com"], sender_emails=["someone@other-business.com"]))
    assert metadata["domains"] == ["example.com", "mail.example.com"]
    assert metadata["website_url"] == "https://example.com"
    assert metadata["website_source"] == "client_profile_domain"
    assert metadata["website_inferred"] and not metadata["website_verified"]
    assert metadata["profile_url"] == "/client-profiles"
    assert "sender_emails" not in metadata
    assert metadata["profile_fingerprint"] == sources.profile_metadata(profile(domains=["example.com", "mail.example.com"]))["profile_fingerprint"]
    missing = sources.profile_metadata(profile(domains=[], sender_emails=["ceo@realcompany.com"]))
    assert missing["website_url"] is None and missing["website_source"] == "missing"
    changed = sources.profile_metadata(profile(domains=["new-company.com"]))
    assert changed["profile_fingerprint"] != metadata["profile_fingerprint"]


def test_profile_website_validation_does_not_create_unsafe_links():
    for value in ("javascript:alert(1)", "http://127.0.0.1/private", "https://user:secret@example.com", "https://localhost/", "https://[broken"):
        assert sources.profile_metadata(profile(domains=[], website_url=value))["website_url"] is None
    metadata = sources.profile_metadata(profile(domains=["gmail.com", "person@example.com", "127.0.0.1", "example.local"]))
    assert metadata["website_url"] is None
    explicit = sources.profile_metadata(profile(website_url="https://example.com/company"))
    assert explicit["website_url"] == "https://example.com/company"
    assert explicit["website_source"] == "client_profile_website" and not explicit["website_verified"]


def test_client_profile_requires_explicit_matching_agency(monkeypatch):
    monkeypatch.setenv("COPY_STUDIO_AGENCY_ID", "agency-1")
    monkeypatch.setattr(sources, "_sb_get", lambda table, params: [profile(agency_id="agency-2")])
    with pytest.raises(integrations.IntegrationUnavailable, match="not mapped"):
        sources.client_profile("client-1")
    monkeypatch.setattr(sources, "_sb_get", lambda table, params: [profile()])
    assert sources.client_profile("client-1")["website_url"] == "https://example.com"


def test_every_mapped_client_call_is_imported_with_assignment_provenance(scoped, monkeypatch):
    profiles = [profile(), profile("client-2", "different.com")]
    meetings = [call("onboarding", is_onboarding=True), call("later-strategy", is_onboarding=False),
                call("ignored", match_status="ignored"), call("ambiguous", match_status="ambiguous"),
                call("wrong-client", client="client-2")]
    queries = source_database(monkeypatch, profiles, meetings)
    result = sources.sync_onboarding("client-1")
    assert {item["source_key"] for item in result["items"]} == {"read_ai:onboarding", "read_ai:later-strategy"}
    assert all(item["content"]["source_assignment"]["method"] == "explicit_client_mapping" for item in result["items"])
    assert result["coverage"]["mapped_calls"] == 2
    assert result["coverage"]["ignored_candidates"] == 1
    assert result["coverage"]["ambiguous_candidates"] == 1
    assert not result["complete"]
    assert result["coverage"]["source_manifest_authoritative"]
    assert result["coverage"]["profile_fingerprint"] == sources.profile_metadata(profiles[0])["profile_fingerprint"]
    assert not result["coverage"]["account_history_verified"]


def test_unique_unassigned_domain_matching_uses_complete_directory_and_no_titles(scoped, monkeypatch):
    profiles = [profile(), profile("client-2", "different.com")]
    meetings = [call("domain-match", client=None),
                call("title-only", client=None, email="personal@gmail.com", title="client-1 onboarding"),
                call("two-clients", client=None, participants=[{"email": "one@example.com"}, {"email": "two@different.com"}]),
                call("previously-ambiguous", client=None, match_status="ambiguous"),
                call("ignored", client=None, match_status="ignored"),
                call("other-client-mapped", client="client-2"),
                call("foreign-agency", client=None, agency_id="agency-2")]
    source_database(monkeypatch, profiles, meetings)
    result = sources.sync_onboarding("client-1")
    assert [item["source_key"] for item in result["items"]] == ["read_ai:domain-match"]
    assignment = result["items"][0]["content"]["source_assignment"]
    assert assignment["method"] == "unique_participant_domain"
    assert assignment["source_client_id"] is None and not assignment["upstream_assignment_changed"]
    assert assignment["full_agency_directory_verified"]
    assert result["coverage"]["recovered_calls"] == 1
    assert result["coverage"]["ambiguous_candidates"] == 2
    assert result["coverage"]["unassigned_unmatched"] == 1
    assert result["coverage"]["wrong_scope_skipped"] == 1
    assert not result["complete"]


def test_full_directory_pagination_prevents_false_unique_match(scoped, monkeypatch):
    monkeypatch.setattr(sources, "PAGE_SIZE", 1)
    source_database(monkeypatch, [profile(), profile("client-2", "example.com")], [call("shared-domain", client=None)])
    result = sources.sync_onboarding("client-1")
    assert result["items"] == []
    assert result["coverage"]["profile_directory_complete"]
    assert result["coverage"]["ambiguous_candidates"] == 1


def test_incomplete_directory_disables_recovery_but_keeps_mapped_calls(scoped, monkeypatch):
    monkeypatch.setattr(sources, "PAGE_SIZE", 1)
    def read(table, params):
        offset = int(params.get("offset", 0))
        if table == "clients":
            if offset:
                raise integrations.IntegrationUnavailable("directory page unavailable")
            return [profile()]
        assert params["client_id"] != "is.null"  # cannot infer uniqueness from a partial directory
        return [call("mapped")] if not offset else []
    monkeypatch.setattr(sources, "_sb_get", read)
    result = sources.sync_onboarding("client-1")
    assert len(result["items"]) == 1
    assert not result["complete"]
    assert not result["coverage"]["source_manifest_authoritative"]
    assert not result["coverage"]["profile_directory_complete"]


def test_missing_transcript_and_remapped_call_are_absent_from_authoritative_manifest(scoped, monkeypatch):
    profiles = [profile(), profile("client-2", "different.com")]
    meetings = [call("current"), call("missing", transcript=[]), call("moved", client="client-2")]
    source_database(monkeypatch, profiles, meetings)
    result = sources.sync_onboarding("client-1")
    assert result["coverage"]["imported_source_keys"] == ["read_ai:current"]
    assert result["coverage"]["source_manifest_authoritative"]
    assert result["coverage"]["missing_transcripts"] == 1
    assert not result["complete"]


def test_relevant_calls_are_paginated_without_truncating_stored_transcripts(scoped, monkeypatch):
    monkeypatch.setattr(sources, "PAGE_SIZE", 2)
    queries = source_database(monkeypatch, [profile()], [call(f"call-{i}") for i in range(5)])
    result = sources.sync_onboarding("client-1")
    assert result["complete"] and result["coverage"]["mapped_calls"] == 5
    assert len(result["coverage"]["imported_source_keys"]) == 5
    offsets = [params["offset"] for table, params in queries if table == "pulse_meetings" and params["client_id"] == "eq.client-1"]
    assert offsets == ["0", "2", "4"]
