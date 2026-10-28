"""Read-only Copy Studio ingestion, separate from the reporting event stream.

Every output is an immutable record candidate. Personalized sent text belongs
in private ``sent_copy`` records; aggregate performance only references their
source keys. Nothing here approves examples or writes to an upstream service.
"""
from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timezone
import hashlib
import json
import os
import re
import time
from urllib.parse import quote

import requests

from app.utils import google_sheets
from app.utils.pulse import smartlead, readai
from app.utils.pulse.normalize import parse_ts
from app.utils.copy_studio.integrations import IntegrationUnavailable
from app.utils.copy_studio.performance import comparison, manual_variants, manual_variant_id
from app.utils.copy_studio.client_context import profile_metadata

PAGE_SIZE = 100
MAX_SEEN_ROWS = 100_000
_ID = re.compile(r"^[a-zA-Z0-9_-]{1,128}$")


def _id(value) -> str:
    value = str(value or "")
    if not _ID.fullmatch(value):
        raise ValueError("A valid source identifier is required.")
    return value


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _hash(value) -> str:
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def _agency() -> str:
    value = os.environ.get("COPY_STUDIO_AGENCY_ID") or os.environ.get("PULSE_AGENCY_ID")
    if not value:
        raise IntegrationUnavailable("Set COPY_STUDIO_AGENCY_ID before synchronising client sources.")
    return _id(value)


def _sb_get(table: str, params: dict) -> list[dict]:
    # New Studio source access never inherits the broad anonymous grants on
    # legacy tables and never creates an agency as a side effect of a read.
    base = os.environ.get("SUPABASE_URL", "").rstrip("/")
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "")
    if not base or not key:
        raise IntegrationUnavailable("Source sync needs SUPABASE_URL and the server SUPABASE_SERVICE_ROLE_KEY.")
    try:
        response = requests.get(f"{base}/rest/v1/{table}", params=params,
                                headers={"apikey": key, "Authorization": f"Bearer {key}"}, timeout=25)
    except requests.RequestException as exc:
        raise IntegrationUnavailable("The source database could not be reached.") from exc
    if response.status_code >= 400:
        raise IntegrationUnavailable(f"Source table {table} is unavailable (HTTP {response.status_code}). Check access and migrations.")
    try:
        rows = response.json()
    except ValueError as exc:
        raise IntegrationUnavailable("The source database returned invalid JSON.") from exc
    if not isinstance(rows, list):
        raise IntegrationUnavailable("The source database returned an unexpected result.")
    return [row for row in rows if isinstance(row, dict)]


def _client_scope(client_id: str) -> dict:
    return {"agency_id": f"eq.{_agency()}", "client_id": f"eq.{_id(client_id)}"}


def _verify_client(client_id: str) -> None:
    # Exact ownership, including when a caller passes a Copy Bank key. NULL
    # agency legacy records require explicit mapping before Studio can use them.
    rows = _sb_get("clients", {"agency_id": f"eq.{_agency()}", "id": f"eq.{_id(client_id)}",
                               "select": "id,agency_id", "limit": "1"})
    if not any(str(row.get("id")) == client_id and str(row.get("agency_id")) == _agency() for row in rows):
        raise IntegrationUnavailable("The client is not mapped to this agency.")


def client_profile(client_id: str) -> dict:
    """The canonical Client Profiles row, with safe, explicitly sourced metadata."""
    client_id = _id(client_id)
    rows = _sb_get("clients", {"id": f"eq.{client_id}", "agency_id": f"eq.{_agency()}",
                               "select": "*", "limit": "1"})
    row = next((row for row in rows if str(row.get("id")) == client_id and str(row.get("agency_id")) == _agency()), None)
    if row is None:
        raise IntegrationUnavailable("The client is not mapped to this agency's Client Profiles.")
    return profile_metadata(row)


def _profile_directory() -> tuple[list[dict], bool, list[str]]:
    """A complete directory is required before calling a domain match unique."""
    profiles, seen_ids, page_hashes = [], set(), set()
    for offset in range(0, 10_000, PAGE_SIZE):
        try:
            rows = _sb_get("clients", {"agency_id": f"eq.{_agency()}", "select": "id,agency_id,name,domains,updated_at",
                                       "order": "id.asc", "offset": str(offset), "limit": str(PAGE_SIZE)})
        except IntegrationUnavailable:
            return profiles, False, ["The full agency Client Profiles directory is unavailable; unassigned calls cannot be matched safely."]
        page_hash = _hash([row.get("id") for row in rows])
        if rows and page_hash in page_hashes:
            return profiles, False, ["Client Profiles pagination repeated; domain matching remains disabled."]
        page_hashes.add(page_hash)
        for row in rows:
            identifier = str(row.get("id") or "")
            if not identifier or str(row.get("agency_id")) != _agency() or identifier in seen_ids:
                return profiles, False, ["Client Profiles scope or pagination could not be verified; unassigned calls were not imported."]
            seen_ids.add(identifier)
            profiles.append(profile_metadata(row))
        if len(rows) < PAGE_SIZE:
            return profiles, True, []
    return profiles, False, ["Client Profiles directory reached the safety limit; unassigned calls were not imported."]


def _verify_campaign(client_id: str, source: str, campaign: dict) -> None:
    """Recheck persisted job mappings after a lease/restart, without changing them."""
    table = "pulse_campaigns" if source == "smartlead" else "campaigns"
    params = {"id": f"eq.{_id(campaign.get('id'))}", "client_id": f"eq.{client_id}", "select": "*", "limit": "1"}
    if source == "smartlead":
        params.update(agency_id=f"eq.{_agency()}", source_tool="eq.smartlead", channel="eq.email")
    rows = _sb_get(table, params)
    if not rows or str(rows[0].get("client_id")) != client_id:
        raise IntegrationUnavailable("The campaign's client mapping changed. Refresh the source list before retrying.")
    row = rows[0]
    if source == "smartlead":
        valid = str(row.get("agency_id")) == _agency() and str(row.get("external_campaign_id")) == str(campaign.get("external_campaign_id"))
    else:
        sheet = lambda value: google_sheets.extract_sheet_id(str(value.get("sheet_id") or value.get("sheet_url") or ""))
        valid = sheet(row) == sheet(campaign)
    if not valid:
        raise IntegrationUnavailable("The campaign's source mapping changed. Refresh the source list before retrying.")


def _item(kind: str, client_id: str, source_key: str, content: dict) -> dict:
    return {"kind": kind, "client_id": client_id, "source_key": source_key, "content": content}


def _result(items=(), *, complete=True, warnings=(), next_cursor=None, **coverage):
    return {"items": list(items), "complete": complete, "next_cursor": next_cursor,
            "warnings": list(warnings), "coverage": {"complete": complete, **coverage}}


def list_campaigns(client_id: str, source: str, *, offset=0, limit=100) -> dict:
    """Enumerate only explicit client + agency mappings; resume by offset."""
    client_id = _id(client_id)
    _verify_client(client_id)
    if source not in {"smartlead", "manual"}:
        raise ValueError("Only Smartlead and manual email campaigns are supported.")
    limit = min(250, max(1, int(limit)))
    offset = max(0, int(offset))
    table = "pulse_campaigns" if source == "smartlead" else "campaigns"
    # Legacy manual campaigns have no agency_id column. Their exact parent
    # client was checked above; ownership is inherited through that mapping.
    scope = _client_scope(client_id) if source == "smartlead" else {"client_id": f"eq.{client_id}"}
    params = {**scope, "select": "*", "order": "id.asc",
              "offset": str(offset), "limit": str(limit)}
    if source == "smartlead":
        params.update(source_tool="eq.smartlead", channel="eq.email")
    rows = _sb_get(table, params)
    # Defend in depth against incorrectly mocked or changed PostgREST filters.
    rows = [row for row in rows if str(row.get("client_id")) == client_id
            and (source == "manual" or str(row.get("agency_id")) == _agency())]
    out = []
    for row in rows:
        raw = row.get("raw_payload") if isinstance(row.get("raw_payload"), dict) else {}
        out.append({**{k: row.get(k) for k in (
            "id", "client_id", "agency_id", "external_campaign_id", "name", "campaign_name", "status",
            "sheet_id", "sheet_url", "created_at", "copy_territory", "copy_industry")}, "agency_id": _agency(),
            "source_created_at": raw.get("created_at") or raw.get("createdAt")})
    return {"campaigns": out, "complete": len(rows) < limit,
            "next_offset": offset + limit if len(rows) == limit else None,
            "scope": "explicitly_mapped_client_campaigns"}


def reconcile_smartlead(client_id: str) -> dict:
    """Compare live IDs with agency mappings, without revealing other clients.

    Imported Pulse campaigns and the provider's current list are not assumed
    to have equal cardinality. New/unmapped campaigns require a reviewed client
    mapping; missing IDs remain visible rather than silently disappearing.
    """
    client_id = _id(client_id)
    _verify_client(client_id)
    live = _rows(_smartlead_get("/campaigns"), keys=("data", "results", "campaigns", "items"))
    live_ids = {str(row.get("id") or row.get("campaign_id")) for row in live if row.get("id") or row.get("campaign_id")}
    if len(live_ids) > 100_000:
        raise IntegrationUnavailable("Smartlead campaign listing exceeded the reconciliation safety limit.")
    mapped_ids, own_ids = set(), set()
    for offset in range(0, 100_000, 1000):
        rows = _sb_get("pulse_campaigns", {"agency_id": f"eq.{_agency()}", "source_tool": "eq.smartlead",
            "channel": "eq.email", "select": "agency_id,client_id,external_campaign_id", "order": "id.asc",
            "limit": "1000", "offset": str(offset)})
        for row in rows:
            if str(row.get("agency_id")) != _agency() or not row.get("client_id") or not row.get("external_campaign_id"):
                continue
            identifier = str(row["external_campaign_id"])
            mapped_ids.add(identifier)
            if str(row["client_id"]) == client_id:
                own_ids.add(identifier)
        if len(rows) < 1000:
            return {"checked_at": _now(), "provider_campaign_count": len(live_ids),
                    "agency_mapped_campaign_count": len(mapped_ids), "client_mapped_campaign_count": len(own_ids),
                    "provider_without_agency_mapping_count": len(live_ids - mapped_ids),
                    "client_mappings_missing_from_provider": sorted(own_ids - live_ids),
                    "scope": "live campaign list versus explicit client mappings; deleted history is retained"}
    raise IntegrationUnavailable("Campaign mapping reconciliation reached its safety limit.")


def _audience_context(territory, industry, *, provenance: str) -> dict:
    """Use explicit source taxonomy, never guess audience from campaign names."""
    territory = str(territory or "").strip()
    industry = str(industry or "").strip()
    audience = {}
    if territory:
        audience["geography"] = territory
    if industry:
        audience["sector"] = industry
    return {"territory": territory, "industry": industry, "audience": audience,
            "context_provenance": provenance, "context_complete": bool(territory and industry)}


def _campaign_context(campaign: dict) -> dict:
    return _audience_context(campaign.get("copy_territory"), campaign.get("copy_industry"),
                             provenance="campaign_copy_mapping")


def sync_copy_bank(client_id: str) -> dict:
    """Client-specific entries only; no legacy Biz Dev fallback or auto-approval."""
    client_id = _id(client_id)
    _verify_client(client_id)
    prefix = f"__c__{client_id}__"
    escaped = prefix.replace("\\", "\\\\").replace("_", "\\_")
    items, offset, seen_keys = [], 0, set()

    def result(*, complete=True, warnings=()):
        # Immutable hashes match store.content_hash. Human review revisions can
        # retain a link to this exact imported version instead of assuming that
        # a mutable Copy Bank array position still contains the same wording.
        versions = [{"source_key": item["source_key"], "content_hash": _hash(item["content"])} for item in items]
        return _result(items, complete=complete, warnings=warnings, imported=len(items), source="copy_bank",
                       imported_source_keys=sorted(item["source_key"] for item in items),
                       imported_versions=versions, source_manifest_authoritative=complete,
                       approval_required=True, approval_status="not_inferred_from_source",
                       scope="selected_client_copy_bank")

    while True:
        try:
            rows = _sb_get("copy_bank_templates", {
                "key": f"like.{escaped}*", "select": "key,content", "order": "key.asc",
                "offset": str(offset), "limit": str(PAGE_SIZE)})
        except IntegrationUnavailable:
            if not offset:
                raise
            return result(complete=False, warnings=["Copy Bank pagination could not finish. Imported examples remain pending review; current-source coverage is partial."])
        for row in rows:
            key = str(row.get("key") or "")
            if not key.startswith(prefix) or not isinstance(row.get("content"), dict):
                continue
            if key in seen_keys:
                return result(complete=False, warnings=["Copy Bank pagination repeated an entry; current-source coverage is partial."])
            seen_keys.add(key)
            territory, _, industry = key[len(prefix):].partition("_")
            context = _audience_context(territory, industry, provenance="copy_bank_key")
            content = row["content"]
            for channel in ("email", "linkedin", "chaser", "flyout"):
                section = content.get(channel) or {}
                if not isinstance(section, dict):
                    continue
                subjects = [s.strip() for s in section.get("subjects", []) if isinstance(s, str) and s.strip()]
                bodies = section.get("steps", []) if channel in {"linkedin", "chaser"} else section.get("variations", [])
                steps, source_positions, unknown_types = [], [], []
                for index, body in enumerate(bodies):
                    if not isinstance(body, dict) or not str(body.get("body") or "").strip():
                        continue
                    kind = "follow_up" if channel == "chaser" else "initial"
                    if channel == "linkedin":
                        supplied = str(body.get("kind") or body.get("type") or "").strip().lower().replace("-", "_").replace(" ", "_")
                        kinds = {"connection": "connection", "connection_request": "connection", "connect": "connection",
                                 "initial": "initial", "initial_message": "initial", "message": "initial",
                                 "follow_up": "follow_up", "followup": "follow_up", "follow_up_message": "follow_up",
                                 "closing": "closing", "closing_message": "closing", "close": "closing"}
                        kind = kinds.get(supplied, "initial" if index == 0 else "follow_up")
                        if supplied and supplied not in kinds:
                            unknown_types.append(supplied)
                    steps.append({"kind": kind, "body": str(body["body"]).strip()})
                    source_positions.append(index + 1)
                if not steps:
                    continue
                # Email variations are alternative opening messages, not steps
                # of a sequence. Never concatenate A, B and C into one outreach
                # sequence. LinkedIn/chaser steps are an actual ordered sequence.
                alternatives = [([step], f":variation:{position}", position)
                                for position, step in zip(source_positions, steps)] if channel in {"email", "flyout"} else [(steps, "", None)]
                for example_steps, suffix, position in alternatives:
                    items.append(_item("example", client_id, f"copy_bank:{key}:{channel}{suffix}", {
                        "source": "copy_bank", "channel": "linkedin" if channel == "linkedin" else "email",
                        "copy_type": channel, "copy_bank_key": key, "source_position": position,
                        "subjects": subjects, "steps": example_steps,
                        "step_kinds": list(dict.fromkeys(step["kind"] for step in example_steps)), **context,
                        "source_url": "/copy-bank", "reusable": True, "sanitized": False,
                        "performance_proven": False, "performance": None,
                        "approval_required": True, "unrecognized_step_types": unknown_types,
                        "provenance_note": "Imported published Copy Bank content. Studio review is required before reuse; no performance status is inferred. Source positions are mutable, not historical attribution IDs.",
                    }))
        if len(rows) < PAGE_SIZE:
            return result()
        offset += PAGE_SIZE
        if offset >= 5000:
            return result(complete=False, warnings=["Copy Bank import reached the safety limit."])


def sync_onboarding(client_id: str) -> dict:
    """Import all relevant stored client calls, not only the onboarding flag.

    The historic function/job name remains compatible with saved jobs. All
    explicitly mapped client calls are eligible unless ignored/ambiguous; an
    unassigned call needs a unique participant-domain match across the COMPLETE
    agency directory. No title matching or source assignment writes occur.
    """
    client_id = _id(client_id)
    _verify_client(client_id)
    profiles, directory_complete, warnings = _profile_directory()
    if not any(profile['id'] == client_id for profile in profiles):
        directory_complete = False
        warnings.append('The selected client was not present in the complete domain directory; unassigned recovery is disabled.')
    items, seen, source_keys = [], set(), []
    counts = {'mapped_calls': 0, 'recovered_calls': 0, 'missing_transcripts': 0,
              'ambiguous_candidates': 0, 'ignored_candidates': 0,
              'wrong_scope_skipped': 0, 'unassigned_examined': 0, 'unassigned_unmatched': 0}
    listing_complete = {'mapped': True, 'unassigned': True}
    fields = ('id,client_id,agency_id,title,report_url,start_time,end_time,transcript,'
              'received_at,participants,match_status,is_onboarding')
    for mode in ('mapped', 'unassigned'):
        # Without a complete directory, even a single apparent match could be
        # shared with another unseen client. No unassigned source is imported.
        if mode == 'unassigned' and not directory_complete:
            listing_complete[mode] = False
            continue
        page_hashes = set()
        for offset in range(0, 10_000, PAGE_SIZE):
            params = {'agency_id': f'eq.{_agency()}',
                      'client_id': f'eq.{client_id}' if mode == 'mapped' else 'is.null',
                      'select': fields, 'order': 'id.asc', 'offset': str(offset), 'limit': str(PAGE_SIZE)}
            try:
                rows = _sb_get('pulse_meetings', params)
            except IntegrationUnavailable as exc:
                warnings.append('Stored call listing is partial. '+str(exc))
                listing_complete[mode] = False
                break
            page_hash = _hash([row.get('id') for row in rows])
            if rows and page_hash in page_hashes:
                warnings.append('Stored call pagination repeated; call coverage is partial.')
                listing_complete[mode] = False
                break
            page_hashes.add(page_hash)
            for row in rows:
                identifier = str(row.get('id') or '')
                if not identifier or str(row.get('agency_id')) != _agency():
                    counts['wrong_scope_skipped'] += 1
                    listing_complete[mode] = False
                    continue
                assigned = str(row.get('client_id') or '')
                # Ignore another client's explicit mapping even if participants
                # share this client's domain; client assignment wins.
                if mode == 'mapped' and assigned != client_id:
                    counts['wrong_scope_skipped'] += 1
                    continue
                if mode == 'unassigned' and assigned:
                    counts['wrong_scope_skipped'] += 1
                    continue
                if identifier in seen:
                    continue
                seen.add(identifier)
                status = str(row.get('match_status') or '').lower()
                participants = row.get('participants') if isinstance(row.get('participants'), list) else []
                match_status, candidates = readai.match_client(participants, profiles) if directory_complete else ('unmatched', [])
                relevant_candidate = assigned == client_id or client_id in candidates
                if mode == 'unassigned':
                    counts['unassigned_examined'] += 1
                if status == 'ignored':
                    counts['ignored_candidates'] += int(relevant_candidate)
                    continue
                if status == 'ambiguous' or (mode == 'unassigned' and match_status == 'ambiguous'):
                    counts['ambiguous_candidates'] += int(relevant_candidate)
                    continue
                if mode == 'unassigned' and (match_status != 'matched' or candidates != [client_id]):
                    counts['unassigned_unmatched'] += int(match_status == 'unmatched')
                    continue
                blocks = row.get('transcript')
                if not isinstance(blocks, list) or not blocks or not any(isinstance(b, dict) and str(b.get('text') or '').strip() for b in blocks):
                    counts['missing_transcripts'] += 1
                    continue
                if any(not isinstance(block, dict) for block in blocks):
                    counts['missing_transcripts'] += 1
                    warnings.append('A relevant stored call has malformed transcript blocks and needs source review.')
                    continue
                source_key = f'read_ai:{identifier}'
                source_keys.append(source_key)
                counts['mapped_calls' if mode == 'mapped' else 'recovered_calls'] += 1
                assignment = {'method': 'explicit_client_mapping' if mode == 'mapped' else 'unique_participant_domain',
                              'client_id': client_id, 'source_client_id': assigned or None,
                              'source_match_status': status, 'upstream_assignment_changed': False}
                if mode == 'unassigned':
                    selected = next(profile for profile in profiles if profile['id'] == client_id)
                    assignment.update(profile_domains=selected['domains'], full_agency_directory_verified=True)
                items.append(_item('transcript', client_id, source_key, {
                    'source': 'read_ai', 'meeting_id': identifier, 'title': row.get('title', ''),
                    'source_url': row.get('report_url') or '', 'transcript': blocks,
                    'start_time': row.get('start_time'), 'end_time': row.get('end_time'),
                    'source_received_at': row.get('received_at'), 'truncated': False,
                    'source_complete': False, 'stored_payload_complete': True,
                    'is_onboarding': bool(row.get('is_onboarding')), 'source_assignment': assignment,
                    'coverage': {'stored_blocks': len(blocks), 'imported_blocks': len(blocks),
                                 'upstream_call_completeness': 'unverified', 'scope': 'stored_webhook_meetings'},
                }))
            if len(rows) < PAGE_SIZE:
                break
        else:
            listing_complete[mode] = False
            warnings.append('Stored call import reached the safety limit; coverage remains partial.')
    authoritative = directory_complete and all(listing_complete.values())
    complete = authoritative and not counts['missing_transcripts'] and not counts['ambiguous_candidates']
    if counts['missing_transcripts']:
        warnings.append('Some relevant stored meetings have no usable transcript; brief coverage is incomplete.')
    if counts['ambiguous_candidates']:
        warnings.append('Some potentially relevant calls have ambiguous client assignments and were not imported.')
    limitations = ['Coverage is limited to Read AI webhook records stored in this dashboard; full Read AI account history is not verified.',
                   'Every stored transcript block is retained, but upstream full-call completeness needs reviewer confirmation.',
                   'Unmatched calls without a unique participant-domain match are not attributed by title.']
    return _result(items, complete=complete, warnings=list(dict.fromkeys(warnings)), imported=len(items),
                   source='read_ai', scope='stored_webhook_meetings', account_history_verified=False,
                   profile_directory_complete=directory_complete, stored_listing_complete=listing_complete,
                   profiles_missing_domains=sum(not profile['domains'] for profile in profiles),
                   selected_client_domains_available=any(profile['id'] == client_id and profile['domains'] for profile in profiles),
                   imported_source_keys=sorted(source_keys), source_manifest_authoritative=authoritative,
                   profile_fingerprint=next((profile['profile_fingerprint'] for profile in profiles if profile['id'] == client_id), None),
                   upstream_call_completeness='unverified', limitations=limitations, **counts)


def _smartlead_get(path: str, params=None, *, sleep=time.sleep) -> object:
    """Bounded GET retry with sanitized errors (API keys are query secrets)."""
    key = smartlead.api_key()
    if not key:
        raise IntegrationUnavailable("Smartlead is not configured on this server.")
    for attempt in range(4):
        try:
            response = requests.get(f"{smartlead.API_BASE}{path}", params={**(params or {}), "api_key": key}, timeout=30)
        except requests.RequestException as exc:
            if attempt == 3:
                raise IntegrationUnavailable("Smartlead could not be reached. The sync can be retried.") from exc
            sleep(min(8, 2**attempt))
            continue
        if response.status_code == 429 or response.status_code >= 500:
            if attempt == 3:
                raise IntegrationUnavailable("Smartlead is rate limited or temporarily unavailable. The sync can be retried.")
            raw = str(response.headers.get("Retry-After", ""))
            wait = min(30, float(raw)) if raw.replace(".", "", 1).isdigit() else min(8, 2**attempt)
            sleep(max(0.1, wait))
            continue
        if response.status_code >= 400:
            raise IntegrationUnavailable(f"Smartlead denied the read (HTTP {response.status_code}). Check its server credential and campaign access.")
        try:
            return response.json()
        except ValueError as exc:
            raise IntegrationUnavailable("Smartlead returned invalid JSON.") from exc
    raise IntegrationUnavailable("Smartlead is unavailable.")


def _rows(payload, keys=("data", "results", "statistics", "sequences", "items")):
    if isinstance(payload, list):
        return [row for row in payload if isinstance(row, dict)]
    if isinstance(payload, dict):
        for key in keys:
            if isinstance(payload.get(key), list):
                return [row for row in payload[key] if isinstance(row, dict)]
    raise IntegrationUnavailable("Smartlead returned an unrecognised source schema; coverage is not complete.")


def _sent_content(client_id, source, campaign_id, step, variant_id, subject, body, sequence_id=""):
    fingerprint = _hash([subject, body])
    sequence_id = str(sequence_id or ("manual" if source == "manual" else "unresolved"))
    key = f"{source}:{campaign_id}:sequence:{sequence_id}:step:{step}:variant:{variant_id}:copy:{fingerprint}"
    return _item("sent_copy", client_id, key, {
        "source": source, "channel": "email", "campaign_id": str(campaign_id),
        "step": step, "variant_id": str(variant_id), "sequence_id": str(sequence_id),
        "subject": subject, "body": body, "content_hash": fingerprint,
        "reusable": False, "sanitized": False, "privacy": "recipient_specific",
        "approval_required": True,
    })


def _number(value, default=0):
    try:
        return max(0, int(value))
    except (ValueError, TypeError):
        return default


def _manual(client_id: str, campaign: dict) -> dict:
    campaign_id = _id(campaign["id"])
    sheet_id = google_sheets.extract_sheet_id(str(campaign.get("sheet_id") or campaign.get("sheet_url") or ""))
    if not sheet_id:
        raise IntegrationUnavailable("This manual campaign has no linked Google Sheet.")
    try:
        records = _read_manual_records(sheet_id)
        # Use the same authoritative aggregation as Pulse with explicitly
        # read-only source records, never its broader write-scoped OAuth cache.
        stats = google_sheets.read_ab_stats(sheet_id, records=records)
        if records and "Send Status" not in records[0]:
            # Older sheets only support allocation denominators. Preserve the
            # reporting fallback, but never label it verified send exposure.
            for row in stats:
                row.pop("sent", None)
    except Exception as exc:
        raise IntegrationUnavailable("The campaign sheet is unavailable. Share it with the configured Google service account as Viewer, then retry.") from exc
    observations = defaultdict(lambda: {"sent": 0, "with_copy": 0, "versions": {}, "meeting_requests": 0,
                                        "unverified_outcome_rows": 0})
    copies = {}
    for row in records:
        label = str(row.get("A/B Variant") or "").strip()
        if not label or "@" not in str(row.get("Recipient Email") or ""):
            continue
        observation = observations[label]
        status = str(row.get("Lead Status") or "").strip().lower()
        if not google_sheets._is_sent(row.get("Send Status", "")):
            if status in {"lead", "interested", "reply", "unsubscribe", "meeting request"}:
                observation["unverified_outcome_rows"] += 1
            continue
        observation["sent"] += 1
        if status == "meeting request":
            observation["meeting_requests"] += 1
        subject, body = str(row.get("Subject Line") or "").strip(), str(row.get("Email Body") or "").strip()
        if not body:
            continue
        item = _sent_content(client_id, "manual", campaign_id, 1, manual_variant_id(label), subject, body)
        item["content"].update(_campaign_context(campaign))
        key = item["source_key"]
        copies[key] = item
        observation["with_copy"] += 1
        observation["versions"][key] = observation["versions"].get(key, 0) + 1
    report = comparison(manual_variants(stats, dict(observations)), source="manual", campaign_id=campaign_id,
        period={"kind": "all_time_status_snapshot", "outcome_dates_available": False},
        campaign_status=campaign.get("status", ""), campaign_name=campaign.get("campaign_name") or campaign.get("name") or "")
    report.update(_campaign_context(campaign))
    report.update(sequence_id="manual", source_url=f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
        attribution_note="Different personalized bodies or edited labels are unresolved at label level; source versions remain separate.")
    items = list(copies.values()) + [_item("performance", client_id, f"manual:{campaign_id}:step:1:performance", report)]
    return _result(items, source="manual", sent_rows=sum(o["sent"] for o in observations.values()),
                   copied_rows=sum(o["with_copy"] for o in observations.values()), unique_copy_versions=len(copies))


def _read_manual_records(sheet_id: str) -> list[dict]:
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2.service_account import Credentials
    sheet_id = _id(sheet_id)
    credentials = Credentials.from_service_account_info(google_sheets._decode_sa_json(),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    session = AuthorizedSession(credentials)
    base = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}"
    try:
        google_sheets._read_limiter.acquire()
        response = session.get(base, params={"fields": "sheets.properties"}, timeout=(5, 25))
        response.raise_for_status()
        sheets = [sheet.get("properties", {}) for sheet in response.json().get("sheets", [])]
        sheets.sort(key=lambda item: item.get("index", 0))
        if not sheets:
            raise IntegrationUnavailable("The manual campaign spreadsheet has no worksheet.")
        title = str(sheets[0].get("title") or "").replace("'", "''")
        # First worksheet is the existing merger's source contract. A complete
        # A:ZZ read avoids silently truncating rows based on stale grid counts.
        google_sheets._read_limiter.acquire()
        response = session.get(base + "/values/" + quote(f"'{title}'!A:ZZ", safe=""),
            params={"valueRenderOption": "UNFORMATTED_VALUE"}, timeout=(5, 25))
        response.raise_for_status()
        grid = response.json().get("values", [])
        if not grid:
            return []
        headers = [str(value).strip() for value in grid[0]]
        required = {"Recipient Email", "A/B Variant", "Lead Status"}
        if not required.issubset(headers) or any(headers.count(name) != 1 for name in required):
            raise IntegrationUnavailable("The sheet needs unique Recipient Email, A/B Variant and Lead Status columns.")
        return [{key: row[index] if index < len(row) else "" for index, key in enumerate(headers) if key}
                for row in grid[1:]]
    finally:
        session.close()


def _templates(client_id: str, external_id: str, sequences: list[dict], context: dict | None = None) -> list[dict]:
    items = []
    for sequence in sequences:
        step = _number(sequence.get("seq_number", sequence.get("sequence_number")))
        sequence_id = str(sequence.get("id") or sequence.get("email_campaign_seq_id") or "")
        variant_content = set()
        for variant in sequence.get("sequence_variants") or []:
            if not isinstance(variant, dict) or variant.get("id") in (None, ""):
                continue
            body = str(variant.get("email_body") or "").strip()
            if not body:
                continue
            subject = str(variant.get("subject") or "").strip()
            variant_content.add((subject, body))
            items.append(_item("example", client_id, f"smartlead:{external_id}:step:{step}:variant:{variant['id']}:template", {
                "source": "smartlead_template", "channel": "email", "campaign_id": external_id,
                "step": step, "sequence_id": sequence_id, "variant_id": str(variant["id"]),
                "variant_label": str(variant.get("variant_label") or ""),
                "subjects": [subject] if subject else [], "steps": [{"kind": "initial" if step == 1 else "follow_up", "body": body}],
                "source_updated_at": variant.get("updated_at"), "reusable": False, "sanitized": False,
                "performance_proven": False, "attribution": "current_template_only",
                "provenance_note": "Current editable template. It is not evidence of what was sent historically.",
            }))
        # Smartlead also returns a sequence's default wording directly on the
        # sequence, including steps with no sequence_variants. Keep this source
        # without inventing a provider variant ID or joining it to history.
        body = str(sequence.get("email_body") or "").strip()
        subject = str(sequence.get("subject") or "").strip()
        if body and (subject, body) not in variant_content:
            items.append(_item("example", client_id,
                f"smartlead:{external_id}:sequence:{sequence_id or 'unresolved'}:step:{step}:default:template", {
                    "source": "smartlead_template", "channel": "email", "campaign_id": external_id,
                    "step": step, "sequence_id": sequence_id, "variant_id": "", "template_kind": "sequence_default",
                    "subjects": [subject] if subject else [],
                    "steps": [{"kind": "initial" if step == 1 else "follow_up", "body": body}],
                    "source_updated_at": sequence.get("updated_at"), "reusable": False, "sanitized": False,
                    "performance_proven": False, "attribution": "current_template_only",
                    "provenance_note": "Current sequence default; no stable variant ID is supplied. Historical attribution is unresolved.",
                }))
    for item in items:
        item["content"].update(context or {})
        item["content"]["step_kinds"] = [step["kind"] for step in item["content"]["steps"]]
    return items


def _smartlead(client_id: str, campaign: dict, cursor: dict | None, max_pages: int) -> dict:
    external_id = _id(campaign.get("external_campaign_id"))
    cursor = dict(cursor or {})
    if cursor and (cursor.get("client_id") != client_id or cursor.get("campaign_id") != external_id):
        raise ValueError("The sync cursor belongs to another campaign or client.")
    # pulse_campaigns.created_at is its import time, not the campaign's first
    # send. Never use it to silently discard historical source rows.
    start = str(cursor.get("from_date") or str(campaign.get("source_created_at") or "2000-01-01")[:10])
    end = str(cursor.get("to_date") or date.today().isoformat())
    try:
        date.fromisoformat(start)
        date.fromisoformat(end)
    except ValueError as exc:
        raise ValueError("Campaign observation dates are invalid.") from exc
    offset = _number(cursor.get("offset"))
    seen = set(cursor.get("seen", []))
    observations = dict(cursor.get("observations", {}))
    page_hashes = list(cursor.get("page_hashes", []))
    unresolved = _number(cursor.get("unresolved_rows"))
    total = _number(cursor.get("rows_seen"))
    items, warnings, complete = [], [], False
    if not cursor:
        sequences = _rows(_smartlead_get(f"/campaigns/{external_id}/sequences"))
        items.extend(_templates(client_id, external_id, sequences, _campaign_context(campaign)))
    for _ in range(min(20, max(1, int(max_pages)))):
        payload = _smartlead_get(f"/campaigns/{external_id}/statistics", {"offset": offset, "limit": PAGE_SIZE})
        rows = _rows(payload)
        page_hash = _hash(rows)
        if rows and page_hash in page_hashes:
            warnings.append("Smartlead repeated a statistics page; attribution coverage remains partial.")
            break
        page_hashes.append(page_hash)
        for row in rows:
            # The hash identifies a source observation without retaining lead
            # names, emails or message bodies in the resumable cursor.
            identity = _hash([external_id, row.get("stats_id")]) if row.get("stats_id") is not None else _hash(row)
            if identity in seen:
                continue
            seen.add(identity)
            sent_time = parse_ts(row.get("sent_time") or row.get("sent_at"))
            if sent_time is None or not (start <= sent_time.date().isoformat() <= end):
                continue
            total += 1
            step = _number(row.get("sequence_number"))
            variant_id = str(row.get("seq_variant_id") or "")
            sequence_id = str(row.get("email_campaign_seq_id") or "")
            body, subject = str(row.get("email_message") or "").strip(), str(row.get("email_subject") or "").strip()
            if not step or not variant_id or not sequence_id or not body:
                unresolved += 1
                continue
            stable = f"{sequence_id}:{step}:{variant_id}"
            observation = observations.setdefault(stable, {"sequence_id": sequence_id, "step": step,
                "variant_id": variant_id, "sent": 0, "versions": {}})
            item = _sent_content(client_id, "smartlead", external_id, step, variant_id, subject, body, sequence_id)
            item["content"].update(_campaign_context(campaign))
            key = item["source_key"]
            observation["sent"] += 1
            observation["versions"][key] = observation["versions"].get(key, 0) + 1
            items.append(item)
        offset += len(rows)
        if len(rows) < PAGE_SIZE:
            complete = True
            break
        if len(seen) >= MAX_SEEN_ROWS:
            warnings.append("Campaign exceeded the source-row safety limit; coverage is partial and needs a narrower history import.")
            break
    # Exact duplicate versions are stored once, even when many sends used them.
    items = list({item["source_key"]: item for item in items}.values())
    state = {"client_id": client_id, "campaign_id": external_id, "offset": offset, "seen": sorted(seen),
             "observations": observations, "rows_seen": total, "unresolved_rows": unresolved,
             "from_date": start, "to_date": end, "page_hashes": page_hashes}
    if complete:
        analytics = _smartlead_get(f"/campaigns/{external_id}/sequence-analytics", {
            "start_date": f"{start}T00:00:00.000Z", "end_date": f"{end}T23:59:59.999Z", "time_zone": "UTC"})
        data = analytics.get("data") if isinstance(analytics, dict) else None
        sequences = data.get("sequences") if isinstance(data, dict) else None
        if not isinstance(sequences, list):
            raise IntegrationUnavailable("Smartlead analytics schema is unavailable; exact attribution has not completed.")
        for sequence in sequences:
            step = _number(sequence.get("seq_number"))
            sequence_id = str(sequence.get("id") or sequence.get("email_campaign_seq_id") or "")
            variants = []
            for raw in sequence.get("variants") or []:
                if not isinstance(raw, dict):
                    continue
                stats = raw.get("stats") or {}
                sent = _number(stats.get("sent_count"))
                if not sent:
                    continue
                variant_id = str(raw.get("id") or raw.get("seq_variant_id") or raw.get("variant_id") or "")
                # Never equate variant A from a template with an analytics A
                # without the stable variant ID, step and sequence identity.
                candidates = [o for o in observations.values() if o["step"] == step and o["variant_id"] == variant_id
                              and (not sequence_id or o["sequence_id"] == sequence_id)]
                observation = candidates[0] if len(candidates) == 1 and variant_id else {}
                versions = observation.get("versions", {})
                exact = len(versions) == 1 and observation.get("sent") == sent
                variants.append({
                    "variant": variant_id or f"unresolved:{raw.get('variant_label', '?')}",
                    "variant_id": variant_id, "variant_label": str(raw.get("variant_label") or ""),
                    "sequence_id": sequence_id or observation.get("sequence_id", ""),
                    "sent": sent, "allocated": sent, "positive": _number(stats.get("positive_reply_count")),
                    "reply": _number(stats.get("reply_count")), "unsubscribe": _number(stats.get("unsubscribed_count")),
                    "attribution": "exact" if exact else "unresolved", "denominator": "sent",
                    "copy_source_keys": sorted(versions), "observed_sent_rows": observation.get("sent", 0),
                    "observed_versions": [{"source_key": key, "sent": count} for key, count in sorted(versions.items())],
                })
            if not variants:
                continue
            report = comparison(variants, source="smartlead", campaign_id=external_id, step=step,
                period={"kind": "provider_window", "from": start, "to": end}, complete=True,
                campaign_name=campaign.get("name") or "", campaign_status=campaign.get("status") or "")
            report.update(_campaign_context(campaign))
            report.update(sequence_id=sequence_id or "unresolved", source_rows=total, unresolved_source_rows=unresolved,
                          attribution_note="Only complete stable-ID matches to one observed sent version are exact. Today's template is not assumed to be historical copy.")
            items.append(_item("performance", client_id, f"smartlead:{external_id}:sequence:{sequence_id or 'unresolved'}:step:{step}:performance:{start}:{end}", report))
    resumable = not complete and not warnings
    return _result(items, complete=complete, next_cursor=state if resumable else None, warnings=warnings,
                   source="smartlead", rows_seen=total, unresolved_rows=unresolved,
                   observed_variants=len(observations), requested_period={"from": start, "to": end})


def sync_campaign(client_id: str, source: str, campaign: dict, cursor=None, max_pages=5) -> dict:
    """One durable-job work unit. Continue ``next_cursor`` until complete.

    A fresh incremental pass starts at zero: providers expose mutable current
    outcomes, so advancing a permanent offset would miss later replies to old
    sends. Immutable version hashes make replay safe and preserve history.
    """
    client_id = _id(client_id)
    _verify_client(client_id)
    if str(campaign.get("client_id")) != client_id or str(campaign.get("agency_id")) != _agency():
        raise ValueError("Campaign ownership does not match the selected client and agency.")
    if source not in {"manual", "smartlead"}:
        raise IntegrationUnavailable("Only Smartlead and manual email results can be synchronised. Meet Alfred remains disabled.")
    _verify_campaign(client_id, source, campaign)
    if source == "manual":
        return _manual(client_id, campaign)
    if source == "smartlead":
        return _smartlead(client_id, campaign, cursor, max_pages)
    raise IntegrationUnavailable("Only Smartlead and manual email results can be synchronised. Meet Alfred remains disabled.")
