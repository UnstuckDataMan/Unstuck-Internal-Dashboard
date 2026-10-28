"""Application use cases for Copy Studio; all callers carry an authorised client.

Imports are pending records, not approved claims. Publishing is an explicit,
reviewer-only action delegated to the atomic Copy Bank bridge.
The durable worker performs slow reads/model calls; HTTP handlers only enqueue.
"""
from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime, timezone

import requests

from .store import StoreError, StoreNotReady, get_store

GLOBAL_CLIENT = '__agency__'
MAX_RECORDS = 1000


def digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(',', ':')).encode()).hexdigest()


def actor(user):
    return str(user.get('email') or '').strip().lower()


def clients(store=None):
    """Strict agency scope. Legacy rows with an unassigned agency are not admitted."""
    store = store or get_store()
    if os.environ.get('COPY_STUDIO_SQLITE_PATH'):
        return [{'id': r['content']['id'], 'name': r['content']['name']}
                for r in store.list('__workspace__', kind='client', limit=1000)
                if r['content'].get('id') and r['content'].get('name')]
    url = os.environ.get('SUPABASE_URL', '').rstrip('/')
    key = os.environ.get('SUPABASE_SERVICE_ROLE_KEY', '')
    agency = os.environ.get('COPY_STUDIO_AGENCY_ID') or os.environ.get('PULSE_AGENCY_ID')
    if not url or not key or not agency:
        raise StoreNotReady('Copy Studio needs a Supabase service credential and an explicit agency ID.')
    rows = []
    try:
        for offset in range(0, 10000, 1000):
            response = requests.get(url+'/rest/v1/clients', headers={'apikey': key, 'Authorization': 'Bearer '+key},
                                    params={'agency_id': 'eq.'+agency, 'select': 'id,name', 'order': 'id.asc',
                                            'offset': offset, 'limit': 1000}, timeout=(5, 20))
            response.raise_for_status()
            page = response.json()
            if not isinstance(page, list):
                raise StoreNotReady('The client directory returned an invalid response.')
            rows.extend({'id': str(c['id']), 'name': str(c.get('name') or c['id'])} for c in page)
            if len(page) < 1000:
                return sorted(rows, key=lambda c: c['name'].casefold())
    except requests.RequestException as exc:
        raise StoreNotReady('Cannot read the agency client directory. Check the service credential and agency mapping.') from exc
    raise StoreNotReady('Client directory exceeds the supported page limit; narrow the agency scope.')


def latest_approved(records, kind):
    matches = [r for r in records if r['kind'] == kind and r['status'] == 'approved']
    return max(matches, key=lambda r: (r.get('created_at', ''), r['id']), default=None)


def source_summary(record):
    content = record['content']
    return {k: record.get(k) for k in ('id', 'kind', 'status', 'source_key', 'content_hash', 'created_at')} | {
        'content': {k: content[k] for k in ('title', 'source', 'source_url', 'channel', 'coverage', 'warnings', 'updated_at',
                                          'source_complete', 'source_completeness_review', 'source_assignment', 'is_onboarding') if k in content}}


def client_profile(client_id, store=None):
    from . import sources
    store = store or get_store()
    if os.environ.get('COPY_STUDIO_SQLITE_PATH'):
        rows = store.iter_records('__workspace__', kind='client')
        row = next((r['content'] for r in rows if str(r['content'].get('id')) == client_id), None)
        if row is None:
            raise LookupError('Client profile not found.')
        return sources.profile_metadata(row)
    return sources.client_profile(client_id)


def context_report(store, client_id):
    return next((r for r in store.iter_records(client_id, kind='sync_report')
                 if r['content'].get('source') == 'onboarding'), None)


def copy_bank_report(store, client_id):
    return next((r for r in store.iter_records(client_id, kind='sync_report')
                 if r['content'].get('source') == 'copy_bank'), None)


def current_examples(store, client_id, *, approved_only=True):
    """Use current approved versions, never a superseded approval or removed copy."""
    report = copy_bank_report(store, client_id)
    coverage = report['content'].get('coverage', {}) if report else {}
    versions = coverage.get('imported_versions')
    manifest = {r['source_key']: r['content_hash'] for r in versions} if isinstance(versions, list) else None
    seen, approved = set(), []
    for record in store.iter_records(client_id, kind='example'):
        if record['source_key'] in seen:
            continue
        seen.add(record['source_key'])
        if approved_only and record['status'] != 'approved':
            continue
        if record['content'].get('source') == 'copy_bank':
            if coverage.get('complete') is False:
                continue
            if manifest is not None:
                expected = manifest.get(record['source_key'])
                ancestor, visited = record, set()
                while ancestor and ancestor['content_hash'] != expected:
                    if ancestor['id'] in visited or len(visited) >= 100:
                        ancestor = None
                        break
                    visited.add(ancestor['id'])
                    ancestor = store.get(ancestor['parent_id'], client_id) if ancestor.get('parent_id') else None
                    if ancestor and (ancestor['kind'] != 'example' or ancestor['source_key'] != record['source_key']):
                        ancestor = None
                if not expected or not ancestor:
                    continue
        approved.append(record)
    return approved


def current_transcripts(store, client_id):
    """Retain audit history, but only use currently attributed conversations."""
    report = context_report(store, client_id)
    coverage = report['content'].get('coverage', {}) if report else {}
    keys = coverage.get('imported_source_keys')
    allowed = set(keys) if isinstance(keys, list) else None
    unique = {}
    for record in store.iter_records(client_id, kind='transcript'):
        if allowed is not None and record['source_key'] not in allowed:
            continue
        unique.setdefault(record['source_key'], record)
    return list(unique.values())


def brief_context_error(store, client_id, brief):
    if store.context_refresh_active(client_id):
        return 'Client context is refreshing. Wait for the source check to finish before generating or approving copy.'
    report = context_report(store, client_id)
    if report and report['content'].get('coverage', {}).get('profile_fingerprint'):
        current_profile = client_profile(client_id, store)
        if current_profile.get('profile_fingerprint') != report['content']['coverage']['profile_fingerprint']:
            return 'The client profile has changed. Refresh the client conversations and review the brief before using this copy.'
    if report and report['content'].get('coverage', {}).get('complete') is False:
        return 'The latest client conversation sync is incomplete. Review its source coverage and refresh before generating copy.'
    if not report or 'imported_source_keys' not in report['content'].get('coverage', {}):
        return ''  # Existing reviewed briefs remain readable before their first context refresh.
    current = {r['id'] for r in current_transcripts(store, client_id)}
    reviewed = {r.get('source_id') for r in brief['content'].get('coverage', {}).get('sources', [])}
    if current != reviewed:
        return 'Client conversations have changed. Build and review an updated brief before generating copy.'
    return ''


def context_refresh_needed(profile, report, *, compare_profile=True, max_age=900):
    if not report or report['content'].get('status') != 'complete':
        return True
    content = report['content']
    if compare_profile and content.get('coverage', {}).get('profile_fingerprint') != profile.get('profile_fingerprint'):
        return True
    try:
        synced = datetime.fromisoformat(content.get('completed_at', '').replace('Z', '+00:00'))
        age = (datetime.now(timezone.utc) - synced).total_seconds()
        return age < 0 or age >= max_age
    except (ValueError, TypeError):
        return True


def campaign_refresh_needed(store, client_id):
    reports = list(store.iter_records(client_id, kind='sync_report'))
    return any(context_refresh_needed({}, next((r for r in reports if r['content'].get('source') == source), None),
        compare_profile=False, max_age=3600) for source in ('smartlead', 'manual'))


def prepare_context(client_id, user):
    """Opening a client refreshes conversations and Copy Bank, without a model call."""
    store = get_store()
    profile = client_profile(client_id, store)
    for job in store.list_jobs(client_id):
        if (job['kind'] == 'sync' and job['status'] in ('queued', 'running')
                and job['payload'].get('source') in ('all', 'context')):
            return job
    bucket = int(datetime.now(timezone.utc).timestamp()) // 900
    fingerprint = profile.get('profile_fingerprint') or digest(profile)
    selected = 'all' if campaign_refresh_needed(store, client_id) else 'context'
    store.save('sync_subscription', client_id, 'subscription', {'enabled': True}, actor(user))
    return store.enqueue(client_id, 'sync', {'source': selected, 'automatic_context': True},
                         actor(user), f'context:{selected}:{bucket}:{fingerprint}')


def client_view(client_id):
    from .performance import refresh_evidence
    from .proofreading import infer_locale
    from .integrations import IntegrationUnavailable
    store = get_store()
    profile, profile_error = None, ''
    try:
        profile = client_profile(client_id, store)
    except IntegrationUnavailable as exc:
        profile_error = str(exc)
    context = context_report(store, client_id)
    # Usage/copy history can be large. Bound each displayed collection separately
    # so model calls never evict the client's brief or reference cards from view.
    collections = {kind: store.list(client_id, kind=kind, limit=200)
                   for kind in ('brief', 'example', 'transcript', 'sync_report', 'sent_copy', 'draft', 'performance')}
    records = [r for rows in collections.values() for r in rows]
    brief = latest_approved(store.iter_records(client_id, kind='brief', status='approved'), 'brief')
    pending_briefs = [r for r in records if r['kind'] == 'brief' and r['status'] != 'approved']
    examples = current_examples(store, client_id, approved_only=False)
    eligible_examples = current_examples(store, client_id)
    transcripts = [r for r in records if r['kind'] == 'transcript']
    coverage = (brief or {}).get('content', {}).get('coverage', {})
    return {
        'profile': profile, 'profile_error': profile_error,
        'locale_suggestion': infer_locale(brief, current_transcripts(store, client_id)),
        'context_sync_needed': bool(profile and (context_refresh_needed(profile, context)
            or context_refresh_needed(profile, copy_bank_report(store, client_id), compare_profile=False)
            or campaign_refresh_needed(store, client_id))),
        'brief_context_error': brief_context_error(store, client_id, brief) if brief else '',
        'brief': brief, 'pending_briefs': pending_briefs,
        'sources': [source_summary(r) for r in records if r['kind'] in ('transcript', 'sync_report', 'sent_copy')],
        'examples': examples, 'eligible_example_ids': [r['id'] for r in eligible_examples],
        'drafts': [r for r in records if r['kind'] == 'draft'],
        'performance': [dict(r, content=refresh_evidence(r['content'])) for r in records if r['kind'] == 'performance'],
        'methodology': store.list(GLOBAL_CLIENT, kind='methodology', limit=100),
        'jobs': store.list_jobs(client_id),
        'completeness': {
            'approved_brief': bool(brief), 'onboarding_sources': len(transcripts),
            'transcript_complete': bool(coverage.get('complete')),
            'approved_email_examples': sum(r['status'] == 'approved' and r['content'].get('channel') == 'email' for r in examples),
            'approved_linkedin_examples': sum(r['status'] == 'approved' and r['content'].get('channel') == 'linkedin' for r in examples),
            'missing_fields': (brief or {}).get('content', {}).get('missing_fields', []),
            'truncated': any(len(rows) == 200 for rows in collections.values()),
        },
    }


def enqueue_sync(client_id, source, user, request_key):
    store = get_store()
    # An explicit refresh opts this client into the worker's ongoing refreshes.
    store.save('sync_subscription', client_id, 'subscription', {'enabled': True}, actor(user))
    return store.enqueue(client_id, 'sync', {'source': source}, actor(user), 'sync:'+request_key)


def enqueue_brief(client_id, user, request_key):
    store = get_store()
    sources = current_transcripts(store, client_id)
    if not sources:
        raise ValueError('Load the client’s Read AI conversations before building a brief.')
    # Latest source revision only; all older immutable versions remain available.
    unique = {}
    for source in sorted(sources, key=lambda r: r.get('created_at', '')):
        unique[source['source_key']] = source['id']
    return store.enqueue(client_id, 'extract_brief', {'source_ids': list(unique.values())}, actor(user), 'brief:'+request_key)


def enqueue_generation(client_id, request, user, request_key):
    from .proofreading import infer_locale
    store = get_store()
    records = current_examples(store, client_id)
    eligible_ids = {r['id'] for r in records}
    brief = latest_approved(store.iter_records(client_id, kind='brief', status='approved'), 'brief')
    if not brief:
        raise ValueError('Review and approve the client brief before generating copy.')
    stale = brief_context_error(store, client_id, brief)
    if stale:
        raise ValueError(stale)
    if not brief['content'].get('coverage', {}).get('complete'):
        raise ValueError('The approved brief has incomplete source coverage. Finish the source review first.')
    request = dict(request)
    if not request.get('locale'):
        request['locale'] = infer_locale(brief, current_transcripts(store, client_id)).get('locale')
    selected = []
    for identifier in request.get('example_ids', []):
        example = store.get(identifier, client_id)
        if not example or example['id'] not in eligible_ids:
            raise ValueError('Select current approved examples belonging to this client. Changed Copy Bank entries need a new review.')
        selected.append(example['id'])
    if not selected:
        selected = [r['id'] for r in records if r['kind'] == 'example' and r['status'] == 'approved'
                    and r['content'].get('channel') == request['channel']]
    if request['mode'] == 'assemble' and not selected:
        raise ValueError('Approve a channel-specific Copy Bank example before assembling copy.')
    if request['channel'] == 'linkedin':
        for identifier in selected:
            example = store.get(identifier, client_id)
            if example['content'].get('source') != 'copy_bank' or example['content'].get('channel') != 'linkedin':
                raise ValueError('LinkedIn references must be approved LinkedIn Copy Bank examples.')
        if not selected:
            raise ValueError('Approve LinkedIn Copy Bank examples before drafting LinkedIn copy.')
    from .writer import retrieve_examples, effective_limits
    effective_limits(request['channel'], request.get('limits'))
    selected = [r['id'] for r in retrieve_examples(client_id, request['channel'],
                [store.get(i, client_id) for i in selected], request.get('audience'),
                performance=list(store.iter_records(client_id, kind='performance')))]
    if request['mode'] == 'assemble' and not selected:
        raise ValueError('The selected references are not eligible for assembly.')
    methodology = latest_approved(store.iter_records(GLOBAL_CLIENT, kind='methodology', status='approved'), 'methodology')
    payload = {'request': request, 'brief_id': brief['id'], 'example_ids': selected,
               'methodology_id': methodology['id'] if methodology else None}
    return store.enqueue(client_id, 'generate', payload, actor(user), 'generate:'+request_key)


def validate_brief(content, client_id, store):
    from .briefs import _source_complete, CATEGORIES, find_contradictions
    if not isinstance(content.get('facts'), list) or not content['facts']:
        raise ValueError('A brief needs evidence-backed facts.')
    if not isinstance(content.get('constraints', []), list):
        raise ValueError('Brief constraints must be a list.')
    ids = set()
    for fact in content['facts']:
        if not isinstance(fact, dict) or not isinstance(fact.get('id'), str) or not fact['id'].strip() or fact['id'] in ids:
            raise ValueError('Every brief fact needs its own unique fact ID.')
        ids.add(fact['id'])
        if fact.get('category') not in CATEGORIES or not isinstance(fact.get('text'), str) or not fact['text'].strip():
            raise ValueError('Every brief fact needs a valid category and statement.')
        if fact.get('decision', 'include') not in ('include', 'exclude'):
            raise ValueError('Choose whether each fact is included or excluded.')
        if not isinstance(fact.get('evidence'), list) or not fact['evidence'] or any(not isinstance(e, dict) for e in fact['evidence']):
            raise ValueError('Every fact needs structured source citations.')
    if find_contradictions(content['facts']):
        raise ValueError('Conflicting facts are still included. Exclude outdated statements or correct them against their evidence before approval.')
    if not content.get('coverage', {}).get('complete'):
        raise ValueError('Incomplete transcript coverage must be resolved before approval.')
    manifest = content.get('coverage', {}).get('sources', [])
    if not manifest:
        raise ValueError('The brief is missing its transcript coverage manifest. Extract the brief from reviewed sources.')
    checked_sources = {}
    for entry in manifest:
        source = store.get(entry.get('source_id', ''), client_id)
        if not source or source['kind'] != 'transcript' or not _source_complete(source['content']):
            raise ValueError('Confirm full-call coverage against the original source before approving this brief.')
        if entry.get('source_version') != source['content_hash'] or not entry.get('stored_text_complete'):
            raise ValueError('Transcript coverage does not match the saved source version.')
        checked_sources[source['id']] = source
    for fact in content['facts']:
        if not fact.get('text') or not fact.get('evidence'):
            raise ValueError('Every fact needs a source and a supporting quote.')
        for evidence in fact['evidence']:
            source = checked_sources.get(evidence.get('source_id', ''))
            if not source:
                raise ValueError('Brief evidence must belong to this client’s source records.')
            if evidence.get('source_version') and evidence['source_version'] != source['content_hash']:
                raise ValueError('A fact citation identifies a different source version.')
            raw = source['content']
            blocks = raw.get('transcript', [])
            quote = evidence.get('quote', '').strip()
            block = evidence.get('block')
            if isinstance(block, int) and 0 <= block < len(blocks):
                text = str(blocks[block].get('text', ''))
            else:
                text = str(raw.get('text', ''))
            if not quote or ' '.join(quote.split()).casefold() not in ' '.join(text.split()).casefold():
                raise ValueError('A brief citation does not match its source text.')


def save_record(client_id, kind, source_key, content, user, parent_id=None):
    store = get_store()
    if parent_id:
        parent = store.get(parent_id, client_id)
        if not parent or parent['kind'] != kind:
            raise ValueError('The previous version does not belong to this client and collection.')
        source_key = parent['source_key']
    # Review, provenance and identity are never accepted from the browser as approvals.
    content = json.loads(json.dumps(content))
    content.pop('approved_by', None)
    content.pop('approved', None)
    if kind == 'brief':
        if not parent_id:
            raise ValueError('Extract a brief from source transcripts before saving an edited version.')
        # Edits may clarify facts, not invent completion of unprocessed sources.
        content['coverage'] = parent['content'].get('coverage', {})
        content['extraction'] = parent['content'].get('extraction', {})
        content['client_id'] = client_id
        if not isinstance(content.get('facts'), list) or any(not isinstance(f, dict) for f in content['facts']):
            raise ValueError('Brief facts must be a list of structured statements.')
        before = {f.get('id'): f for f in parent['content'].get('facts', [])}
        for fact in content.get('facts', []):
            if fact != before.get(fact.get('id')):
                fact.update(origin='human', edited_by=actor(user))
    if kind == 'example':
        content['human_edited'] = True
        if parent_id:
            for key in ('source', 'source_url', 'channel', 'copy_bank_key', 'source_position',
                        'original_sent_record_id', 'original_sent_source_key', 'campaign_id', 'variant_id',
                        'sequence_id', 'step', 'performance_links', 'performance_proven', 'attribution',
                        'territory', 'industry', 'audience', 'context_provenance'):
                if key in parent['content']:
                    content[key] = parent['content'][key]
                else:
                    content.pop(key, None)
        else:
            content['source'] = 'manual_reference'
        if content.get('channel') == 'linkedin' and content.get('source') != 'copy_bank':
            raise ValueError('Import LinkedIn examples from Copy Bank, then save reviewed revisions.')
    return store.save(kind, client_id, source_key, content, actor(user), parent_id=parent_id)


def review_record(client_id, record_id, decision, notes, user):
    store = get_store()
    record = store.get(record_id, client_id)
    if not record:
        raise LookupError('Version not found.')
    if record['kind'] not in ('brief', 'example', 'draft', 'methodology'):
        raise ValueError('This source record is not an approvable reference.')
    if decision == 'approved':
        if record['kind'] == 'brief':
            validate_brief(record['content'], client_id, store)
            stale = brief_context_error(store, client_id, record)
            if stale:
                raise ValueError(stale)
        elif record['kind'] == 'example':
            content = record['content']
            if not content.get('sanitized'):
                raise ValueError('Confirm that recipient-specific personal data has been removed before approving this reusable example.')
            if content.get('channel') not in ('email', 'linkedin'):
                raise ValueError('Choose email or LinkedIn for this example.')
            steps = content.get('steps', [])
            if not isinstance(steps, list) or any(not isinstance(s, dict) or s.get('kind') not in ('connection', 'initial', 'follow_up', 'closing')
                                                or not isinstance(s.get('body'), str) or not s['body'].strip() for s in steps):
                raise ValueError('Every reference step needs a valid channel step and non-empty body.')
            if not steps:
                raise ValueError('Empty copy cannot be approved.')
            if content.get('channel') == 'linkedin' and content.get('source') != 'copy_bank':
                raise ValueError('LinkedIn references must originate in Copy Bank.')
        elif record['kind'] == 'draft':
            validate_saved_draft(client_id, record)
    return store.review(record_id, client_id, actor(user), decision, notes)


def draft_context(client_id, draft):
    store = get_store()
    context = draft['content'].get('studio_context', {})
    brief = store.get(context.get('brief_id', ''), client_id)
    if not brief or brief['status'] != 'approved':
        raise ValueError('The brief used for this draft is no longer approved.')
    stale = brief_context_error(store, client_id, brief)
    if stale:
        raise ValueError(stale)
    current = latest_approved(store.iter_records(client_id, kind='brief', status='approved'), 'brief')
    if current and current['id'] != brief['id']:
        raise ValueError('The client brief has changed. Generate a new draft from the current approved brief.')
    examples = [store.get(i, client_id) for i in context.get('example_ids', [])]
    eligible_ids = {r['id'] for r in current_examples(store, client_id)}
    if any(not r or r['id'] not in eligible_ids for r in examples):
        raise ValueError('A reference used for this draft is no longer current and approved.')
    if context.get('methodology_id'):
        methodology = store.get(context['methodology_id'], GLOBAL_CLIENT)
        if not methodology or methodology['status'] != 'approved':
            raise ValueError('The methodology used for this draft is no longer approved.')
    return brief, examples, context.get('request', {})


def validate_saved_draft(client_id, draft):
    from . import writer
    from .proofreading import final_check_current
    brief, examples, request = draft_context(client_id, draft)
    report = writer.validate_draft(client_id, brief, examples, request, draft['content'])
    # The writer returns a review report; semantic accuracy still needs human review.
    issues = report.get('issues', [])
    if any(i.get('severity') == 'error' for i in issues):
        raise ValueError('Resolve the draft’s quality checks before approval.')
    if draft['content'].get('final_check') and not final_check_current(draft['content']):
        raise ValueError('The final spelling and copy check is incomplete or out of date. Check this version again before approval.')
    return report


def revise_draft(client_id, identifier, content, user):
    from . import writer
    store = get_store()
    parent = store.get(identifier, client_id)
    if not parent or parent['kind'] != 'draft':
        raise LookupError('Draft not found.')
    revised = dict(parent['content'])
    revised['angles'] = content.get('angles', [])
    revised['human_edited'] = True
    revised.pop('final_check', None)
    brief, examples, request = draft_context(client_id, parent)
    revised['review'] = writer.validate_draft(client_id, brief, examples, request, revised)
    return store.save('draft', client_id, parent['source_key'], revised, actor(user), parent_id=identifier)


def enqueue_draft_check(client_id, identifier, user, request_key, *, locale=None):
    from .proofreading import infer_locale
    store = get_store()
    record = store.get(identifier, client_id)
    if not record or record['kind'] != 'draft':
        raise LookupError('Draft not found.')
    brief, _, writing_request = draft_context(client_id, record)
    selected_locale = locale or writing_request.get('locale') or infer_locale(brief, current_transcripts(store, client_id)).get('locale')
    if selected_locale not in ('en-GB', 'en-US'):
        raise ValueError('Choose UK or US English before checking the final copy.')
    return store.enqueue(client_id, 'review_draft', {'record_id': identifier, 'locale': selected_locale},
                         actor(user), 'check:'+request_key)


def prepare_reference(client_id, identifier, subject, body, sanitized, user):
    store = get_store()
    sent = store.get(identifier, client_id)
    if not sent or sent['kind'] != 'sent_copy':
        raise LookupError('Sent copy version not found.')
    original = sent['content']
    if not sanitized or not body.strip():
        raise ValueError('Remove recipient-specific data and confirm sanitisation before preparing a reusable reference.')
    performance = []
    for record in store.iter_records(client_id, kind='performance'):
        for variant in record['content'].get('variants', []):
            if sent['source_key'] in variant.get('copy_source_keys', []):
                performance.append({'record_id': record['id'], 'content_hash': record['content_hash'],
                                    'variant_id': variant.get('variant_id'), 'attribution': variant.get('attribution')})
    content = {'source': original['source'], 'channel': 'email', 'subjects': [subject] if subject else [],
               'steps': [{'kind': 'initial' if str(original.get('step', 1)) == '1' else 'follow_up', 'body': body}],
               'sanitized': True, 'reusable': True,
               'original_sent_record_id': sent['id'], 'original_sent_source_key': sent['source_key'],
               'campaign_id': original.get('campaign_id'), 'variant_id': original.get('variant_id'),
               'step': original.get('step'), 'sequence_id': original.get('sequence_id'),
               'performance_links': performance, 'performance_proven': False,
               'source_url': original.get('source_url', ''), 'approval_required': True,
               'attribution': 'review_required', 'title': 'Sent copy reference for review',
               **{k: original[k] for k in ('audience', 'territory', 'industry', 'context_provenance') if k in original}}
    return store.save('example', client_id, 'reference:'+sent['id'], content, actor(user))
