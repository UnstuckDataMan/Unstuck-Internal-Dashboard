"""Explicit, reviewed append to Copy Bank. No campaign sending or notifications.

The database RPC owns the final compare-and-append transaction. A preview is not
an approval, and a network retry cannot append the same selection twice.
"""
from __future__ import annotations

from copy import deepcopy
import json
import uuid

from .store import StoreNotReady, get_store


class PublishingError(ValueError):
    pass


class PublishingConflict(PublishingError):
    pass


_KINDS = {'draft', 'brief', 'example', 'transcript', 'sync_report'}
_NAMESPACE = uuid.UUID('6b4cb4cc-b982-4210-b542-fec8b13c2a8a')
_SECTIONS = {'email': 'Opening email', 'chaser': 'Email follow-ups',
             'linkedin': 'LinkedIn sequence', 'flyout': 'BizDev flyout'}


def _checked(response):
    if not isinstance(response, dict):
        raise StoreNotReady('Copy Bank returned an invalid response.')
    if response.get('ok') is False:
        error = PublishingConflict if response.get('code') == 'conflict' else PublishingError
        raise error(response.get('message') or 'Copy Bank publishing is unavailable.')
    return response


class _Gateway:
    def __init__(self, store):
        if not callable(getattr(store, '_rpc', None)):
            raise StoreNotReady('Copy Bank publishing requires the configured Supabase database and publishing migration. Local records are not published to a live Copy Bank.')
        self.store = store

    def destinations(self, client_id):
        return _checked(self.store._rpc('copy_bank_destinations', p_client_id=client_id))

    def destination(self, client_id, selection):
        return _checked(self.store._rpc('copy_bank_destination', p_client_id=client_id,
                                       p_selection=selection))

    def publish(self, **kwargs):
        return _checked(self.store._rpc('publish_copy_bank', **{'p_'+k: v for k, v in kwargs.items()}))


class _Snapshot:
    def __init__(self, records):
        self.records = sorted(records, key=lambda r: (r['created_at'], r['id']), reverse=True)

    def get(self, identifier, client_id):
        return next((r for r in self.records if r['id'] == identifier and r['client_id'] == client_id), None)

    def iter_records(self, client_id, kind=None, status=None):
        return (r for r in self.records if r['client_id'] == client_id
                and (kind is None or r['kind'] == kind) and (status is None or r['status'] == status))


def _draft_and_context(store, client_id, draft_id):
    """Validate against exactly the records/reviews later compared inside SQL."""
    from . import service
    from .proofreading import final_check_current
    records = [r for r in store.iter_records(client_id) if r['kind'] in _KINDS]
    draft = next((r for r in records if r['id'] == draft_id and r['kind'] == 'draft'), None)
    if not draft or draft.get('agency_id') != store.agency_id or draft.get('client_id') != client_id:
        raise PublishingError('This saved draft was not found for the selected client.')
    if draft['status'] != 'approved':
        raise PublishingError('Approve this exact saved draft before publishing it.')
    latest = max((r for r in records if r['kind'] == 'draft' and r['source_key'] == draft['source_key']),
                 key=lambda r: (r['created_at'], r['id']))
    if latest['id'] != draft_id:
        raise PublishingError('A newer edited version exists. Check and approve that version before publishing.')
    if not final_check_current(draft['content']):
        raise PublishingError('Complete the final copy and spelling check on this exact saved version before publishing.')
    context = draft['content'].get('studio_context', {})
    if context.get('methodology_id'):
        methodology = store.get(context['methodology_id'], '__agency__')
        if not methodology or methodology['kind'] != 'methodology' or methodology['status'] != 'approved':
            raise PublishingError('The methodology used for this draft is no longer approved.')
        records.append(methodology)
    snapshot = _Snapshot(records)
    brief = snapshot.get(context.get('brief_id'), client_id)
    if not brief or brief['kind'] != 'brief' or brief['status'] != 'approved':
        raise PublishingError('The brief used for this draft is no longer approved.')
    latest_brief = service.latest_approved(snapshot.iter_records(client_id, kind='brief'), 'brief')
    if latest_brief['id'] != brief['id']:
        raise PublishingError('The client brief has changed. Generate from the current approved brief.')
    # This also checks the current profile and active refresh jobs. The SQL
    # transaction compares the profile/version and rejects active refresh jobs;
    # the immutable corpus signature closes the gap between these reads.
    context_error = service.brief_context_error(store, client_id, brief)
    if context_error:
        raise PublishingError(context_error)
    current_ids = {r['id'] for r in service.current_examples(snapshot, client_id)}
    if any(identifier not in current_ids for identifier in context.get('example_ids', [])):
        raise PublishingError('A reference used for this draft is no longer current and approved.')
    signature = {r['id']: {'status': r['status'], 'review_id': (r.get('review') or {}).get('id')}
                 for r in records}
    return draft, signature


def destinations(client_id, *, store=None, gateway=None):
    store = store or get_store()
    result = (gateway or _Gateway(store)).destinations(client_id)
    _checked(result)
    profile = result.get('profile') or {}
    if str(profile.get('client_id')) != client_id:
        raise PublishingError('Copy Bank has no exact profile for this client. Add its profile in Copy Bank first.')
    sections = [{'id': key, 'label': label} for key, label in _SECTIONS.items()
                if key != 'flyout' or profile.get('type') == 'bizdev']
    return {'client_id': client_id, 'profile': {k: profile.get(k) for k in ('client_id', 'name', 'type')},
            'territories': result.get('territories', []), 'industries': result.get('industries', []),
            'sections': sections, 'configured': bool(result.get('territories') and result.get('industries'))}


def _selection(content, profile, selection):
    if not isinstance(selection, dict):
        raise PublishingError('Choose the destination and angle to publish.')
    chosen = {key: selection.get(key) for key in ('territory', 'industry', 'section', 'angle_index')}
    for key in ('territory', 'industry'):
        if not isinstance(chosen[key], str) or not chosen[key].strip():
            raise PublishingError('Choose a territory and industry from this client’s Copy Bank profile.')
    section, index = chosen['section'], chosen['angle_index']
    channel = content.get('channel') or content.get('request', {}).get('channel') or content.get('studio_context', {}).get('request', {}).get('channel')
    if (channel == 'linkedin' and section != 'linkedin') or (channel == 'email' and section not in ('email', 'chaser', 'flyout')) or channel not in ('email', 'linkedin'):
        raise PublishingError('This Copy Bank section does not match the draft channel.')
    if section == 'flyout' and profile.get('type') != 'bizdev':
        raise PublishingError('Flyout copy can only be published to an explicitly mapped BizDev profile.')
    angles = content.get('angles', [])
    if type(index) is not int or not 0 <= index < len(angles):
        raise PublishingError('Choose one saved draft angle.')
    angle = angles[index]
    allowed = {'initial'} if section in ('email', 'flyout') else ({'follow_up', 'closing'} if section == 'chaser' else {'connection', 'initial', 'follow_up', 'closing'})
    steps = angle.get('steps', [])
    indices = selection.get('step_indices')
    if indices is None:
        indices = [i for i, step in enumerate(steps) if step.get('kind') in allowed]
    if (not isinstance(indices, list) or not indices or any(type(i) is not int or not 0 <= i < len(steps) for i in indices)
            or len(set(indices)) != len(indices)):
        raise PublishingError('Choose at least one valid, distinct step for the selected section.')
    indices = sorted(indices)
    selected = [steps[i] for i in indices]
    if any(step.get('kind') not in allowed or not isinstance(step.get('body'), str) or not step['body'].strip() for step in selected):
        raise PublishingError('The selected steps are empty or incompatible with this section.')
    subjects = angle.get('subjects', []) if section in ('email', 'flyout') else []
    if not isinstance(subjects, list) or any(not isinstance(s, str) or not s.strip() for s in subjects):
        raise PublishingError('Saved subject variants must contain non-empty text.')
    chosen['step_indices'] = indices
    return chosen, {'subjects': subjects, 'steps': [{'kind': s['kind'], 'body': s['body']} for s in selected]}


def _section_counts(content, section):
    existing = content.get(section, {})
    if isinstance(existing, list):
        return sum(bool(row.get('subject')) for row in existing if isinstance(row, dict)), len(existing)
    if not isinstance(existing, dict):
        raise PublishingError('The destination section has an unsupported format. Review it in Copy Bank first.')
    subjects = existing.get('subjects', [])
    bodies = existing.get('variations' if section in ('email', 'flyout') else 'steps', [])
    if not isinstance(subjects, list) or not isinstance(bodies, list):
        raise PublishingError('The destination section has an unsupported format. Review it in Copy Bank first.')
    return len(subjects), len(bodies)


def _prepare(client_id, draft_id, selection, store, gateway):
    destination = _checked(gateway.destination(client_id, selection))
    if str((destination.get('profile') or {}).get('client_id')) != client_id:
        raise PublishingError('Copy Bank has no exact profile for this client.')
    draft, context = _draft_and_context(store, client_id, draft_id)
    if destination.get('pending'):
        raise PublishingConflict('This Copy Bank destination has a pending approval. Resolve it there before appending copy.')
    normalized, appended = _selection(draft['content'], destination['profile'], selection)
    identity = {'agency_id': store.agency_id, 'client_id': client_id, 'draft_id': draft_id,
                'destination_key': destination['key'], 'selection': normalized}
    publication_id = str(uuid.uuid5(_NAMESPACE, json.dumps(identity, sort_keys=True, separators=(',', ':'), ensure_ascii=False)))
    return draft, context, destination, normalized, appended, publication_id


def preview(client_id, draft_id, selection, *, store=None, gateway=None):
    store = store or get_store()
    gateway = gateway or _Gateway(store)
    _, _, dest, chosen, appended, identifier = _prepare(client_id, draft_id, selection, store, gateway)
    subject_count, body_count = _section_counts(dest.get('content') or {}, chosen['section'])
    warnings = []
    if chosen['section'] in ('linkedin', 'chaser') and body_count:
        warnings.append('These messages will be appended after the existing steps in this section.')
    return {'publication_id': identifier, 'draft_id': draft_id, 'destination_key': dest['key'],
            'destination_version': dest['version'], 'selection': chosen, 'appended_copy': deepcopy(appended),
            'append_summary': {'subjects_added': len(appended['subjects']), 'bodies_added': len(appended['steps']),
                               'existing_subjects': subject_count, 'existing_bodies': body_count},
            'warnings': warnings, 'already_published': False}


def publish(client_id, draft_id, selection, expected_version, actor, *, store=None, gateway=None):
    store = store or get_store()
    gateway = gateway or _Gateway(store)
    if not isinstance(actor, str) or not actor.strip() or len(actor) > 320:
        raise PublishingError('An authenticated reviewer identity is required.')
    if not isinstance(expected_version, str) or not expected_version or len(expected_version) > 100:
        raise PublishingError('Preview the destination before publishing.')
    _, context, _, chosen, _, identifier = _prepare(client_id, draft_id, selection, store, gateway)
    result = gateway.publish(client_id=client_id, draft_id=draft_id, selection=chosen,
                             expected_version=expected_version, actor=actor, publication_id=identifier,
                             context_snapshot=context)
    return _checked(result)
