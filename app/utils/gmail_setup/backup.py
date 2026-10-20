"""Durable sheet intents, fast request scheduling and merged worker results."""
import copy
import re
import threading
from datetime import datetime, timezone

from fastapi import HTTPException

from . import service as s, sheets, store

_running = set()
_running_lock = threading.Lock()
_ROW_FIELDS = ('tab', 'sheet_id', 'row', 'url')


class ChangedIntent(HTTPException):
    def __init__(self):
        super().__init__(409, 'A newer sheet backup is waiting. The latest state will be backed up.')


def identity(state, job):
    saved = job.get('sheet_identity')
    if saved:
        return copy.deepcopy(saved)
    operator = next((member for member in state['members'] if member['id'] == job.get('operator_id')), {})
    owner = job.get('sheet_owner') or (job['client_name'] if job['section'] == 'client' else operator.get('name', ''))
    master = next((row for row in state['setups'] if row['id'] == job.get('master_domain_id')
                   and row.get('client_id') == job['client_id'] and not row.get('archived')
                   and row.get('master_domain_id') == row['id']), None)
    display = job['email']
    if master:
        display += ' (Master)' if master['id'] == job['id'] else f' (Secondary: {master["domain"]} is Master)'
    return {'email': job['email'], 'domain': job['domain'], 'section': job['section'], 'owner': owner,
            'sheet_display_email': display, 'master_domain_id': master['id'] if master else None,
            'master_domain': master['domain'] if master else None}


def queue(state, job, live, force=False):
    old = job.get('sheet_intent')
    if old and old['live'] is live and not force:
        return
    saved = identity(state, job)
    if saved['email']:
        job['sheet_identity'] = copy.deepcopy(saved)
    generation = s.new_id()
    snapshot = {'id': job['id'], 'email': saved['email'], 'domain': saved['domain'], 'section': saved['section'],
                'order': job.get('order'), 'sheet_display_email': saved['sheet_display_email']}
    job['sheet_intent'] = {'generation': generation, 'live': live,
                           'stage': s.STAGES[2] if live else s.GMAIL_DELETED,
                           'snapshot': snapshot, 'owner': saved['owner']}
    previous = job.get('sync') or {}
    job['sync'] = {**{key: previous[key] for key in _ROW_FIELDS if key in previous},
                   'status': 'pending', 'generation': generation, 'live': live,
                   'operation': 'activate' if live else 'deactivate', 'stage': job['sheet_intent']['stage']}


def queue_retry(user, job_id):
    for attempt in range(5):
        record = s.board(user)
        state = copy.deepcopy(record['state'])
        s.unlocked(state, allow_sheet=True)
        job = s.find(state, 'setups', job_id, True)
        s.find(state, 'clients', job['client_id'], True)
        desired = job.get('sheet_intent', {}).get('live')
        if desired is None:
            if job['stage'] not in (s.STAGES[2], s.GMAIL_DELETED):
                raise HTTPException(409, 'Complete setup or mark Gmail deleted before creating a sheet backup.')
            desired = job['stage'] == s.STAGES[2]
        queue(state, job, desired, force=True)
        s.audit(state, user, 'sheet_backup_queued')
        try:
            return store.save(user, state, record['revision'])
        except HTTPException as exc:
            if exc.status_code != 409 or attempt == 4:
                raise


def queue_client(user, client_id, expected):
    record = s.board(user)
    if record['revision'] != expected:
        raise HTTPException(409, 'The team tracker changed. Refresh before backing up this client.', headers=store.REVISION_CONFLICT_HEADERS)
    state = copy.deepcopy(record['state'])
    s.unlocked(state, allow_sheet=True)
    s.find(state, 'clients', client_id, True)
    for job in state['setups']:
        if job.get('client_id') == client_id and not job.get('archived') and job['stage'] in (s.STAGES[2], s.GMAIL_DELETED):
            queue(state, job, job['stage'] == s.STAGES[2], force=True)
    s.audit(state, user, 'client_sheet_backup_queued')
    return store.save(user, state, expected)


def pending(state):
    active_clients = {client['id'] for client in state.get('clients', []) if not client.get('archived')}
    return [job for job in state['setups'] if not job.get('archived') and job.get('client_id') in active_clients
            and job.get('sheet_intent') and job.get('sync', {}).get('status') == 'pending']


def _matches(job, generation):
    return job.get('sheet_intent', {}).get('generation') == generation


def sync_once(user, job_id):
    record = s.board(user)
    state = copy.deepcopy(record['state'])
    job = s.find(state, 'setups', job_id, True)
    s.find(state, 'clients', job['client_id'], True)
    if not job.get('sheet_intent'):
        if job.get('sync', {}).get('status') == 'synced' and job['stage'] == s.STAGES[2]:
            return record
        if job['stage'] not in (s.STAGES[2], s.GMAIL_DELETED):
            raise HTTPException(409, 'Complete setup or mark Gmail deleted before backing up to the strategy sheet.')
        queue(state, job, job['stage'] == s.STAGES[2])
    intent = copy.deepcopy(job['sheet_intent'])
    generation = intent['generation']
    if job.get('sync', {}).get('status') == 'synced' and job['sync'].get('generation') == generation:
        return record
    s.unlocked(state)
    token = s.new_id()
    state['sync_lock'] = {'kind': 'sheet', 'job_id': job_id, 'token': token, 'generation': generation, 'at': s.now()}
    job['sync'] = {**job['sync'], 'status': 'pending', 'in_progress': True}
    store.save(user, state, record['revision'])

    def guard():
        latest = store.load(user)
        lock = (latest or {}).get('state', {}).get('sync_lock') or {}
        current = next((row for row in (latest or {}).get('state', {}).get('setups', []) if row['id'] == job_id), None)
        age = (datetime.now(timezone.utc)-datetime.fromisoformat(lock.get('at') or s.now())).total_seconds()
        if lock.get('token') != token or not 0 <= age < s.LEASE_SECONDS-25:
            raise ChangedIntent()
        if not current or current.get('archived') or not _matches(current, generation):
            raise ChangedIntent()

    base = {key: job['sync'][key] for key in (*_ROW_FIELDS, 'generation', 'live', 'operation', 'stage') if key in job['sync']}
    outcome = None
    try:
        guard()
        if intent['live']:
            if not s.email_domain(intent['snapshot']['email'], intent['snapshot']['domain']):
                raise HTTPException(400, 'Add the sending email before backing up this completed domain.')
            if type(intent['snapshot'].get('order')) is not int or intent['snapshot']['order'] < 1:
                raise HTTPException(400, 'This completed domain needs a valid completion order before sheet backup.')
            outcome = sheets.publish(intent['snapshot'], intent['owner'], guard=guard)
        else:
            outcome = sheets.deactivate(intent['snapshot'], intent['owner'], guard=guard)
        if outcome.get('missing'):
            base = {key: value for key, value in base.items() if key not in ('row', 'url')}
        status = {**base, **outcome, 'status': 'synced', 'at': s.now(), 'live': intent['live']}
    except ChangedIntent:
        status = None
    except HTTPException as exc:
        status = {**base, 'status': 'error', 'at': s.now(), 'message': str(exc.detail)}
        match = re.search(r'https://docs\.google\.com/spreadsheets/d/[A-Za-z0-9_-]+/edit#gid=\d+&range=A\d+', str(exc.detail))
        if match:
            status['url'] = match.group(0)
    except Exception:
        status = {**base, 'status': 'error', 'at': s.now(), 'message': 'The sheet backup could not be verified. Retry to check the existing row before adding anything.'}
    for attempt in range(8):
        latest = s.board(user)
        final = copy.deepcopy(latest['state'])
        current = s.find(final, 'setups', job_id)
        lock = final.get('sync_lock') or {}
        if lock.get('token') != token:
            return latest
        if _matches(current, generation) and status is not None:
            current['sync'] = status
        elif not _matches(current, generation) and outcome:
            # A newer intent keeps its pending status, while retaining any
            # location found by the older request for the next reconciliation.
            current.setdefault('sync', {}).update({key: outcome[key] for key in _ROW_FIELDS if key in outcome})
        final['sync_lock'] = None
        s.audit(final, user, 'sheet_backup_'+(status['status'] if status is not None and _matches(current, generation) else 'superseded'))
        try:
            return store.save(user, final, latest['revision'])
        except HTTPException as exc:
            if exc.status_code != 409 or attempt == 7:
                raise


def drain(user):
    for _ in range(5000):
        record = s.board(user)
        jobs = pending(record['state'])
        if not jobs:
            return
        try:
            sync_once(user, jobs[0]['id'])
        except HTTPException:
            # A DNS claim or another process owns this workspace. Durable
            # pending intents resume on the next board request or save.
            return


def resume(user):
    workspace = store.workspace(user)
    with _running_lock:
        if workspace in _running:
            return
        _running.add(workspace)
    try:
        drain(user)
    except Exception:
        # Do not reflect provider errors or private board contents into logs.
        return
    finally:
        with _running_lock:
            _running.discard(workspace)
