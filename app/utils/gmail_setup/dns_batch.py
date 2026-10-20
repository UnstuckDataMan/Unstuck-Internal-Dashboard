"""Client-scoped, resumable DNS setup; each request processes one domain."""
import copy
from datetime import datetime, timezone
from fastapi import HTTPException
from . import dns_setup as dns, service as s, store

TERMINAL = ('verified', 'needs review')
RUNNING_SECONDS = s.LEASE_SECONDS+60


def running(batch):
    stamp = batch.get('current_started_at')
    return bool(batch.get('current_job_id') and stamp and
                (datetime.now(timezone.utc)-datetime.fromisoformat(stamp)).total_seconds() < RUNNING_SECONDS)


def start(user, client_id, revision, confirmed_purchased, destination):
    dns.require_setup_access(user)
    if confirmed_purchased is not True:
        raise HTTPException(400, 'Confirm that all domain purchases are complete before proceeding.')
    record = s.board(user)
    if record['revision'] != revision:
        raise HTTPException(409, 'The client changed. Refresh and review its domains before proceeding.')
    state = copy.deepcopy(record['state'])
    s.unlocked(state)
    client = s.find(state, 'clients', client_id, True)
    if running(client.get('dns_setup_batch') or {}):
        raise HTTPException(409, 'A domain is already being processed for this client. Wait for it to finish before resuming.')
    jobs = [j for j in state['setups'] if j.get('client_id') == client_id and not j.get('archived') and j['stage'] not in s.INACTIVE_STAGES]
    if not jobs:
        raise HTTPException(400, 'Add actual domains to this client before running setup.')
    for job in jobs:
        destination = dns.destination_url(destination, job['domain'])
    client['dns_setup_batch'] = {'id': s.new_id(), 'created_at': s.now(), 'destination': destination,
        'items': [{'id': j['id'], 'domain': j['domain'], 'email': j['email']} for j in jobs], 'results': {}}
    s.audit(state, user, 'dns_batch_start')
    return store.save(user, state, revision)


def finish(user, client_id, batch_id, job_id, result):
    # A normal tracker edit between domain completion and this result must not
    # discard either change. No external action is repeated during CAS retries.
    for attempt in range(3):
        latest = s.board(user)
        state = copy.deepcopy(latest['state'])
        s.unlocked(state)
        client = s.find(state, 'clients', client_id, True)
        batch = client.get('dns_setup_batch')
        if not batch or batch['id'] != batch_id or batch.get('current_job_id') != job_id:
            raise HTTPException(409, 'The batch changed. Refresh to see its saved results.')
        batch['results'][job_id] = result
        batch.pop('current_job_id', None)
        batch.pop('current_started_at', None)
        if result.get('rate_limited'):
            batch['paused_reason'] = 'Porkbun rate limit reached. Wait before resuming this batch.'
        if all(batch['results'].get(i['id'], {}).get('status') in TERMINAL for i in batch['items']):
            batch['completed_at'] = s.now()
        s.audit(state, user, 'dns_batch_domain_'+result['status'].replace(' ', '_'))
        try:
            return store.save(user, state, latest['revision'])
        except HTTPException as exc:
            if exc.status_code != 409 or attempt == 2:
                raise


def step(user, client_id, batch_id, job_id, revision):
    dns.require_setup_access(user)
    record = s.board(user)
    if record['revision'] != revision:
        raise HTTPException(409, 'The tracker changed. Refresh and resume the saved batch.')
    state = copy.deepcopy(record['state'])
    s.unlocked(state)
    client = s.find(state, 'clients', client_id, True)
    batch = client.get('dns_setup_batch')
    if not batch or batch['id'] != batch_id:
        raise HTTPException(409, 'This batch was replaced. Refresh and review the current batch.')
    item = next((i for i in batch['items'] if i['id'] == job_id), None)
    if not item:
        raise HTTPException(400, 'Only domains in this client’s confirmed batch can be processed.')
    if batch['results'].get(job_id, {}).get('status') in TERMINAL:
        return record
    if running(batch):
        raise HTTPException(409, 'A domain is already being processed. Wait for its result before resuming.')
    pending = [i for i in batch['items'] if batch['results'].get(i['id'], {}).get('status') not in TERMINAL]
    if not pending or pending[0]['id'] != job_id:
        raise HTTPException(409, 'Process the next waiting domain first to keep the batch sequential.')
    batch.update(current_job_id=job_id, current_started_at=s.now())
    batch.pop('paused_reason', None)
    batch['results'][job_id] = {'status': 'running', 'checked_at': s.now()}
    claimed = store.save(user, state, revision)
    try:
        job = s.find(state, 'setups', job_id, True)
        if job.get('client_id') != client_id or job['domain'] != item['domain'] or job['email'] != item['email'] or job['stage'] in s.INACTIVE_STAGES:
            raise HTTPException(409, 'This domain changed after confirmation, was removed or burned, or its Gmail was deleted. Review it before running setup again.')
        if not s.email_domain(job['email'], job['domain']):
            raise HTTPException(400, 'Add the sending email for this domain before running setup again.')
        # Fresh record review per domain avoids expiring a large batch's DNS
        # snapshots while earlier domains are still being configured.
        prepared = dns.preview(user, job_id, claimed['revision'])
        preview_id = s.find(prepared['state'], 'setups', job_id)['dns_setup_preview']['id']
        applied = dns.apply(user, job_id, preview_id, prepared['revision'], True, batch['destination'])
        outcome = s.find(applied['state'], 'setups', job_id)['dns_setup']
        result = {k: outcome[k] for k in ('status', 'message', 'checked_at', 'rate_limited') if k in outcome}
    except HTTPException as exc:
        result = {'status': 'needs review', 'message': str(exc.detail), 'checked_at': s.now(), 'rate_limited': exc.status_code == 429}
    except Exception:
        result = {'status': 'needs review', 'message': 'This domain could not be verified. Review its records before retrying.', 'checked_at': s.now()}
    if result.get('rate_limited'):
        result['status'] = 'pending'
    return finish(user, client_id, batch_id, job_id, result)
