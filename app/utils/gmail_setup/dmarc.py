"""Client-scoped DMARC creation, with purchase confirmation and durable claims."""
import copy
from fastapi import HTTPException
from app.utils.domain_cluster import provisioning as p
from . import service as s, store


def require_admin(user):
    if user.get('role') != 'admin':
        raise HTTPException(403, 'An administrator must configure DMARC through the shared Porkbun account.')


def apply_record(domain, email, key, client=None):
    """Create only a missing record; verify account, delegation and readback."""
    email = s.email_domain(email, domain)
    if not email:
        raise HTTPException(400, 'Add the sending email address before setting up DMARC.')
    client = client or p.Porkbun()
    check = client.check()
    if not all(check.get(field) for field in ('credentials_valid', 'read', 'dns')):
        raise HTTPException(403, 'Porkbun needs valid credentials with read and DNS permissions.')
    options = {'forward': False, 'dmarc': True, 'rua': email, 'ruf': email}
    approved = p.inspect(domain, options, client)
    if approved['status'] != 'ready':
        raise HTTPException(409, ' '.join(approved['conflicts']))
    result = p.apply_one(domain, approved, options, key, client)
    return {**result, 'domain': domain, 'email': email, 'host': '_dmarc.'+domain,
            'content': p.dmarc_value(domain, options), 'ttl': 600,
            'message': ('DMARC verified in Porkbun. Public DNS propagation may still be pending.'
                        if result['status'] == 'verified' else result['message'])}


def start(user, client_id, revision, confirmed_purchased, setup_ids=None):
    require_admin(user)
    if confirmed_purchased is not True:
        raise HTTPException(400, 'Confirm that domain purchases are complete before proceeding.')
    record = s.board(user)
    if record['revision'] != revision:
        raise HTTPException(409, 'The client changed. Refresh and review its domains before proceeding.')
    state = copy.deepcopy(record['state'])
    s.unlocked(state)
    client = s.find(state, 'clients', client_id, True)
    jobs = [j for j in state['setups'] if j.get('client_id') == client_id and not j.get('archived') and j['stage'] not in s.INACTIVE_STAGES]
    if setup_ids is not None:
        selected = set(setup_ids)
        if not selected or selected - {j['id'] for j in jobs}:
            raise HTTPException(400, 'Select active domains belonging to this client only.')
        jobs = [j for j in jobs if j['id'] in selected]
    if not jobs:
        raise HTTPException(400, 'Add domains to this client first.')
    batch = {'id': s.new_id(), 'created_at': s.now(), 'items': [], 'results': {}}
    for job in jobs:
        item = {'id': job['id'], 'domain': job['domain'], 'email': job['email']}
        batch['items'].append(item)
        try:
            if s.normalize(item['domain']) != item['domain'] or not s.email_domain(item['email'], item['domain']):
                raise ValueError()
        except (ValueError, HTTPException):
            result = {**item, 'status': 'needs review', 'message': 'Add a valid sending email for this domain, then retry DMARC.', 'checked_at': s.now()}
            batch['results'][job['id']] = result
            job['dmarc'] = result
    client['dmarc_batch'] = batch
    s.audit(state, user, 'dmarc_start')
    return store.save(user, state, revision)


def apply(user, client_id, batch_id, job_id, revision):
    require_admin(user)
    record = s.board(user)
    if record['revision'] != revision:
        raise HTTPException(409, 'The tracker changed. Refresh before retrying DMARC.')
    state = copy.deepcopy(record['state'])
    s.unlocked(state)
    client = s.find(state, 'clients', client_id, True)
    batch = client.get('dmarc_batch')
    if not batch or batch['id'] != batch_id:
        raise HTTPException(409, 'Confirm domain purchases again to start a current DMARC request.')
    item = next((i for i in batch['items'] if i['id'] == job_id), None)
    job = s.find(state, 'setups', job_id, True)
    if job['stage'] in s.INACTIVE_STAGES:
        status = 'burned' if job['stage'] == s.BURNED else 'marked Gmail deleted'
        raise HTTPException(409, f'This domain is {status}. Reopen it before setting up DMARC.')
    if not item or job.get('client_id') != client_id:
        raise HTTPException(400, 'Only domains belonging to this client can be configured.')
    if job['domain'] != item['domain'] or job['email'] != item['email']:
        raise HTTPException(409, 'The domain or sending email changed. Review it and restart DMARC setup.')
    if not item['email'] or batch['results'].get(job_id, {}).get('status') == 'verified':
        return record
    token = s.new_id()
    state['sync_lock'] = {'kind': 'dmarc', 'token': token, 'job_id': job_id, 'at': s.now()}
    batch['results'][job_id] = {**item, 'status': 'running', 'checked_at': s.now()}
    store.save(user, state, revision)  # Claim before any external write.
    try:
        key = p.fingerprint(['gmail-set-dmarc', store.workspace(user), job_id, item['domain'], item['email']])
        result = apply_record(item['domain'], item['email'], key)
    except HTTPException as exc:
        result = {**item, 'status': 'needs review', 'message': str(exc.detail), 'checked_at': s.now()}
    except Exception:
        # Do not echo credential-bearing provider exceptions. Persist uncertainty
        # and release the claim so a read-before-write retry can reconcile it.
        result = {**item, 'status': 'needs review', 'message': 'DMARC could not be verified. Retry to check whether the record was created.', 'checked_at': s.now()}
    latest = store.load(user)
    if not latest or (latest['state'].get('sync_lock') or {}).get('token') != token:
        raise HTTPException(409, 'The DMARC request changed. Refresh and retry; existing records will be checked before writing.')
    final = copy.deepcopy(latest['state'])
    s.find(final, 'clients', client_id)['dmarc_batch']['results'][job_id] = result
    s.find(final, 'setups', job_id)['dmarc'] = result
    final['sync_lock'] = None
    s.audit(final, user, 'dmarc_'+result['status'].replace(' ', '_'))
    return store.save(user, final, latest['revision'])
