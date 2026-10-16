"""Confirmed DNS cleanup, tracking, DMARC and website forwarding for one domain."""
import copy
import re
from datetime import datetime, timezone
from urllib.parse import urlsplit
from fastapi import HTTPException
from app.utils.domain_cluster import provisioning as p
from . import service as s, store

PREVIEW_SECONDS = 900
MAX_RECORDS = 10


def require_admin(user):
    if user.get('role') != 'admin':
        raise HTTPException(403, 'An administrator must configure DNS through the shared Porkbun account.')


def canonical(record, domain):
    name = str(record.get('name', '')).lower().rstrip('.')
    if name in ('', '@'):
        name = domain
    elif name in ('_acme-challenge', 'acme-challenge', '_dmarc', 't'):
        name += '.'+domain
    return {'id': str(record.get('id', '')), 'name': name,
            'type': str(record.get('type', '')).upper(), 'content': str(record.get('content', '')),
            'ttl': str(record.get('ttl', '')), 'prio': str(record.get('prio', ''))}


def matches(record, domain):
    r = canonical(record, domain)
    if r['name'] == domain:
        if r['type'] == 'MX':
            return r['content'].strip().lower().rstrip('.') in ('fwd1.porkbun.com', 'fwd2.porkbun.com')
        if r['type'] == 'TXT':
            value = re.sub(r'\s+', ' ', r['content'].strip().strip('"')).lower()
            return value == 'v=spf1 include:_spf.porkbun.com ~all'
    return r['type'] == 'TXT' and r['name'] in ('_acme-challenge.'+domain, 'acme-challenge.'+domain)


def inspect_zone(domain, client):
    if s.normalize(domain) != domain:
        raise HTTPException(400, 'This setup needs a valid root domain.')
    check = client.check()
    if not all(check.get(k) for k in ('credentials_valid', 'read', 'dns')):
        raise HTTPException(403, 'Porkbun needs valid credentials with read and DNS permissions.')
    # Existing account, Active status, API access and nameserver checks.
    p.inspect(domain, {'forward': False, 'dmarc': False}, client)
    rows = client.call('get', '/dns/retrieve/'+domain).get('records')
    if not isinstance(rows, list) or any(not isinstance(r, dict) for r in rows):
        raise HTTPException(502, 'Porkbun returned an unreadable DNS record list. No deletion was approved.')
    records = [canonical(r, domain) for r in rows]
    if len({r['id'] for r in records}) != len(records):
        raise HTTPException(502, 'Porkbun returned ambiguous DNS record IDs. No deletion was approved.')
    targets = [r for r in records if matches(r, domain)]
    if len(targets) > MAX_RECORDS or any(not re.fullmatch(r'[0-9]{1,30}', r['id']) for r in targets):
        raise HTTPException(409, 'This domain needs a manual DNS review before cleanup.')
    return records, targets


def active_job(state, job_id):
    job = s.find(state, 'setups', job_id, True)
    s.find(state, 'clients', job['client_id'], True)
    if job['stage'] == s.BURNED:
        raise HTTPException(409, 'This domain is burned. Reopen it before setting up DNS.')
    return job


def additions(records, domain, email):
    email = s.email_domain(email, domain)
    if not email:
        raise HTTPException(400, 'Add the sending email on this domain before setting up DNS and DMARC.')
    desired = [
        {'name': 't', 'type': 'CNAME', 'content': 'open.sleadtrack.com', 'ttl': 600},
        {'name': '_dmarc', 'type': 'TXT', 'content': p.dmarc_value(domain, {'rua': email, 'ruf': email}), 'ttl': 600}]
    actions = []
    for payload in desired:
        host = payload['name']+'.'+domain
        existing = [r for r in records if canonical(r, domain)['name'] == host]
        compact = lambda v: re.sub(r'\s+', '', str(v).replace('"', '')).lower()
        same = len(existing) == 1 and existing[0]['type'] == payload['type'] and str(existing[0]['ttl']) == '600'
        if payload['type'] == 'CNAME':
            same = same and existing[0]['content'].lower().rstrip('.') == payload['content']
        else:
            same = same and compact(existing[0]['content']) == compact(payload['content'])
        if existing and not same:
            raise HTTPException(409, 'Existing '+host+' record differs. DNS setup stopped; review it in Porkbun before retrying.')
        actions.append({'payload': payload, 'action': 'already configured' if same else 'create'})
    return actions


def destination_url(raw, domain):
    value = str(raw or '').strip()
    if value and '://' not in value:
        value = 'https://'+value
    try:
        p.validate_options({'forward': True, 'dmarc': False, 'destination': value})
        host = (urlsplit(value).hostname or '').lower().rstrip('.')
        if host == domain or host.endswith('.'+domain):
            raise ValueError('Choose the client website, not this setup domain: forwarding would loop.')
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from None
    return value


def forwarding(domain, destination, client):
    rows = client.call('get', '/domain/getUrlForwarding/'+domain).get('forwards')
    if not isinstance(rows, list) or any(not isinstance(r, dict) for r in rows):
        raise HTTPException(502, 'Porkbun returned an unreadable forwarding list. Review and retry.')
    desired = {'subdomain': '', 'location': destination, 'type': 'temporary',
               'redirectType': '302', 'includePath': 'no', 'wildcard': 'yes'}
    root = [r for r in rows if not r.get('subdomain')]
    same = len(root) == 1 and all(str(root[0].get(k, '')) == v for k, v in desired.items() if k not in ('subdomain', 'redirectType'))
    same = same and str(root[0].get('redirectType', '302')) == '302'
    if root and not same:
        raise HTTPException(409, 'Existing website forwarding differs. Review it in Porkbun before proceeding.')
    return {'payload': desired, 'action': 'already configured' if same else 'create'}


def preview(user, job_id, revision):
    require_admin(user)
    record = s.board(user)
    if record['revision'] != revision:
        raise HTTPException(409, 'The tracker changed. Refresh before reviewing DNS records.')
    state = copy.deepcopy(record['state'])
    s.unlocked(state)
    job = active_job(state, job_id)
    records, targets = inspect_zone(job['domain'], p.Porkbun())
    actions = additions(records, job['domain'], job['email'])
    job['dns_setup_preview'] = {'id': s.new_id(), 'domain': job['domain'], 'email': job['email'],
                                'created_at': s.now(), 'records': targets, 'additions': actions}
    s.audit(state, user, 'dns_setup_preview')
    return store.save(user, state, revision)


def configure_records(domain, email, destination, approved, preview_id, guard, client=None):
    client = client or p.Porkbun()
    records, targets = inspect_zone(domain, client)
    saved = {r['id']: r for r in approved}
    live = {r['id']: r for r in records}
    # Missing approved records are already removed; changed or newly added
    # records require a fresh confirmation rather than expanding deletion.
    if any(r['id'] not in saved for r in targets) or any(i in live and live[i] != r for i, r in saved.items()):
        raise HTTPException(409, 'DNS records changed after review. Open Set up DNS again to review the current records.')
    # Check both destination hosts before any deletion, including on retries.
    additions(records, domain, email)
    forwarding(domain, destination, client)
    for record in targets:
        guard()
        client.call('post', '/dns/delete/'+domain+'/'+record['id'], {},
                    idempotency=p.fingerprint(['gmail-dns-cleanup', preview_id, domain, record]))
    for index in range(2):
        current = client.call('get', '/dns/retrieve/'+domain).get('records')
        if not isinstance(current, list) or any(not isinstance(r, dict) for r in current):
            raise HTTPException(502, 'DNS records could not be checked before adding the new records. Review and retry.')
        actions = additions([canonical(r, domain) for r in current], domain, email)
        action = actions[index]
        if action['action'] == 'create':
            guard()
            response = client.call('post', '/dns/create/'+domain, action['payload'],
                                   idempotency=p.fingerprint(['gmail-dns-setup', preview_id, domain, action['payload']]))
            if not isinstance(response.get('id'), (str, int)):
                raise HTTPException(502, 'Porkbun did not confirm the new DNS record ID. Review current records before retrying.')
    remaining = client.call('get', '/dns/retrieve/'+domain).get('records')
    if not isinstance(remaining, list) or any(not isinstance(r, dict) for r in remaining):
        raise HTTPException(502, 'DNS setup was requested but readback could not be verified. Review the records again.')
    if any(str(r.get('id')) in saved or matches(r, domain) for r in remaining):
        raise HTTPException(502, 'Some default DNS records remain. Review the current records before retrying.')
    if any(a['action'] != 'already configured' for a in additions([canonical(r, domain) for r in remaining], domain, email)):
        raise HTTPException(502, 'Tracking CNAME or DMARC could not be verified. Review current records before retrying.')
    forward = forwarding(domain, destination, client)
    if forward['action'] == 'create':
        guard()
        client.call('post', '/domain/addUrlForward/'+domain, forward['payload'],
                    idempotency=p.fingerprint(['gmail-dns-forward', preview_id, domain, forward['payload']]))
    if forwarding(domain, destination, client)['action'] != 'already configured':
        raise HTTPException(502, 'Website forwarding could not be verified. Review current forwarding before retrying.')
    # Forwarding can affect DNS; verify the mail and tracking records once more.
    final = client.call('get', '/dns/retrieve/'+domain).get('records')
    if not isinstance(final, list) or any(not isinstance(r, dict) for r in final) or any(matches(r, domain) for r in final):
        raise HTTPException(502, 'DNS changed during forwarding setup. Review current records before retrying.')
    if any(a['action'] != 'already configured' for a in additions([canonical(r, domain) for r in final], domain, email)):
        raise HTTPException(502, 'Tracking CNAME or DMARC changed during forwarding setup. Review and retry.')
    return {'domain': domain, 'status': 'verified', 'removed': len(approved),
            'email': email, 'destination': destination,
            'message': 'Setup verified: default DNS cleared, tracking CNAME, DMARC and website forwarding ready.',
            'checked_at': s.now()}


def apply(user, job_id, preview_id, revision, confirmed_purchased, destination):
    require_admin(user)
    if confirmed_purchased is not True:
        raise HTTPException(400, 'Confirm that domain purchase is complete before proceeding.')
    record = s.board(user)
    if record['revision'] != revision:
        raise HTTPException(409, 'The tracker changed. Review the DNS records again before proceeding.')
    state = copy.deepcopy(record['state'])
    s.unlocked(state)
    job = active_job(state, job_id)
    destination = destination_url(destination, job['domain'])
    approved = job.get('dns_setup_preview')
    if not approved or approved['id'] != preview_id or approved['domain'] != job['domain'] or approved['email'] != job['email']:
        raise HTTPException(409, 'Review this domain’s DNS records before proceeding.')
    age = (datetime.now(timezone.utc)-datetime.fromisoformat(approved['created_at'])).total_seconds()
    if not 0 <= age < PREVIEW_SECONDS:
        raise HTTPException(409, 'The DNS preview expired. Review the current records again.')
    token = s.new_id()
    state['sync_lock'] = {'kind': 'dns_setup', 'token': token, 'job_id': job_id, 'at': s.now()}
    job['dns_setup'] = {'domain': job['domain'], 'email': job['email'], 'status': 'running', 'checked_at': s.now()}
    store.save(user, state, revision)

    def guard():
        latest = store.load(user)
        lock = (latest or {}).get('state', {}).get('sync_lock') or {}
        age = (datetime.now(timezone.utc)-datetime.fromisoformat(lock.get('at') or approved['created_at'])).total_seconds()
        if lock.get('token') != token or not 0 <= age < s.LEASE_SECONDS-10:
            raise HTTPException(409, 'The DNS cleanup claim expired. Review the current records again.')

    try:
        result = configure_records(job['domain'], job['email'], destination, approved['records'], preview_id, guard)
    except HTTPException as exc:
        result = {'domain': job['domain'], 'email': job['email'], 'status': 'needs review', 'message': str(exc.detail), 'checked_at': s.now()}
    except Exception:
        result = {'domain': job['domain'], 'email': job['email'], 'status': 'needs review', 'message': 'DNS setup could not be verified. Review the current records before retrying.', 'checked_at': s.now()}
    # Keep the exact reviewed values alongside the outcome for later reference,
    # including when a provider response was lost after a partial deletion.
    result['records'] = copy.deepcopy(approved['records'])
    latest = store.load(user)
    if not latest or (latest['state'].get('sync_lock') or {}).get('token') != token:
        raise HTTPException(409, 'The cleanup request changed. Review current DNS records before retrying.')
    final = copy.deepcopy(latest['state'])
    final_job = s.find(final, 'setups', job_id)
    final_job['dns_setup'] = result
    if result['status'] == 'verified':
        s.find(final, 'clients', job['client_id'])['website_url'] = destination
        final_job['dmarc'] = {'domain': job['domain'], 'email': job['email'], 'status': 'verified',
                              'message': 'DMARC verified in Porkbun.', 'content': p.dmarc_value(job['domain'], {'rua': job['email'], 'ruf': job['email']}),
                              'ttl': 600, 'checked_at': result['checked_at']}
    final['sync_lock'] = None
    s.audit(final, user, 'dns_setup_'+result['status'].replace(' ', '_'))
    return store.save(user, final, latest['revision'])
