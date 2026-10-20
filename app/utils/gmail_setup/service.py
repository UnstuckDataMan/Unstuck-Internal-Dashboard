"""Server-authoritative stages, device history, reserve counts and sync leases."""
import copy
import re
from datetime import datetime, timezone
from uuid import uuid4, uuid5, NAMESPACE_URL
from fastapi import HTTPException
from app.utils.domain_cluster.engine import normalize
from . import store

STAGES = ('Not started', 'Purchase made', 'Setup complete')
STARTED = 'Started'
PLANNING_STAGES = (STAGES[0], STARTED)
BURNED = 'Burned'
GMAIL_DELETED = 'Gmail deleted'
INACTIVE_STAGES = (BURNED, GMAIL_DELETED)
CLOSED_STAGES = (STAGES[2], *INACTIVE_STAGES)
FAILURES = {'Failed at QR': 'QR code', 'Failed at payment': 'Payment'}
DEVICE_STATUSES = ('Unused', 'Used', 'Blocked')
RESERVE_CONTINGENCY = 10
DEFAULT_CAPACITY = {'Phone': 2, 'PC': 1}
MAX_MASTER_DOMAINS = 5
LEASE_SECONDS = 180


def now():
    return datetime.now(timezone.utc).isoformat()


def new_id():
    return str(uuid4())


def initial(user):
    members = []
    for row in store.team_directory(user):
        email = row['email'].strip().lower()
        members.append({'id': str(uuid5(NAMESPACE_URL, 'gmail-set:'+email)),
                        'name': (row.get('name') or email.split('@')[0].replace('.', ' ').title())[:100],
                        'email': email, 'active': True})
    return {'members': members, 'devices': [], 'clients': [], 'setups': [], 'events': [], 'sync_lock': None}


def board(user):
    record = store.load(user) or {'state': initial(user), 'revision': 0, 'updated_at': None}
    record = copy.deepcopy(record)
    ensure_clients(record['state'])
    ensure_device_usage(record['state'])
    return record


def client_key(section, name):
    return section+':'+re.sub(r'\s+', ' ', name.strip()).casefold()


def ensure_clients(state):
    # Older flat setup records become stable client groups without touching
    # their domains, stages, purchase history or published sheet rows.
    clients = state.setdefault('clients', [])
    by_key = {client_key(c['section'], c['name']): c for c in clients}
    by_id = {c['id']: c for c in clients}
    for job in state['setups']:
        if job.get('client_id') in by_id:
            continue
        section = job['section']
        name = job.get('client_name', '').strip() or 'Biz Dev'
        key = client_key(section, name)
        if key not in by_key:
            client = {'id': str(uuid5(NAMESPACE_URL, 'gmail-client:'+key)), 'name': name,
                      'section': section, 'account_manager_id': job['account_manager_id'],
                      'operator_id': job.get('operator_id'), 'created_at': job.get('created_at', now())}
            clients.append(client)
            by_key[key] = client
        job['client_id'] = by_key[key]['id']
    # Upgrade once, before clearing the old placeholder. The original demand
    # survives burned/deleted rows; replacements fill it instead of adding to it.
    history_counts = {}
    for job in state['setups']:
        history_counts[job['client_id']] = history_counts.get(job['client_id'], 0)+1
    for client in clients:
        if 'domain_target' not in client:
            client['domain_target'] = max(client.get('planned_domains', 0), history_counts.get(client['id'], 0))
        if history_counts.get(client['id']):
            client['planned_domains'] = 0


def client_progress(state, client, jobs=None):
    jobs = [job for job in (state['setups'] if jobs is None else jobs)
            if job.get('client_id') == client['id'] and not job.get('archived')]
    complete = sum(job['stage'] == STAGES[2] for job in jobs)
    actual_todo = sum(job['stage'] not in CLOSED_STAGES for job in jobs)
    target = client.get('domain_target', 0)
    remaining = max(0, target-complete)
    replacements = max(0, remaining-actual_todo)
    return {'target': target, 'complete': complete, 'actual_todo': actual_todo,
            'todo': remaining, 'remaining': remaining, 'replacements_needed': replacements}


def client_progress_map(state):
    grouped = {}
    for job in state['setups']:
        grouped.setdefault(job.get('client_id'), []).append(job)
    return {client['id']: client_progress(state, client, grouped.get(client['id'], []))
            for client in state['clients'] if not client.get('archived')}


def grow_client_target(state, client):
    healthy = sum(job.get('client_id') == client['id'] and not job.get('archived')
                  and job['stage'] not in INACTIVE_STAGES for job in state['setups'])
    client['domain_target'] = max(client.get('domain_target', 0), healthy)
    if healthy:
        client['planned_domains'] = 0


def detach_master_children(state, master):
    for job in state['setups']:
        if job['id'] != master['id'] and job.get('client_id') == master['client_id'] \
                and not job.get('archived') and job.get('master_domain_id') == master['id']:
            job.update(master_domain_id=None, updated_at=now())
            job.pop('master_color', None)


def set_master_domain(state, job, master_id):
    """Change grouping labels without changing a domain's setup workflow."""
    find(state, 'clients', job['client_id'], True)
    roots = [row for row in state['setups'] if row.get('client_id') == job['client_id']
             and not row.get('archived') and row.get('master_domain_id') == row['id']]
    if master_id == job['id']:
        existing = job.get('master_domain_id') == job['id']
        if not existing and len(roots) >= MAX_MASTER_DOMAINS:
            raise HTTPException(409, 'Each client can have up to 5 master domains. Remove a Master label before adding another.')
        color = job.get('master_color') if existing else None
        if type(color) is not int or not 0 <= color < MAX_MASTER_DOMAINS:
            used = {root.get('master_color') for root in roots if root['id'] != job['id']}
            color = next((value for value in range(MAX_MASTER_DOMAINS) if value not in used), None)
            if color is None:
                raise HTTPException(409, 'No master domain colour is available. Remove a Master label before adding another.')
        job.update(master_domain_id=job['id'], master_color=color, updated_at=now())
        return
    if master_id is not None:
        master = find(state, 'setups', master_id, True)
        if master.get('client_id') != job['client_id']:
            raise HTTPException(409, 'Choose a master domain from this client.')
        if master.get('master_domain_id') != master['id']:
            raise HTTPException(409, 'Choose an existing master domain. Domains cannot link through another child domain.')
    # Validate the new destination before touching the previous group's labels.
    if job.get('master_domain_id') == job['id']:
        detach_master_children(state, job)
    job.update(master_domain_id=master_id, updated_at=now())
    job.pop('master_color', None)


def find(state, collection, identifier, active=False):
    row = next((r for r in state[collection] if r['id'] == identifier), None)
    if not row or (active and (row.get('archived') or row.get('active') is False)):
        raise HTTPException(404, collection.rstrip('s').capitalize()+' not found or removed.')
    return row


def device_uses(device):
    return max(device.get('use_count', 0), int(bool(device.get('used_at') or device['status'] == 'Used')))


def device_capacity(state, device):
    return state.get('device_capacity', {}).get(device['kind'], DEFAULT_CAPACITY[device['kind']])


def ensure_device_usage(state):
    """Recover known historical uses without counting a stage transition twice."""
    known = {}
    reversed_uses = {}
    for job in state['setups']:
        for device_id in job.get('device_usage_resets', {}):
            reversed_uses.setdefault(device_id, set()).add(job['id'])
        snapshots = [job.get('purchase_device') or {}]
        snapshots += job.get('purchase_device_history', [])
        snapshots += [failure.get('device') or {} for failure in job.get('device_failures', [])]
        for snapshot in snapshots:
            if snapshot.get('id') and snapshot['id'] not in job.get('device_usage_resets', {}):
                known.setdefault(snapshot['id'], set()).add(job['id'])
    for d in state['devices']:
        previous_ids = set(d.get('setup_ids', []))
        removed = previous_ids & reversed_uses.get(d['id'], set())
        setup_ids = (previous_ids | known.get(d['id'], set())) - reversed_uses.get(d['id'], set())
        d['setup_ids'] = sorted(setup_ids)
        d['use_count'] = max(device_uses(d)-len(removed), len(setup_ids))
        if d['use_count'] == 0:
            d['used_at'] = None
            if d['status'] == 'Used':
                d['status'] = 'Unused'
    # Repair older resets which detached the phone but left its failed state
    # behind. This is a read-time upgrade; the next CAS mutation persists it.
    for job in state['setups']:
        if job['stage'] in PLANNING_STAGES and not job.get('archived') and not job.get('device_id'):
            failure = next((f for f in reversed(job.get('device_failures', [])) if f.get('device', {}).get('id') and not f.get('reverted_at')), None)
            if failure:
                device = next((d for d in state['devices'] if d['id'] == failure['device']['id']), None)
                if device:
                    restore_failed_device(state, job, device)


def restore_failed_device(state, job, device):
    failures = [f for f in job.get('device_failures', []) if f.get('device', {}).get('id') == device['id'] and not f.get('reverted_at')]
    if not failures:
        return
    stamp = now()
    # A previous successful completion remains a real use; otherwise undo
    # this setup's failed attempt without removing any other domain's use.
    successful_use = bool(job.get('completed_at') and (job.get('purchase_device') or {}).get('id') == device['id'])
    if not successful_use:
        ids = device.get('setup_ids', [])
        device['use_count'] = max(0, device_uses(device)-int(job['id'] in ids))
        device['setup_ids'] = [identifier for identifier in ids if identifier != job['id']]
        job.setdefault('device_usage_resets', {})[device['id']] = stamp
        if not device['use_count']:
            device['used_at'] = None
    others = [s for s in state['setups'] if s['id'] != job['id'] and failure_stage(s) and s.get('device_id') == device['id'] and device['id'] not in s.get('device_usage_resets', {})]
    own_block = not device.get('manual_block') and (job['id'] in device.get('block_job_ids', []) or device.get('block_reason') == FAILURES.get(failures[-1]['stage']))
    device['block_job_ids'] = [identifier for identifier in device.get('block_job_ids', []) if identifier != job['id']]
    if device['status'] == 'Blocked' and own_block and not others:
        device.update(status='Used' if device['use_count'] else 'Unused', block_reason=None, blocked_at=None)
    elif device['status'] == 'Blocked' and others and not device.get('manual_block'):
        device['block_reason'] = FAILURES[failure_stage(others[-1])]
    elif device['status'] == 'Used' and not device['use_count']:
        device['status'] = 'Unused'
    for failure in failures:
        failure['reverted_at'] = stamp


def allocation_available(state, device, job_id=None):
    if device.get('archived') or device['status'] == 'Blocked':
        return False
    reserved = sum(j['id'] != job_id and not j.get('archived') and j['stage'] in PLANNING_STAGES
        and j.get('device_id') == device['id'] and j['id'] not in device.get('setup_ids', []) for j in state['setups'])
    return device_uses(device) + reserved < device_capacity(state, device)


def consume_device_use(state, device, job):
    job_id = job['id']
    job.get('device_usage_resets', {}).pop(device['id'], None)
    if job_id not in device.get('setup_ids', []):
        capacity = device_capacity(state, device)
        if device_uses(device) >= capacity:
            count = {1: 'one', 2: 'two'}.get(capacity, str(capacity))
            raise HTTPException(409, f'This {device["kind"]} has already reached its limit of {count} {"setup" if capacity == 1 else "setups"}. Choose another device or allocate a device later.')
        device['use_count'] = device_uses(device) + 1
        device.setdefault('setup_ids', []).append(job_id)
    device['used_at'] = device.get('used_at') or now()


def unlocked(state, allow_sheet=False):
    lock = state.get('sync_lock')
    if lock and (datetime.now(timezone.utc)-datetime.fromisoformat(lock['at'])).total_seconds() < LEASE_SECONDS:
        if allow_sheet and lock.get('kind') == 'sheet':
            return
        if lock.get('kind') == 'dns_setup':
            raise HTTPException(409, 'DNS records are being configured. Retry in a moment.')
        if lock.get('kind') == 'dmarc':
            raise HTTPException(409, 'A DMARC record is being configured. Retry in a moment.')
        raise HTTPException(409, 'A completed setup is syncing to Google Sheets. Retry in a moment.')


def audit(state, user, action):
    state['events'] = (state['events']+[{'at': now(), 'by': user['email'], 'action': action}])[-200:]


def reserve_plan(state):
    """Compare setup slots with to do domains plus ten contingency setups."""
    # Operations reads the stored board directly. Resolve legacy client links
    # before forecasting so it agrees with the client view and never adds both.
    ensure_clients(state)
    progress = client_progress_map(state)
    jobs = [j for j in state['setups'] if j.get('client_id') in progress
            and not j.get('archived') and j['stage'] not in CLOSED_STAGES]
    planned = sum(client['replacements_needed'] for client in progress.values())
    todo = sum(client['remaining'] for client in progress.values())
    usable = {d['id']: d for d in state['devices'] if not d.get('archived') and d['status'] != 'Blocked'}
    remaining = {kind: sum(max(0, device_capacity(state, d)-device_uses(d))
        for d in usable.values() if d['kind'] == kind) for kind in DEFAULT_CAPACITY}
    # A purchase-stage domain can finish on its already-consumed slot. A domain
    # explicitly reset to Not started is conservatively treated as new work.
    purchased = sum(j['stage'] == STAGES[1] and j.get('device_id') in usable
        and j['id'] in usable[j['device_id']].get('setup_ids', [])
        and (j.get('purchase_device') or {}).get('id') == j['device_id'] for j in jobs)
    order_kind = state.get('order_device_kind', 'Phone')
    per_new = device_capacity(state, {'kind': order_kind})
    target = todo+RESERVE_CONTINGENCY
    available = sum(remaining.values())+purchased
    missing = max(0, target-available)
    return {'todo': todo, 'actual_todo': len(jobs), 'planned_todo': planned,
            'contingency': RESERVE_CONTINGENCY, 'target': target,
            'available_setups': available, 'setup_shortfall': missing,
            'phone_available': remaining['Phone'], 'pc_available': remaining['PC'],
            'purchased_setups': purchased, 'shortfall': (missing+per_new-1)//per_new,
            'phone_setups': device_capacity(state, {'kind': 'Phone'}),
            'pc_setups': device_capacity(state, {'kind': 'PC'}),
            'order_kind': order_kind, 'order_capacity': per_new}


def summary(state):
    ensure_device_usage(state)
    devices = [d for d in state['devices'] if not d.get('archived')]
    unused = sum(d['status'] == 'Unused' and not device_uses(d) for d in devices)
    plan = reserve_plan(state)
    return {**plan, 'unused': unused,
            'low_stock': plan['setup_shortfall'] > 0, 'devices': len(devices),
            'blocked': sum(d['status'] == 'Blocked' for d in devices),
            'exhausted_phones': sum(d['kind'] == 'Phone' and device_uses(d) >= device_capacity(state, d) for d in devices),
            'not_started': plan['planned_todo']+sum(s['stage'] == STAGES[0] or s['stage'] in FAILURES and not s.get('purchase_at') for s in state['setups'] if not s.get('archived')),
            'started': sum(s['stage'] == STARTED for s in state['setups'] if not s.get('archived')),
            'purchased': sum(s['stage'] == STAGES[1] or s['stage'] in FAILURES and bool(s.get('purchase_at')) for s in state['setups'] if not s.get('archived')),
            'complete': sum(s['stage'] == STAGES[2] for s in state['setups'] if not s.get('archived')),
            'burned': sum(s['stage'] == BURNED for s in state['setups'] if not s.get('archived')),
            'gmail_deleted': sum(s['stage'] == GMAIL_DELETED for s in state['setups'] if not s.get('archived')),
            'sync_errors': sum(s.get('sync', {}).get('status') == 'error' for s in state['setups'] if not s.get('archived') and s['stage'] not in INACTIVE_STAGES)}


def result(record):
    state = copy.deepcopy(record['state'])
    # Lease tokens remain server-side; public state is not accepted as a mutation.
    state.pop('sync_lock', None)
    from . import passwords
    passwords.scrub_public_clients(state)
    for job in state['setups']:
        job['sheet_identity_locked'] = sheet_identity_locked(job)
        job.pop('sheet_identity', None)
        job.pop('sheet_intent', None)
    plan = summary(state)
    return {'state': state, 'revision': record['revision'], 'updated_at': record['updated_at'], 'summary': plan,
            'client_progress': client_progress_map(state),
            **({'import_summary': record['import_summary']} if 'import_summary' in record else {})}


def reveal_client_password(user, client_id):
    from . import passwords
    client = find(board(user)['state'], 'clients', client_id, True)
    if not client.get('password_ciphertext'):
        raise HTTPException(404, 'No password is saved for this client.')
    return passwords.reveal(user, client_id, client['password_ciphertext'])


def sheet_identity_locked(job):
    return bool(job.get('sheet_identity') or job.get('sync', {}).get('status') == 'synced' and not job.get('sync', {}).get('missing'))


def text(value, label, required=True):
    value = str(value or '').strip()
    if required and not value:
        raise HTTPException(400, label+' is required.')
    return value


def email_domain(raw, domain):
    raw = text(raw, 'Sending email', False).lower()
    if not raw:
        return ''
    if not re.fullmatch(r'[a-z0-9][a-z0-9.!#$%&\x27*+/=?^_`{|}~-]{0,63}@[a-z0-9.-]+', raw) or raw.count('@') != 1:
        raise HTTPException(400, 'Enter a valid sending email address, such as me@your-domain.com.')
    try:
        matches = normalize(raw.rsplit('@', 1)[1]) == domain
    except ValueError:
        matches = False
    if not matches:
        raise HTTPException(400, 'The sending email must use the tracked domain.')
    return raw


def device_slot(state, member_id, kind, exclude=None):
    if member_id:
        find(state, 'members', member_id, True)
    occupied = {d['number'] for d in state['devices'] if not d.get('archived') and d['member_id'] == member_id and d['kind'] == kind and d['id'] != exclude}
    number = next((n for n in range(1, 11) if n not in occupied), None)
    if number is None:
        raise HTTPException(400, 'A team member can have up to 10 phones and 10 PCs. Remove or reassign a device first.')
    return number


def mutate(user, expected, action, data):
    record = board(user)
    if record['revision'] != expected:
        raise HTTPException(409, 'Someone updated the team tracker. Refresh before saving again.', headers=store.REVISION_CONFLICT_HEADERS)
    state = copy.deepcopy(record['state'])
    unlocked(state, allow_sheet=True)
    lock = state.get('sync_lock') or {}
    if lock.get('kind') != 'sheet' or (datetime.now(timezone.utc)-datetime.fromisoformat(lock['at'])).total_seconds() >= LEASE_SECONDS:
        state['sync_lock'] = None
    outcome = None
    if action == 'reserve_set':
        raise HTTPException(409, 'The device target is automatic. Change the phone or laptop setup limits instead, then refresh.')
    elif action == 'capacity_set':
        limits = {'Phone': data['phone_setups'], 'PC': data['pc_setups']}
        if any(type(v) is not int or not 1 <= v <= 100 for v in limits.values()) or data['order_kind'] not in limits:
            raise HTTPException(400, 'Choose setup limits from 1 to 100 and a device type to order.')
        state.update(device_capacity=limits, order_device_kind=data['order_kind'])
    elif action == 'client_save':
        current = find(state, 'clients', data['id'], True) if data.get('id') else None
        data = {'section': (current or {}).get('section', 'client'), **data}
        target = data.get('domain_target')
        if target is None:
            target = data.get('planned_domains')
        if target is None:
            target = (current or {}).get('domain_target', 0)
        if type(target) is not int or not 0 <= target <= 5000:
            raise HTTPException(400, 'Enter a domain target from 0 to 5000.')
        healthy = sum(job.get('client_id') == current['id'] and not job.get('archived') and job['stage'] not in INACTIVE_STAGES
                      for job in state['setups']) if current else 0
        if target < healthy:
            raise HTTPException(400, f'Target must cover {healthy} current domains.')
        metadata = {}
        if 'sender_full_name' in data:
            metadata['sender_full_name'] = text(data['sender_full_name'], 'Sender full name', False)
        if 'account_manager_email' in data:
            email = text(data['account_manager_email'], 'Account manager email', False).lower()
            if email and not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+', email):
                raise HTTPException(400, 'Enter a valid account manager email address.')
            metadata['account_manager_email'] = email
        if 'website_url' in data:
            from .dns_setup import destination_url
            website = text(data['website_url'], 'Client website', False)
            metadata['website_url'] = destination_url(website, '') if website else ''
        planned = 0 if current and any(j.get('client_id') == current['id'] for j in state['setups']) else target
        client_name = re.sub(r'\s+', ' ', text(data['name'], 'Client name'))
        key = client_key(data['section'], client_name)
        if any(not c.get('archived') and c['id'] != (current or {}).get('id') and client_key(c['section'], c['name']) == key for c in state['clients']):
            raise HTTPException(409, 'This client already exists. Open the client to add another domain.')
        if current is None and len(state['clients']) >= 1000:
            raise HTTPException(400, 'Client limit reached.')
        find(state, 'members', data['account_manager_id'], data['account_manager_id'] != (current or {}).get('account_manager_id'))
        if data.get('operator_id'):
            find(state, 'members', data['operator_id'], True)
        if current:
            if current['section'] != data['section']:
                raise HTTPException(409, 'The sheet section cannot change. Create a separate client for a different section.')
            for job in state['setups']:
                if job.get('client_id') == current['id']:
                    if job['stage'] == STAGES[2]:
                        job.setdefault('sheet_owner', job['client_name'] if job['section'] == 'client' else find(state, 'members', job['operator_id'])['name'])
                    else:
                        job['account_manager_id'] = data['account_manager_id']
                    job['client_name'] = client_name
            current.update(name=client_name, account_manager_id=data['account_manager_id'],
                           operator_id=data.get('operator_id', current.get('operator_id')), planned_domains=planned,
                           domain_target=target, updated_at=now(), **metadata)
        else:
            state['clients'].append({'id': new_id(), 'name': client_name, 'section': data['section'],
                                     'account_manager_id': data['account_manager_id'], 'operator_id': data.get('operator_id'),
                                     'planned_domains': planned, 'domain_target': target, 'created_at': now(), **metadata})
    elif action == 'client_password':
        from . import passwords
        client = find(state, 'clients', data['id'], True)
        ciphertext = passwords.seal(user, client['id'], data['password'])
        if ciphertext is None:
            client.pop('password_ciphertext', None)
        else:
            client['password_ciphertext'] = ciphertext
        client['password_updated_at'] = now()
    elif action == 'client_remove':
        client = find(state, 'clients', data['id'], True)
        client.update(archived=True, deleted_at=now())
        for job in state['setups']:
            if job.get('client_id') == client['id']:
                job.update(archived=True, deleted_at=now())
    elif action == 'client_complete':
        client = find(state, 'clients', data['id'], True)
        jobs = [job for job in state['setups'] if job.get('client_id') == client['id']
                and not job.get('archived') and job['stage'] not in CLOSED_STAGES]
        # Validate the entire client before the single CAS save. A missing
        # email, owner or valid purchase device must never partially complete it.
        for job in jobs:
            try:
                set_stage(state, job, STAGES[2])
            except HTTPException as exc:
                raise HTTPException(exc.status_code, f'{job["domain"]}: {exc.detail}') from None
        outcome = {'completed': len(jobs)}
    elif action == 'setup_bulk':
        client = find(state, 'clients', data['client_id'], True)
        sender = text(data.get('sender_name'), 'Sender name', False).lower().rstrip('@')
        if sender:
            try:
                email_domain(sender+'@example.com', 'example.com')
            except HTTPException:
                raise HTTPException(400, 'Enter the sender name before @, such as ashley or ashley.smith.') from None
        operator_id = data.get('operator_id')
        if operator_id:
            find(state, 'members', operator_id, True)
        parsed = []
        for line, content in enumerate(data['domains'].splitlines(), 1):
            for token in re.split(r'[,;\s]+', content.strip()):
                if not token:
                    continue
                token = token.lower().rstrip('.')
                try:
                    raw_domain = token.rsplit('@', 1)[-1]
                    domain = normalize(raw_domain)
                    if domain != raw_domain:
                        raise ValueError()
                    email = email_domain(token, domain) if '@' in token else sender+'@'+domain if sender else ''
                except (ValueError, HTTPException):
                    raise HTTPException(400, f'Line {line}: enter a valid root domain or email address.') from None
                parsed.append((domain, email))
        if not 1 <= len(parsed) <= 500:
            raise HTTPException(400, 'Paste between 1 and 500 domains or email addresses.')
        existing = {j['domain'] for j in state['setups'] if not j.get('archived')}
        added = 0
        for domain, email in parsed:
            if domain in existing:
                continue
            if len(state['setups']) >= 5000:
                raise HTTPException(400, 'Setup history limit reached.')
            state['setups'].append({'id': new_id(), 'client_id': client['id'], 'client_name': client['name'],
                'section': client['section'], 'account_manager_id': client['account_manager_id'], 'operator_id': operator_id,
                'domain': domain, 'email': email, 'device_id': None, 'order': None, 'notes': '', 'stage': STAGES[0],
                'created_at': now(), 'updated_at': now(), 'purchase_at': None, 'purchase_device': None,
                'completed_at': None, 'sync': {'status': 'not_ready'}, 'archived': False})
            existing.add(domain)
            added += 1
        if added:
            grow_client_target(state, client)
        outcome = {'added': added, 'skipped': len(parsed)-added}
    elif action == 'member_save':
        if data.get('id'):
            member = find(state, 'members', data['id'], True)
        else:
            if len(state['members']) >= 500:
                raise HTTPException(400, 'Team limit reached.')
            member = {'id': new_id(), 'active': True}
            state['members'].append(member)
        name = text(data.get('name'), 'Team member name')
        email = text(data.get('email'), 'Team email', False).lower()
        if email and not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+', email):
            raise HTTPException(400, 'Enter a valid team email address.')
        if any(m['id'] != member['id'] and m['active'] and (m['name'].casefold() == name.casefold() or email and m.get('email') == email) for m in state['members']):
            raise HTTPException(400, 'This team member already exists.')
        member.update(name=name, email=email)
    elif action == 'member_remove':
        member = find(state, 'members', data['id'], True)
        if any(s['stage'] not in CLOSED_STAGES and not s.get('archived') and member['id'] in (s['account_manager_id'], s['operator_id']) for s in state['setups']):
            raise HTTPException(409, 'Reassign this member’s open setups before removing them.')
        member['active'] = False
        for d in state['devices']:
            if d['member_id'] == member['id']:
                d['number'] = 1+max((x['number'] for x in state['devices'] if x['member_id'] is None and x['kind'] == d['kind']), default=0)
                d['member_id'] = None
                # Physical IDs and purchase snapshots stay unchanged.
        
    elif action == 'devices_add':
        member_id = data.get('member_id')
        if member_id:
            find(state, 'members', member_id, True)
        count = data['phones']+data['pcs']
        if count < 1:
            raise HTTPException(400, 'Choose at least one device to add.')
        if len(state['devices'])+count > 5000:
            raise HTTPException(400, 'Device history limit reached.')
        status = data.get('status', 'Unused')
        use_count = data.get('use_count')
        if use_count is None:
            use_count = 1 if status == 'Used' else 0
        if type(use_count) is not int or not 0 <= use_count <= 100 or status == 'Unused' and use_count or status == 'Used' and not use_count:
            raise HTTPException(400, 'Choose a usage count that matches the device condition.')
        for kind, amount in [('Phone', data['phones']), ('PC', data['pcs'])]:
            for _ in range(amount):
                number = device_slot(state, member_id, kind) if member_id else 1+max((d['number'] for d in state['devices'] if d['member_id'] is None and d['kind'] == kind), default=0)
                state['devices'].append({'id': new_id(), 'asset_tag': 'DEV-'+uuid4().hex[:8].upper(), 'kind': kind, 'number': number,
                    'member_id': member_id, 'status': status, 'block_reason': (data.get('block_reason') or 'Other') if status == 'Blocked' else None, 'used_at': now() if use_count else None,
                    'blocked_at': now() if status == 'Blocked' else None, 'archived': False, 'added_at': now(),
                    'use_count': use_count, 'setup_ids': [], 'manual_block': status == 'Blocked'})
    elif action == 'device_update':
        d = find(state, 'devices', data['id'], True)
        if 'member_id' in data and data['member_id'] != d['member_id']:
            if any(s['stage'] in PLANNING_STAGES and not s.get('archived') and s.get('device_id') == d['id'] for s in state['setups']):
                raise HTTPException(409, 'Reassign this device on its pending setup before transferring it.')
            member_id = data['member_id']
            number = device_slot(state, member_id, d['kind'], d['id']) if member_id else 1+max((x['number'] for x in state['devices'] if x['member_id'] is None and x['kind'] == d['kind']), default=0)
            d.update(member_id=member_id, number=number)
        if 'status' in data:
            status = data['status']
            if d['status'] == 'Blocked' and status != 'Blocked':
                raise HTTPException(409, 'Blocked devices cannot be returned to service. Add a different physical device.')
            if status == 'Unused' and device_uses(d):
                raise HTTPException(409, 'A used device cannot count as new stock again.')
            if status == 'Blocked' and (d['status'] != 'Blocked' or data.get('block_reason', d.get('block_reason')) != d.get('block_reason')):
                d['manual_block'] = True
            d['status'] = status
            if status == 'Blocked':
                d['blocked_at'] = d['blocked_at'] or now()
                d['block_reason'] = data.get('block_reason') or d.get('block_reason') or 'Other'
            if status == 'Used':
                d['used_at'] = d['used_at'] or now()
                d['use_count'] = max(device_uses(d), 1)
        if 'asset_tag' in data:
            tag = text(data['asset_tag'], 'Asset label')
            if any(x['id'] != d['id'] and x['asset_tag'].casefold() == tag.casefold() for x in state['devices']):
                raise HTTPException(400, 'That physical asset label already exists.')
            d['asset_tag'] = tag
    elif action == 'device_remove':
        d = find(state, 'devices', data['id'], True)
        if any(s['stage'] not in CLOSED_STAGES and not s.get('archived') and s.get('device_id') == d['id'] for s in state['setups']):
            raise HTTPException(409, 'This device is linked to an open setup. Reassign or archive that setup first.')
        d['archived'] = True
    elif action in {'setup_save', 'setup_assign'}:
        current = find(state, 'setups', data['id'], True) if data.get('id') else None
        confirm_gmail_deleted(current, data.get('stage'), data.get('confirmed_auto_renew_stopped', False))
        if current and current['stage'] in INACTIVE_STAGES and data.get('stage') not in {STAGES[0], *INACTIVE_STAGES}:
            raise HTTPException(409, 'Choose Not started to reopen this inactive domain before editing it.')
        if current and sheet_identity_locked(current):
            if action == 'setup_save' and data.get('stage'):
                same = all(data.get(key) == current.get(key) for key in ('account_manager_id', 'operator_id', 'device_id'))
                same = same and text(data.get('email'), 'Sending email', False).lower() == current['email']
                same = same and text(data.get('notes'), 'Notes', False) == current.get('notes', '')
                same = same and data.get('section') == current['section']
                same = same and data.get('client_id') in (None, current['client_id'])
                same = same and (not data.get('domain') or str(data['domain']).lower().rstrip('.') == current['domain'])
                if same:
                    set_stage(state, current, data['stage'], data.get('confirmed_auto_renew_stopped', False))
                    grow_client_target(state, find(state, 'clients', current['client_id'], True))
                    audit(state, user, action)
                    return store.save(user, state, expected)
            if action != 'setup_assign' or current['stage'] not in PLANNING_STAGES:
                raise HTTPException(409, 'A published setup is locked to its sheet record. Correct the sheet directly if needed.')
            # Reopening releases the workflow assignment, not its published
            # identity. Only a new setup person or pending device can change.
            if set(data) - {'id', 'operator_id', 'device_id', 'email'} or 'email' in data and data['email'] != current['email']:
                raise HTTPException(409, 'The published sending email and sheet record cannot be changed. Only the setup person and device can be reassigned after reopening.')
            operator_id = data.get('operator_id', current.get('operator_id'))
            operator = find(state, 'members', operator_id, True) if operator_id else None
            device_id = data.get('device_id', None if operator_id != current.get('operator_id') else current.get('device_id'))
            if device_id:
                device = find(state, 'devices', device_id, True)
                historical = current.get('device_id') == device_id and (current.get('purchase_device') or {}).get('id') == device_id
                if device['status'] == 'Blocked' or not historical and not allocation_available(state, device, current['id']):
                    raise HTTPException(409, 'This device is blocked, fully used or reserved. Choose another device or allocate one later.')
                if not operator or device['member_id'] != operator['id']:
                    raise HTTPException(409, 'Allocate the device to the setup owner first.')
            current.update(operator_id=operator_id, device_id=device_id, updated_at=now())
            audit(state, user, action)
            return store.save(user, state, expected)
        if action == 'setup_assign':
            if current is None:
                raise HTTPException(404, 'Setup not found.')
            patch = data
            if 'email' in patch and patch['email'] is None:
                raise HTTPException(400, 'Enter an email address or leave it blank while planning.')
            if 'operator_id' in patch and patch['operator_id'] != current.get('operator_id'):
                if current.get('purchase_device') and current['stage'] in STAGES[1:]:
                    raise HTTPException(409, 'The purchase owner is recorded in history. It cannot be reassigned.')
                patch = {'device_id': None, **patch}
            data = {**current, **patch}
            if current['stage'] in FAILURES and data.get('device_id') != current.get('device_id'):
                data['stage'] = STAGES[1] if current.get('purchase_at') else STAGES[0]
        client = None
        if data.get('client_id'):
            client = find(state, 'clients', data['client_id'], True)
            if current and current.get('client_id') != client['id']:
                raise HTTPException(409, 'Edit this domain within its existing client.')
            data = {**data, 'client_name': client['name'], 'section': client['section']}
        # Existing client assignments survive removal from the setup team.
        # New manager assignments and all operators must still be active.
        existing_manager = (client or current or {}).get('account_manager_id')
        manager = find(state, 'members', data['account_manager_id'], data['account_manager_id'] != existing_manager)
        operator = find(state, 'members', data['operator_id'], True) if data.get('operator_id') else None
        try:
            supplied_domain = text(data.get('domain'), 'Domain', False).lower().rstrip('.')
            raw_email = text(data.get('email'), 'Sending email', False).lower()
            if not supplied_domain and ('@' not in raw_email or raw_email.count('@') != 1):
                raise HTTPException(400, 'Enter a valid sending email address.')
            raw_domain = supplied_domain or raw_email.rsplit('@', 1)[1]
            domain = normalize(raw_domain)
            if supplied_domain and domain != raw_domain:
                raise ValueError()
        except ValueError:
            raise HTTPException(400, 'Enter the root domain only, such as your-company.com.') from None
        email = email_domain(data.get('email'), domain)
        if data['section'] == 'client' and not text(data.get('client_name'), 'Client name', False):
            raise HTTPException(400, 'Enter the client name for the client sheet section.')
        if any(s['id'] != (current or {}).get('id') and not s.get('archived') and (email and s['email'] == email or not email and not s['email'] and s['domain'] == domain) for s in state['setups']):
            raise HTTPException(409, 'This sending email or domain-only draft already has a setup. Update that record instead.')
        device_id = data.get('device_id')
        if current and current.get('purchase_device') and current['stage'] in STAGES[1:] and data.get('operator_id') != current.get('operator_id'):
            raise HTTPException(409, 'The purchase owner is recorded in history. It cannot be reassigned.')
        if current and current.get('purchase_device') and current['stage'] in STAGES[1:] and device_id != current.get('device_id'):
            raise HTTPException(409, 'The purchase device is recorded in history and cannot be changed after purchase.')
        if device_id:
            historical = current and current.get('device_id') == device_id and (current['stage'] not in PLANNING_STAGES or (current.get('purchase_device') or {}).get('id') == device_id)
            device = find(state, 'devices', device_id, not historical)
            if not historical:
                if not allocation_available(state, device, (current or {}).get('id')):
                    raise HTTPException(409, 'This device is blocked, fully used or reserved. Choose another device or allocate one later.')
                if not operator or device['member_id'] != operator['id']:
                    raise HTTPException(409, 'Allocate the device to the setup owner first.')
        if current is None:
            if len(state['setups']) >= 5000:
                raise HTTPException(400, 'Setup history limit reached.')
            current = {'id': new_id(), 'stage': STAGES[0], 'created_at': now(), 'purchase_at': None, 'completed_at': None,
                       'purchase_device': None, 'order': None, 'sync': {'status': 'not_ready'}, 'archived': False}
            state['setups'].append(current)
        if (current['stage'] not in PLANNING_STAGES or current.get('purchase_at')) and domain != current['domain']:
            raise HTTPException(409, 'The purchased domain cannot be changed. Create a new setup for a different domain.')
        client = next((c for c in state['clients'] if c['id'] == data.get('client_id')), None)
        if client is None:
            # Compatibility for previously opened forms and legacy API callers.
            display_name = text(data.get('client_name'), 'Client name', False) or 'Biz Dev'
            key = client_key(data['section'], display_name)
            client = next((c for c in state['clients'] if not c.get('archived') and client_key(c['section'], c['name']) == key), None)
            if client is None:
                client = {'id': new_id(), 'name': display_name, 'section': data['section'], 'account_manager_id': manager['id'], 'operator_id': operator['id'] if operator else None, 'created_at': now()}
                state['clients'].append(client)
        if current.get('client_id') and current['client_id'] != client['id']:
            # Legacy forms can resolve a different client by name. Group labels
            # must never follow that move across client boundaries.
            if current.get('master_domain_id') == current['id']:
                detach_master_children(state, current)
            current['master_domain_id'] = None
            current.pop('master_color', None)
        current['client_id'] = client['id']
        client['planned_domains'] = 0
        current.update(domain=domain, email=email, client_name=text(data.get('client_name'), 'Client name', False),
                       section=data['section'], account_manager_id=manager['id'], operator_id=operator['id'] if operator else None,
                       device_id=device_id, notes=text(data.get('notes'), 'Notes', False), updated_at=now())
        set_stage(state, current, data.get('stage') or current['stage'], data.get('confirmed_auto_renew_stopped', False))
        grow_client_target(state, client)
    elif action == 'setup_device':
        # A missing device can be recorded later, including after a sheet sync,
        # without permitting edits to the published email, owner or sheet row.
        job = find(state, 'setups', data['id'], True)
        if job['stage'] in INACTIVE_STAGES:
            raise HTTPException(409, 'Choose Not started to reopen this inactive domain before allocating a device.')
        if job.get('device_id'):
            raise HTTPException(409, 'This setup already has a device. Use Edit details to change a pending allocation.')
        if sheet_identity_locked(job) and job['stage'] not in PLANNING_STAGES and data['operator_id'] != job['operator_id']:
            raise HTTPException(409, 'Keep the published setup owner. Allocate the device to that team member first.')
        operator = find(state, 'members', data['operator_id'], True)
        device = find(state, 'devices', data['device_id'], True)
        if not allocation_available(state, device, job['id']) or device['member_id'] != operator['id']:
            raise HTTPException(409, 'Choose an available device allocated to the setup owner, or add a device later.')
        if job['stage'] in FAILURES:
            raise HTTPException(409, 'Use Edit details to choose a replacement device and retry stage.')
        job.update(device_id=device['id'], operator_id=operator['id'], updated_at=now())
        if job['stage'] in STAGES[1:]:
            record_purchase_device(state, job)
    elif action == 'setup_stage':
        job = find(state, 'setups', data['id'], True)
        set_stage(state, job, data['stage'], data.get('confirmed_auto_renew_stopped', False))
        grow_client_target(state, find(state, 'clients', job['client_id'], True))
    elif action == 'setup_notes':
        job = find(state, 'setups', data['id'], True)
        if not isinstance(data.get('notes'), str) or len(data['notes']) > 1000:
            raise HTTPException(400, 'Keep notes to 1,000 characters or fewer.')
        job.update(notes=data['notes'], updated_at=now())
    elif action == 'setup_master':
        job = find(state, 'setups', data['id'], True)
        set_master_domain(state, job, data['master_domain_id'])
    elif action == 'setup_remove':
        job = find(state, 'setups', data['id'], True)
        if job.get('master_domain_id') == job['id']:
            detach_master_children(state, job)
        job.update(archived=True, deleted_at=now())
    else:
        raise HTTPException(400, 'Unknown tracker action.')
    audit(state, user, action)
    saved = store.save(user, state, expected)
    if outcome is not None:
        saved['import_summary'] = outcome
    return saved


def device_snapshot(state, device):
    member = find(state, 'members', device['member_id']) if device.get('member_id') else {'name': 'Unallocated'}
    return {'id': device['id'], 'asset_tag': device['asset_tag'],
            'label': device['kind']+' '+str(device['number']), 'member': member['name']}


def record_purchase_device(state, job):
    d = find(state, 'devices', job['device_id'], True)
    if d['status'] == 'Blocked':
        raise HTTPException(409, 'The allocated purchase device is blocked. Select another device.')
    if d['member_id'] != job['operator_id']:
        raise HTTPException(409, 'Allocate the device to the setup owner before recording purchase.')
    consume_device_use(state, d, job)
    d['status'] = 'Used'
    previous = job.get('purchase_device')
    if previous and previous['id'] != d['id']:
        job.setdefault('purchase_device_history', []).append(previous)
    job['purchase_device'] = device_snapshot(state, d)


def failure_stage(job):
    """Keep failed-device ownership through inactive labels and legacy boards."""
    stage = job['stage']
    if stage in INACTIVE_STAGES and job.get('inactive_failure_stage') in FAILURES:
        return job['inactive_failure_stage']
    visited = set()
    while stage in INACTIVE_STAGES and stage not in visited:
        visited.add(stage)
        stage = job.get('burned_from' if stage == BURNED else 'gmail_deleted_from')
    return stage if stage in FAILURES else None


def confirm_gmail_deleted(job, stage, confirmed):
    if stage == GMAIL_DELETED and (job or {}).get('stage') != GMAIL_DELETED and confirmed is not True:
        raise HTTPException(400, 'Confirm that auto-renew has been stopped in Porkbun before marking this Gmail deleted.')


def set_stage(state, job, stage, confirmed_auto_renew_stopped=False):
    if stage not in (*STAGES, *FAILURES, STARTED, *INACTIVE_STAGES):
        raise HTTPException(400, 'Choose a valid setup stage.')
    old = job['stage']
    confirm_gmail_deleted(job, stage, confirmed_auto_renew_stopped)
    if old in CLOSED_STAGES and stage not in {old, STAGES[0], *INACTIVE_STAGES}:
        raise HTTPException(409, 'Choose Not started to reopen this setup before changing its stage.')
    if stage in INACTIVE_STAGES:
        # Retain real purchase, failure and sheet history without consuming a
        # new device slot or treating the abandoned domain as unfinished work.
        if old != stage:
            stamp = now()
            job['inactive_failure_stage'] = failure_stage(job)
            if stage == BURNED:
                job.update(burned_from=old, burned_at=stamp)
            else:
                job.update(gmail_deleted_from=old, gmail_deleted_at=stamp, auto_renew_stopped_at=stamp)
        job.update(stage=stage, updated_at=now())
        if stage == GMAIL_DELETED:
            from . import backup
            backup.queue(state, job, False, force=old != stage)
        return
    if stage in PLANNING_STAGES:
        # The workflow can be reopened independently of recorded purchases,
        # completed sheet writes and the physical phone's usage/block state.
        device = find(state, 'devices', job['device_id']) if job.get('device_id') else None
        was_blocked = device and device['status'] == 'Blocked'
        if failure_stage(job) and device:
            restore_failed_device(state, job, device)
        if device and (was_blocked or device.get('archived')):
            job['device_id'] = None
        job.update(stage=stage, updated_at=now())
        return
    if stage == STAGES[2] and not job['email']:
        raise HTTPException(400, 'Add the sending email address before completing setup.')
    if stage == STAGES[2] and not job.get('operator_id'):
        raise HTTPException(400, 'Choose the setup person before completing setup.')
    if stage == STAGES[2] and old != stage:
        try:
            find(state, 'members', job['operator_id'], True)
        except HTTPException:
            raise HTTPException(400, 'Choose an active setup person before completing setup.') from None
    if stage in FAILURES:
        reason = FAILURES[stage]
        d = find(state, 'devices', job['device_id'], True) if job.get('device_id') else None
        if d:
            if d['status'] != 'Blocked' and d['member_id'] != job['operator_id']:
                raise HTTPException(409, 'Allocate the device to the setup owner before recording failure.')
            consume_device_use(state, d, job)
            if d['status'] != 'Blocked':
                d['manual_block'] = False
            d['block_job_ids'] = sorted(set(d.get('block_job_ids', [])) | {job['id']})
            d.update(status='Blocked', block_reason=d.get('block_reason') if d.get('manual_block') else reason, blocked_at=d.get('blocked_at') or now())
        # Keep the failed physical asset even when the retry uses a replacement.
        history = job.setdefault('device_failures', [])
        if old != stage or not history or history[-1].get('device', {}).get('id') != job.get('device_id'):
            history.append({'stage': stage, 'at': now(), 'device': device_snapshot(state, d) if d else {}})
            job['device_failures'] = history[-200:]
    elif stage in STAGES[1:]:
        # Inventory is optional: tracking and sheet publishing must work with
        # zero stock. Only a real recorded device changes reserve counts.
        if old in FAILURES and job.get('device_id') and find(state, 'devices', job['device_id'], True)['status'] == 'Blocked':
            raise HTTPException(409, 'Choose a replacement device or allocate one later before retrying.')
        if job.get('device_id') and ((job.get('purchase_device') or {}).get('id') != job['device_id'] or job['device_id'] in job.get('device_usage_resets', {})):
            record_purchase_device(state, job)
        job['purchase_at'] = job.get('purchase_at') or now()
    if stage == STAGES[2] and not job.get('completed_at'):
        client = find(state, 'clients', job['client_id'], True)
        completed = [j for j in state['setups'] if j.get('client_id') == client['id'] and j.get('completed_at')]
        number = max(client.get('next_order', 1), len(completed)+1, max((j.get('order') or 0 for j in completed), default=0)+1)
        client['next_order'] = number+1
        job.update(completed_at=now(), order=number,
            sheet_owner=job['client_name'] if job['section'] == 'client' else find(state, 'members', job['operator_id'])['name'])
    job['stage'] = stage
    job['updated_at'] = now()
    if stage == STAGES[2]:
        from . import backup
        backup.queue(state, job, True, force=old != stage)


def sync_completed(user, job_id):
    """Synchronous worker primitive; HTTP handlers only schedule the worker."""
    from . import backup
    return backup.sync_once(user, job_id)
