"""Server-authoritative stages, device history, reserve counts and sync leases."""
import copy
import re
from datetime import datetime, timezone
from uuid import uuid4, uuid5, NAMESPACE_URL
from fastapi import HTTPException
from app.utils.domain_cluster.engine import normalize
from . import store

STAGES = ('Not started', 'Purchase made', 'Setup complete')
FAILURES = {'Failed at QR': 'QR code', 'Failed at payment': 'Payment'}
DEVICE_STATUSES = ('Unused', 'Used', 'Blocked')
RESERVE_TARGET = 20
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


def find(state, collection, identifier, active=False):
    row = next((r for r in state[collection] if r['id'] == identifier), None)
    if not row or (active and (row.get('archived') or row.get('active') is False)):
        raise HTTPException(404, collection.rstrip('s').capitalize()+' not found or removed.')
    return row


def device_uses(device):
    return max(device.get('use_count', 0), int(bool(device.get('used_at') or device['status'] == 'Used')))


def ensure_device_usage(state):
    """Recover known historical uses without counting a stage transition twice."""
    known = {}
    for job in state['setups']:
        snapshots = [job.get('purchase_device') or {}]
        snapshots += [failure.get('device') or {} for failure in job.get('device_failures', [])]
        for snapshot in snapshots:
            if snapshot.get('id'):
                known.setdefault(snapshot['id'], set()).add(job['id'])
    for d in state['devices']:
        setup_ids = set(d.get('setup_ids', [])) | known.get(d['id'], set())
        d['setup_ids'] = sorted(setup_ids)
        d['use_count'] = max(device_uses(d), len(setup_ids))


def allocation_available(state, device, job_id=None):
    if device.get('archived') or device['status'] == 'Blocked':
        return False
    reserved = sum(j['id'] != job_id and not j.get('archived') and j['stage'] == STAGES[0]
        and j.get('device_id') == device['id'] for j in state['setups'])
    if device['kind'] == 'Phone':
        return device_uses(device) + reserved < 2
    return reserved == 0


def consume_device_use(device, job_id):
    if job_id not in device.get('setup_ids', []):
        if device['kind'] == 'Phone' and device_uses(device) >= 2:
            raise HTTPException(409, 'This phone has already been used for two setups. Choose another phone or allocate a device later.')
        device['use_count'] = device_uses(device) + 1
        device.setdefault('setup_ids', []).append(job_id)
    device['used_at'] = device.get('used_at') or now()


def unlocked(state):
    lock = state.get('sync_lock')
    if lock and (datetime.now(timezone.utc)-datetime.fromisoformat(lock['at'])).total_seconds() < LEASE_SECONDS:
        raise HTTPException(409, 'A completed setup is syncing to Google Sheets. Retry in a moment.')


def audit(state, user, action):
    state['events'] = (state['events']+[{'at': now(), 'by': user['email'], 'action': action}])[-200:]


def summary(state):
    ensure_device_usage(state)
    devices = [d for d in state['devices'] if not d.get('archived')]
    unused = sum(d['status'] == 'Unused' and not device_uses(d) for d in devices)
    return {'unused': unused, 'target': RESERVE_TARGET, 'shortfall': max(0, RESERVE_TARGET-unused),
            'low_stock': unused < RESERVE_TARGET, 'devices': len(devices),
            'blocked': sum(d['status'] == 'Blocked' for d in devices),
            'exhausted_phones': sum(d['kind'] == 'Phone' and device_uses(d) >= 2 for d in devices),
            'not_started': sum(s['stage'] == STAGES[0] or s['stage'] in FAILURES and not s.get('purchase_at') for s in state['setups'] if not s.get('archived')),
            'purchased': sum(s['stage'] == STAGES[1] or s['stage'] in FAILURES and bool(s.get('purchase_at')) for s in state['setups'] if not s.get('archived')),
            'complete': sum(s['stage'] == STAGES[2] for s in state['setups'] if not s.get('archived')),
            'sync_errors': sum(s.get('sync', {}).get('status') == 'error' for s in state['setups'] if not s.get('archived'))}


def result(record):
    state = copy.deepcopy(record['state'])
    # Lease tokens remain server-side; public state is not accepted as a mutation.
    state.pop('sync_lock', None)
    return {'state': state, 'revision': record['revision'], 'updated_at': record['updated_at'], 'summary': summary(state)}


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
        raise HTTPException(409, 'Someone updated the team tracker. Refresh before saving again.')
    state = copy.deepcopy(record['state'])
    unlocked(state)
    state['sync_lock'] = None
    if action == 'client_save':
        client_name = re.sub(r'\s+', ' ', text(data['name'], 'Client name'))
        key = client_key(data['section'], client_name)
        if any(client_key(c['section'], c['name']) == key for c in state['clients']):
            raise HTTPException(409, 'This client already exists. Open the client to add another domain.')
        if len(state['clients']) >= 1000:
            raise HTTPException(400, 'Client limit reached.')
        find(state, 'members', data['account_manager_id'], True)
        if data.get('operator_id'):
            find(state, 'members', data['operator_id'], True)
        state['clients'].append({'id': new_id(), 'name': client_name, 'section': data['section'],
                                 'account_manager_id': data['account_manager_id'], 'operator_id': data.get('operator_id'),
                                 'created_at': now()})
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
        if any(s['stage'] != STAGES[2] and not s.get('archived') and member['id'] in (s['account_manager_id'], s['operator_id']) for s in state['setups']):
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
        for kind, amount in [('Phone', data['phones']), ('PC', data['pcs'])]:
            for _ in range(amount):
                number = device_slot(state, member_id, kind) if member_id else 1+max((d['number'] for d in state['devices'] if d['member_id'] is None and d['kind'] == kind), default=0)
                status = data.get('status', 'Unused')
                state['devices'].append({'id': new_id(), 'asset_tag': 'DEV-'+uuid4().hex[:8].upper(), 'kind': kind, 'number': number,
                    'member_id': member_id, 'status': status, 'block_reason': (data.get('block_reason') or 'Other') if status == 'Blocked' else None, 'used_at': now() if status == 'Used' else None,
                    'blocked_at': now() if status == 'Blocked' else None, 'archived': False, 'added_at': now(),
                    'use_count': 1 if status == 'Used' else 0, 'setup_ids': []})
    elif action == 'device_update':
        d = find(state, 'devices', data['id'], True)
        if 'member_id' in data and data['member_id'] != d['member_id']:
            if any(s['stage'] == STAGES[0] and not s.get('archived') and s.get('device_id') == d['id'] for s in state['setups']):
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
        if any(s['stage'] != STAGES[2] and not s.get('archived') and s.get('device_id') == d['id'] for s in state['setups']):
            raise HTTPException(409, 'This device is linked to an open setup. Reassign or archive that setup first.')
        d['archived'] = True
    elif action == 'setup_save':
        current = find(state, 'setups', data['id'], True) if data.get('id') else None
        if current and current.get('sync', {}).get('status') == 'synced':
            raise HTTPException(409, 'A published setup is locked to its sheet record. Correct the sheet directly if needed.')
        client = None
        if data.get('client_id'):
            client = find(state, 'clients', data['client_id'])
            if current and current.get('client_id') != client['id']:
                raise HTTPException(409, 'Edit this domain within its existing client.')
            data = {**data, 'client_name': client['name'], 'section': client['section']}
        # Existing client assignments survive removal from the setup team.
        # New manager assignments and all operators must still be active.
        existing_manager = (client or current or {}).get('account_manager_id')
        manager = find(state, 'members', data['account_manager_id'], data['account_manager_id'] != existing_manager)
        operator = find(state, 'members', data['operator_id'], True)
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
        if current and current.get('purchase_device') and current['stage'] not in FAILURES and device_id != current.get('device_id'):
            raise HTTPException(409, 'The purchase device is recorded in history and cannot be changed after purchase.')
        if device_id:
            historical = current and current['stage'] != STAGES[0] and current.get('device_id') == device_id
            device = find(state, 'devices', device_id, not historical)
            if not historical:
                if not allocation_available(state, device, (current or {}).get('id')):
                    raise HTTPException(409, 'This device is blocked, fully used or reserved. Choose another device or allocate one later.')
                if device['member_id'] != operator['id']:
                    raise HTTPException(409, 'Allocate the device to the setup owner first.')
        if current is None:
            if len(state['setups']) >= 5000:
                raise HTTPException(400, 'Setup history limit reached.')
            current = {'id': new_id(), 'stage': STAGES[0], 'created_at': now(), 'purchase_at': None, 'completed_at': None,
                       'purchase_device': None, 'sync': {'status': 'not_ready'}, 'archived': False}
            state['setups'].append(current)
        if current['stage'] != STAGES[0] and domain != current['domain']:
            raise HTTPException(409, 'The purchased domain cannot be changed. Create a new setup for a different domain.')
        client = next((c for c in state['clients'] if c['id'] == data.get('client_id')), None)
        if client is None:
            # Compatibility for previously opened forms and legacy API callers.
            display_name = text(data.get('client_name'), 'Client name', False) or 'Biz Dev'
            key = client_key(data['section'], display_name)
            client = next((c for c in state['clients'] if client_key(c['section'], c['name']) == key), None)
            if client is None:
                client = {'id': new_id(), 'name': display_name, 'section': data['section'], 'account_manager_id': manager['id'], 'operator_id': operator['id'], 'created_at': now()}
                state['clients'].append(client)
        current['client_id'] = client['id']
        current.update(domain=domain, email=email, client_name=text(data.get('client_name'), 'Client name', False),
                       section=data['section'], account_manager_id=manager['id'], operator_id=operator['id'],
                       device_id=device_id, order=data.get('order', 1), notes=text(data.get('notes'), 'Notes', False), updated_at=now())
        set_stage(state, current, data.get('stage') or current['stage'])
    elif action == 'setup_device':
        # A missing device can be recorded later, including after a sheet sync,
        # without permitting edits to the published email, owner or sheet row.
        job = find(state, 'setups', data['id'], True)
        if job.get('device_id'):
            raise HTTPException(409, 'This setup already has a device. Use Edit details to change a pending allocation.')
        if job.get('sync', {}).get('status') == 'synced' and data['operator_id'] != job['operator_id']:
            raise HTTPException(409, 'Keep the published setup owner. Allocate the device to that team member first.')
        operator = find(state, 'members', data['operator_id'], True)
        device = find(state, 'devices', data['device_id'], True)
        if not allocation_available(state, device, job['id']) or device['member_id'] != operator['id']:
            raise HTTPException(409, 'Choose an available device allocated to the setup owner, or add a device later.')
        if job['stage'] in FAILURES:
            raise HTTPException(409, 'Use Edit details to choose a replacement device and retry stage.')
        job.update(device_id=device['id'], operator_id=operator['id'], updated_at=now())
        if job.get('purchase_at') or job['stage'] in STAGES[1:]:
            record_purchase_device(state, job)
    elif action == 'setup_stage':
        job = find(state, 'setups', data['id'], True)
        set_stage(state, job, data['stage'])
    elif action == 'setup_remove':
        job = find(state, 'setups', data['id'], True)
        if job.get('sync', {}).get('status') == 'synced':
            raise HTTPException(409, 'Published setups remain in history. Use the stage filter to hide completed setups.')
        job['archived'] = True
    else:
        raise HTTPException(400, 'Unknown tracker action.')
    audit(state, user, action)
    return store.save(user, state, expected)


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
    consume_device_use(d, job['id'])
    d['status'] = 'Used'
    job['purchase_device'] = device_snapshot(state, d)


def set_stage(state, job, stage):
    if stage not in (*STAGES, *FAILURES):
        raise HTTPException(400, 'Choose a valid setup stage.')
    old = job['stage']
    if (old == STAGES[2] and stage != old or
            stage == STAGES[0] and (job.get('purchase_at') or old == STAGES[1])):
        raise HTTPException(409, 'Purchase and Live sheet history cannot be reset. Select the current or a later stage.')
    if stage == STAGES[2] and not job['email']:
        raise HTTPException(400, 'Add the sending email address before completing setup.')
    if stage in FAILURES:
        reason = FAILURES[stage]
        d = find(state, 'devices', job['device_id'], True) if job.get('device_id') else None
        if d:
            if d['status'] != 'Blocked' and d['member_id'] != job['operator_id']:
                raise HTTPException(409, 'Allocate the device to the setup owner before recording failure.')
            consume_device_use(d, job['id'])
            d.update(status='Blocked', block_reason=reason, blocked_at=d.get('blocked_at') or now())
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
        if job.get('device_id') and (job.get('purchase_device') or {}).get('id') != job['device_id']:
            record_purchase_device(state, job)
        job['purchase_at'] = job.get('purchase_at') or now()
    elif job.get('device_id') and find(state, 'devices', job['device_id'], True)['status'] == 'Blocked':
        raise HTTPException(409, 'Choose a replacement device or allocate one later before retrying.')
    if stage == STAGES[2] and old != stage:
        job.update(completed_at=now(), sync={'status': 'pending'})
    job['stage'] = stage
    job['updated_at'] = now()


def sync_completed(user, job_id):
    """Persistent workspace lease serializes writes across processes and retries."""
    from . import sheets
    record = board(user)
    state = copy.deepcopy(record['state'])
    job = find(state, 'setups', job_id, True)
    if job['stage'] != STAGES[2]:
        raise HTTPException(409, 'Complete setup before publishing it to the strategy sheet.')
    if job.get('sync', {}).get('status') == 'synced':
        return record
    unlocked(state)
    token = new_id()
    state['sync_lock'] = {'job_id': job_id, 'token': token, 'at': now()}
    claimed = store.save(user, state, record['revision'])
    try:
        owner = job['client_name'] if job['section'] == 'client' else find(state, 'members', job['operator_id'])['name']
        synced = sheets.publish(job, owner)
        status = {'status': 'synced', 'at': now(), **synced}
    except HTTPException as exc:
        status = {'status': 'error', 'at': now(), 'message': exc.detail}
        match = re.search(r'https://docs\.google\.com/spreadsheets/d/[A-Za-z0-9_-]+/edit#gid=\d+&range=A\d+', str(exc.detail))
        if match:
            status['url'] = match.group(0)
    # Fail closed if a newer lease has replaced this writer. The stable sheet
    # metadata and exact-email lookup reconcile any committed external write.
    latest = store.load(user)
    if not latest or (latest['state'].get('sync_lock') or {}).get('token') != token:
        raise HTTPException(409, 'The sync claim changed. Refresh and retry; the sheet write will be reconciled.')
    final = copy.deepcopy(latest['state'])
    find(final, 'setups', job_id)['sync'] = status
    final['sync_lock'] = None
    audit(final, user, 'sheet_sync_'+status['status'])
    return store.save(user, final, latest['revision'])
