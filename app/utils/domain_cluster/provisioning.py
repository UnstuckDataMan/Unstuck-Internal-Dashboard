"""Read-only inspection and narrowly scoped, idempotent new-domain setup.

There are deliberately no registration, delete, nameserver-change or DNS-edit APIs.
Only missing redirects and one missing DMARC TXT record may be created, on the
current server-recorded batch of Planned domains owned by the connected account.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from urllib.parse import urlsplit

import requests
from fastapi import HTTPException
from . import engine as e
from .enrichment import configured

BASE = 'https://api.porkbun.com/api/json/v3'


class Porkbun:
    def call(self, method, path, data=None, idempotency=None):
        if not configured():
            raise HTTPException(503, 'Porkbun API key and matching secret must be configured on the server.')
        headers = {'X-API-Key':os.environ['PORKBUN_API_KEY'], 'X-Secret-API-Key':os.environ['PORKBUN_SECRET_API_KEY']}
        if idempotency:
            headers['Idempotency-Key'] = idempotency
        try:
            response = getattr(requests, method)(BASE+path, headers=headers, **({'json':data or {}} if method=='post' else {}), timeout=(3,6), allow_redirects=False)
            if response.status_code == 429:
                raise HTTPException(429, 'Porkbun lookup/write limit reached. Wait before retrying.')
            if not response.ok:
                # Provider messages may contain credentials or untrusted data.
                raise HTTPException(502, 'Porkbun could not complete the operation. Check ownership, API access and read/DNS permissions.')
            body = response.json()
            if body.get('status') != 'SUCCESS':
                raise HTTPException(502, 'Porkbun did not confirm success. No completion was assumed.')
            if str(response.headers.get('X-Porkbun-Mock','')).lower()=='true' or body.get('mock') is True or body.get('sandbox') is True:
                raise HTTPException(502, 'Porkbun returned a mock or sandbox response; live setup was not verified.')
            if body.get('warnings'):
                raise HTTPException(409, 'Porkbun reported a DNS/forwarding warning. Verify the authoritative provider before continuing.')
            return body
        except (requests.RequestException, ValueError, TypeError):
            raise HTTPException(502, 'Porkbun request timed out or returned an unreadable response. Recheck the domain before retrying; a write may have completed.') from None

    def check(self):
        data = self.call('post','/ping')
        scopes = data.get('scopes',[])
        valid = data.get('credentialsValid') is True
        return {'credentials_valid':valid,'read':valid and bool({'read','full'} & set(scopes)), 'dns':valid and bool({'dns','full'} & set(scopes)), 'checked_at':e.now()}


def validate_options(raw):
    options = {**raw}
    if not options.get('forward') and not options.get('dmarc'):
        raise ValueError('Choose forwarding, DMARC, or both.')
    if options.get('forward'):
        url = options.get('destination','').strip()
        parsed = urlsplit(url)
        if parsed.scheme not in {'https','http'} or not parsed.hostname or parsed.username or parsed.password or any(c.isspace() for c in url):
            raise ValueError('Enter a complete http:// or https:// forwarding destination without credentials or spaces.')
        e.normalize(url)
        options['destination'] = url
    if options.get('dmarc'):
        for key in ('rua','ruf'):
            template = (options.get(key) or options.get('report_user','me')+'@{domain}').strip()
            sample = template.replace('{domain}','example.com')
            if '{' in sample or '}' in sample or not re.fullmatch(r'[A-Za-z0-9.!#$%&\x27*+/=?^_`|~-]+@[A-Za-z0-9.-]+',sample):
                raise ValueError('DMARC recipients must be a mailbox address or a template such as me@{domain}.')
            e.normalize(sample)
            if key in raw: options[key] = template
    for name in [options.get('report_user','me'), *options.get('report_users',{}).values()]:
        if not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9._+\-]{0,63}',name):
            raise ValueError('Report mailbox name must be a valid local part such as me or richard. Do not include @ or spaces.')
    return options


def dmarc_value(domain, options):
    name = options.get('report_users',{}).get(domain) or options.get('report_user','me')
    rua = (options.get('rua') or name+'@{domain}').replace('{domain}',domain)
    ruf = (options.get('ruf') or name+'@{domain}').replace('{domain}',domain)
    return f'v=DMARC1; p=none; pct=100; rua=mailto:{rua}; ruf=mailto:{ruf}; fo=1'


def fingerprint(value):
    return hashlib.sha256(json.dumps(value,sort_keys=True,separators=(',',':')).encode()).hexdigest()


def latest_batch(state, batch_id):
    batches = state.get('setup_batches',[])
    if not batches or batches[-1]['id'] != batch_id:
        raise HTTPException(409, 'Select the current batch of newly Planned domains. Earlier batches and the full portfolio cannot be configured here.')
    batch = batches[-1]
    records = {r['domain']:r for r in state['domains']}
    for domain in batch['domains']:
        row = records.get(domain)
        if not row or row['status'] != 'Planned' or row.get('sheet_sources'):
            raise HTTPException(409, 'This batch now contains an existing, active or removed domain. Reload and review its inventory first.')
    return batch


def inspect(domain, options, client):
    meta = client.call('get','/domain/get/'+domain).get('domain',{})
    if meta.get('domain','').lower() != domain or str(meta.get('status','')).upper() != 'ACTIVE':
        raise HTTPException(409, 'Domain is not Active in the connected Porkbun account. Register it first.')
    if meta.get('apiAccess') not in (1,True,'1','yes'):
        raise HTTPException(409, 'Enable API access for this domain in Porkbun before setup.')
    nameservers = sorted(str(v).lower().rstrip('.') for v in client.call('get','/domain/getNs/'+domain).get('ns',[]))
    if not nameservers or not all(n.endswith('.porkbun.com') for n in nameservers):
        raise HTTPException(409, 'Domain is not delegated exclusively to Porkbun nameservers. Configure its authoritative DNS provider separately.')
    forwards = client.call('get','/domain/getUrlForwarding/'+domain).get('forwards',[]) if options['forward'] else []
    records = client.call('get','/dns/retrieve/'+domain).get('records',[]) if options['dmarc'] else []
    if not isinstance(forwards,list) or not isinstance(records,list):
        raise HTTPException(502, 'Porkbun returned an unexpected record list. No setup was approved.')
    forwards = [{**r,'subdomain':r.get('subdomain') or ''} for r in forwards]
    records = [r for r in records if r.get('name','').lower().rstrip('.')=='_dmarc.'+domain]
    actions, conflicts = [], []
    for host in (['','www'] if options.get('include_www') else ['']) if options['forward'] else []:
        matching = [r for r in forwards if (r.get('subdomain') or '')==host]
        desired = {'subdomain':host,'location':options['destination'],'type':'permanent' if options['redirect_code']=='301' else 'temporary','redirectType':options['redirect_code'],'includePath':'yes' if options['include_path'] else 'no','wildcard':'no'}
        same = len(matching)==1 and all(str(matching[0].get(k,''))==str(v) for k,v in desired.items() if k!='redirectType') and str(matching[0].get('redirectType','301' if matching[0].get('type')=='permanent' else '302'))==options['redirect_code']
        action = 'already configured' if same else 'conflict' if matching else 'create'
        if action=='conflict': conflicts.append('Existing forwarding for '+(host+'.' if host else '')+domain+' differs from this request.')
        actions.append({'kind':'forward','host':host,'action':action,'payload':desired})
    if options['dmarc']:
        content = dmarc_value(domain,options)
        compact = lambda s: re.sub(r'\s+','',str(s).replace('"','')).lower()
        same = len(records)==1 and records[0].get('type')=='TXT' and compact(records[0].get('content',''))==compact(content) and str(records[0].get('ttl'))=='600'
        action = 'already configured' if same else 'conflict' if records else 'create'
        if action=='conflict': conflicts.append('Existing DMARC TXT record differs or has a different TTL. It will not be overwritten or duplicated.')
        actions.append({'kind':'dmarc','host':'_dmarc','action':action,'payload':{'name':'_dmarc','type':'TXT','content':content,'ttl':600}})
    relevant = {'nameservers':nameservers,'forwards':sorted(forwards,key=lambda r:json.dumps(r,sort_keys=True)),'dmarc':sorted(records,key=lambda r:json.dumps(r,sort_keys=True))}
    return {'domain':domain,'status':'blocked' if conflicts else 'ready','actions':actions,'conflicts':conflicts,'fingerprint':fingerprint(relevant)}


def preview(domains, options, client=None):
    client = client or Porkbun()
    check = client.check()
    if not check['credentials_valid'] or not check['read'] or not check['dns']:
        raise HTTPException(403, 'Porkbun credentials need both read and DNS permissions for forwarding/DMARC setup.')
    def read(domain):
        try:
            if options['forward'] and e.normalize(options['destination']) == domain:
                raise HTTPException(400,'The forwarding destination would loop back to this domain.')
            return inspect(domain,options,client)
        except HTTPException as exc:
            return {'domain':domain,'status':'blocked','actions':[],'conflicts':[str(exc.detail)]}
    with ThreadPoolExecutor(max_workers=3) as pool:
        return list(pool.map(read,domains))


def apply_one(domain, approved, options, preview_id, client=None):
    client = client or Porkbun()
    current = inspect(domain,options,client)
    if current['status']!='ready':
        raise HTTPException(409,'Domain configuration has conflicts. Create a fresh preview.')
    if current['fingerprint'] != approved['fingerprint']:
        # A timed-out earlier write may already have converged. Permit a pure
        # read-only reconciliation, never continue a changed partial plan.
        if all(a['action']=='already configured' for a in current['actions']):
            return {'domain':domain,'status':'verified','message':'Already configured; no write needed.','checked_at':e.now()}
        raise HTTPException(409,'DNS or forwarding changed after preview. Create a fresh preview; no further writes were made.')
    performed=[]
    try:
        for action in current['actions']:
            if action['action']!='create':
                continue
            key = fingerprint([preview_id,domain,action['kind'],action['host'],action['payload']])
            path = '/domain/addUrlForward/'+domain if action['kind']=='forward' else '/dns/create/'+domain
            result = client.call('post',path,action['payload'],idempotency=key)
            if action['kind']=='dmarc' and not isinstance(result.get('id'),(str,int)):
                raise HTTPException(502,'Porkbun did not return a usable DNS record ID. Verify this domain before retrying.')
            performed.append(action['kind']+':'+action['host'])
        verified = inspect(domain,options,client)
        if verified['status']!='ready' or any(a['action']!='already configured' for a in verified['actions']):
            raise HTTPException(502,'Porkbun accepted the request but readback did not confirm every setting yet. Recheck before retrying.')
        return {'domain':domain,'status':'verified','message':'Forwarding and/or DMARC verified through Porkbun API readback. Public DNS propagation may still be pending.','performed':performed,'checked_at':e.now()}
    except HTTPException as exc:
        return {'domain':domain,'status':'needs review','message':str(exc.detail),'performed':performed,'checked_at':e.now()}
