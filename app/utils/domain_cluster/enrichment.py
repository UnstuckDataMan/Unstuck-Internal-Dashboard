"""Optional bounded public lookups. No registrar purchases or HTTP redirects."""
from __future__ import annotations

import copy
import os
import re
import threading
import time
from collections import deque
from concurrent.futures import ThreadPoolExecutor

import dns.exception
import dns.resolver
import requests
from fastapi import HTTPException
from .engine import now, normalize

_lock = threading.Lock()
_cache = {}
_hits = {}
_cooldown = 0.0
PORKBUN_ENDPOINT = 'https://api.porkbun.com/api/json/v3/domain/checkDomain'


def configured():
    return bool(os.environ.get('PORKBUN_API_KEY') and os.environ.get('PORKBUN_SECRET_API_KEY'))


def rate_limit(key, count=3, window=60):
    with _lock:
        stamp = time.monotonic()
        hits = _hits.setdefault(key, deque())
        while hits and hits[0] < stamp - window:
            hits.popleft()
        if len(hits) >= count:
            raise HTTPException(429, 'Lookup limit reached. Wait a minute before retrying.', headers={'Retry-After': str(window)})
        hits.append(stamp)
        # Bound per-process rate state even for long-lived multi-user services.
        if len(_hits) > 5000:
            for old in list(_hits):
                if not _hits[old] or _hits[old][-1] < stamp - 600:
                    del _hits[old]


def _cached(key, ttl):
    with _lock:
        entry = _cache.get(key)
        if entry and time.monotonic() - entry[0] < ttl:
            return copy.deepcopy(entry[1])
    return None


def _put(key, data):
    with _lock:
        if len(_cache) > 2000:
            _cache.clear()
        _cache[key] = (time.monotonic(), copy.deepcopy(data))


def porkbun(domains):
    global _cooldown
    if not configured():
        raise HTTPException(503, 'Porkbun is not connected. Set PORKBUN_API_KEY and PORKBUN_SECRET_API_KEY in Render environment variables.')
    if not 1 <= len(domains) <= 25 or any(len(d) > 67 or not re.fullmatch('[a-z]+\\.com', d) for d in domains):
        raise HTTPException(400, 'Check up to 25 proposed letters-only .com names.')
    if time.monotonic() < _cooldown:
        raise HTTPException(429, 'Porkbun requested a pause. Retry later.', headers={'Retry-After': str(int(_cooldown - time.monotonic()) + 1)})
    output, missing = {}, []
    for d in domains:
        hit = _cached(('porkbun', d), 300)
        if hit:
            output[d] = {**hit, 'cached': True}
        else:
            missing.append(d)
    if not missing:
        return output
    rate_limit('porkbun-account', count=6)
    try:
        response = requests.post(PORKBUN_ENDPOINT, json={'apikey': os.environ['PORKBUN_API_KEY'], 'secretapikey': os.environ['PORKBUN_SECRET_API_KEY'], 'domains': missing}, timeout=(5, 15), allow_redirects=False)
        if response.status_code == 429:
            try:
                delay = min(86400, max(60, int(response.headers.get('Retry-After', '60'))))
            except ValueError:
                delay = 60
            _cooldown = time.monotonic() + delay
            raise HTTPException(429, 'Porkbun rate limit reached. Retry after the indicated pause.', headers={'Retry-After': str(delay)})
        if not response.ok:
            raise HTTPException(502, 'Porkbun check failed. Verify the server API credentials and read permissions, then retry.')
        body = response.json()
        if body.get('status') != 'SUCCESS' or not isinstance(body.get('domains'), dict):
            raise HTTPException(502, 'Porkbun could not verify this batch. Availability remains unconfirmed.')
        mock = str(response.headers.get('X-Porkbun-Mock', '')).lower() == 'true' or body.get('mock') is True
        for domain in missing:
            row = body['domains'].get(domain, {})
            avail = row.get('avail')
            verified = not mock and avail in {'yes', 'no'}
            result = {'availability': 'Available on Porkbun' if verified and avail == 'yes' else 'Unavailable on Porkbun' if verified else 'Availability not verified', 'available': avail == 'yes' if verified else None, 'verified': verified, 'checked_at': now() if verified else None, 'price_usd': str(row['price'])[:30] if verified and 'price' in row else None, 'premium': row.get('premium') == 'yes' if verified else None, 'cached': False, 'search_url': 'https://porkbun.com/checkout/search?q=' + domain}
            output[domain] = result
            if verified:
                _put(('porkbun', domain), result)
        return output
    except (requests.RequestException, ValueError, TypeError):
        # Never log requests/errors containing credential-bearing payloads.
        raise HTTPException(502, 'Porkbun did not return a usable response. No availability was assumed.') from None


def dns_domain(domain, selector=''):
    domain = normalize(domain)
    key = ('dns', domain, selector)
    hit = _cached(key, 600)
    if hit:
        return {**hit, 'cached': True}
    resolver = dns.resolver.Resolver()
    resolver.timeout = 1
    resolver.lifetime = 2
    queries = [('MX', domain, 'MX'), ('SPF', domain, 'TXT'), ('DMARC', '_dmarc.'+domain, 'TXT'), ('NS', domain, 'NS')]
    if selector:
        queries.append(('DKIM', selector+'._domainkey.'+domain, 'TXT'))
    records = {}
    for name, host, kind in queries:
        try:
            answers = [str(r)[:1000] for r in resolver.resolve(host, kind, search=False)]
            if name == 'SPF':
                answers = [v for v in answers if 'v=spf1' in v.lower()]
            elif name == 'DMARC':
                answers = [v for v in answers if 'v=dmarc1' in v.lower()]
            records[name] = {'values': answers[:20]}
        except dns.resolver.NXDOMAIN:
            records[name] = {'error': 'Name does not exist'}
        except dns.resolver.NoAnswer:
            records[name] = {'error': 'No record returned'}
        except dns.exception.DNSException:
            records[name] = {'error': 'Lookup failed or timed out; status unknown'}
    result = {'domain': domain, 'records': records, 'checked_at': now(), 'cached': False}
    _put(key, result)
    return result


def dns_batch(domains, selector):
    if selector and not re.fullmatch('[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?', selector):
        raise HTTPException(400, 'DKIM selector must be a single valid DNS label.')
    with ThreadPoolExecutor(max_workers=5) as pool:
        results = list(pool.map(lambda d: dns_domain(d, selector), domains[:20]))
    providers = {}
    for row in results:
        for record in ('MX', 'NS'):
            for value in row['records'][record].get('values', []):
                providers.setdefault((record, value), []).append(row['domain'])
    commonality = [{'type': key[0], 'value': key[1], 'domains': ds} for key, ds in providers.items() if len(ds) > 1]
    return {'results': results, 'commonality': commonality, 'note': 'Shared public providers are common and do not establish harmful clustering. No redirects, ports or private hosts were probed. Results cached for up to ten minutes.'}
