"""Private, owner-scoped Supabase documents with optimistic concurrency.

One database row is the transaction boundary for a portfolio and its undo history.
No partial domain/cluster writes and no silent last-writer-wins updates.
"""
from __future__ import annotations

import copy
import hashlib
import uuid

import requests
from fastapi import HTTPException
from app.utils import supabase as sb
from .engine import now, analyse

TABLE = 'domain_cluster_portfolios'


def scope(user):
    email = user.get('email', '').strip().lower()
    if not email:
        raise HTTPException(401, 'Sign in to use saved portfolios.')
    return hashlib.sha256(email.encode()).hexdigest(), hashlib.sha256(email.rsplit('@', 1)[-1].encode()).hexdigest()


def _call(method, user, *, portfolio_id=None, body=None, revision=None, listing=False):
    if not sb.sb_service_configured():
        raise HTTPException(503, 'Portfolio storage is not configured. Set SUPABASE_SERVICE_ROLE_KEY and apply the domain cluster migration. Analysis remains available.')
    owner, workspace = scope(user)
    params = {'owner_id': f'eq.{owner}', 'workspace_id': f'eq.{workspace}'}
    if portfolio_id:
        params['id'] = f'eq.{portfolio_id}'
    if revision is not None:
        params['revision'] = f'eq.{revision}'
    if method == 'get':
        params.update(select='id,name,revision,updated_at' if listing else '*', order='updated_at.desc', limit='100')
    payload = body
    if method == 'post':
        payload = {**body, 'owner_id': owner, 'workspace_id': workspace}
    try:
        response = getattr(requests, method)(f'{sb.SUPABASE_URL}/rest/v1/{TABLE}', params=params, headers=sb.sb_service_headers('return=representation'), **({'json': payload} if payload is not None else {}), timeout=20)
        if response.status_code == 409:
            raise HTTPException(409, 'Portfolio already exists. Reload it before saving.')
        if not response.ok:
            raise HTTPException(503, 'Portfolio database unavailable. Check configuration and apply the domain cluster migration; your working analysis is unchanged.')
        return response.json()
    except (requests.RequestException, ValueError):
        raise HTTPException(503, 'Could not reach portfolio storage. Retry without closing this page.') from None


def list_portfolios(user):
    return _call('get', user, listing=True)


def load(user, portfolio_id):
    rows = _call('get', user, portfolio_id=portfolio_id)
    if not rows:
        raise HTTPException(404, 'Portfolio not found.')
    return rows[0]


def save(user, portfolio_id, name, state, expected_revision):
    if not portfolio_id:
        if expected_revision != 0:
            raise HTTPException(409, 'Create a new portfolio with revision zero.')
        payload = {'id': str(uuid.uuid4()), 'name': name, 'state': state, 'revision': 1, 'updated_at': now()}
        rows = _call('post', user, body=payload)
    else:
        payload = {'name': name, 'state': state, 'revision': expected_revision + 1, 'updated_at': now()}
        rows = _call('patch', user, portfolio_id=portfolio_id, body=payload, revision=expected_revision)
    if not rows:
        raise HTTPException(409, 'This portfolio changed in another tab. Reload before saving to avoid losing changes.')
    return rows[0]


def checkpoint(state, action):
    snapshot = copy.deepcopy({k: state[k] for k in ('domains', 'clusters', 'settings', 'recommendations')})
    state['history'] = (state.get('history', []) + [{'at': now(), 'action': action, 'before': snapshot}])[-10:]


def snapshot(state):
    report = analyse(state)
    state['snapshots'] = (state.get('snapshots', []) + [{'at': now(), 'rule_version': report['rule_version'], 'summary': report['summary']}])[-30:]
