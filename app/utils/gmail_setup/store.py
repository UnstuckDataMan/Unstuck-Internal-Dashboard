"""One workspace document, conditional revisions and no browser database grants."""
import hashlib
import requests
from fastapi import HTTPException
from app.utils import supabase as sb

TABLE = 'gmail_setup_boards'


def workspace(user):
    email = user.get('email', '').strip().lower()
    if '@' not in email:
        raise HTTPException(401, 'Sign in to use Gmail Set.')
    return hashlib.sha256(email.rsplit('@', 1)[1].encode()).hexdigest()


def call(method, table, params, body=None):
    if not sb.sb_service_configured():
        raise HTTPException(503, 'Gmail Set storage is not configured.')
    try:
        r = requests.request(method, sb.SUPABASE_URL+'/rest/v1/'+table, params=params,
                             headers=sb.sb_service_headers('return=representation'),
                             **({'json': body} if body is not None else {}), timeout=(3, 12))
        if r.status_code == 409:
            raise HTTPException(409, 'The team tracker changed. Refresh and retry.')
        if not r.ok:
            raise HTTPException(503, 'Gmail Set storage is unavailable. Your changes were not saved.')
        return r.json()
    except (requests.RequestException, ValueError):
        raise HTTPException(503, 'Could not reach Gmail Set storage. Refresh before retrying.') from None


def load(user):
    rows = call('get', TABLE, {'workspace_id': 'eq.'+workspace(user), 'select': '*', 'limit': '1'})
    return rows[0] if rows else None


def team_directory(user):
    domain = user['email'].rsplit('@', 1)[1].lower()
    rows = call('get', 'app_users', {'select': 'email,name', 'active': 'eq.true', 'limit': '1000'})
    # Only the signed-in workspace's internal directory is exposed in this tool.
    return [r for r in rows if r.get('email', '').lower().endswith('@'+domain)]


def save(user, state, revision):
    from .service import now
    body = {'state': state, 'revision': revision+1, 'updated_at': now()}
    if revision == 0:
        rows = call('post', TABLE, {}, {**body, 'workspace_id': workspace(user)})
    else:
        rows = call('patch', TABLE, {'workspace_id': 'eq.'+workspace(user), 'revision': 'eq.'+str(revision)}, body)
    if not rows:
        raise HTTPException(409, 'Someone updated the team tracker. Refresh before saving again.')
    return rows[0]
