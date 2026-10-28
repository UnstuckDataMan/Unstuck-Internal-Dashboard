"""Internal, client-scoped Copy Studio. No publishing or campaign sending routes."""
from __future__ import annotations

import secrets
import uuid
import os
from datetime import datetime, timezone
from typing import Literal

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel, ConfigDict, Field

from app import auth
from app.deps import templates
from app.utils.copy_studio import integrations, providers, service
from app.utils.copy_studio.store import StoreError, StoreNotReady, get_store

router = APIRouter()


class Strict(BaseModel):
    model_config = ConfigDict(extra='forbid')


class Sync(Strict):
    source: Literal['all', 'copy_bank', 'onboarding', 'smartlead', 'manual'] = 'all'


class Audience(Strict):
    subsector: str = Field('', max_length=300)
    geography: str = Field('', max_length=200)
    company_size: str = Field('', max_length=100)
    role: str = Field('', max_length=200)
    offer: str = Field('', max_length=500)


class Generate(Strict):
    channel: Literal['email', 'linkedin']
    mode: Literal['assemble', 'draft']
    audience: Audience
    objective: str = Field('', max_length=1000)
    angles: int = Field(3, ge=1, le=3)
    example_ids: list[str] = Field(default_factory=list, max_length=12)
    limits: dict[str, int] = Field(default_factory=dict)
    variables: dict[str, str] = Field(default_factory=dict)
    template_id: str | None = None


class Record(Strict):
    kind: Literal['brief', 'example']
    source_key: str = Field(min_length=1, max_length=500)
    content: dict
    parent_id: str | None = None


class Revision(Strict):
    content: dict


class Review(Strict):
    decision: Literal['approved', 'rejected', 'needs_changes']
    notes: str = Field('', max_length=3000)


class ConfirmTranscript(Strict):
    confirmed: Literal[True]
    notes: str = Field(min_length=10, max_length=3000)


class SentReference(Strict):
    subject: str = Field('', max_length=500)
    body: str = Field(min_length=1, max_length=10000)
    sanitized: Literal[True]


def access(request: Request):
    user = auth.require_login(request)
    if user.get('role') not in ('admin', 'reviewer', 'sdr'):
        raise HTTPException(403, 'Copy Studio is currently an internal staff pilot.')
    if 'copy_studio' not in user.get('tools', []):
        raise HTTPException(403, 'Copy Studio access is required.')
    if not user.get('email'):
        raise HTTPException(403, 'A verified session identity is required.')
    agency = os.environ.get('COPY_STUDIO_AGENCY_ID') or os.environ.get('PULSE_AGENCY_ID')
    if user.get('agency_id') and str(user['agency_id']) != str(agency):
        raise HTTPException(403, 'This session belongs to a different agency.')
    return user


def mutate(request: Request, user=Depends(access)):
    if auth.auth_enabled():
        token = request.session.get('copy_studio_csrf', '')
        sent = request.headers.get('X-Copy-Studio-CSRF', '')
        if not token or not secrets.compare_digest(token, sent):
            raise HTTPException(403, 'Refresh Copy Studio before saving: session verification expired.')
        if request.headers.get('sec-fetch-site') == 'cross-site':
            raise HTTPException(403, 'Cross-site writes are not permitted.')
    return user


def reviewer(user):
    if user.get('role') not in auth.COPY_APPROVER_ROLES:
        raise HTTPException(403, 'An admin or reviewer must approve this version.')


def permitted_clients(user):
    store = get_store()
    if user.get('agency_id') and str(user['agency_id']) != str(store.agency_id):
        raise HTTPException(403, 'This session belongs to a different agency.')
    rows = service.clients(store)
    # Future restricted staff assignments fail closed when an explicit list exists.
    if 'client_ids' in user:
        allowed = set(user.get('client_ids') or [])
        rows = [c for c in rows if c['id'] in allowed]
    return rows


def check_client(client_id, user):
    if client_id not in {c['id'] for c in permitted_clients(user)}:
        raise HTTPException(404, 'Client not found or not assigned to this account.')


def result(call):
    try:
        return JSONResponse(call(), headers={'Cache-Control': 'no-store'})
    except StoreNotReady as exc:
        return JSONResponse({'error': str(exc)}, status_code=503, headers={'Cache-Control': 'no-store'})
    except StoreError as exc:
        return JSONResponse({'error': str(exc)}, status_code=409, headers={'Cache-Control': 'no-store'})
    except ValueError as exc:
        return JSONResponse({'error': str(exc)}, status_code=422, headers={'Cache-Control': 'no-store'})
    except LookupError:
        return JSONResponse({'error': 'Version not found.'}, status_code=404, headers={'Cache-Control': 'no-store'})


def request_key(request):
    value = request.headers.get('Idempotency-Key') or str(uuid.uuid4())
    if len(value) > 120:
        raise HTTPException(422, 'Idempotency key is too long.')
    return value


@router.get('/copy-studio')
def page(request: Request, user=Depends(access)):
    request.session.setdefault('copy_studio_csrf', secrets.token_urlsafe(32))
    return templates.TemplateResponse(request, 'copy_studio.html', {
        'csrf': request.session['copy_studio_csrf'], 'can_review': user['role'] in auth.COPY_APPROVER_ROLES,
    }, headers={'Cache-Control': 'no-store'})


@router.get('/api/copy-studio/bootstrap')
def bootstrap(request: Request, user=Depends(access)):
    request.session.setdefault('copy_studio_csrf', secrets.token_urlsafe(32))
    info = {'capabilities': integrations.capabilities(), 'model': providers.provider_status(),
            'can_review': user['role'] in auth.COPY_APPROVER_ROLES, 'csrf': request.session['copy_studio_csrf'],
            'ready': False, 'clients': [], 'pilot': 'internal', 'sending_enabled': False}
    try:
        info['clients'] = permitted_clients(user)
        # Check migration/readiness even when the agency currently has no clients.
        get_store().list(service.GLOBAL_CLIENT, kind='methodology', limit=1)
        info['ready'] = True
        from app.utils.copy_studio import costs
        info['costs'] = costs.summarize(get_store(), [c['id'] for c in info['clients']])
    except StoreError as exc:
        info['error'] = str(exc)
    return JSONResponse(info, headers={'Cache-Control': 'no-store'})


@router.get('/api/copy-studio/costs')
def usage_costs(user=Depends(access)):
    from app.utils.copy_studio import costs
    return result(lambda: costs.summarize(get_store(), [c['id'] for c in permitted_clients(user)]))


@router.get('/api/copy-studio/clients/{client_id}')
def client(client_id: str, user=Depends(access)):
    def run():
        check_client(client_id, user)
        return service.client_view(client_id)
    return result(run)


@router.post('/api/copy-studio/clients/{client_id}/sync')
def sync(client_id: str, body: Sync, request: Request, user=Depends(mutate)):
    def run():
        check_client(client_id, user)
        return service.enqueue_sync(client_id, body.source, user, request_key(request))
    return result(run)


@router.post('/api/copy-studio/clients/{client_id}/extract-brief')
def extract_brief(client_id: str, request: Request, user=Depends(mutate)):
    def run():
        check_client(client_id, user)
        return service.enqueue_brief(client_id, user, request_key(request))
    return result(run)


@router.post('/api/copy-studio/clients/{client_id}/generate')
def generate(client_id: str, body: Generate, request: Request, user=Depends(mutate)):
    def run():
        check_client(client_id, user)
        return service.enqueue_generation(client_id, body.model_dump(), user, request_key(request))
    return result(run)


@router.get('/api/copy-studio/clients/{client_id}/jobs/{identifier}')
def job(client_id: str, identifier: str, user=Depends(access)):
    def run():
        check_client(client_id, user)
        found = get_store().get_job(identifier, client_id)
        if not found:
            raise LookupError()
        return found
    return result(run)


@router.get('/api/copy-studio/clients/{client_id}/records/{identifier}')
def record(client_id: str, identifier: str, user=Depends(access)):
    def run():
        check_client(client_id, user)
        found = get_store().get(identifier, client_id)
        if not found:
            raise LookupError()
        return found
    return result(run)


@router.get('/api/copy-studio/clients/{client_id}/records')
def history(client_id: str, kind: Literal['brief', 'example', 'transcript', 'sync_report', 'sent_copy', 'draft', 'performance'],
            cursor: str | None = None, user=Depends(access)):
    def run():
        check_client(client_id, user)
        return get_store().list_page(client_id, kind=kind, limit=100, cursor=cursor)
    return result(run)


@router.post('/api/copy-studio/clients/{client_id}/records')
def save_record(client_id: str, body: Record, user=Depends(mutate)):
    def run():
        check_client(client_id, user)
        return service.save_record(client_id, body.kind, body.source_key, body.content, user, body.parent_id)
    return result(run)


@router.post('/api/copy-studio/clients/{client_id}/records/{identifier}/review')
def review(client_id: str, identifier: str, body: Review, user=Depends(mutate)):
    def run():
        check_client(client_id, user)
        reviewer(user)
        return service.review_record(client_id, identifier, body.decision, body.notes, user)
    return result(run)


@router.post('/api/copy-studio/clients/{client_id}/drafts/{identifier}/revisions')
def revision(client_id: str, identifier: str, body: Revision, user=Depends(mutate)):
    def run():
        check_client(client_id, user)
        return service.revise_draft(client_id, identifier, body.content, user)
    return result(run)


@router.post('/api/copy-studio/clients/{client_id}/drafts/{identifier}/check')
def check_draft(client_id: str, identifier: str, request: Request, user=Depends(mutate)):
    def run():
        check_client(client_id, user)
        return service.enqueue_draft_check(client_id, identifier, user, request_key(request))
    return result(run)


@router.post('/api/copy-studio/clients/{client_id}/records/{identifier}/reference')
def prepare_reference(client_id: str, identifier: str, body: SentReference, user=Depends(mutate)):
    def run():
        check_client(client_id, user)
        return service.prepare_reference(client_id, identifier, body.subject, body.body, body.sanitized, user)
    return result(run)


@router.post('/api/copy-studio/clients/{client_id}/transcripts/{identifier}/confirm-complete')
def confirm_transcript(client_id: str, identifier: str, body: ConfirmTranscript, user=Depends(mutate)):
    def run():
        check_client(client_id, user)
        reviewer(user)
        store = get_store()
        previous = store.get(identifier, client_id)
        if not previous or previous['kind'] != 'transcript':
            raise LookupError()
        content = dict(previous['content'])
        content['source_complete'] = True
        content['source_completeness_review'] = {
            'actor': service.actor(user), 'notes': body.notes, 'confirmed_at': datetime.now(timezone.utc).isoformat()}
        return store.save('transcript', client_id, previous['source_key'], content, service.actor(user), parent_id=identifier)
    return result(run)


@router.post('/api/copy-studio/methodology/seed')
def seed_methodology(user=Depends(mutate)):
    reviewer(user)
    from app.utils.copy_studio.methodology import default_methodology
    return result(lambda: get_store().save('methodology', service.GLOBAL_CLIENT, 'build-brief-principles',
                                           default_methodology(), service.actor(user)))


@router.post('/api/copy-studio/methodology/{identifier}/review')
def review_methodology(identifier: str, body: Review, user=Depends(mutate)):
    reviewer(user)
    return result(lambda: service.review_record(service.GLOBAL_CLIENT, identifier, body.decision, body.notes, user))
