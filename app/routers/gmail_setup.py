"""Shared Gmail setup board, protected by dashboard tool access and CSRF."""
import secrets
from typing import Literal
from uuid import UUID
from fastapi import APIRouter, Depends, Request
from pydantic import BaseModel, ConfigDict, Field
from app.deps import templates
from app.routers.domain_cluster import access, mutation_access
from app.utils.gmail_setup import service, sheets

router = APIRouter()


class Strict(BaseModel):
    model_config = ConfigDict(extra='forbid')


class Member(Strict):
    id: str | None = Field(None, max_length=36)
    name: str = Field(min_length=1, max_length=100)
    email: str = Field('', max_length=254)


class Devices(Strict):
    member_id: str | None = Field(None, max_length=36)
    phones: int = Field(0, ge=0, le=10)
    pcs: int = Field(0, ge=0, le=10)
    status: Literal['Unused', 'Used', 'Blocked'] = 'Unused'
    block_reason: Literal['QR code', 'Payment', 'Other'] | None = None


class Device(Strict):
    id: str = Field(max_length=36)
    member_id: str | None = Field(None, max_length=36)
    status: Literal['Unused', 'Used', 'Blocked'] | None = None
    block_reason: Literal['QR code', 'Payment', 'Other'] | None = None
    asset_tag: str | None = Field(None, min_length=1, max_length=80)


class Client(Strict):
    name: str = Field(min_length=1, max_length=150)
    section: Literal['client', 'bizdev'] = 'client'
    account_manager_id: str = Field(max_length=36)
    operator_id: str | None = Field(None, max_length=36)


class Setup(Strict):
    client_id: str | None = Field(None, max_length=36)
    stage: Literal['Not started', 'Purchase made', 'Setup complete'] | None = None
    id: str | None = Field(None, max_length=36)
    account_manager_id: str = Field(max_length=36)
    operator_id: str = Field(max_length=36)
    client_name: str = Field('', max_length=150)
    domain: str = Field('', max_length=253)
    email: str = Field('', max_length=254)
    section: Literal['client', 'bizdev'] = 'client'
    device_id: str | None = Field(None, max_length=36)
    order: int = Field(1, ge=1, le=1000)
    notes: str = Field('', max_length=1000)


class Stage(Strict):
    id: str = Field(max_length=36)
    stage: Literal['Not started', 'Purchase made', 'Setup complete']


class Identifier(Strict):
    id: str = Field(max_length=36)


class Change(Strict):
    revision: int = Field(ge=0)
    action: Literal['client_save', 'member_save', 'member_remove', 'devices_add', 'device_update', 'device_remove', 'setup_save', 'setup_stage', 'setup_remove']
    data: dict


@router.get('/gmail-set')
@router.get('/domain-cluster-builder', include_in_schema=False)
def page(request: Request, user=Depends(access)):
    request.session.setdefault('domain_csrf', secrets.token_urlsafe(32))
    return templates.TemplateResponse(request, 'gmail_setup.html', {'csrf': request.session['domain_csrf']})


@router.get('/api/gmail-setup')
def get_board(user=Depends(access)):
    return service.result(service.board(user))


@router.get('/api/gmail-setup/summary')
def get_summary(user=Depends(access)):
    from app.utils.gmail_setup import store
    record = store.load(user)
    return service.summary(record['state'] if record else {'devices': [], 'setups': []})


@router.post('/api/gmail-setup/change')
def change(body: Change, user=Depends(mutation_access)):
    models = {'client_save': Client, 'member_save': Member, 'member_remove': Identifier, 'devices_add': Devices,
              'device_update': Device, 'device_remove': Identifier, 'setup_save': Setup,
              'setup_stage': Stage, 'setup_remove': Identifier}
    from pydantic import ValidationError
    from fastapi import HTTPException
    try:
        data = models[body.action].model_validate(body.data).model_dump(exclude_unset=body.action == 'device_update')
        if body.action == 'device_update' and 'status' in data and data['status'] is None:
            raise HTTPException(400, 'Choose a device status.')
    except ValidationError:
        raise HTTPException(400, 'Check the required fields, device counts and selected values.') from None
    record = service.mutate(user, body.revision, body.action, data)
    if body.action in {'setup_stage', 'setup_save'}:
        job = next((j for j in record['state']['setups'] if j['id'] == data.get('id')), None) if data.get('id') else record['state']['setups'][-1]
        if job and job['stage'] == 'Setup complete' and job.get('sync', {}).get('status') != 'synced':
            record = service.sync_completed(user, job['id'])
    return service.result(record)


@router.post('/api/gmail-setup/setups/{job_id}/sync')
def sync(job_id: UUID, user=Depends(mutation_access)):
    return service.result(service.sync_completed(user, str(job_id)))


@router.post('/api/gmail-setup/check-sheet')
def check_sheet(user=Depends(mutation_access)):
    return sheets.readiness()
