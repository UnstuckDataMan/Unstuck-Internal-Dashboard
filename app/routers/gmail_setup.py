"""Shared Gmail setup board, protected by dashboard tool access and CSRF."""
import secrets
import hashlib
from pathlib import Path
from typing import Literal
from uuid import UUID
from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel, ConfigDict, Field
from app.deps import templates
from app.routers.domain_cluster import access, mutation_access
from app.utils.gmail_setup import service

router = APIRouter()
_STATIC = Path(__file__).resolve().parents[1] / 'static'
ASSET_VERSION = hashlib.sha256(b''.join((_STATIC / name).read_bytes() for name in
    ('gmail-setup.js', 'gmail-setup.css'))).hexdigest()[:12]


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
    use_count: int | None = Field(None, ge=0, le=100, strict=True)


class Device(Strict):
    id: str = Field(max_length=36)
    member_id: str | None = Field(None, max_length=36)
    status: Literal['Unused', 'Used', 'Blocked'] | None = None
    block_reason: Literal['QR code', 'Payment', 'Other'] | None = None
    asset_tag: str | None = Field(None, min_length=1, max_length=80)


class Client(Strict):
    id: str | None = Field(None, max_length=36)
    name: str = Field(min_length=1, max_length=150)
    section: Literal['client', 'bizdev'] = 'client'
    account_manager_id: str = Field(max_length=36)
    operator_id: str | None = Field(None, max_length=36)
    planned_domains: int | None = Field(None, ge=0, le=5000, strict=True)
    domain_target: int | None = Field(None, ge=0, le=5000, strict=True)
    sender_full_name: str | None = Field(None, max_length=150)
    website_url: str | None = Field(None, max_length=2048)
    account_manager_email: str | None = Field(None, max_length=254)


class Setup(Strict):
    client_id: str | None = Field(None, max_length=36)
    stage: Literal['Not started', 'Started', 'Purchase made', 'Setup complete', 'Failed at QR', 'Failed at payment', 'Burned'] | None = None
    id: str | None = Field(None, max_length=36)
    account_manager_id: str = Field(max_length=36)
    operator_id: str | None = Field(None, max_length=36)
    client_name: str = Field('', max_length=150)
    domain: str = Field('', max_length=253)
    email: str = Field('', max_length=254)
    section: Literal['client', 'bizdev'] = 'client'
    device_id: str | None = Field(None, max_length=36)
    order: int = Field(1, ge=1, le=1000)
    notes: str = Field('', max_length=1000)


class Stage(Strict):
    id: str = Field(max_length=36)
    stage: Literal['Not started', 'Started', 'Purchase made', 'Setup complete', 'Failed at QR', 'Failed at payment', 'Burned']


class SetupDevice(Strict):
    id: str = Field(max_length=36)
    operator_id: str = Field(max_length=36)
    device_id: str = Field(max_length=36)


class Identifier(Strict):
    id: str = Field(max_length=36)


class Reserve(Strict):
    target: int = Field(ge=1, le=5000)


class Capacity(Strict):
    phone_setups: int = Field(ge=1, le=100, strict=True)
    pc_setups: int = Field(ge=1, le=100, strict=True)
    order_kind: Literal['Phone', 'PC'] = 'Phone'


class Bulk(Strict):
    client_id: str = Field(max_length=36)
    domains: str = Field(min_length=1, max_length=50000)
    sender_name: str = Field('', max_length=64)
    operator_id: str | None = Field(None, max_length=36)


class Assignment(Strict):
    id: str = Field(max_length=36)
    operator_id: str | None = Field(None, max_length=36)
    device_id: str | None = Field(None, max_length=36)
    email: str | None = Field(None, max_length=254)


class Notes(Strict):
    id: str = Field(max_length=36)
    notes: str = Field('', max_length=1000)


class Change(Strict):
    revision: int = Field(ge=0)
    action: Literal['reserve_set', 'capacity_set', 'client_save', 'client_remove', 'client_complete', 'setup_bulk', 'setup_assign', 'setup_notes', 'member_save', 'member_remove', 'devices_add', 'device_update', 'device_remove', 'setup_save', 'setup_stage', 'setup_device', 'setup_remove']
    data: dict
    defer_sync: bool = False


@router.get('/gmail-set')
@router.get('/domain-cluster-builder', include_in_schema=False)
def page(request: Request, user=Depends(access)):
    request.session.setdefault('domain_csrf', secrets.token_urlsafe(32))
    return templates.TemplateResponse(request, 'gmail_setup.html', {
        'csrf': request.session['domain_csrf'], 'gmail_asset_version': ASSET_VERSION,
        'gmail_can_manage_dns': user.get('role') == 'admin'})


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
    models = {'reserve_set': Reserve, 'capacity_set': Capacity, 'client_save': Client, 'client_remove': Identifier, 'client_complete': Identifier, 'setup_bulk': Bulk, 'setup_assign': Assignment, 'setup_notes': Notes,
              'member_save': Member, 'member_remove': Identifier, 'devices_add': Devices,
              'device_update': Device, 'device_remove': Identifier, 'setup_save': Setup,
              'setup_stage': Stage, 'setup_device': SetupDevice, 'setup_remove': Identifier}
    from pydantic import ValidationError
    from fastapi import HTTPException
    try:
        data = models[body.action].model_validate(body.data).model_dump(exclude_unset=body.action in {'device_update', 'setup_assign', 'client_save'})
        if body.action == 'device_update' and 'status' in data and data['status'] is None:
            raise HTTPException(400, 'Choose a device status.')
    except ValidationError:
        raise HTTPException(400, 'Check the required fields, device counts and selected values.') from None
    record = service.mutate(user, body.revision, body.action, data)
    return service.result(record)


@router.post('/api/gmail-setup/setups/{job_id}/sync')
def sync(job_id: UUID, user=Depends(mutation_access)):
    raise HTTPException(410, 'Master Strategy sheet updates are manual. Add completed sending emails to the sheet yourself.')


@router.post('/api/gmail-setup/check-sheet')
def check_sheet(user=Depends(mutation_access)):
    raise HTTPException(410, 'Master Strategy sheet updates are manual. Add completed sending emails to the sheet yourself.')


class DmarcStart(Strict):
    revision: int = Field(ge=0)
    confirmed_purchased: Literal[True]
    setup_ids: list[UUID] | None = Field(None, min_length=1, max_length=5000)


class DmarcApply(Strict):
    revision: int = Field(ge=1)


@router.post('/api/gmail-setup/clients/{client_id}/dmarc/start')
def start_dmarc(client_id: UUID, body: DmarcStart, user=Depends(mutation_access)):
    from app.utils.gmail_setup import dmarc
    return service.result(dmarc.start(user, str(client_id), body.revision, body.confirmed_purchased,
        [str(identifier) for identifier in body.setup_ids] if body.setup_ids is not None else None))


@router.post('/api/gmail-setup/clients/{client_id}/dmarc/{batch_id}/{job_id}')
def apply_dmarc(client_id: UUID, batch_id: UUID, job_id: UUID, body: DmarcApply, user=Depends(mutation_access)):
    from app.utils.gmail_setup import dmarc
    return service.result(dmarc.apply(user, str(client_id), str(batch_id), str(job_id), body.revision))


class DnsSetupPreview(Strict):
    revision: int = Field(ge=0)


class DnsSetupApply(DnsSetupPreview):
    preview_id: UUID
    confirmed_purchased: Literal[True]
    destination: str = Field(min_length=1, max_length=2048)


@router.post('/api/gmail-setup/setups/{job_id}/dns-setup/preview')
def preview_dns_setup(job_id: UUID, body: DnsSetupPreview, user=Depends(mutation_access)):
    from app.utils.gmail_setup import dns_setup
    return service.result(dns_setup.preview(user, str(job_id), body.revision))


@router.post('/api/gmail-setup/setups/{job_id}/dns-setup/apply')
def apply_dns_setup(job_id: UUID, body: DnsSetupApply, user=Depends(mutation_access)):
    from app.utils.gmail_setup import dns_setup
    return service.result(dns_setup.apply(user, str(job_id), str(body.preview_id), body.revision, body.confirmed_purchased, body.destination))


class DnsBatchStart(DnsSetupPreview):
    confirmed_purchased: Literal[True]
    destination: str = Field(min_length=1, max_length=2048)


@router.post('/api/gmail-setup/clients/{client_id}/dns-setup/start')
def start_dns_batch(client_id: UUID, body: DnsBatchStart, user=Depends(mutation_access)):
    from app.utils.gmail_setup import dns_batch
    return service.result(dns_batch.start(user, str(client_id), body.revision, body.confirmed_purchased, body.destination))


@router.post('/api/gmail-setup/clients/{client_id}/dns-setup/{batch_id}/{job_id}')
def step_dns_batch(client_id: UUID, batch_id: UUID, job_id: UUID, body: DmarcApply, user=Depends(mutation_access)):
    from app.utils.gmail_setup import dns_batch
    return service.result(dns_batch.step(user, str(client_id), str(batch_id), str(job_id), body.revision))
