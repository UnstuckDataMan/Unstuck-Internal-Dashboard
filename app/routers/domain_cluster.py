"""Gmail Set / Domain Cluster Builder. All portfolio reads are owner scoped."""
from __future__ import annotations

import copy
import re
import secrets
from typing import Literal
from uuid import UUID, uuid4

from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel, Field, ConfigDict

from app import auth
from app.deps import templates
from app.utils.domain_cluster import engine as e, store, enrichment, sheets, provisioning

router = APIRouter()


class Strict(BaseModel):
    model_config = ConfigDict(extra='forbid')


class Settings(Strict):
    capacity: int = Field(10, ge=1, le=10)
    watch: int = Field(7, ge=1, le=10)
    near: int = Field(9, ge=1, le=10)
    sensitivity: float = Field(.72, ge=.3, le=.99)
    dictionary: dict[str, list[str]] = Field(default_factory=dict)
    themes: list[str] = Field(default_factory=list, max_length=11)
    industry: str = Field('', max_length=100)
    geography: str = Field('', max_length=100)
    default_import: Literal['add', 'replace'] = 'add'


class Domain(Strict):
    domain: str = Field(max_length=253)
    status: Literal['Active', 'Paused', 'Planned', 'Retired'] = 'Active'
    owner: str = Field('', max_length=100)
    platform: str = Field('', max_length=100)
    notes: str = Field('', max_length=1000)
    added_at: str = Field('', max_length=50)
    last_analysed_at: str = Field('', max_length=50)
    override: str | None = Field(None, max_length=80)
    reviewed: bool = False
    sheet_sources: list[Literal['client', 'bizdev']] = Field(default_factory=list, max_length=2)
    sheet_owners: list[str] = Field(default_factory=list, max_length=100)
    sheet_at_risk: bool = False
    sheet_needs_deleting: bool = False
    sheet_replacing: bool = False
    source_checked_at: str = Field('', max_length=50)


class Cluster(Strict):
    name: str = Field(min_length=1, max_length=100)


class WorkingState(Strict):
    domains: list[Domain] = Field(default_factory=list, max_length=e.MAX_DOMAINS)
    clusters: dict[str, Cluster] = Field(default_factory=dict)
    settings: Settings = Field(default_factory=Settings)


class Analysis(Strict):
    state: WorkingState = Field(default_factory=WorkingState)


class Import(Analysis):
    text: str = Field('', max_length=e.MAX_INPUT)
    mode: Literal['add', 'replace'] = 'add'


class Generate(Analysis):
    batch: int = Field(0, ge=0, le=1000000)
    require_available: bool = False
    live_sheet: bool = False


class Save(Analysis):
    portfolio_id: UUID | None = None
    name: str = Field('My domain portfolio', min_length=1, max_length=100)
    revision: int = Field(0, ge=0)
    action: str = Field('Save portfolio', max_length=100)


class Mutation(Strict):
    revision: int = Field(ge=1)
    action: Literal['import', 'domain', 'cluster', 'merge', 'planned', 'settings', 'undo', 'reanalyse']
    text: str = Field('', max_length=e.MAX_INPUT)
    mode: Literal['add', 'replace'] = 'add'
    domain: str = Field('', max_length=253)
    changes: dict = Field(default_factory=dict)
    cluster: str = Field('', max_length=80)
    name: str = Field('', max_length=100)
    target: str = Field('', max_length=80)
    candidates: list[dict] = Field(default_factory=list, max_length=10)
    settings: Settings | None = None


def access(request: Request):
    # Even AUTH_DISABLED cannot expose saved private inventories on the internet.
    user = auth.get_session_user(request)
    if not user:
        raise HTTPException(401, 'Sign in to use Gmail Set.')
    if 'gmail_set' not in user.get('tools', []):
        raise HTTPException(403, 'Gmail Set access required.')
    return user


def mutation_access(request: Request, user=Depends(access)):
    token = request.headers.get('X-Domain-CSRF', '')
    expected = request.session.get('domain_csrf', '')
    if not token or not expected or not secrets.compare_digest(token, expected):
        raise HTTPException(403, 'Session verification failed. Refresh the page and retry.', headers={'X-Gmail-Session': 'refresh'})
    return user


def working(model):
    state = e.empty_state()
    data = model.model_dump()
    e.settings(data['settings'])
    if len(data['clusters']) > 100:
        raise ValueError('Limit is 100 custom clusters.')
    dictionary = data['settings']['dictionary']
    if sum(len(v) for v in dictionary.values()) > 300 or any(not re.fullmatch('[a-z]{3,25}', w) for words in dictionary.values() for w in words):
        raise ValueError('Custom dictionary allows up to 300 lowercase words, each 3–25 letters.')
    seen = set()
    for row in data['domains']:
        root = e.normalize(row['domain'])
        if root in seen:
            raise ValueError('Duplicate domain in portfolio.')
        seen.add(root)
        row['domain'] = root
        row['added_at'] = row['added_at'] or e.now()
    state.update(data)
    e.dictionary(data['settings'])
    e.analyse(state)  # Validate overrides and dictionary, including empty states.
    return state


def response(state, **extra):
    # Raw paste and invalid mailbox lines never enter persisted state/history.
    return {'state': state, 'analysis': e.analyse(state), **extra}


@router.get('/domain-cluster-builder/legacy', include_in_schema=False)
def page(request: Request, user=Depends(access)):
    request.session.setdefault('domain_csrf', secrets.token_urlsafe(32))
    return templates.TemplateResponse(request, 'domain_cluster.html', {'csrf': request.session['domain_csrf'], 'families': {k: v[0] for k, v in e.FAMILIES.items()}, 'can_setup':user.get('role')=='admin'})


@router.get('/api/domain-cluster/portfolios')
def portfolios(user=Depends(access)):
    return {'portfolios': store.list_portfolios(user)}


@router.get('/api/domain-cluster/portfolios/{portfolio_id}')
def portfolio(portfolio_id: UUID, user=Depends(access)):
    record = store.load(user, str(portfolio_id))
    return response(record['state'], portfolio={k: record[k] for k in ('id', 'name', 'revision', 'updated_at')})


@router.post('/api/domain-cluster/analyse')
def analyse(body: Analysis, user=Depends(mutation_access)):
    try:
        return response(working(body.state))
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from None


@router.post('/api/domain-cluster/parse')
def parse(body: Import, user=Depends(mutation_access)):
    try:
        state = working(body.state)
        parsed = e.parse(body.text)
        imported, preview = e.apply_import(state, parsed, body.mode)
        return response(imported, parsed=parsed, preview=preview)
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from None


@router.post('/api/domain-cluster/generate')
def generate(body: Generate, user=Depends(mutation_access)):
    try:
        state = working(body.state)
        source = None
        if body.live_sheet:
            source = sheets.read_live()
            sheets.require_valid(source)
            state, _ = sheets.refresh(state, source)
        extra = response(state, source=source) if source else {}
        enrichment.rate_limit(('generation', store.scope(user)[0]), count=20)
        if not body.require_available:
            return {**e.generate(state, body.batch), **extra}
        accepted, checked, excluded = [], [], set()
        # At most 30 checks per request; no endless registry scanning.
        for offset in range(3):
            result = e.generate(state, body.batch + offset, excluded=excluded)
            candidates = result['recommendations']
            if not candidates:
                break
            availability = enrichment.porkbun([c['domain'] for c in candidates])
            for c in candidates:
                excluded.add(c['domain'])
                c.update(availability[c['domain']])
                checked.append(c)
                if c.get('available') is True and e.validate_candidate(c, state, accepted):
                    accepted.append(c)
                if len(accepted) == 10:
                    break
            if len(accepted) == 10:
                break
        if len(accepted) != 10:
            return {**extra, 'recommendations': [], 'shortfall': 10-len(accepted), 'qualifying_count': len(accepted), 'checked': checked, 'message': 'Verified search did not find ten available names within this search budget. Try another search. No incomplete batch was approved.'}
        counts = {c['id']: c['projected'] for c in e.analyse(state)['clusters']}
        for c in accepted:
            counts[c['cluster']] = counts.get(c['cluster'], 0) + 1
        for c in accepted:
            c['projected'] = counts[c['cluster']]
        return {**extra, 'recommendations': accepted, 'shortfall': 0, 'checked_count': len(checked)}
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from None


@router.post('/api/domain-cluster/portfolios')
def save(body: Save, user=Depends(mutation_access)):
    try:
        state = working(body.state)
        if body.portfolio_id:
            previous = store.load(user, str(body.portfolio_id))
            if previous['revision'] != body.revision:
                raise HTTPException(409, 'Portfolio changed. Reload before saving.')
            # Keep server-side history, never trust submitted histories.
            old_state = copy.deepcopy(previous['state'])
            store.checkpoint(old_state, body.action)
            state['history'] = old_state['history']
            state['snapshots'] = old_state.get('snapshots', [])
            state['setup_batches'] = copy.deepcopy(previous['state'].get('setup_batches', []))
            state['recommendations'] = copy.deepcopy(previous['state'].get('recommendations', []))
            ensure_setup_idle(previous['state'])
        stamp = e.now()
        for r in state['domains']:
            r['last_analysed_at'] = stamp
        store.snapshot(state)
        record = store.save(user, str(body.portfolio_id) if body.portfolio_id else None, body.name, state, body.revision)
        return response(record['state'], portfolio={k: record[k] for k in ('id', 'name', 'revision', 'updated_at')})
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from None


@router.post('/api/domain-cluster/portfolios/{portfolio_id}/update')
def update(portfolio_id: UUID, body: Mutation, user=Depends(mutation_access)):
    record = store.load(user, str(portfolio_id))
    if record['revision'] != body.revision:
        raise HTTPException(409, 'Portfolio changed. Reload before editing.')
    state = copy.deepcopy(record['state'])
    ensure_setup_idle(state)
    try:
        if body.action == 'undo':
            if not state['history']:
                raise ValueError('No update to undo.')
            entry = state['history'].pop()
            state.update(entry['before'])
        else:
            store.checkpoint(state, body.action)
            if body.action == 'import':
                parsed = e.parse(body.text)
                if not parsed['domains']:
                    raise ValueError('No valid domains to import.')
                state, _ = e.apply_import(state, parsed, body.mode)
            elif body.action == 'domain':
                allowed = {'status', 'owner', 'platform', 'notes', 'override', 'reviewed'}
                if set(body.changes) - allowed:
                    raise ValueError('Unknown domain field.')
                row = next((r for r in state['domains'] if r['domain'] == body.domain), None)
                if row is None:
                    raise ValueError('Domain not found.')
                row.update(body.changes)
            elif body.action == 'cluster':
                if not body.name.strip():
                    raise ValueError('Cluster name required.')
                key = body.cluster or 'manual-' + uuid4().hex[:12]
                if key not in e.FAMILIES and key != 'review' and not key.startswith('manual-'):
                    raise ValueError('Invalid cluster identifier.')
                state['clusters'][key] = {'name': body.name.strip()}
            elif body.action == 'merge':
                report = e.analyse(state)
                known = {c['id'] for c in report['clusters']} | set(e.FAMILIES)
                if body.cluster == body.target or body.cluster not in known or body.target not in known:
                    raise ValueError('Select two different existing clusters.')
                members = {r['domain'] for r in report['domains'] if r['cluster'] == body.cluster}
                for r in state['domains']:
                    if r['domain'] in members:
                        r['override'] = body.target
                        r['reviewed'] = True
            elif body.action == 'planned':
                if any(r.get('sheet_sources') for r in state['domains']):
                    source = sheets.read_live()
                    sheets.require_valid(source)
                    current, _ = sheets.refresh(state, source)
                    inventory = lambda s: sorted((r['domain'], r['status']) for r in s['domains'])
                    if inventory(current) != inventory(state):
                        raise HTTPException(409, 'Live sheet inventory changed since your last analysis. Analyse the sheet, review and save the changes, then generate a fresh batch.')
                accepted = []
                for candidate in body.candidates:
                    if not e.validate_candidate(candidate, state, accepted):
                        raise ValueError('A selected domain violates naming, duplication or capacity policy. Generate a fresh batch.')
                    accepted.append(candidate)
                for candidate in accepted:
                    state['domains'].append(Domain(domain=candidate['domain'], status='Planned', added_at=e.now()).model_dump())
                if not accepted:
                    raise ValueError('Select at least one new domain.')
                state['setup_batches'] = (state.get('setup_batches',[]) + [{'id':str(uuid4()),'domains':[c['domain'] for c in accepted],'created_at':e.now(),'preview':None}])[-10:]
                # Preserve construction/evidence, never a client-asserted registrar verdict.
                state['recommendations'] = [
                    {'domain': c['domain'], 'words': c['words'],
                     'cluster': e.classify(c['domain'], state['settings'])['cluster'],
                     'availability': 'Availability not checked', 'checked_at': None}
                    for c in accepted
                ]
            elif body.action == 'settings':
                if body.settings is None:
                    raise ValueError('Settings required.')
                state['settings'] = body.settings.model_dump()
            # reanalyse intentionally retains overrides/reviewed decisions.
        validated = working(WorkingState.model_validate({k: state[k] for k in ('domains', 'clusters', 'settings')}))
        state.update({k: validated[k] for k in ('domains', 'clusters', 'settings')})
        store.snapshot(state)
        record = store.save(user, str(portfolio_id), record['name'], state, body.revision)
        return response(record['state'], portfolio={k: record[k] for k in ('id', 'name', 'revision', 'updated_at')})
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from None



class Availability(Strict):
    domains: list[str] = Field(min_length=1, max_length=10)


class DNS(Strict):
    selector: str = Field('', max_length=63)


@router.get('/api/domain-cluster/integrations')
def integrations(user=Depends(access)):
    return {'porkbun': enrichment.configured()}


@router.post('/api/domain-cluster/availability')
def availability(body: Availability, user=Depends(mutation_access)):
    enrichment.rate_limit(('availability', store.scope(user)[0]), count=6)
    return {'results': enrichment.porkbun(body.domains)}


@router.post('/api/domain-cluster/portfolios/{portfolio_id}/dns')
def dns_lookup(portfolio_id: UUID, body: DNS, user=Depends(mutation_access)):
    enrichment.rate_limit(('dns', store.scope(user)[0]), count=2)
    record = store.load(user, str(portfolio_id))
    domains = [r['domain'] for r in record['state']['domains'] if r['status'] == 'Active']
    return enrichment.dns_batch(domains, body.selector)


@router.post('/api/domain-cluster/strategy-sheet')
def strategy_sheet(body: Analysis, user=Depends(mutation_access)):
    try:
        state = working(body.state)
        source = sheets.read_live()
        state, preview = sheets.refresh(state, source)
        parsed = {'unique': len({r['domain'] for r in source['records']}), 'duplicates': preview['duplicates'], 'invalid': source['invalid']}
        return response(state, source=source, preview=preview, parsed=parsed)
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from None


class SetupOptions(Strict):
    forward: bool = True
    destination: str = Field('', max_length=2048)
    redirect_code: Literal['301','302','307'] = '301'
    include_path: bool = True
    include_www: bool = True
    dmarc: bool = True
    report_user: str = Field('me', min_length=1, max_length=64)
    report_users: dict[str,str] = Field(default_factory=dict)


class SetupPreview(Strict):
    batch_id: UUID
    revision: int = Field(ge=1)
    domains: list[str] = Field(min_length=1, max_length=10)
    options: SetupOptions = Field(default_factory=SetupOptions)


class SetupApply(Strict):
    batch_id: UUID
    preview_id: UUID
    revision: int = Field(ge=1)
    domain: str = Field(max_length=253)


def setup_access(user=Depends(mutation_access)):
    if user.get('role') != 'admin':
        raise HTTPException(403,'An administrator must configure forwarding and DNS through the shared Porkbun account.')
    return user


def ensure_setup_idle(state):
    batches=state.get('setup_batches',[])
    if batches and batches[-1].get('preview'):
        preview=batches[-1]['preview']
        if any(r.get('status')=='running' for r in preview.get('results',{}).values()):
            from datetime import datetime, timezone
            if (datetime.now(timezone.utc)-datetime.fromisoformat(preview['started_at'])).total_seconds()<300:
                raise HTTPException(409,'A domain is still being configured. Wait before changing this portfolio.')


@router.post('/api/domain-cluster/integrations/check')
def check_integrations(user=Depends(mutation_access)):
    from app.utils.domain_cluster.diagnostics import check
    return check()


@router.post('/api/domain-cluster/portfolios/{portfolio_id}/setup/preview')
def setup_preview(portfolio_id: UUID, body: SetupPreview, user=Depends(setup_access)):
    enrichment.rate_limit(('setup-preview',store.scope(user)[0]),count=3)
    record=store.load(user,str(portfolio_id))
    if record['revision'] != body.revision:
        raise HTTPException(409,'Portfolio changed. Reload before previewing setup.')
    state=copy.deepcopy(record['state'])
    ensure_setup_idle(state)
    batch=provisioning.latest_batch(state,str(body.batch_id))
    domains=sorted(set(body.domains))
    if any(d not in batch['domains'] for d in domains):
        raise HTTPException(400,'Only selected domains in the current new-domain batch may be configured.')
    try:
        options=provisioning.validate_options(body.options.model_dump())
        if set(options['report_users']) - set(domains):
            raise ValueError('Report-name overrides must refer only to selected domains.')
    except ValueError as exc:
        raise HTTPException(400,str(exc)) from None
    items=provisioning.preview(domains,options)
    batch['preview']={'id':str(uuid4()),'created_at':e.now(),'options':options,'items':items,'results':{}}
    record=store.save(user,str(portfolio_id),record['name'],state,body.revision)
    return response(record['state'],portfolio={k:record[k] for k in ('id','name','revision','updated_at')})


@router.post('/api/domain-cluster/portfolios/{portfolio_id}/setup/apply')
def setup_apply(portfolio_id: UUID, body: SetupApply, user=Depends(setup_access)):
    from datetime import datetime, timezone
    record=store.load(user,str(portfolio_id))
    if record['revision'] != body.revision:
        raise HTTPException(409,'Portfolio changed. Reload before applying setup.')
    state=copy.deepcopy(record['state'])
    batch=provisioning.latest_batch(state,str(body.batch_id))
    preview=batch.get('preview')
    if not preview or preview['id'] != str(body.preview_id):
        raise HTTPException(409,'Create a current setup preview first.')
    if (datetime.now(timezone.utc)-datetime.fromisoformat(preview['created_at'])).total_seconds()>1800:
        raise HTTPException(409,'Setup preview expired. Refresh it before applying changes.')
    item=next((r for r in preview['items'] if r['domain']==body.domain),None)
    if item is None or item['status']!='ready' or body.domain not in batch['domains']:
        raise HTTPException(400,'This domain was not approved in the current preview.')
    if preview['results'].get(body.domain,{}).get('status')=='verified':
        return response(state,portfolio={k:record[k] for k in ('id','name','revision','updated_at')})
    ensure_setup_idle(state)
    preview['started_at']=e.now()
    preview['results'][body.domain]={'domain':body.domain,'status':'running','checked_at':e.now()}
    # Persist the claim before any external write; a concurrent tab loses the CAS.
    claimed=store.save(user,str(portfolio_id),record['name'],state,body.revision)
    try:
        result=provisioning.apply_one(body.domain,item,preview['options'],preview['id'])
    except HTTPException as exc:
        result={'domain':body.domain,'status':'needs review','message':str(exc.detail),'checked_at':e.now()}
    preview['results'][body.domain]=result
    record=store.save(user,str(portfolio_id),record['name'],state,claimed['revision'])
    return response(record['state'],portfolio={k:record[k] for k in ('id','name','revision','updated_at')})
