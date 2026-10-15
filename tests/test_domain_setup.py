"""Current-batch setup: no purchases, unrelated domains or destructive DNS changes."""
import copy

import pytest
from fastapi import HTTPException
from app.utils.domain_cluster import engine as e, provisioning as p
# Reuse the realistic owner-filtered, revision-checked private storage harness.
from test_domain_cluster import signed_client, portfolio_db, state_for, api_working


def options(**changes):
    return {'forward':True,'destination':'https://example.org','redirect_code':'301','include_www':True,'include_path':True,'dmarc':True,'report_user':'me','report_users':{},**changes}


class Registrar:
    def __init__(self):
        self.calls=[]; self.forwards={}; self.records={}; self.fail_www=False; self.readback_bad=False
    def check(self): return {'credentials_valid':True,'read':True,'dns':True}
    def call(self,method,path,data=None,idempotency=None):
        self.calls.append((method,path,copy.deepcopy(data),idempotency))
        domain=path.split('/')[-1]
        if path.startswith('/domain/get/'):
            return {'status':'SUCCESS','domain':{'domain':domain,'status':'ACTIVE','apiAccess':1}}
        if path.startswith('/domain/getNs/'):
            return {'status':'SUCCESS','ns':['curitiba.ns.porkbun.com','maceio.ns.porkbun.com']}
        if path.startswith('/domain/getUrlForwarding/'):
            return {'status':'SUCCESS','forwards':copy.deepcopy(self.forwards.get(domain,[]))}
        if path.startswith('/dns/retrieve/'):
            return {'status':'SUCCESS','records':copy.deepcopy(self.records.get(domain,[]))}
        if path.startswith('/domain/addUrlForward/'):
            if self.fail_www and data['subdomain']=='www': raise HTTPException(502,'Provider timeout; verify before retrying.')
            self.forwards.setdefault(domain,[]).append({**data,'id':str(len(self.calls))})
            return {'status':'SUCCESS'}
        if path.startswith('/dns/create/'):
            self.records.setdefault(domain,[]).append({**data,'name':data['name']+'.'+domain,'id':'123'})
            return {'status':'SUCCESS','id':'123'}
        raise AssertionError(path)


def test_dmarc_exact_default_and_known_names():
    assert p.dmarc_value('summitgrowth.com',options()) == 'v=DMARC1; p=none; pct=100; rua=mailto:me@summitgrowth.com; ruf=mailto:me@summitgrowth.com; fo=1'
    named=options(report_user='richard')
    assert 'rua=mailto:richard@summitgrowth.com' in p.dmarc_value('summitgrowth.com',named)
    assert 'ruf=mailto:richard@summitgrowth.com' in p.dmarc_value('summitgrowth.com',named)
    named['report_users']={'summitgrowth.com':'alex'}
    assert 'mailto:alex@summitgrowth.com' in p.dmarc_value('summitgrowth.com',named)
    assert 'mailto:richard@other.com' in p.dmarc_value('other.com',named)


@pytest.mark.parametrize('name',['me@domain.com','x; rua=evil','two names','', '<script>'])
def test_bad_report_names(name):
    with pytest.raises(ValueError):p.validate_options(options(report_user=name))


@pytest.mark.parametrize('url',['javascript:alert(1)','https://user:pw@example.com','not-a-url','https://example.com/with spaces'])
def test_bad_forwarding_urls(url):
    with pytest.raises(ValueError):p.validate_options(options(destination=url))


def test_create_then_readback_and_safe_retry():
    fake=Registrar(); plan=p.preview(['cedargrowth.com'],options(),fake)[0]
    assert all(a['action']=='create' for a in plan['actions'])
    assert not any(m=='post' for m,*_ in fake.calls)
    result=p.apply_one('cedargrowth.com',plan,options(),'preview-one',fake)
    assert result['status']=='verified'
    writes=[c for c in fake.calls if c[0]=='post']
    assert len(writes)==3
    assert all(c[3] for c in writes)
    dns=next(c for c in writes if c[1].startswith('/dns/create'))
    assert dns[2]['name']=='_dmarc' and dns[2]['ttl']==600 and dns[2]['type']=='TXT'
    count=len(writes)
    again=p.apply_one('cedargrowth.com',plan,options(),'preview-one',fake)
    assert again['status']=='verified'
    assert len([c for c in fake.calls if c[0]=='post'])==count
    assert not any(any(word in path for word in ('delete','updateNs','/domain/create','/dns/edit')) for _,path,_,_ in fake.calls)


def test_dmarc_conflicts_and_cname_not_overwritten():
    for record in ({'type':'TXT','content':'v=DMARC1; p=reject','ttl':600},{'type':'CNAME','content':'elsewhere.example.org','ttl':600}):
        fake=Registrar();fake.records['cedargrowth.com']=[{**record,'name':'_dmarc.cedargrowth.com','id':'456'}]
        plan=p.preview(['cedargrowth.com'],options(),fake)[0]
        assert plan['status']=='blocked'
        with pytest.raises(HTTPException):p.apply_one('cedargrowth.com',plan,options(),'preview',fake)
        assert not any(c[0]=='post' for c in fake.calls)


def test_existing_forward_conflict_not_duplicated():
    fake=Registrar();fake.forwards['cedargrowth.com']=[{'subdomain':'','location':'https://other.example.org','type':'permanent','includePath':'yes','wildcard':'no'}]
    plan=p.preview(['cedargrowth.com'],options(),fake)[0]
    assert plan['status']=='blocked' and plan['conflicts']
    assert not any(c[0]=='post' for c in fake.calls)


def test_dns_change_since_preview_blocks_every_write():
    fake=Registrar();plan=p.preview(['cedargrowth.com'],options(),fake)[0]
    fake.forwards['cedargrowth.com']=[{'subdomain':'elsewhere','location':'https://example.net'}]
    with pytest.raises(HTTPException):p.apply_one('cedargrowth.com',plan,options(),'preview',fake)
    assert not any(c[0]=='post' for c in fake.calls)


def test_partial_failure_can_be_repreviewed_without_duplicate_root_forward():
    fake=Registrar();plan=p.preview(['cedargrowth.com'],options(),fake)[0];fake.fail_www=True
    result=p.apply_one('cedargrowth.com',plan,options(),'preview',fake)
    assert result['status']=='needs review' and result['performed']==['forward:']
    assert not fake.records
    fake.fail_www=False;plan=p.preview(['cedargrowth.com'],options(),fake)[0]
    assert next(a for a in plan['actions'] if a['kind']=='forward' and a['host']=='')['action']=='already configured'
    result=p.apply_one('cedargrowth.com',plan,options(),'second-preview',fake)
    assert result['status']=='verified'
    assert len(fake.forwards['cedargrowth.com'])==2 and len(fake.records['cedargrowth.com'])==1


def test_external_dns_or_missing_api_access_blocks_domain():
    fake=Registrar();original=fake.call
    def external(method,path,*a,**kw):
        if '/getNs/' in path:return {'ns':['someone.ns.cloudflare.com']}
        return original(method,path,*a,**kw)
    fake.call=external
    plan=p.preview(['cedargrowth.com'],options(),fake)[0]
    assert plan['status']=='blocked' and 'nameservers' in plan['conflicts'][0]


@pytest.fixture
def batch(signed_client,portfolio_db,monkeypatch):
    client,user=signed_client;user['role']='admin'
    state=state_for(['existinggrowth.com'])
    saved=client.post('/api/domain-cluster/portfolios',json={'state':api_working(state)}).json()
    candidates=e.generate(state)['recommendations'][:2]
    result=client.post('/api/domain-cluster/portfolios/'+saved['portfolio']['id']+'/update',json={'revision':1,'action':'planned','candidates':candidates}).json()
    fake=Registrar();monkeypatch.setattr(p,'Porkbun',lambda:fake)
    return client,user,result,fake


def test_api_current_selected_batch_only_and_receipts_survive_save(batch):
    client,user,result,fake=batch
    pid=result['portfolio']['id'];current=result['state']['setup_batches'][-1]
    target=current['domains'][0]
    root=f'/api/domain-cluster/portfolios/{pid}/setup'
    bad=client.post(root+'/preview',json={'revision':2,'batch_id':current['id'],'domains':['existinggrowth.com'],'options':options()})
    assert bad.status_code==400 and fake.calls==[]
    response=client.post(root+'/preview',json={'revision':2,'batch_id':current['id'],'domains':[target],'options':options()})
    assert response.status_code==200,response.text
    plan=response.json();preview=plan['state']['setup_batches'][-1]['preview']
    response=client.post(root+'/apply',json={'revision':plan['portfolio']['revision'],'batch_id':current['id'],'preview_id':preview['id'],'domain':target})
    assert response.status_code==200,response.text
    applied=response.json();assert applied['state']['setup_batches'][-1]['preview']['results'][target]['status']=='verified'
    assert all(path.endswith('/'+target) for _,path,_,_ in fake.calls)
    again=client.post(root+'/apply',json={'revision':applied['portfolio']['revision'],'batch_id':current['id'],'preview_id':preview['id'],'domain':current['domains'][1]})
    assert again.status_code==400
    save=client.post('/api/domain-cluster/portfolios',json={'portfolio_id':pid,'revision':applied['portfolio']['revision'],'state':api_working(applied['state'])})
    assert save.status_code==200
    assert save.json()['state']['setup_batches']==applied['state']['setup_batches']


def test_api_admin_and_revision_gate_precedes_registrar_changes(batch):
    client,user,result,fake=batch
    pid=result['portfolio']['id'];current=result['state']['setup_batches'][-1]
    body={'revision':2,'batch_id':current['id'],'domains':[current['domains'][0]],'options':options()}
    user['role']='sdr'
    assert client.post(f'/api/domain-cluster/portfolios/{pid}/setup/preview',json=body).status_code==403
    user['role']='admin';body['revision']=1
    assert client.post(f'/api/domain-cluster/portfolios/{pid}/setup/preview',json=body).status_code==409
    assert fake.calls==[]


def test_preview_requires_dns_scope():
    fake=Registrar();fake.check=lambda:{'credentials_valid':True,'read':True,'dns':False}
    with pytest.raises(HTTPException) as exc:p.preview(['cedargrowth.com'],options(),fake)
    assert exc.value.status_code==403 and fake.calls==[]


def test_provider_errors_expose_machine_code_not_secret_or_raw_message(monkeypatch):
    from conftest import FakeResponse
    monkeypatch.setenv('PORKBUN_API_KEY','key');monkeypatch.setenv('PORKBUN_SECRET_API_KEY','private-secret')
    monkeypatch.setattr(p.requests,'post',lambda *a,**kw:FakeResponse(400,{'status':'ERROR','code':'INVALID_API_KEYS_002','message':'reflected private-secret'}))
    with pytest.raises(HTTPException) as exc:p.Porkbun().check()
    assert 'INVALID_API_KEYS_002' in exc.value.detail
    assert 'private-secret' not in exc.value.detail
