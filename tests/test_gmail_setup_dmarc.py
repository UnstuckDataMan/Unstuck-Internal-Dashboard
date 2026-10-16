"""Confirmed, client-scoped DMARC writes with readback and safe retries."""
import copy
import re
import pytest
from fastapi import HTTPException
from app import auth
from app.utils.gmail_setup import dmarc as d, service as s, store
from app.utils.domain_cluster import provisioning as p
from test_gmail_setup import board_db, USER, OTHER, people, change, create_setup
from test_domain_setup import Registrar


@pytest.fixture
def registrar(monkeypatch):
    client=Registrar()
    monkeypatch.setattr(p,'Porkbun',lambda:client)
    return client


def writes(client):return [c for c in client.calls if c[0]=='post']


def start(client_id):
    return d.start(USER,client_id,s.board(USER)['revision'],True)


def apply(record,client_id,job_id):
    batch=next(c for c in record['state']['clients'] if c['id']==client_id)['dmarc_batch']
    return d.apply(USER,client_id,batch['id'],job_id,record['revision'])


def test_exact_policy_uses_sending_email_and_existing_record_is_idempotent(registrar):
    result=d.apply_record('newclient.com','ashley@newclient.com','stable-key')
    assert result['status']=='verified'
    assert result['content']=='v=DMARC1; p=none; pct=100; rua=mailto:ashley@newclient.com; ruf=mailto:ashley@newclient.com; fo=1'
    assert result['host']=='_dmarc.newclient.com' and result['ttl']==600
    assert len(writes(registrar))==1
    assert writes(registrar)[0][1]=='/dns/create/newclient.com'
    assert writes(registrar)[0][2]=={'name':'_dmarc','type':'TXT','content':result['content'],'ttl':600}
    assert writes(registrar)[0][3]
    assert d.apply_record('newclient.com','ashley@newclient.com','stable-key')['status']=='verified'
    assert len(writes(registrar))==1
    assert all('/getUrlForwarding/' not in path for _,path,*_ in registrar.calls)


@pytest.mark.parametrize('record',[
    {'type':'TXT','content':'v=DMARC1; p=reject','ttl':600},
    {'type':'CNAME','content':'elsewhere.example.org','ttl':600},
])
def test_conflicting_dns_is_not_overwritten_or_duplicated(registrar,record):
    registrar.records['newclient.com']=[{**record,'name':'_dmarc.newclient.com','id':'existing'}]
    with pytest.raises(HTTPException,match='not be overwritten'):
        d.apply_record('newclient.com','me@newclient.com','key')
    assert not writes(registrar)


@pytest.mark.parametrize('problem',['ownership','api_access','nameservers'])
def test_domain_must_be_owned_api_enabled_and_on_porkbun_dns(registrar,problem):
    original=registrar.call
    def call(method,path,*args,**kwargs):
        result=original(method,path,*args,**kwargs)
        if path.startswith('/domain/get/'):
            if problem=='ownership':result['domain']['domain']='anotherclient.com'
            if problem=='api_access':result['domain']['apiAccess']=0
        if path.startswith('/domain/getNs/') and problem=='nameservers':result['ns']=['ns1.other.example']
        return result
    registrar.call=call
    with pytest.raises(HTTPException):d.apply_record('newclient.com','me@newclient.com','key')
    assert not writes(registrar)


def test_confirmation_scope_missing_email_and_refresh_persistence(board_db,registrar):
    owner=people()['Dylan'];record,jid=create_setup(owner)
    client_id=record['state']['setups'][0]['client_id']
    _,missing=create_setup(owner,domain='missingemail.com',email='')
    _,deleted=create_setup(owner,domain='deletedclient.com',email='me@deletedclient.com')
    change('setup_remove',{'id':deleted})
    another=change('client_save',{'name':'Other client','section':'client','account_manager_id':owner})['state']['clients'][-1]
    change('setup_bulk',{'client_id':another['id'],'domains':'unrelated.com','sender_name':'me'})
    unrelated=s.board(USER)['state']['setups'][-1]['id']
    before=s.board(USER)
    with pytest.raises(HTTPException):d.start(USER,client_id,before['revision'],False)
    with pytest.raises(HTTPException):d.start(USER,client_id,0,True)
    with pytest.raises(HTTPException):d.start(OTHER,client_id,before['revision'],True)
    assert s.board(USER)==before and not registrar.calls
    record=start(client_id)
    batch=next(c for c in record['state']['clients'] if c['id']==client_id)['dmarc_batch']
    assert {i['id'] for i in batch['items']}=={jid,missing}
    assert batch['results'][missing]['status']=='needs review'
    assert not registrar.calls # Opening/confirming the batch itself cannot write DNS.
    with pytest.raises(HTTPException):apply(record,client_id,unrelated)
    record=apply(record,client_id,jid)
    assert record['state']['setups'][0]['dmarc']['status']=='verified'
    assert record['state']['setups'][0]['stage']=='Not started'
    assert len(writes(registrar))==1
    assert s.board(USER)['state']['clients'][0]['dmarc_batch']['results'][jid]['status']=='verified'
    assert 'sync_lock' not in s.result(record)['state']
    apply(record,client_id,jid)
    assert len(writes(registrar))==1


def test_email_change_after_confirmation_prevents_old_recipient_write(board_db,registrar):
    owner=people()['Dylan'];record,jid=create_setup(owner)
    client_id=record['state']['setups'][0]['client_id'];start(client_id)
    record=change('setup_assign',{'id':jid,'email':'ashley@cedarclients.com'})
    with pytest.raises(HTTPException,match='sending email changed'):apply(record,client_id,jid)
    assert not registrar.calls


def test_dns_claim_serializes_mutations_and_provider_errors_release_it(board_db,monkeypatch):
    owner=people()['Dylan'];record,jid=create_setup(owner)
    client_id=record['state']['setups'][0]['client_id'];record=start(client_id)
    def request(*args):
        stored=store.load(USER)
        assert stored['state']['sync_lock']['kind']=='dmarc'
        with pytest.raises(HTTPException,match='DMARC'):
            change('capacity_set',{'phone_setups':3,'pc_setups':2,'order_kind':'Phone'})
        raise HTTPException(503,'Porkbun credentials need attention.')
    monkeypatch.setattr(d,'apply_record',request)
    record=apply(record,client_id,jid)
    assert record['state']['sync_lock'] is None
    assert record['state']['setups'][0]['dmarc']['status']=='needs review'
    assert 'credentials' in record['state']['setups'][0]['dmarc']['message']


def test_lost_write_response_can_retry_without_duplicate_record(board_db,registrar):
    original=registrar.call
    def call(method,path,*args,**kwargs):
        result=original(method,path,*args,**kwargs)
        if path.startswith('/dns/create/'):
            raise HTTPException(502,'Provider response lost; check whether write completed.')
        return result
    registrar.call=call
    owner=people()['Dylan'];record,jid=create_setup(owner)
    client_id=record['state']['setups'][0]['client_id']
    first=apply(start(client_id),client_id,jid)
    assert first['state']['setups'][0]['dmarc']['status']=='needs review'
    second=apply(start(client_id),client_id,jid)
    assert second['state']['setups'][0]['dmarc']['status']=='verified'
    assert len(writes(registrar))==1


def test_dmarc_route_requires_csrf_admin_and_explicit_confirmation(client,board_db,registrar,monkeypatch):
    owner=people()['Dylan'];record,jid=create_setup(owner)
    client_id=record['state']['setups'][0]['client_id']
    user={**USER};monkeypatch.setattr(auth,'get_session_user',lambda request:user)
    page=client.get('/gmail-set');csrf=re.search('data-csrf="([^"]+)"',page.text)[1]
    endpoint=f'/api/gmail-setup/clients/{client_id}/dmarc/start'
    payload={'revision':record['revision'],'confirmed_purchased':True}
    assert client.post(endpoint,json=payload).status_code==403
    client.headers['X-Domain-CSRF']=csrf
    assert client.post(endpoint,json={**payload,'confirmed_purchased':False}).status_code==422
    user['role']='sdr';assert client.post(endpoint,json=payload).status_code==403
    user['role']='admin';r=client.post(endpoint,json=payload)
    assert r.status_code==200 and not registrar.calls
    current=r.json();batch=current['state']['clients'][0]['dmarc_batch']
    result=client.post(f'/api/gmail-setup/clients/{client_id}/dmarc/{batch["id"]}/{jid}',json={'revision':current['revision']})
    assert result.status_code==200 and result.json()['state']['setups'][0]['dmarc']['status']=='verified'
