"""Saved sequential batch progress and registrar safety with synthetic zones."""
import copy
import re
from datetime import datetime, timedelta, timezone
import pytest
from fastapi import HTTPException
from app import auth
from app.utils.gmail_setup import dns_batch as b, dns_setup as dns, service as s, store
from app.utils.domain_cluster import provisioning as p
from test_gmail_setup import board_db, USER, OTHER, people, change
from test_gmail_setup_dns_setup import SetupRegistrar, defaults, retained


@pytest.fixture
def batch_case(board_db,monkeypatch):
    owner=people()['Dylan']
    client=change('client_save',{'name':'Batch client','section':'client','account_manager_id':owner})['state']['clients'][0]
    record=change('setup_bulk',{'client_id':client['id'],'domains':'batchfirst.com\nbatchsecond.com\nbatchburned.com\nbatchdeleted.com','sender_name':'me'})
    jobs=record['state']['setups']
    change('setup_stage',{'id':jobs[2]['id'],'stage':'Burned'});change('setup_remove',{'id':jobs[3]['id']})
    other=change('client_save',{'name':'Unrelated client','section':'client','account_manager_id':owner})['state']['clients'][-1]
    foreign=change('setup_bulk',{'client_id':other['id'],'domains':'batchforeign.com','sender_name':'me'})['state']['setups'][-1]
    registrar=SetupRegistrar()
    for job in jobs+[foreign]:registrar.records[job['domain']]=defaults(job['domain'])+retained(job['domain'])
    monkeypatch.setattr(p,'Porkbun',lambda:registrar)
    return client['id'],jobs,foreign,registrar


def start(cid):return b.start(USER,cid,s.board(USER)['revision'],True,'https://clientwebsite.com')
def batch(record,cid):return next(c for c in record['state']['clients'] if c['id']==cid)['dns_setup_batch']
def step(record,cid,jid):return b.step(USER,cid,batch(record,cid)['id'],jid,record['revision'])


def test_all_domains_saved_progress_scope_sequential_completion_and_replay(batch_case):
    cid,jobs,foreign,registrar=batch_case
    before=s.summary(s.board(USER)['state']);unrelated=copy.deepcopy(registrar.records['batchforeign.com'])
    record=start(cid);plan=batch(record,cid)
    assert [i['id'] for i in plan['items']]==[jobs[0]['id'],jobs[1]['id']]
    assert not registrar.calls
    with pytest.raises(HTTPException,match='next waiting'):step(record,cid,jobs[1]['id'])
    with pytest.raises(HTTPException):step(record,cid,foreign['id'])
    record=step(record,cid,jobs[0]['id'])
    assert batch(record,cid)['results'][jobs[0]['id']]['status']=='verified'
    assert jobs[1]['id'] not in batch(record,cid)['results']
    calls=len(registrar.calls);record=step(record,cid,jobs[0]['id']);assert len(registrar.calls)==calls
    record=step(record,cid,jobs[1]['id'])
    assert all(r['status']=='verified' for r in batch(record,cid)['results'].values())
    assert batch(record,cid)['completed_at'] and not batch(record,cid).get('current_job_id')
    assert registrar.records['batchforeign.com']==unrelated
    assert registrar.records['batchburned.com']==defaults('batchburned.com')+retained('batchburned.com')
    assert s.summary(record['state'])==before


@pytest.mark.parametrize('issue',['missing-email','conflicting-dns','burned','deleted','changed-email'])
def test_one_domain_needs_attention_does_not_block_the_rest(batch_case,issue):
    cid,jobs,foreign,registrar=batch_case
    if issue=='missing-email':change('setup_assign',{'id':jobs[0]['id'],'email':''})
    if issue=='conflicting-dns':registrar.records['batchfirst.com'].append({'id':'999','name':'t.batchfirst.com','type':'CNAME','content':'other-tracker.com','ttl':'600'})
    record=start(cid)
    if issue=='burned':record=change('setup_stage',{'id':jobs[0]['id'],'stage':'Burned'})
    if issue=='deleted':record=change('setup_remove',{'id':jobs[0]['id']})
    if issue=='changed-email':record=change('setup_assign',{'id':jobs[0]['id'],'email':'changed@batchfirst.com'})
    record=step(record,cid,jobs[0]['id'])
    assert batch(record,cid)['results'][jobs[0]['id']]['status']=='needs review'
    assert not any(c[0]=='post' and 'batchfirst.com' in c[1] for c in registrar.calls)
    record=step(record,cid,jobs[1]['id'])
    assert batch(record,cid)['results'][jobs[1]['id']]['status']=='verified'
    assert batch(record,cid)['completed_at']


@pytest.mark.parametrize('phase',['preview','apply'])
def test_rate_limit_pauses_and_can_resume_without_losing_results(batch_case,monkeypatch,phase):
    cid,jobs,foreign,registrar=batch_case;record=start(cid)
    original=dns.preview if phase=='preview' else dns.configure_records
    def limited(*args,**kwargs):raise HTTPException(429,'Porkbun lookup/write limit reached. Wait before retrying.')
    monkeypatch.setattr(dns,'preview' if phase=='preview' else 'configure_records',limited)
    record=step(record,cid,jobs[0]['id'])
    assert batch(record,cid)['results'][jobs[0]['id']]['status']=='pending'
    assert batch(record,cid)['paused_reason'] and not batch(record,cid).get('current_job_id')
    assert not any(c[0]=='post' for c in registrar.calls)
    monkeypatch.setattr(dns,'preview' if phase=='preview' else 'configure_records',original)
    record=step(record,cid,jobs[0]['id'])
    assert batch(record,cid)['results'][jobs[0]['id']]['status']=='verified'
    assert 'paused_reason' not in batch(record,cid)


def test_busy_and_interrupted_domain_claims_prevent_parallel_writes_then_reconcile(batch_case,board_db):
    cid,jobs,foreign,registrar=batch_case;record=start(cid)
    stored=board_db[store.workspace(USER)];plan=batch(stored,cid)
    plan.update(current_job_id=jobs[0]['id'],current_started_at=s.now())
    plan['results'][jobs[0]['id']]={'status':'running'}
    with pytest.raises(HTTPException,match='already being processed'):step(s.board(USER),cid,jobs[0]['id'])
    with pytest.raises(HTTPException,match='already being processed'):start(cid)
    assert not registrar.calls
    plan['current_started_at']=(datetime.now(timezone.utc)-timedelta(seconds=b.RUNNING_SECONDS+1)).isoformat()
    record=step(s.board(USER),cid,jobs[0]['id'])
    assert batch(record,cid)['results'][jobs[0]['id']]['status']=='verified'


def test_legacy_preview_age_does_not_expire_a_waiting_batch(batch_case,board_db):
    cid,jobs,foreign,registrar=batch_case;record=start(cid)
    batch(board_db[store.workspace(USER)],cid)['created_at']=(datetime.now(timezone.utc)-timedelta(days=1)).isoformat()
    record=step(s.board(USER),cid,jobs[0]['id'])
    assert batch(record,cid)['results'][jobs[0]['id']]['status']=='verified'
    assert next(j for j in record['state']['setups'] if j['id']==jobs[0]['id'])['dns_setup_preview']['created_at']>batch(record,cid)['created_at']


def test_start_and_step_require_admin_confirmation_current_batch_and_csrf(client,batch_case,monkeypatch):
    cid,jobs,foreign,registrar=batch_case;record=s.board(USER)
    with pytest.raises(HTTPException):b.start(OTHER,cid,record['revision'],True,'https://clientwebsite.com')
    with pytest.raises(HTTPException):b.start(USER,cid,record['revision'],False,'https://clientwebsite.com')
    user={**USER};monkeypatch.setattr(auth,'get_session_user',lambda request:user)
    page=client.get('/gmail-set');csrf=re.search('data-csrf="([^"]+)"',page.text)[1]
    url=f'/api/gmail-setup/clients/{cid}/dns-setup/start'
    body={'revision':record['revision'],'confirmed_purchased':True,'destination':'https://clientwebsite.com'}
    assert client.post(url,json=body).status_code==403
    client.headers['X-Domain-CSRF']=csrf
    assert client.post(url,json={**body,'confirmed_purchased':False}).status_code==422
    user['role']='sdr';assert client.post(url,json=body).status_code==403
    user['role']='admin';record=client.post(url,json=body).json()
    endpoint=f'/api/gmail-setup/clients/{cid}/dns-setup/{batch(record,cid)["id"]}/{jobs[0]["id"]}'
    assert client.post(endpoint,json={'revision':record['revision']}).json()['state']['clients'][0]['dns_setup_batch']['results'][jobs[0]['id']]['status']=='verified'
    old=record;record=start(cid)
    with pytest.raises(HTTPException):b.step(USER,cid,batch(old,cid)['id'],jobs[1]['id'],record['revision'])
