"""DNS cleanup is confirmed, exact-record scoped, and verified without live writes."""
import copy
import re
from datetime import datetime, timedelta, timezone
import pytest
from fastapi import HTTPException
from app import auth
from app.utils.gmail_setup import dns_setup as d, service as s, store
from app.utils.domain_cluster import provisioning as p
from test_gmail_setup import board_db, USER, OTHER, people, change, create_setup
from test_domain_setup import Registrar


def defaults(domain):
    return [dict(id=str(i+1),name=name,type=kind,content=value,ttl='600',prio='0') for i,(name,kind,value) in enumerate([
        (domain,'MX','fwd1.porkbun.com'),(domain,'MX','fwd2.porkbun.com'),
        (domain,'TXT','v=spf1 include:_spf.porkbun.com ~all'),
        ('_acme-challenge.'+domain,'TXT','challenge-one'),('_acme-challenge.'+domain,'TXT','challenge-two')])]


def retained(domain):
    return [dict(id=str(i+50),name=name,type=kind,content=value,ttl='600',prio='0') for i,(name,kind,value) in enumerate([
        (domain,'MX','smtp.google.com'),(domain,'TXT','v=spf1 include:_spf.google.com ~all'),
        (domain,'TXT','v=spf1 include:_spf.porkbun.com include:_spf.google.com ~all'),
        (domain,'TXT','google-site-verification=keep'),('_dmarc.'+domain,'TXT',p.dmarc_value(domain,{'rua':'me@'+domain,'ruf':'me@'+domain})),
        ('google._domainkey.'+domain,'TXT','v=DKIM1; p=keep'),(domain,'A','1.2.3.4'),
        ('www.'+domain,'CNAME','example.org'),('other.'+domain,'MX','fwd1.porkbun.com'),
        ('_acme-challenge.other.'+domain,'TXT','keep-subdomain-challenge')])]


class SetupRegistrar(Registrar):
    def call(self,method,path,data=None,idempotency=None):
        if path.startswith('/dns/create/'):
            self.calls.append((method,path,copy.deepcopy(data),idempotency))
            domain=path.split('/')[-1];identifier=str(1000+len(self.calls))
            self.records.setdefault(domain,[]).append({**data,'name':data['name']+'.'+domain,'id':identifier,'ttl':str(data['ttl']),'prio':'0'})
            return {'status':'SUCCESS','id':identifier}
        if path.startswith('/dns/delete/'):
            self.calls.append((method,path,copy.deepcopy(data),idempotency))
            assert method=='post'
            _,_,_,domain,identifier=path.split('/')
            self.records[domain]=[r for r in self.records[domain] if str(r['id'])!=identifier]
            return {'status':'SUCCESS'}
        return super().call(method,path,data,idempotency)


@pytest.fixture
def registrar(monkeypatch):
    fake=SetupRegistrar();fake.records['cedarclients.com']=defaults('cedarclients.com')+retained('cedarclients.com')
    fake.records['neighbor.com']=defaults('neighbor.com')
    monkeypatch.setattr(p,'Porkbun',lambda:fake)
    return fake


def deletes(fake):return [c for c in fake.calls if c[1].startswith('/dns/delete/')]


def preview(jid):return d.preview(USER,jid,s.board(USER)['revision'])


def apply(record,jid):
    job=next(j for j in record['state']['setups'] if j['id']==jid)
    return d.apply(USER,jid,job['dns_setup_preview']['id'],record['revision'],True,'https://clientwebsite.com')


def test_preview_and_confirm_remove_only_the_five_reviewed_records(board_db,registrar):
    owner=people()['Dylan'];_,jid=create_setup(owner)
    before=s.summary(s.board(USER)['state'])
    record=preview(jid);approved=record['state']['setups'][0]['dns_setup_preview']
    assert len(approved['records'])==5 and not deletes(registrar)
    original_neighbor=copy.deepcopy(registrar.records['neighbor.com'])
    result=apply(record,jid)
    assert len(deletes(registrar))==5
    assert all(c[1].startswith('/dns/delete/cedarclients.com/') and c[3] for c in deletes(registrar))
    assert [r for r in registrar.records['cedarclients.com'] if r['type']!='CNAME' or r['name']!='t.cedarclients.com']==retained('cedarclients.com')
    assert registrar.records['neighbor.com']==original_neighbor
    job=result['state']['setups'][0]
    assert job['dns_setup']['status']=='verified' and job['dns_setup']['removed']==5
    assert job['stage']=='Not started' and s.summary(result['state'])==before
    assert result['state']['sync_lock'] is None
    apply(result,jid) # Lost response / replay never deletes additional records.
    assert len(deletes(registrar))==5


@pytest.mark.parametrize('record',[
    {'name':'@','type':'MX','content':'FWD1.PORKBUN.COM.'},
    {'name':'','type':'TXT','content':'"v=spf1  include:_spf.porkbun.com ~all"'},
    {'name':'_acme-challenge','type':'TXT','content':'challenge'},
    {'name':'acme-challenge.cedarclients.com','type':'TXT','content':'challenge'},
])
def test_matching_handles_provider_name_and_spf_formats(record):
    assert d.matches(record,'cedarclients.com')


@pytest.mark.parametrize('change_kind',['changed-content','changed-ttl','added-record'])
def test_changed_records_require_a_fresh_preview_before_any_delete(board_db,registrar,change_kind):
    _,jid=create_setup(people()['Dylan']);record=preview(jid)
    if change_kind=='changed-content':registrar.records['cedarclients.com'][0]['content']='smtp.google.com'
    elif change_kind=='changed-ttl':registrar.records['cedarclients.com'][0]['ttl']='3600'
    else:registrar.records['cedarclients.com'].append({**defaults('cedarclients.com')[0],'id':'888'})
    result=apply(record,jid)
    assert result['state']['setups'][0]['dns_setup']['status']=='needs review'
    assert 'changed after review' in result['state']['setups'][0]['dns_setup']['message']
    assert not deletes(registrar) and result['state']['sync_lock'] is None


def test_partial_lost_response_is_reconciled_from_current_records(board_db,registrar):
    _,jid=create_setup(people()['Dylan']);record=preview(jid)
    original=registrar.call
    def lost_response(method,path,*args,**kwargs):
        result=original(method,path,*args,**kwargs)
        if path.startswith('/dns/delete/') and len(deletes(registrar))==1:
            raise HTTPException(502,'Response lost; review before retrying.')
        return result
    registrar.call=lost_response
    first=apply(record,jid)
    assert first['state']['setups'][0]['dns_setup']['status']=='needs review'
    assert len(deletes(registrar))==1 and first['state']['sync_lock'] is None
    second=apply(first,jid)
    assert second['state']['setups'][0]['dns_setup']['status']=='verified'
    assert len(deletes(registrar))==5
    assert [r for r in registrar.records['cedarclients.com'] if r['type']!='CNAME' or r['name']!='t.cedarclients.com']==retained('cedarclients.com')


def test_unconfirmed_expired_foreign_burned_and_archived_requests_never_delete(board_db,registrar):
    _,jid=create_setup(people()['Dylan']);record=preview(jid)
    identifier=record['state']['setups'][0]['dns_setup_preview']['id']
    with pytest.raises(HTTPException):d.apply(USER,jid,identifier,record['revision'],False,'https://clientwebsite.com')
    with pytest.raises(HTTPException):d.apply(OTHER,jid,identifier,record['revision'],True,'https://clientwebsite.com')
    with pytest.raises(HTTPException):d.preview({'email':'admin@another-company.com','role':'admin'},jid,0)
    stored=board_db[store.workspace(USER)]
    stored['state']['setups'][0]['dns_setup_preview']['created_at']=(datetime.now(timezone.utc)-timedelta(minutes=16)).isoformat()
    with pytest.raises(HTTPException,match='expired'):apply(s.board(USER),jid)
    record=change('setup_stage',{'id':jid,'stage':'Burned'})
    with pytest.raises(HTTPException,match='burned'):preview(jid)
    with pytest.raises(HTTPException,match='burned'):apply(record,jid)
    change('setup_remove',{'id':jid})
    with pytest.raises(HTTPException):preview(jid)
    assert not deletes(registrar)


@pytest.mark.parametrize('problem',['denied','delegation','inactive','malformed','bad-id','too-many'])
def test_preflight_failure_never_approves_deletion(board_db,registrar,problem):
    _,jid=create_setup(people()['Dylan']);original=registrar.call
    if problem=='denied':registrar.check=lambda:{'credentials_valid':True,'read':True,'dns':False}
    elif problem=='malformed':registrar.records['cedarclients.com']='not a list'
    elif problem=='bad-id':registrar.records['cedarclients.com'][0]['id']='../other.com/99'
    elif problem=='too-many':registrar.records['cedarclients.com']=[{**defaults('cedarclients.com')[3],'id':str(i)} for i in range(11)]
    def call(method,path,*args,**kwargs):
        if problem=='delegation' and path.startswith('/domain/getNs/'):
            return {'ns':['ns1.other-provider.com']}
        if problem=='inactive' and path.startswith('/domain/get/'):
            return {'domain':{'domain':'cedarclients.com','status':'PENDING','apiAccess':1}}
        return original(method,path,*args,**kwargs)
    registrar.call=call
    with pytest.raises(HTTPException):preview(jid)
    assert not deletes(registrar)


def test_claim_blocks_other_mutations_and_missing_readback_is_not_success(board_db,registrar):
    owner=people()['Dylan'];_,jid=create_setup(owner);record=preview(jid)
    original=registrar.call
    def call(method,path,*args,**kwargs):
        if path.startswith('/dns/delete/'):
            lock=board_db[store.workspace(USER)]['state']['sync_lock']
            assert lock['kind']=='dns_setup'
            with pytest.raises(HTTPException,match='DNS records are being configured'):
                change('member_save',{'id':owner,'name':'Changed','email':'dylan@unstuck-agency.com'})
            # Provider claims success but does not remove anything.
            registrar.calls.append((method,path,{},kwargs.get('idempotency')))
            return {'status':'SUCCESS'}
        return original(method,path,*args,**kwargs)
    registrar.call=call
    result=apply(record,jid)
    assert result['state']['setups'][0]['dns_setup']['status']=='needs review'
    assert 'remain' in result['state']['setups'][0]['dns_setup']['message']
    assert result['state']['sync_lock'] is None


def test_routes_require_csrf_admin_confirmation_and_server_selected_record_ids(client,board_db,registrar,monkeypatch):
    _,jid=create_setup(people()['Dylan']);user={**USER}
    monkeypatch.setattr(auth,'get_session_user',lambda request:user)
    page=client.get('/gmail-set');csrf=re.search('data-csrf="([^"]+)"',page.text)[1]
    base=f'/api/gmail-setup/setups/{jid}/dns-setup'
    payload={'revision':s.board(USER)['revision']}
    assert client.post(base+'/preview',json=payload).status_code==403
    client.headers['X-Domain-CSRF']=csrf
    user['role']='sdr';assert client.post(base+'/preview',json=payload).status_code==403
    user['role']='admin';record=client.post(base+'/preview',json=payload).json()
    assert not deletes(registrar)
    payload={'revision':record['revision'],'preview_id':record['state']['setups'][0]['dns_setup_preview']['id'],'confirmed_purchased':True,'destination':'https://clientwebsite.com'}
    assert client.post(base+'/apply',json={**payload,'confirmed_purchased':False}).status_code==422
    assert client.post(base+'/apply',json={**payload,'records':[{'id':'50'}]}).status_code==422
    assert client.post(base+'/apply',json=payload).json()['state']['setups'][0]['dns_setup']['status']=='verified'


def test_missing_tracking_and_dmarc_created_with_exact_policy_then_wildcard_forward(board_db,registrar):
    _,jid=create_setup(people()['Dylan'])
    registrar.records['cedarclients.com']=[r for r in registrar.records['cedarclients.com'] if r['name']!='_dmarc.cedarclients.com']
    result=apply(preview(jid),jid)
    creates=[c for c in registrar.calls if c[1].startswith('/dns/create/')]
    assert [c[2]['name'] for c in creates]==['t','_dmarc']
    assert creates[0][2]=={'name':'t','type':'CNAME','content':'open.sleadtrack.com','ttl':600}
    assert creates[1][2]=={'name':'_dmarc','type':'TXT','content':'v=DMARC1; p=none; pct=100; rua=mailto:me@cedarclients.com; ruf=mailto:me@cedarclients.com; fo=1','ttl':600}
    forwards=[c for c in registrar.calls if c[1].startswith('/domain/addUrlForward/')]
    assert len(forwards)==1 and forwards[0][2]=={'subdomain':'','location':'https://clientwebsite.com','type':'temporary','redirectType':'302','includePath':'no','wildcard':'yes'}
    assert max(registrar.calls.index(c) for c in deletes(registrar)) < min(registrar.calls.index(c) for c in creates)
    assert result['state']['setups'][0]['dmarc']['status']=='verified'
    assert result['state']['clients'][0]['website_url']=='https://clientwebsite.com'
    apply(result,jid)
    assert len([c for c in registrar.calls if c[1].startswith('/dns/create/')])==2
    assert len([c for c in registrar.calls if c[1].startswith('/domain/addUrlForward/')])==1


@pytest.mark.parametrize('conflict',['tracking','dmarc','forwarding'])
def test_conflicts_at_any_destination_stop_all_deletion(board_db,registrar,conflict):
    _,jid=create_setup(people()['Dylan'])
    if conflict=='tracking':registrar.records['cedarclients.com'].append({'id':'444','name':'t.cedarclients.com','type':'CNAME','content':'other-tracker.com','ttl':'600'})
    if conflict=='dmarc':
        next(r for r in registrar.records['cedarclients.com'] if r['name']=='_dmarc.cedarclients.com')['content']='v=DMARC1; p=reject;'
    if conflict=='forwarding':registrar.forwards['cedarclients.com']=[{'subdomain':'','location':'https://differentclient.com','type':'temporary','includePath':'no','wildcard':'yes'}]
    if conflict=='forwarding':
        result=apply(preview(jid),jid)
        assert result['state']['setups'][0]['dns_setup']['status']=='needs review'
    else:
        with pytest.raises(HTTPException,match='differs'):preview(jid)
    assert not any(c[0]=='post' for c in registrar.calls)


@pytest.mark.parametrize('url',['','javascript:alert(1)','https://user:pass@client.com','https://cedarclients.com','http://www.cedarclients.com','https://client.com/a b'])
def test_invalid_or_looping_url_stops_before_any_write(board_db,registrar,url):
    _,jid=create_setup(people()['Dylan']);record=preview(jid)
    identifier=record['state']['setups'][0]['dns_setup_preview']['id']
    with pytest.raises(HTTPException):d.apply(USER,jid,identifier,record['revision'],True,url)
    assert not any(c[0]=='post' for c in registrar.calls)


@pytest.mark.parametrize('endpoint',['/dns/create/','/domain/addUrlForward/'])
def test_lost_creation_or_forward_response_reconciles_without_duplicates(board_db,registrar,endpoint):
    _,jid=create_setup(people()['Dylan']);record=preview(jid);original=registrar.call;failed=False
    def call(method,path,*args,**kwargs):
        nonlocal failed
        result=original(method,path,*args,**kwargs)
        if path.startswith(endpoint) and not failed:
            failed=True
            raise HTTPException(502,'Response lost after write.')
        return result
    registrar.call=call
    first=apply(record,jid);assert first['state']['setups'][0]['dns_setup']['status']=='needs review'
    second=apply(first,jid);assert second['state']['setups'][0]['dns_setup']['status']=='verified'
    assert len([r for r in registrar.records['cedarclients.com'] if r['name']=='t.cedarclients.com'])==1
    assert len(registrar.forwards['cedarclients.com'])==1
    assert len(deletes(registrar))==5

