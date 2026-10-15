"""Shared setup journey, real device accounting and retryable sheet publishing."""
import copy
from datetime import datetime, timedelta, timezone
from uuid import uuid4
import pytest
import requests
from fastapi import HTTPException
from app import auth
from app.utils.gmail_setup import service as s, store, sheets
from app.utils import supabase as sb
from conftest import FakeResponse

USER={'email':'dylan@unstuck-agency.com','role':'admin','tools':['gmail_set']}
OTHER={'email':'matthew@unstuck-agency.com','role':'sdr','tools':['gmail_set']}


@pytest.fixture
def board_db(monkeypatch):
    rows={}
    monkeypatch.setattr(store,'team_directory',lambda user:[{'email':'dylan@unstuck-agency.com','name':'Dylan'},{'email':'matthew@unstuck-agency.com','name':'Matthew'},{'email':'robyn@unstuck-agency.com','name':'Robyn'}])
    def load(user):return copy.deepcopy(rows.get(store.workspace(user)))
    def save(user,state,revision):
        key=store.workspace(user);old=rows.get(key)
        if (old or {'revision':0})['revision']!=revision:raise HTTPException(409,'Stale revision')
        record={'state':copy.deepcopy(state),'revision':revision+1,'updated_at':s.now(),'workspace_id':key}
        rows[key]=record
        return copy.deepcopy(record)
    monkeypatch.setattr(store,'load',load);monkeypatch.setattr(store,'save',save)
    return rows


def change(action,data,user=USER,revision=None):
    record=s.board(user)
    return s.mutate(user,record['revision'] if revision is None else revision,action,data)


def people():return {m['name']:m['id'] for m in s.board(USER)['state']['members']}


def devices(owner,phones=1,pcs=0,status='Unused',reason=None):
    return change('devices_add',{'member_id':owner,'phones':phones,'pcs':pcs,'status':status,'block_reason':reason})


def create_setup(owner,device_id=None,domain='cedarclients.com',email='me@cedarclients.com',section='client'):
    record=change('setup_save',{'account_manager_id':people()['Robyn'],'operator_id':owner,'client_name':'Cedar client','domain':domain,'email':email,'section':section,'device_id':device_id,'order':1})
    return record,record['state']['setups'][-1]['id']


def test_shared_board_roster_scope_and_no_fake_stock(board_db):
    initial=s.board(USER)
    assert len(initial['state']['members'])==3 and initial['state']['devices']==[]
    assert s.summary(initial['state'])['shortfall']==30
    change('reserve_set',{'target':20})
    assert not s.result(initial)['state'].get('sync_lock')
    ids=people();devices(ids['Dylan'],phones=10,pcs=10)
    assert len(s.board(OTHER)['state']['devices'])==20
    assert s.summary(s.board(OTHER)['state'])['low_stock'] is False
    stranger={'email':'another@different-workspace.com'}
    assert s.board(stranger)['state']['devices']==[]
    with pytest.raises(HTTPException) as exc:devices(ids['Dylan'])
    assert exc.value.status_code==400
    with pytest.raises(HTTPException) as exc:change('member_save',{'name':'Other','email':''},revision=0)
    assert exc.value.status_code==409


@pytest.mark.parametrize('reason',['QR code','Payment'])
def test_block_reasons_excluded_red_state_and_reserve(board_db,reason):
    owner=people()['Dylan'];record=devices(owner,phones=10,pcs=10)
    d=record['state']['devices'][0]
    record=change('device_update',{'id':d['id'],'status':'Blocked','block_reason':reason})
    assert record['state']['devices'][0]['block_reason']==reason
    assert s.summary(record['state'])['unused']==19 and s.summary(record['state'])['shortfall']==11
    with pytest.raises(HTTPException):create_setup(owner,d['id'])
    with pytest.raises(HTTPException):change('device_update',{'id':d['id'],'status':'Unused'})


def test_purchase_consumes_fresh_stock_once_and_preserves_device_history(board_db):
    ids=people();record=devices(ids['Dylan'],phones=2);d=record['state']['devices'][0]
    record,jid=create_setup(ids['Dylan'],d['id'])
    assert s.summary(record['state'])['unused']==2 # allocation is not purchase
    record=change('setup_stage',{'id':jid,'stage':'Purchase made'})
    assert s.summary(record['state'])['unused']==1
    original=record['state']['setups'][0]['purchase_device']
    record=change('device_update',{'id':d['id'],'member_id':ids['Matthew']})
    assert record['state']['setups'][0]['purchase_device']==original
    with pytest.raises(HTTPException):change('device_update',{'id':d['id'],'status':'Unused'})
    reset=change('setup_stage',{'id':jid,'stage':'Not started'})
    assert reset['state']['setups'][0]['stage']=='Not started'
    assert reset['state']['devices'][0]['use_count']==1
    # Reusing an unblocked used device does not pretend another new device was consumed.
    record,jid2=create_setup(ids['Matthew'],d['id'],'otherclient.com','me@otherclient.com')
    record=change('setup_stage',{'id':jid2,'stage':'Purchase made'})
    assert s.summary(record['state'])['unused']==1
    assert record['state']['setups'][1]['purchase_device']['member']=='Matthew'


def test_pending_device_reservation_and_block_after_allocation(board_db):
    owner=people()['Dylan'];d=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,d['id'])
    create_setup(owner,d['id'],'otherclient.com','me@otherclient.com')
    with pytest.raises(HTTPException):create_setup(owner,d['id'],'thirdclient.com','me@thirdclient.com')
    change('device_update',{'id':d['id'],'status':'Blocked','block_reason':'Payment'})
    with pytest.raises(HTTPException):change('setup_stage',{'id':jid,'stage':'Purchase made'})
    assert s.board(USER)['state']['setups'][0]['stage']=='Not started'


def test_remove_member_and_device_preserves_physical_state(board_db):
    ids=people();record=devices(ids['Dylan'],phones=1,pcs=1,status='Blocked',reason='QR code')
    d=record['state']['devices'][0]
    record=change('member_remove',{'id':ids['Dylan']})
    assert all(d['member_id'] is None and d['status']=='Blocked' for d in record['state']['devices'])
    assert s.summary(record['state'])['unused']==0
    record=change('device_remove',{'id':d['id']})
    assert s.summary(record['state'])['devices']==1
    assert record['state']['devices'][0]['archived'] is True


def test_required_sending_email_root_domain_and_open_member_removal(board_db):
    owner=people()['Dylan'];d=devices(owner)['state']['devices'][0]
    with pytest.raises(HTTPException):create_setup(owner,d['id'],email='me@another.com')
    with pytest.raises(HTTPException):create_setup(owner,d['id'],domain='https://cedarclients.com')
    record,jid=create_setup(owner,d['id'],email='')
    with pytest.raises(HTTPException):change('member_remove',{'id':owner})
    change('setup_stage',{'id':jid,'stage':'Purchase made'})
    with pytest.raises(HTTPException):change('setup_stage',{'id':jid,'stage':'Setup complete'})


class SheetSession:
    def __init__(self,section='client'):
        self.gid,self.title=sheets.TABS[section]
        self.headers=['Email','Order']+(['Warming up'] if section=='bizdev' else [])+['Live','At risk','Needs deleting','Being replaced','Biz Dev lead' if section=='bizdev' else 'Client']
        self.rows=[self.headers, ['old@existing.com',1]+([False] if section=='bizdev' else [])+[True,False,False,False,'Existing']]
        self.markers=[];self.calls=[];self.inserts=0;self.lose_response=False
    def range_values(self,r):
        import re
        cells=r.rsplit('!',1)[1]
        match=re.fullmatch(r'([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?',cells)
        first,row,last,end=match.groups();index=lambda letters:sum((ord(c)-64)*26**i for i,c in enumerate(letters[::-1]))-1
        start=index(first);finish=index(last or first);row=int(row)-1;end=int(end or row+1)
        return [r[start:finish+1] for r in self.rows[row:end]]
    def get(self,url,params=None,**kw):
        self.calls.append(('get',url,params))
        if url==sheets.BASE:return FakeResponse(200,{'sheets':[{'properties':{'sheetId':self.gid,'title':self.title,'gridProperties':{'rowCount':100,'columnCount':30}}}]})
        if url.endswith('/values:batchGet'):
            return FakeResponse(200,{'valueRanges':[{'values':self.range_values(v)} for k,v in params if k=='ranges']})
        if '/values/' in url:return FakeResponse(200,{'values':self.range_values(url.split('/values/')[1])})
        raise AssertionError(url)
    def post(self,url,json,**kw):
        self.calls.append(('post',url,json))
        if url.endswith('/developerMetadata:search'):
            value=json['dataFilters'][0]['developerMetadataLookup']['metadataValue']
            return FakeResponse(200,{'matchedDeveloperMetadata':[{'developerMetadata':m} for m in self.markers if m['metadataValue']==value]})
        assert url==sheets.BASE+':batchUpdate'
        for request in json['requests']:
            if 'insertDimension' in request:
                idx=request['insertDimension']['range']['startIndex'];self.rows.insert(idx,['']*len(self.headers));self.inserts+=1
                for m in self.markers:
                    loc=m['location']['dimensionRange']
                    if loc['startIndex']>=idx:loc.update(startIndex=loc['startIndex']+1,endIndex=loc['endIndex']+1)
            if 'updateCells' in request:
                item=request['updateCells'];r=item['range'];val=item['rows'][0]['values'][0]['userEnteredValue']
                self.rows[r['startRowIndex']][r['startColumnIndex']]=next(iter(val.values()))
            if 'createDeveloperMetadata' in request:self.markers.append(copy.deepcopy(request['createDeveloperMetadata']['developerMetadata']))
        if self.lose_response:
            self.lose_response=False;raise requests.Timeout()
        return FakeResponse(200,{})


@pytest.fixture
def sheet_session(monkeypatch):
    monkeypatch.setattr(sheets.gs._read_limiter,'acquire',lambda:None)
    monkeypatch.setattr(sheets.gs._write_limiter,'acquire',lambda:None)
    def make(section='client'):
        session=SheetSession(section);monkeypatch.setattr(sheets.gs,'_authed_session',lambda:session);return session
    return make


@pytest.mark.parametrize('section',['client','bizdev'])
def test_complete_publishes_typed_live_to_correct_tab_and_retries_without_duplicate(board_db,sheet_session,section):
    session=sheet_session(section);owner=people()['Dylan'];d=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,d['id'],section=section)
    change('setup_stage',{'id':jid,'stage':'Purchase made'})
    change('setup_stage',{'id':jid,'stage':'Setup complete'})
    session.lose_response=True # Google commits, but response is lost before DB ack
    first=s.sync_completed(USER,jid)
    assert first['state']['setups'][0]['sync']['status']=='error'
    assert session.inserts==1
    second=s.sync_completed(OTHER,jid)
    job=second['state']['setups'][0]
    assert job['sync']['status']=='synced' and session.inserts==1
    assert session.rows[2][session.headers.index('Live')] is True
    if section=='bizdev':assert session.rows[2][session.headers.index('Warming up')] is False
    assert session.rows[2][-1]==('Dylan' if section=='bizdev' else 'Cedar client')
    assert '#gid='+str(session.gid)+'&range=A3' in job['sync']['url']
    assert s.sync_completed(USER,jid)['revision']==second['revision']
    assert second['state']['sync_lock'] is None
    deleted=change('setup_remove',{'id':jid})
    assert deleted['state']['setups'][0]['archived'] is True
    assert deleted['state']['devices'][0]['use_count']==1
    # New-row writes contain typed literals, not formula interpretation or secrets.
    body=[data for method,url,data in session.calls if method=='post' and url.endswith(':batchUpdate')][0]
    assert not any('formulaValue' in str(r) for r in body['requests'])
    assert len(session.rows)==3


def test_existing_email_conflicts_are_row_specific_and_no_write(sheet_session):
    session=sheet_session();job={'id':str(uuid4()),'section':'client','email':'old@existing.com','order':1}
    with pytest.raises(HTTPException) as exc:sheets.publish(job,'Different client')
    assert 'row 2' in exc.value.detail and 'owner/client differs' in exc.value.detail
    assert session.inserts==0
    session.rows[1][-1]='Cedar client';session.rows[1][session.headers.index('At risk')]=True
    with pytest.raises(HTTPException) as exc:sheets.publish(job,'Cedar client')
    assert 'row 2' in exc.value.detail and 'At Risk' in exc.value.detail


def test_sheet_lease_prevents_overlapping_work_and_expires_safely(board_db,monkeypatch):
    owner=people()['Dylan'];d=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,d['id']);change('setup_stage',{'id':jid,'stage':'Purchase made'});change('setup_stage',{'id':jid,'stage':'Setup complete'})
    def publish(job,owner):
        with pytest.raises(HTTPException) as exc:change('member_save',{'name':'New','email':''},user=OTHER)
        assert exc.value.status_code==409
        with pytest.raises(HTTPException):s.sync_completed(OTHER,jid)
        return {'tab':'Client','row':3,'url':'https://docs.google.com/example'}
    monkeypatch.setattr(sheets,'publish',publish)
    assert s.sync_completed(USER,jid)['state']['setups'][0]['sync']['status']=='synced'
    record=s.board(USER);record['state']['sync_lock']={'token':'expired','job_id':jid,'at':(datetime.now(timezone.utc)-timedelta(seconds=181)).isoformat()};store.save(USER,record['state'],record['revision'])
    assert change('member_save',{'name':'New','email':''})['state']['sync_lock'] is None


def test_tracker_routes_auth_csrf_schema_and_shared_visibility(client,board_db,monkeypatch):
    monkeypatch.setattr(auth,'get_session_user',lambda request:None)
    assert client.get('/api/gmail-setup').status_code==401
    user={**USER,'tools':[]};monkeypatch.setattr(auth,'get_session_user',lambda request:user)
    assert client.get('/gmail-set').status_code==403
    user['tools']=['gmail_set']
    page=client.get('/gmail-set');assert 'GPT domain clustering plugin in Notion' in page.text
    assert 'dc-app' not in page.text and 'gs-app' in page.text
    assert page.headers['cache-control']=='no-store'
    from app.routers.gmail_setup import ASSET_VERSION
    assert f'/static/gmail-setup.js?v={ASSET_VERSION}' in page.text
    assert client.get('/static/gmail-setup.js?v=old').headers['cache-control']=='no-store'
    assert client.post('/api/gmail-setup/change',json={'revision':0,'action':'member_remove','data':{'id':'x'}}).status_code==403
    import re
    client.headers['X-Domain-CSRF']=re.search('data-csrf="([^"]+)"',page.text)[1]
    initial=client.get('/api/gmail-setup').json();owner=initial['state']['members'][0]['id']
    r=client.post('/api/gmail-setup/change',json={'revision':0,'action':'devices_add','data':{'member_id':owner,'phones':10,'pcs':10}})
    assert r.status_code==200 and r.json()['summary']['unused']==20
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'device_update','data':{'id':r.json()['state']['devices'][0]['id'],'status':None}}).status_code==400
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'devices_add','data':{'member_id':owner,'phones':11}}).status_code==400
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'setup_save','data':{'workspace_id':'injected'}}).status_code==400
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'reserve_set','data':{'target':0}}).status_code==400
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'reserve_set','data':{'target':5001}}).status_code==400
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'reserve_set','data':{'target':35}}).status_code==200
    user['email']=OTHER['email'];assert client.get('/api/gmail-setup/summary').json()['unused']==20
    assert client.get('/api/gmail-setup/summary').json()['target']==35
    assert 'sync_lock' not in client.get('/api/gmail-setup').json()['state']


def test_store_filters_workspace_and_uses_revision_lease(monkeypatch):
    calls=[];monkeypatch.setattr(sb,'SUPABASE_URL','https://example.supabase.co');monkeypatch.setattr(sb,'SUPABASE_SERVICE_ROLE_KEY','server-key')
    def request(method,url,**kw):calls.append((method,url,kw));return FakeResponse(200,[{'state':{},'revision':3}])
    monkeypatch.setattr(store.requests,'request',request)
    store.load(USER);store.save(USER,{},2)
    assert calls[0][2]['params']['workspace_id']=='eq.'+store.workspace(USER)
    assert calls[1][2]['params']['revision']=='eq.2'
    assert calls[1][2]['json']['revision']==3
    assert 'owner_id' not in calls[1][2]['params'] # explicitly shared within workspace


def test_missing_header_and_write_denial_point_to_exact_tab_and_row(board_db,sheet_session):
    session=sheet_session('bizdev');job={'id':str(uuid4()),'section':'bizdev','email':'me@cedarclients.com','order':1}
    session.headers[3]='Not Live'
    with pytest.raises(HTTPException) as exc:sheets.publish(job,'Dylan')
    assert 'row 1' in exc.value.detail and 'range=A1' in exc.value.detail
    session.headers[3]='Live'
    original=session.post
    def denied(url,json,**kwargs):
        if url==sheets.BASE+':batchUpdate':return FakeResponse(403,{'error':'protected'})
        return original(url,json,**kwargs)
    session.post=denied
    owner=people()['Dylan'];d=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,d['id'],section='bizdev')
    change('setup_stage',{'id':jid,'stage':'Purchase made'});change('setup_stage',{'id':jid,'stage':'Setup complete'})
    record=s.sync_completed(USER,jid)
    sync=record['state']['setups'][0]['sync']
    assert sync['status']=='error' and 'row 3' in sync['message']
    assert sync['url'].endswith('#gid=838806676&range=A3')
    assert 'protected' in sync['message']
    assert session.inserts==0


def test_client_creation_grouping_and_existing_board_upgrade(board_db):
    owner=people()['Dylan']
    client=change('client_save',{'name':'  Cedar   Client ','section':'client','account_manager_id':people()['Robyn'],'operator_id':owner})['state']['clients'][0]
    assert client['name']=='Cedar Client'
    with pytest.raises(HTTPException):change('client_save',{'name':'cedar client','section':'client','account_manager_id':people()['Robyn']})
    record,jid=create_setup(owner,domain='firstcedar.com',email='me@firstcedar.com')
    record,jid2=create_setup(owner,domain='secondcedar.com',email='me@secondcedar.com')
    assert len(record['state']['clients'])==1
    assert {j['client_id'] for j in record['state']['setups']}=={client['id']}
    # A pre-update flat board groups safely on read without publishing or
    # changing records, purchase snapshots, revisions or device counts.
    legacy=copy.deepcopy(record);legacy['state'].pop('clients')
    for job in legacy['state']['setups']:job.pop('client_id')
    board_db[store.workspace(USER)]=legacy
    first=s.board(USER);second=s.board(USER)
    assert first['revision']==record['revision']
    assert len(first['state']['clients'])==1
    assert first['state']['clients']==second['state']['clients']
    assert all(j['client_id']==first['state']['clients'][0]['id'] for j in first['state']['setups'])
    assert all(j['stage']=='Not started' and not j['purchase_device'] for j in first['state']['setups'])


def test_explicit_client_is_authoritative_and_direct_complete_is_atomic(board_db):
    owner=people()['Dylan'];d=devices(owner)['state']['devices'][0]
    client=change('client_save',{'name':'Cedar Client','section':'bizdev','account_manager_id':people()['Robyn'],'operator_id':owner})['state']['clients'][0]
    data={'client_id':client['id'],'account_manager_id':people()['Robyn'],'operator_id':owner,
          'client_name':'Wrong client','section':'client','domain':'cedarclients.com','email':'',
          'device_id':d['id'],'order':1,'stage':'Setup complete'}
    before=s.board(USER)
    with pytest.raises(HTTPException):change('setup_save',data)
    assert s.board(USER)==before # failed completion neither saves nor consumes device
    record=change('setup_save',{**data,'email':'me@cedarclients.com'})
    job=record['state']['setups'][0]
    assert job['client_name']=='Cedar Client' and job['section']=='bizdev'
    assert job['client_id']==client['id'] and job['stage']=='Setup complete'
    assert job['purchase_device']['id']==d['id'] and job['purchase_at'] and job['completed_at']
    assert s.summary(record['state'])['unused']==0
    reset=change('setup_stage',{'id':job['id'],'stage':'Not started'})
    assert reset['state']['setups'][0]['stage']=='Not started'
    assert reset['state']['setups'][0]['completed_at']==job['completed_at']


@pytest.mark.parametrize('section',['client','bizdev'])
def test_new_sheet_row_is_immediately_below_last_entry_with_live_checkbox(sheet_session,section):
    session=sheet_session(section)
    n=len(session.headers)
    session.rows.extend([['']*n,['']*n,['lastdomain.com',10]+[False]*(n-3)+['Last client'],['']*n,['']*n])
    before=copy.deepcopy(session.rows)
    job={'id':str(uuid4()),'section':section,'email':'me@newclient.com','order':1}
    result=sheets.publish(job,'New client')
    assert result['row']==6 # last actual domain is row 5, despite blank/formatted tail
    assert session.rows[:5]==before[:5]
    assert session.rows[5][0]=='me@newclient.com'
    assert session.rows[5][session.headers.index('Live')] is True
    batch=next(data for method,url,data in session.calls if method=='post' and url==sheets.BASE+':batchUpdate')
    checkbox=next(r['setDataValidation'] for r in batch['requests'] if 'setDataValidation' in r and r['setDataValidation']['range']['startColumnIndex']==session.headers.index('Live'))
    assert checkbox['rule']['condition']['type']=='BOOLEAN'


@pytest.mark.parametrize('email,domain',[('me@cedarclients.com','cedarclients.com'),(' Person@team.example.co.uk ','example.co.uk')])
def test_email_only_form_derives_domain_and_assigns_client(board_db,email,domain):
    owner=people()['Dylan']
    client=change('client_save',{'name':'Cedar','section':'client','account_manager_id':people()['Robyn']})['state']['clients'][0]
    record=change('setup_save',{'client_id':client['id'],'account_manager_id':people()['Robyn'],'operator_id':owner,'email':email,'section':'client','order':1,'stage':'Not started'})
    job=record['state']['setups'][0]
    assert job['domain']==domain and job['email']==email.strip().lower()
    assert job['client_id']==client['id'] and job['client_name']=='Cedar'


@pytest.mark.parametrize('stage',[*s.STAGES, *s.FAILURES])
def test_zero_inventory_never_blocks_setup_creation_or_stage(board_db,stage):
    owner=people()['Dylan']
    record=change('setup_save',{'account_manager_id':people()['Robyn'],'operator_id':owner,
        'client_name':'No inventory','email':'me@waitingforphones.com','section':'client','stage':stage})
    job=record['state']['setups'][0]
    assert job['stage']==stage and job['device_id'] is None and job['purchase_device'] is None
    assert record['state']['devices']==[] and s.summary(record['state'])['unused']==0
    assert sum(s.summary(record['state'])[key] for key in ('not_started','purchased','complete'))==1


@pytest.mark.parametrize('failure,reason',s.FAILURES.items())
@pytest.mark.parametrize('purchased',[False,True])
def test_failure_blocks_real_device_and_retry_retains_failed_asset(board_db,failure,reason,purchased):
    owner=people()['Dylan'];stock=devices(owner,phones=2)['state']['devices'];first,replacement=stock
    _,jid=create_setup(owner,first['id'])
    if purchased:change('setup_stage',{'id':jid,'stage':'Purchase made'})
    record=change('setup_stage',{'id':jid,'stage':failure})
    job=record['state']['setups'][0];d=record['state']['devices'][0]
    assert job['stage']==failure and d['status']=='Blocked' and d['block_reason']==reason
    assert s.summary(record['state'])['unused']==1
    assert job['device_failures'][0]['device']['asset_tag']==first['asset_tag']
    assert job['sync']['status']=='not_ready' # failures never publish a Live row
    with pytest.raises(HTTPException):change('setup_stage',{'id':jid,'stage':'Purchase made'})
    with pytest.raises(HTTPException):change('device_update',{'id':first['id'],'status':'Unused'})
    record=change('setup_save',{'id':jid,'account_manager_id':people()['Robyn'],'operator_id':owner,
        'client_name':'Cedar client','email':job['email'],'section':'client','device_id':replacement['id'],'stage':'Purchase made'})
    job=record['state']['setups'][0]
    assert job['purchase_device']['id']==replacement['id']
    assert job['device_failures'][0]['device']['id']==first['id']
    assert s.summary(record['state'])['unused']==0
    assert record['state']['devices'][0]['status']=='Blocked'


def test_add_missing_device_after_sheet_sync_keeps_live_row_and_purchase_time(board_db,sheet_session):
    session=sheet_session();owner=people()['Dylan']
    _,jid=create_setup(owner)
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    purchase_at=record['state']['setups'][0]['purchase_at']
    record=s.sync_completed(USER,jid);sync=record['state']['setups'][0]['sync']
    d=devices(owner)['state']['devices'][0]
    with pytest.raises(HTTPException):change('setup_device',{'id':jid,'operator_id':people()['Matthew'],'device_id':d['id']})
    record=change('setup_device',{'id':jid,'operator_id':owner,'device_id':d['id']})
    job=record['state']['setups'][0]
    assert job['sync']==sync and job['purchase_at']==purchase_at and job['purchase_device']['id']==d['id']
    assert s.summary(record['state'])['unused']==0 and session.inserts==1
    with pytest.raises(HTTPException):change('setup_device',{'id':jid,'operator_id':owner,'device_id':d['id']})
    with pytest.raises(HTTPException):change('setup_stage',{'id':jid,'stage':'Failed at QR'})


def test_later_device_cannot_take_reserved_or_blocked_stock(board_db):
    owner=people()['Dylan'];d=devices(owner)['state']['devices'][0]
    create_setup(owner,d['id'])
    create_setup(owner,d['id'],'secondreserved.com','me@secondreserved.com')
    _,jid=create_setup(owner,domain='laterdevice.com',email='me@laterdevice.com')
    with pytest.raises(HTTPException):change('setup_device',{'id':jid,'operator_id':owner,'device_id':d['id']})
    change('device_update',{'id':d['id'],'status':'Blocked','block_reason':'Payment'})
    with pytest.raises(HTTPException):change('setup_device',{'id':jid,'operator_id':owner,'device_id':d['id']})


def test_removed_client_manager_does_not_block_planning_without_devices(board_db):
    ids=people()
    client=change('client_save',{'name':'Alab Group','section':'client',
        'account_manager_id':ids['Matthew'],'operator_id':ids['Dylan']})['state']['clients'][0]
    change('member_remove',{'id':ids['Matthew']})
    data={'client_id':client['id'],'account_manager_id':ids['Matthew'],'operator_id':ids['Dylan'],
        'email':'ashley@alabgrouppr.com','section':'client','stage':'Not started'}
    record=change('setup_save',data)
    job=record['state']['setups'][0]
    assert job['device_id'] is None and job['stage']=='Not started'
    assert record['state']['devices']==[] and job['account_manager_id']==ids['Matthew']
    # Removal is still enforced for setup operators and unrelated assignments.
    with pytest.raises(HTTPException):change('setup_save',{**data,'id':job['id'],'operator_id':ids['Matthew']})
    with pytest.raises(HTTPException):change('setup_save',{**data,'client_id':None,'email':'me@unrelatedclient.com','client_name':'Unrelated'})


def test_phone_supports_two_distinct_setups_and_never_counts_completion_twice(board_db):
    ids=people();phone=devices(ids['Dylan'])['state']['devices'][0]
    assert phone['use_count']==0
    _,first=create_setup(ids['Dylan'],phone['id'])
    record=change('setup_stage',{'id':first,'stage':'Purchase made'})
    assert record['state']['devices'][0]['use_count']==1
    record=change('setup_stage',{'id':first,'stage':'Setup complete'})
    assert record['state']['devices'][0]['use_count']==1
    # Moving the same physical phone does not reset its history/capacity.
    change('device_update',{'id':phone['id'],'member_id':ids['Matthew']})
    _,second=create_setup(ids['Matthew'],phone['id'],'secondphoneuse.com','me@secondphoneuse.com')
    record=change('setup_stage',{'id':second,'stage':'Purchase made'})
    assert record['state']['devices'][0]['use_count']==2
    assert s.summary(record['state'])['exhausted_phones']==1
    with pytest.raises(HTTPException):create_setup(ids['Matthew'],phone['id'],'thirdphoneuse.com','me@thirdphoneuse.com')
    with pytest.raises(HTTPException):change('device_update',{'id':phone['id'],'status':'Unused'})
    # Existing second setup can still complete; a third cannot be allocated.
    record=change('setup_stage',{'id':second,'stage':'Setup complete'})
    assert record['state']['devices'][0]['use_count']==2
    _,third=create_setup(ids['Matthew'],domain='thirdphoneuse.com',email='me@thirdphoneuse.com')
    with pytest.raises(HTTPException):change('setup_device',{'id':third,'operator_id':ids['Matthew'],'device_id':phone['id']})


def test_two_phone_slots_can_be_reserved_and_used_without_overbooking(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,first=create_setup(owner,phone['id'])
    _,second=create_setup(owner,phone['id'],'secondreserved.com','me@secondreserved.com')
    with pytest.raises(HTTPException):create_setup(owner,phone['id'],'thirdreserved.com','me@thirdreserved.com')
    change('setup_stage',{'id':first,'stage':'Purchase made'})
    with pytest.raises(HTTPException):create_setup(owner,phone['id'],'thirdreserved.com','me@thirdreserved.com')
    record=change('setup_stage',{'id':second,'stage':'Purchase made'})
    assert record['state']['devices'][0]['use_count']==2


def test_pc_allows_one_setup_and_keeps_capacity_after_transfer_and_reopening(board_db):
    ids=people();pc=devices(ids['Dylan'],phones=0,pcs=1)['state']['devices'][0]
    _,first=create_setup(ids['Dylan'],pc['id'])
    # A planned setup reserves the PC's only slot without consuming new stock.
    assert s.summary(s.board(USER)['state'])['unused']==1
    with pytest.raises(HTTPException):create_setup(ids['Dylan'],pc['id'],'secondpc.com','me@secondpc.com')
    change('setup_stage',{'id':first,'stage':'Purchase made'})
    record=change('setup_stage',{'id':first,'stage':'Setup complete'})
    assert record['state']['devices'][0]['use_count']==1
    assert not s.allocation_available(record['state'],record['state']['devices'][0])
    assert s.summary(record['state'])['unused']==0
    change('setup_stage',{'id':first,'stage':'Not started'})
    record=change('setup_stage',{'id':first,'stage':'Setup complete'})
    assert record['state']['devices'][0]['use_count']==1
    change('device_update',{'id':pc['id'],'member_id':ids['Matthew']})
    with pytest.raises(HTTPException):create_setup(ids['Matthew'],pc['id'],'secondpc.com','me@secondpc.com')
    _,second=create_setup(ids['Matthew'],domain='secondpc.com',email='me@secondpc.com')
    with pytest.raises(HTTPException):change('setup_device',{'id':second,'operator_id':ids['Matthew'],'device_id':pc['id']})
    with pytest.raises(HTTPException):change('setup_assign',{'id':second,'device_id':pc['id']})


@pytest.mark.parametrize('failure',list(s.FAILURES))
def test_failed_pc_reset_restores_its_single_slot(board_db,failure):
    owner=people()['Dylan'];pc=devices(owner,phones=0,pcs=1)['state']['devices'][0]
    _,first=create_setup(owner,pc['id'])
    change('setup_stage',{'id':first,'stage':failure})
    reset=change('setup_stage',{'id':first,'stage':'Not started'})
    d=reset['state']['devices'][0]
    assert d['status']=='Unused' and d['use_count']==0
    assert s.allocation_available(reset['state'],d)
    _,second=create_setup(owner,pc['id'],'retrypc.com','me@retrypc.com')
    record=change('setup_stage',{'id':second,'stage':'Setup complete'})
    assert record['state']['devices'][0]['use_count']==1
    with pytest.raises(HTTPException):create_setup(owner,pc['id'],'thirdpc.com','me@thirdpc.com')


def test_legacy_used_pc_and_overbooked_allocation_cannot_consume_another_use(board_db):
    owner=people()['Dylan'];pc=devices(owner,phones=0,pcs=1,status='Used')['state']['devices'][0]
    _,jid=create_setup(owner)
    stored=board_db[store.workspace(USER)]['state']
    stored['devices'][0].pop('use_count')
    # An old pending allocation must not bypass the new capacity at purchase.
    stored['setups'][0]['device_id']=pc['id']
    before=s.board(USER)
    assert before['state']['devices'][0]['use_count']==1
    assert not s.allocation_available(before['state'],before['state']['devices'][0])
    with pytest.raises(HTTPException,match='one setup'):
        change('setup_stage',{'id':jid,'stage':'Purchase made'})
    assert s.board(USER)==before


def test_legacy_phone_usage_recovered_from_history_without_reset(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    for i in range(2):
        _,jid=create_setup(owner,phone['id'],f'legacyphone{i}.com',f'me@legacyphone{i}.com')
        change('setup_stage',{'id':jid,'stage':'Setup complete'})
    stored=board_db[store.workspace(USER)]
    stored['state']['devices'][0].pop('use_count')
    stored['state']['devices'][0].pop('setup_ids')
    revision=stored['revision']
    migrated=s.board(USER)
    assert migrated['revision']==revision
    assert migrated['state']['devices'][0]['use_count']==2
    assert len(migrated['state']['devices'][0]['setup_ids'])==2
    assert s.summary(migrated['state'])['unused']==0
    # Reading old records does not itself modify the persisted workspace.
    assert 'use_count' not in stored['state']['devices'][0]


def test_bulk_planning_atomic_line_errors_duplicates_and_shared_manager(board_db):
    ids=people()
    client=change('client_save',{'name':'Bulk client','section':'client','account_manager_id':ids['Robyn']})['state']['clients'][0]
    before=s.board(USER)
    with pytest.raises(HTTPException) as exc:
        change('setup_bulk',{'client_id':client['id'],'domains':'validone.com\nnot-a-domain\nvalidtwo.com'})
    assert 'Line 2' in exc.value.detail and s.board(USER)==before
    pasted='\n'.join(f'bulk{i}.com' for i in range(12))+'\nme@bulkmixed.com, bulk0.com'
    record=change('setup_bulk',{'client_id':client['id'],'domains':pasted})
    assert record['import_summary']=={'added':13,'skipped':1}
    jobs=record['state']['setups']
    assert all(j['account_manager_id']==ids['Robyn'] and j['operator_id'] is None and j['device_id'] is None and j['order'] is None for j in jobs)
    assert jobs[-1]['email']=='me@bulkmixed.com'
    again=change('setup_bulk',{'client_id':client['id'],'domains':pasted,'operator_id':ids['Dylan']})
    assert again['import_summary']=={'added':0,'skipped':14} and len(again['state']['setups'])==13


@pytest.mark.parametrize('section',['client','bizdev'])
def test_completion_order_follows_completion_not_creation_and_survives_retry_delete(board_db,sheet_session,section):
    session=sheet_session(section);ids=people()
    client=change('client_save',{'name':'Order client','section':section,'account_manager_id':ids['Robyn']})['state']['clients'][0]
    record=change('setup_bulk',{'client_id':client['id'],'domains':'firstcreated.com\nsecondcreated.com\nthirdcreated.com','operator_id':ids['Dylan']})
    first,second,third=record['state']['setups']
    # Completion without a sending email must not consume an order number.
    with pytest.raises(HTTPException):change('setup_stage',{'id':second['id'],'stage':'Setup complete'})
    for job in (second,first,third):
        change('setup_assign',{'id':job['id'],'email':'me@'+job['domain']})
    change('setup_stage',{'id':second['id'],'stage':'Setup complete'})
    session.lose_response=True
    failed=s.sync_completed(USER,second['id'])
    assert next(j for j in failed['state']['setups'] if j['id']==second['id'])['order']==1
    change('setup_stage',{'id':first['id'],'stage':'Setup complete'})
    s.sync_completed(USER,first['id']);s.sync_completed(USER,second['id'])
    change('setup_remove',{'id':second['id']})
    record=change('setup_stage',{'id':third['id'],'stage':'Setup complete'})
    byid={j['id']:j for j in record['state']['setups']}
    assert [byid[j['id']]['order'] for j in (second,first,third)]==[1,2,3]
    stable=change('setup_stage',{'id':third['id'],'stage':'Setup complete'})
    assert stable['state']['clients'][0]['next_order']==4
    s.sync_completed(USER,third['id'])
    assert [r[1] for r in session.rows[2:]]==[1,2,3]
    assert session.inserts==3
    other=change('client_save',{'name':'Other client','section':section,'account_manager_id':ids['Robyn']})['state']['clients'][-1]
    record=change('setup_save',{'client_id':other['id'],'account_manager_id':ids['Robyn'],'operator_id':ids['Dylan'],
        'email':'me@anotherorder.com','section':section,'stage':'Setup complete','order':999})
    assert record['state']['setups'][-1]['order']==1 # input order is ignored


def test_rename_and_delete_client_preserve_sheet_owner_sequence_and_device_history(board_db,sheet_session):
    session=sheet_session();ids=people();phone=devices(ids['Dylan'])['state']['devices'][0]
    record,jid=create_setup(ids['Dylan'],phone['id']);client=record['state']['clients'][0]
    change('setup_stage',{'id':jid,'stage':'Setup complete'})
    session.lose_response=True;s.sync_completed(USER,jid)
    record=change('client_save',{**client,'name':'Renamed client','account_manager_id':ids['Robyn']})
    assert record['state']['setups'][0]['client_name']=='Renamed client'
    assert record['state']['setups'][0]['sheet_owner']=='Cedar client'
    synced=s.sync_completed(USER,jid)
    assert synced['state']['setups'][0]['sync']['status']=='synced' and session.inserts==1
    assert session.rows[2][-1]=='Cedar client'
    record=change('client_remove',{'id':client['id']})
    assert record['state']['clients'][0]['archived'] and record['state']['setups'][0]['archived']
    assert record['state']['devices'][0]['use_count']==1
    assert len(session.rows)==3 and session.rows[2][session.headers.index('Live')] is True
    assert s.summary(record['state'])['complete']==0
    with pytest.raises(HTTPException):change('setup_bulk',{'client_id':client['id'],'domains':'notallowed.com'})
    # Reusing a deleted client's name creates a new identity; deleted work stays deleted.
    replacement=change('client_save',{'name':'Renamed client','section':'client','account_manager_id':ids['Robyn']})
    assert replacement['state']['clients'][-1]['id']!=client['id']


def test_inline_assignment_tracks_owner_and_device_without_resetting_usage(board_db):
    ids=people();phone=devices(ids['Dylan'])['state']['devices'][0]
    client=change('client_save',{'name':'Inline','section':'client','account_manager_id':ids['Robyn']})['state']['clients'][0]
    record=change('setup_bulk',{'client_id':client['id'],'domains':'inlineone.com\ninlinetwo.com'})
    first,second=record['state']['setups']
    record=change('setup_assign',{'id':first['id'],'operator_id':ids['Dylan'],'email':'me@inlineone.com'})
    change('setup_assign',{'id':first['id'],'device_id':phone['id']})
    record=change('setup_assign',{'id':first['id'],'operator_id':ids['Matthew']})
    assert record['state']['setups'][0]['device_id'] is None # releases incompatible pending allocation
    assert record['state']['devices'][0]['use_count']==0
    change('setup_assign',{'id':second['id'],'operator_id':ids['Dylan'],'device_id':phone['id'],'email':'me@inlinetwo.com'})
    change('setup_stage',{'id':second['id'],'stage':'Failed at QR'})
    replacement=devices(ids['Dylan'])['state']['devices'][-1]
    record=change('setup_assign',{'id':second['id'],'device_id':replacement['id']})
    assert record['state']['setups'][1]['stage']=='Not started'
    assert record['state']['setups'][1]['device_failures'][0]['device']['id']==phone['id']
    assert record['state']['devices'][0]['status']=='Blocked'
    record=change('setup_stage',{'id':second['id'],'stage':'Setup complete'})
    assert record['state']['devices'][-1]['use_count']==1
    with pytest.raises(HTTPException):change('setup_assign',{'id':second['id'],'operator_id':ids['Matthew']})


@pytest.mark.parametrize('failure',list(s.FAILURES))
@pytest.mark.parametrize('purchased',[False,True])
def test_failed_setup_can_reset_and_use_a_replacement_without_erasing_phone_history(board_db,failure,purchased):
    owner=people()['Dylan'];phone,replacement=devices(owner,phones=2)['state']['devices']
    _,jid=create_setup(owner,phone['id'])
    if purchased:change('setup_stage',{'id':jid,'stage':'Purchase made'})
    failed=change('setup_stage',{'id':jid,'stage':failure})['state']['setups'][0]
    record=change('setup_stage',{'id':jid,'stage':'Not started'})
    job=record['state']['setups'][0]
    assert job['stage']=='Not started' and job['device_id'] is None
    assert [f['device'] for f in job['device_failures']]==[f['device'] for f in failed['device_failures']]
    assert all(f.get('reverted_at') for f in job['device_failures'])
    assert record['state']['devices'][0]['status']=='Unused' and record['state']['devices'][0]['use_count']==0
    record=change('setup_device',{'id':jid,'operator_id':owner,'device_id':replacement['id']})
    assert record['state']['devices'][1]['use_count']==0 # allocation for a retry is not a new purchase
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    assert record['state']['setups'][0]['order']==1 and record['state']['devices'][1]['use_count']==1
    assert record['state']['devices'][0]['status']=='Unused'
    if purchased:assert record['state']['setups'][0]['purchase_device_history'][0]['id']==phone['id']


@pytest.mark.parametrize('synced',[False,True])
def test_completed_setup_can_reopen_without_renumbering_or_duplicate_sheet_writes(board_db,sheet_session,synced):
    session=sheet_session();owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,phone['id'])
    change('setup_stage',{'id':jid,'stage':'Setup complete'})
    session.lose_response=not synced
    original=s.sync_completed(USER,jid)['state']['setups'][0]
    assert session.inserts==1
    record=change('setup_stage',{'id':jid,'stage':'Not started'})
    job=record['state']['setups'][0]
    assert job['stage']=='Not started' and job['order']==original['order']==1
    assert job['completed_at']==original['completed_at'] and job['sync']==original['sync']
    assert record['state']['devices'][0]['use_count']==1
    # Reopening an already-counted use must not reserve a second phone slot.
    _,second=create_setup(owner,phone['id'],'secondafterreset.com','me@secondafterreset.com')
    change('setup_stage',{'id':second,'stage':'Setup complete'})
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    assert record['state']['devices'][0]['use_count']==2
    assert record['state']['clients'][0]['next_order']==3
    assert record['state']['setups'][0]['order']==1
    result=s.sync_completed(USER,jid)
    assert result['state']['setups'][0]['sync']['status']=='synced' and session.inserts==1
    assert session.rows[2][1]==1


@pytest.mark.parametrize('sender',[' Ashley ','ashley@','ashley.smith'])
def test_bulk_sender_name_autofills_bare_domains_and_preserves_explicit_emails(board_db,sender):
    ids=people();client=change('client_save',{'name':'Sender client','section':'client','account_manager_id':ids['Robyn']})['state']['clients'][0]
    record=change('setup_bulk',{'client_id':client['id'],'domains':'senderone.com\ncustom@sendertwo.com\nsenderthree.com','sender_name':sender})
    prefix=sender.strip().lower().rstrip('@')
    assert [j['email'] for j in record['state']['setups']]==[prefix+'@senderone.com','custom@sendertwo.com',prefix+'@senderthree.com']
    assert all(j['operator_id'] is None and j['device_id'] is None for j in record['state']['setups'])
    before=s.board(USER)
    with pytest.raises(HTTPException) as exc:change('setup_bulk',{'client_id':client['id'],'domains':'newdomain.com','sender_name':'invalid sender'})
    assert 'sender name' in exc.value.detail and s.board(USER)==before


def test_reserve_target_is_shared_editable_and_defaults_to_thirty(board_db):
    assert s.summary(s.board(USER)['state'])['target']==30
    record=change('reserve_set',{'target':40})
    assert s.summary(record['state'])['shortfall']==40
    record=devices(people()['Dylan'],phones=3)
    assert s.summary(record['state'])['shortfall']==37
    assert s.summary(s.board(OTHER)['state'])['target']==40
    assert s.summary(s.board({'email':'person@separate-workspace.com'})['state'])['target']==30
    record=change('reserve_set',{'target':2})
    assert s.summary(record['state'])['shortfall']==0 and not s.summary(record['state'])['low_stock']


def test_reset_failure_restores_remaining_phone_capacity_and_can_reuse_same_phone(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,first=create_setup(owner,phone['id']);change('setup_stage',{'id':first,'stage':'Setup complete'})
    _,second=create_setup(owner,phone['id'],'failedsecond.com','me@failedsecond.com')
    change('setup_stage',{'id':second,'stage':'Failed at payment'})
    record=change('setup_stage',{'id':second,'stage':'Not started'})
    device=record['state']['devices'][0]
    assert device['status']=='Used' and device['use_count']==1 and device['setup_ids']==[first]
    assert record['state']['setups'][1]['device_usage_resets'][phone['id']]
    for _ in range(3):assert s.board(USER)['state']['devices'][0]['use_count']==1
    change('setup_device',{'id':second,'operator_id':owner,'device_id':phone['id']})
    record=change('setup_stage',{'id':second,'stage':'Setup complete'})
    assert record['state']['devices'][0]['use_count']==2
    assert phone['id'] not in record['state']['setups'][1].get('device_usage_resets',{})


def test_legacy_reset_repairs_stale_block_without_erasing_audit_history(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,phone['id']);change('setup_stage',{'id':jid,'stage':'Failed at payment'})
    stored=board_db[store.workspace(USER)]
    stored['state']['setups'][0].update(stage='Not started',device_id=None)
    original_revision=stored['revision']
    migrated=s.board(USER)
    assert migrated['revision']==original_revision
    assert migrated['state']['devices'][0]['status']=='Unused' and migrated['state']['devices'][0]['use_count']==0
    assert migrated['state']['setups'][0]['device_failures'][0]['device']['id']==phone['id']
    assert migrated['state']['setups'][0]['device_failures'][0]['reverted_at']
    assert stored['state']['devices'][0]['status']=='Blocked' # read upgrade is not an unleased write
    saved=change('reserve_set',{'target':30})
    assert saved['state']['devices'][0]['status']=='Unused'
    for _ in range(3):assert s.board(USER)['state']['devices'][0]['use_count']==0


def test_reset_one_failure_keeps_other_domain_block_and_manual_blocks(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,first=create_setup(owner,phone['id'])
    _,second=create_setup(owner,phone['id'],'otherfailure.com','me@otherfailure.com')
    change('setup_stage',{'id':first,'stage':'Failed at QR'})
    change('setup_stage',{'id':second,'stage':'Failed at payment'})
    record=change('setup_stage',{'id':first,'stage':'Not started'})
    assert record['state']['devices'][0]['status']=='Blocked' and record['state']['devices'][0]['use_count']==1
    record=change('setup_stage',{'id':second,'stage':'Not started'})
    assert record['state']['devices'][0]['status']=='Unused' and record['state']['devices'][0]['use_count']==0
    change('setup_device',{'id':first,'operator_id':owner,'device_id':phone['id']})
    change('setup_stage',{'id':first,'stage':'Failed at QR'})
    change('device_update',{'id':phone['id'],'status':'Blocked','block_reason':'Other'})
    record=change('setup_stage',{'id':first,'stage':'Not started'})
    assert record['state']['devices'][0]['status']=='Blocked' and record['state']['devices'][0]['manual_block']
    assert record['state']['devices'][0]['block_reason']=='Other'


def test_reset_failed_flag_keeps_previous_completed_use(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,phone['id']);change('setup_stage',{'id':jid,'stage':'Setup complete'})
    change('setup_stage',{'id':jid,'stage':'Not started'})
    change('setup_stage',{'id':jid,'stage':'Failed at QR'})
    record=change('setup_stage',{'id':jid,'stage':'Not started'})
    assert record['state']['devices'][0]['status']=='Used' and record['state']['devices'][0]['use_count']==1
    assert record['state']['setups'][0]['order']==1
