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
    assert s.summary(initial['state'])['shortfall']==20
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
    assert s.summary(record['state'])['unused']==19 and s.summary(record['state'])['shortfall']==1
    with pytest.raises(HTTPException):create_setup(owner,d['id'])
    with pytest.raises(HTTPException):change('device_update',{'id':d['id'],'status':'Unused'})


def test_purchase_consumes_fresh_stock_once_and_preserves_device_history(board_db):
    ids=people();record=devices(ids['Dylan'],phones=2);d=record['state']['devices'][0]
    record,jid=create_setup(ids['Dylan'],d['id'])
    assert s.summary(record['state'])['unused']==2 # allocation is not purchase
    with pytest.raises(HTTPException):change('setup_stage',{'id':jid,'stage':'Setup complete'})
    record=change('setup_stage',{'id':jid,'stage':'Purchase made'})
    assert s.summary(record['state'])['unused']==1
    original=record['state']['setups'][0]['purchase_device']
    record=change('device_update',{'id':d['id'],'member_id':ids['Matthew']})
    assert record['state']['setups'][0]['purchase_device']==original
    with pytest.raises(HTTPException):change('device_update',{'id':d['id'],'status':'Unused'})
    with pytest.raises(HTTPException):change('setup_stage',{'id':jid,'stage':'Not started'})
    # Reusing an unblocked used device does not pretend another new device was consumed.
    record,jid2=create_setup(ids['Matthew'],d['id'],'otherclient.com','me@otherclient.com')
    record=change('setup_stage',{'id':jid2,'stage':'Purchase made'})
    assert s.summary(record['state'])['unused']==1
    assert record['state']['setups'][1]['purchase_device']['member']=='Matthew'


def test_pending_device_reservation_and_block_after_allocation(board_db):
    owner=people()['Dylan'];d=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,d['id'])
    with pytest.raises(HTTPException):create_setup(owner,d['id'],'otherclient.com','me@otherclient.com')
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
    with pytest.raises(HTTPException):change('setup_remove',{'id':jid})
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
    assert client.post('/api/gmail-setup/change',json={'revision':0,'action':'member_remove','data':{'id':'x'}}).status_code==403
    import re
    client.headers['X-Domain-CSRF']=re.search('data-csrf="([^"]+)"',page.text)[1]
    initial=client.get('/api/gmail-setup').json();owner=initial['state']['members'][0]['id']
    r=client.post('/api/gmail-setup/change',json={'revision':0,'action':'devices_add','data':{'member_id':owner,'phones':10,'pcs':10}})
    assert r.status_code==200 and r.json()['summary']['unused']==20
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'device_update','data':{'id':r.json()['state']['devices'][0]['id'],'status':None}}).status_code==400
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'devices_add','data':{'member_id':owner,'phones':11}}).status_code==400
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'setup_save','data':{'workspace_id':'injected'}}).status_code==400
    user['email']=OTHER['email'];assert client.get('/api/gmail-setup/summary').json()['unused']==20
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
