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


def master_client(name='Grouping client',count=7):
    owner=people()['Dylan']
    client=change('client_save',{'name':name,'account_manager_id':owner,'domain_target':count})['state']['clients'][-1]
    stem=name.lower().replace(' ','')
    record=change('setup_bulk',{'client_id':client['id'],'domains':'\n'.join(f'{stem}{i}.com' for i in range(count)),
        'sender_name':'me','operator_id':owner})
    return client,[job for job in record['state']['setups'] if job['client_id']==client['id']]


def mark_master(job_id,master_id):return change('setup_master',{'id':job_id,'master_domain_id':master_id})


def delete_gmail(job_id):return change('setup_stage',{'id':job_id,'stage':s.GMAIL_DELETED,'confirmed_auto_renew_stopped':True})


@pytest.fixture
def password_key(monkeypatch):
    monkeypatch.setenv('GMAIL_SETUP_PASSWORD_KEY','test-only-client-password-key')


def test_client_password_is_encrypted_at_rest_replaceable_removable_and_read_only_on_reveal(board_db,password_key):
    import json
    client,jobs=master_client('Password client',count=1)
    original=s.board(USER);secret='Synthetic-client-credential-123'
    record=change('client_password',{'id':client['id'],'password':secret})
    saved=record['state']['clients'][0]
    assert saved['password_ciphertext']!=secret and saved['password_updated_at']
    assert secret not in json.dumps(board_db[store.workspace(USER)])
    assert all('password' not in key for key in saved if key not in {'password_ciphertext','password_updated_at'})
    public=s.result(record)['state']['clients'][0]
    assert public['has_password'] is True and 'password_ciphertext' not in public and 'password' not in public
    before=s.board(USER)
    assert s.reveal_client_password(OTHER,client['id'])==secret
    assert s.board(USER)==before # reveal never advances revisions or writes audit/state
    replaced=change('client_password',{'id':client['id'],'password':'Replacement-test-credential'})
    assert replaced['state']['clients'][0]['password_ciphertext']!=saved['password_ciphertext']
    assert s.reveal_client_password(USER,client['id'])=='Replacement-test-credential'
    removed=change('client_password',{'id':client['id'],'password':None})
    assert 'password_ciphertext' not in removed['state']['clients'][0]
    assert s.result(removed)['state']['clients'][0]['has_password'] is False
    with pytest.raises(HTTPException) as exc:s.reveal_client_password(USER,client['id'])
    assert exc.value.status_code==404


@pytest.mark.parametrize('secret',['  spaces matter  ','\t\n','A synthetic Unicode credential 🔑','x'*1024])
def test_client_password_preserves_whitespace_unicode_and_maximum_length(board_db,password_key,secret):
    client,jobs=master_client('Exact password',count=1)
    change('client_password',{'id':client['id'],'password':secret})
    assert s.reveal_client_password(USER,client['id'])==secret


def test_client_password_missing_or_changed_key_fails_closed_and_removal_can_recover(board_db,password_key,monkeypatch):
    client,jobs=master_client('Password key',count=1)
    secret='Key-recovery-test-credential'
    change('client_password',{'id':client['id'],'password':secret})
    original=s.board(USER)
    monkeypatch.setenv('GMAIL_SETUP_PASSWORD_KEY','changed-test-only-key')
    with pytest.raises(HTTPException) as exc:s.reveal_client_password(USER,client['id'])
    assert exc.value.status_code==503 and secret not in exc.value.detail
    assert original['state']['clients'][0]['password_ciphertext'] not in exc.value.detail
    monkeypatch.delenv('GMAIL_SETUP_PASSWORD_KEY');monkeypatch.delenv('SESSION_SECRET')
    for operation in (lambda:s.reveal_client_password(USER,client['id']),
                      lambda:change('client_password',{'id':client['id'],'password':'Must not save without a key'})):
        with pytest.raises(HTTPException) as exc:operation()
        assert exc.value.status_code==503 and s.board(USER)==original
    removed=change('client_password',{'id':client['id'],'password':None})
    assert 'password_ciphertext' not in removed['state']['clients'][0]


def test_client_password_key_falls_back_to_session_secret_and_explicit_key_takes_priority(board_db,monkeypatch):
    monkeypatch.delenv('GMAIL_SETUP_PASSWORD_KEY',raising=False)
    monkeypatch.setenv('SESSION_SECRET','synthetic-session-secret-for-password-test')
    client,jobs=master_client('Fallback password key',count=1)
    change('client_password',{'id':client['id'],'password':'Fallback-test-value'})
    assert s.reveal_client_password(USER,client['id'])=='Fallback-test-value'
    monkeypatch.setenv('GMAIL_SETUP_PASSWORD_KEY','explicit-test-key-overrides-session')
    with pytest.raises(HTTPException):s.reveal_client_password(USER,client['id'])
    monkeypatch.delenv('GMAIL_SETUP_PASSWORD_KEY')
    assert s.reveal_client_password(USER,client['id'])=='Fallback-test-value'


def test_client_password_ciphertext_is_bound_to_client_and_workspace(board_db,password_key):
    first,jobs=master_client('First password client',count=1)
    second,others=master_client('Second password client',count=1)
    record=change('client_password',{'id':first['id'],'password':'Bound-test-credential'})
    ciphertext=record['state']['clients'][0]['password_ciphertext']
    record['state']['clients'][1]['password_ciphertext']=ciphertext
    store.save(USER,record['state'],record['revision'])
    with pytest.raises(HTTPException) as exc:s.reveal_client_password(USER,second['id'])
    assert exc.value.status_code==503 and 'Bound-test-credential' not in exc.value.detail
    assert s.reveal_client_password(USER,first['id'])=='Bound-test-credential'
    outside={'email':'person@other-workspace.com','role':'sdr','tools':['gmail_set']}
    with pytest.raises(HTTPException) as exc:s.reveal_client_password(outside,first['id'])
    assert exc.value.status_code==404
    copied=copy.deepcopy(s.board(USER));copied['workspace_id']=store.workspace(outside)
    board_db[store.workspace(outside)]=copied
    with pytest.raises(HTTPException) as exc:s.reveal_client_password(outside,first['id'])
    assert exc.value.status_code==503 and ciphertext not in exc.value.detail


def test_client_password_changes_do_not_touch_setup_device_dns_sheet_history_or_math(board_db,password_key,monkeypatch):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    record,jid=create_setup(owner,phone['id']);client=record['state']['clients'][0]
    mark_master(jid,jid);change('setup_notes',{'id':jid,'notes':'Keep this account history'})
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    record['state']['setups'][0]['dns_setup']={'status':'verified','destination':'https://example.com'}
    record['state']['setups'][0]['sync']={'status':'synced','row':42}
    store.save(USER,record['state'],record['revision']);delete_gmail(jid)
    before=s.board(USER);plan=s.summary(copy.deepcopy(before['state']));progress=s.result(before)['client_progress']
    monkeypatch.setattr(sheets,'publish',lambda *args:pytest.fail('Password changes cannot publish sheet rows'))
    monkeypatch.setattr(s,'set_stage',lambda *args:pytest.fail('Password changes cannot change setup stages'))
    record=change('client_password',{'id':client['id'],'password':'History-preserving-test-value'})
    assert record['state']['setups']==before['state']['setups'] and record['state']['devices']==before['state']['devices']
    assert s.summary(record['state'])==plan and s.result(record)['client_progress']==progress
    private={k:v for k,v in record['state']['clients'][0].items() if k not in {'password_ciphertext','password_updated_at'}}
    assert private==before['state']['clients'][0]


def test_client_password_routes_require_explicit_authenticated_csrf_reveal_and_scrub_all_boards(client,board_db,password_key,monkeypatch):
    import json
    from app.utils.gmail_setup import dns_setup
    group,jobs=master_client('Revealed password',count=1);secret='Explicit-reveal-test-credential'
    change('client_password',{'id':group['id'],'password':secret})
    before=s.board(USER);user=None
    monkeypatch.setattr(auth,'get_session_user',lambda request:user)
    endpoint=f'/api/gmail-setup/clients/{group["id"]}/password/reveal'
    assert client.post(endpoint).status_code==401
    user={**OTHER,'tools':[]};assert client.post(endpoint).status_code==403
    user={**OTHER};assert client.post(endpoint).status_code==403
    assert client.get(endpoint).status_code==405
    client.headers['X-Domain-CSRF']=client.get('/api/gmail-setup/session').json()['csrf']
    response=client.post(endpoint)
    assert response.status_code==200 and response.json()=={'password':secret}
    assert response.headers['cache-control']=='no-store' and s.board(USER)==before
    public=client.get('/api/gmail-setup')
    assert public.json()['state']['clients'][0]['has_password'] is True
    assert 'password_ciphertext' not in public.text and secret not in public.text
    changed=client.post('/api/gmail-setup/change',json={'revision':before['revision'],'action':'setup_notes',
        'data':{'id':jobs[0]['id'],'notes':'An unrelated safe update'}})
    assert changed.status_code==200 and secret not in changed.text and 'password_ciphertext' not in changed.text
    user={**USER}
    private=s.board(USER)
    monkeypatch.setattr(dns_setup,'preview',lambda *args:copy.deepcopy(private))
    dns=client.post(f'/api/gmail-setup/setups/{jobs[0]["id"]}/dns-setup/preview',json={'revision':private['revision']})
    assert dns.status_code==200 and secret not in dns.text and 'password_ciphertext' not in dns.text
    assert s.board(USER)['state']['clients'][0]['password_ciphertext']==before['state']['clients'][0]['password_ciphertext']
    change('client_remove',{'id':group['id']})
    assert client.post(endpoint).status_code==404


@pytest.mark.parametrize('data_patch',[{}, {'password':''},{'password':'x'*1025},{'password':123},{'password':False},
    {'password':'test','password_ciphertext':'injected'}])
def test_client_password_schema_rejects_missing_empty_oversize_nonstring_or_extra_fields(client,board_db,password_key,monkeypatch,data_patch):
    monkeypatch.setattr(auth,'get_session_user',lambda request:USER)
    client.headers['X-Domain-CSRF']=client.get('/api/gmail-setup/session').json()['csrf']
    group,jobs=master_client('Validated password',count=1);before=s.board(USER)
    response=client.post('/api/gmail-setup/change',json={'revision':before['revision'],'action':'client_password',
        'data':{'id':group['id'],**data_patch}})
    assert response.status_code==400 and s.board(USER)==before


def test_public_client_scrubbing_removes_legacy_password_secrets_without_mutating_storage(board_db,password_key):
    group,jobs=master_client('Scrubbed password',count=1)
    record=change('client_password',{'id':group['id'],'password':'Private-test-secret'})
    record['state']['clients'][0].update(password='legacy-plaintext',password_plaintext='old-plaintext',password_key='old-key')
    private=store.save(USER,record['state'],record['revision'])
    public=s.result(private)['state']['clients'][0]
    assert public['has_password'] is True
    assert not set(public).intersection({'password_ciphertext','password','password_plaintext','password_key'})
    assert s.board(USER)==private


@pytest.mark.parametrize('starting_stage',['Not started','Started','Purchase made','Setup complete','Failed at QR','Failed at payment','Burned'])
def test_gmail_deleted_requires_confirmation_and_preserves_history_without_archiving(board_db,starting_stage):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,phone['id'])
    change('setup_notes',{'id':jid,'notes':'Keep account history and notes'})
    mark_master(jid,jid)
    record=change('setup_stage',{'id':jid,'stage':starting_stage})
    if starting_stage=='Setup complete':
        record['state']['setups'][0]['sync']={'status':'synced','row':42}
        record=store.save(USER,record['state'],record['revision'])
    before=s.board(USER)
    with pytest.raises(HTTPException,match='auto-renew'):
        change('setup_stage',{'id':jid,'stage':s.GMAIL_DELETED})
    assert s.board(USER)==before
    record=delete_gmail(jid);job=record['state']['setups'][0]
    assert job['stage']==s.GMAIL_DELETED and job['archived'] is False
    assert job['gmail_deleted_from']==starting_stage and job['gmail_deleted_at']==job['auto_renew_stopped_at']
    for field in ('notes','master_domain_id','master_color','purchase_at','purchase_device','purchase_device_history',
                  'device_failures','completed_at','order','sheet_owner','sync','device_id','operator_id'):
        assert job.get(field)==before['state']['setups'][0].get(field)
    assert record['state']['devices']==before['state']['devices'] and record['state']['clients']==before['state']['clients']
    plan=s.summary(record['state'])
    assert plan['gmail_deleted']==1 and plan['burned']==plan['complete']==plan['actual_todo']==0
    assert plan['todo']==plan['planned_todo']==1 and plan['target']==11
    repeated=change('setup_stage',{'id':jid,'stage':s.GMAIL_DELETED})
    for field in ('gmail_deleted_from','gmail_deleted_at','auto_renew_stopped_at'):
        assert repeated['state']['setups'][0][field]==job[field]
    with pytest.raises(HTTPException,match='Not started'):
        change('setup_stage',{'id':jid,'stage':'Setup complete'})
    change('setup_notes',{'id':jid,'notes':'Still editable after deletion'})
    assert s.board(USER)['state']['setups'][0]['notes']=='Still editable after deletion'


@pytest.mark.parametrize('action',['setup_stage','setup_save'])
def test_gmail_deleted_confirmation_is_enforced_by_api_for_stage_and_details(client,board_db,monkeypatch,action):
    monkeypatch.setattr(auth,'get_session_user',lambda request:USER)
    client.headers['X-Domain-CSRF']=client.get('/api/gmail-setup/session').json()['csrf']
    record,jid=create_setup(people()['Dylan']);job=record['state']['setups'][0]
    data={'id':jid,'stage':s.GMAIL_DELETED}
    if action=='setup_save':
        data.update(client_id=job['client_id'],account_manager_id=job['account_manager_id'],operator_id=job['operator_id'],
            email=job['email'],section=job['section'],notes='')
    payload={'revision':record['revision'],'action':action,'data':data}
    before=s.board(USER)
    for confirmation in ({},{'confirmed_auto_renew_stopped':False},{'confirmed_auto_renew_stopped':'true'},
                         {'confirmed_auto_renew_stopped':1},{'confirmed_auto_renew_stopped':None}):
        rejected=client.post('/api/gmail-setup/change',json={**payload,'data':{**data,**confirmation}})
        assert rejected.status_code==400 and s.board(USER)==before
    saved=client.post('/api/gmail-setup/change',json={**payload,'data':{**data,'confirmed_auto_renew_stopped':True}})
    assert saved.status_code==200 and saved.json()['state']['setups'][0]['stage']==s.GMAIL_DELETED
    # Selecting the unchanged stage again needs no new attestation.
    repeated=client.post('/api/gmail-setup/change',json={**payload,'revision':saved.json()['revision']})
    assert repeated.status_code==200
    assert repeated.json()['state']['setups'][0]['auto_renew_stopped_at']==saved.json()['state']['setups'][0]['auto_renew_stopped_at']


def test_gmail_deleted_guard_runs_before_a_new_details_save_creates_or_changes_anything(board_db):
    before=s.board(USER)
    with pytest.raises(HTTPException,match='auto-renew'):
        change('setup_save',{'account_manager_id':people()['Dylan'],'operator_id':people()['Dylan'],
            'client_name':'Never created','email':'me@nevercreated.com','section':'client','stage':s.GMAIL_DELETED})
    assert s.board(USER)==before and board_db=={}


def test_gmail_deleted_replacement_demand_bulk_completion_and_target_floor_are_consistent(board_db):
    owner=people()['Dylan'];client,jobs=master_client('Deleted replacement target',count=4)
    change('setup_stage',{'id':jobs[0]['id'],'stage':'Setup complete'})
    delete_gmail(jobs[1]['id'])
    change('setup_stage',{'id':jobs[2]['id'],'stage':'Burned'})
    record=s.result(s.board(USER))
    assert record['client_progress'][client['id']]=={'target':4,'complete':1,'actual_todo':1,'todo':3,'remaining':3,'replacements_needed':2}
    completed=change('client_complete',{'id':client['id']})
    assert [job['stage'] for job in completed['state']['setups']]==['Setup complete',s.GMAIL_DELETED,'Burned','Setup complete']
    assert s.summary(completed['state'])['todo']==2 and s.summary(completed['state'])['gmail_deleted']==1
    assert completed['state']['clients'][0]['domain_target']==4
    record=change('client_save',{'id':client['id'],'name':client['name'],'account_manager_id':owner,'domain_target':2})
    assert s.summary(record['state'])['todo']==0 # inactive rows do not raise the minimum target
    record=change('setup_stage',{'id':jobs[1]['id'],'stage':'Not started'})
    assert record['state']['clients'][0]['domain_target']==3 and s.summary(record['state'])['todo']==1


def test_gmail_deleted_reopening_keeps_successful_device_use_and_old_sheet_metadata(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,phone['id'])
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    record['state']['setups'][0]['sync']={'status':'synced','row':42};store.save(USER,record['state'],record['revision'])
    deleted=delete_gmail(jid)['state']['setups'][0]
    reopened=change('setup_stage',{'id':jid,'stage':'Not started'})
    assert reopened['state']['devices'][0]['use_count']==1
    assert reopened['state']['setups'][0]['sync']==deleted['sync']
    assert reopened['state']['setups'][0]['order']==deleted['order']
    change('setup_assign',{'id':jid,'operator_id':people()['Matthew']})
    assert s.board(USER)['state']['setups'][0]['device_id'] is None
    with pytest.raises(HTTPException,match='auto-renew'):
        change('setup_stage',{'id':jid,'stage':s.GMAIL_DELETED})


@pytest.mark.parametrize('failure',list(s.FAILURES))
def test_failure_reset_survives_burned_deleted_wrappers_and_repeated_inactive_transitions(board_db,failure):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,phone['id'])
    change('setup_stage',{'id':jid,'stage':failure})
    change('setup_stage',{'id':jid,'stage':'Burned'})
    delete_gmail(jid)
    change('setup_stage',{'id':jid,'stage':'Burned'})
    deleted=delete_gmail(jid)
    assert deleted['state']['setups'][0]['gmail_deleted_from']=='Burned'
    assert s.failure_stage(deleted['state']['setups'][0])==failure
    reset=change('setup_stage',{'id':jid,'stage':'Not started'})
    assert reset['state']['setups'][0]['device_id'] is None
    assert reset['state']['devices'][0]['status']=='Unused' and reset['state']['devices'][0]['use_count']==0
    assert reset['state']['setups'][0]['device_failures'][0]['reverted_at']


def test_other_deleted_failure_still_holds_shared_device_block_until_it_is_reopened(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,first=create_setup(owner,phone['id'])
    _,second=create_setup(owner,phone['id'],'otherdeletedfailure.com','me@otherdeletedfailure.com')
    change('setup_stage',{'id':first,'stage':'Failed at QR'})
    change('setup_stage',{'id':second,'stage':'Failed at payment'})
    change('setup_stage',{'id':first,'stage':'Burned'})
    delete_gmail(first);delete_gmail(second)
    reset=change('setup_stage',{'id':first,'stage':'Not started'})
    assert reset['state']['devices'][0]['status']=='Blocked' and reset['state']['devices'][0]['block_reason']=='Payment'
    assert reset['state']['devices'][0]['use_count']==1
    reset=change('setup_stage',{'id':second,'stage':'Not started'})
    assert reset['state']['devices'][0]['status']=='Unused' and reset['state']['devices'][0]['use_count']==0


def test_deleted_legacy_burned_failure_provenance_and_sync_error_counts_remain_correct(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,phone['id'])
    change('setup_stage',{'id':jid,'stage':'Failed at QR'})
    record=change('setup_stage',{'id':jid,'stage':'Burned'})
    record['state']['setups'][0].pop('inactive_failure_stage')
    record['state']['setups'][0]['sync']={'status':'error','message':'Old historical failure'}
    store.save(USER,record['state'],record['revision'])
    record=delete_gmail(jid)
    assert s.summary(record['state'])['sync_errors']==0
    reset=change('setup_stage',{'id':jid,'stage':'Not started'})
    assert reset['state']['devices'][0]['use_count']==0 and reset['state']['devices'][0]['status']=='Unused'


def test_master_domains_limit_is_per_client_stable_and_includes_burned_roots(board_db):
    first,jobs=master_client()
    assert all('master_domain_id' not in job and 'master_color' not in job for job in jobs)
    for job in jobs[:5]:mark_master(job['id'],job['id'])
    record=mark_master(jobs[2]['id'],jobs[2]['id']) # existing roots remain valid at the limit
    assert [job['master_color'] for job in record['state']['setups'][:5]]==list(range(5))
    change('setup_stage',{'id':jobs[0]['id'],'stage':'Burned'})
    mark_master(jobs[5]['id'],jobs[1]['id'])
    before=s.board(USER)
    with pytest.raises(HTTPException,match='up to 5') as exc:mark_master(jobs[5]['id'],jobs[5]['id'])
    assert exc.value.status_code==409 and not exc.value.headers
    assert s.board(USER)==before
    second,others=master_client('Independent grouping')
    for job in others[:5]:mark_master(job['id'],job['id'])
    record=s.board(USER)
    assert [job['master_color'] for job in record['state']['setups'] if job['client_id']==second['id'] and job.get('master_domain_id')==job['id']]==list(range(5))
    assert record['state']['setups'][0]['master_color']==0 and record['state']['setups'][0]['stage']=='Burned'


def test_master_domain_reassignment_and_ungrouping_are_reversible_without_nested_groups(board_db):
    client,jobs=master_client(count=5)
    first,second,child,other,unused=jobs
    mark_master(first['id'],first['id']);mark_master(second['id'],second['id'])
    mark_master(child['id'],first['id']);mark_master(other['id'],first['id'])
    record=mark_master(child['id'],second['id'])
    assert record['state']['setups'][2]['master_domain_id']==second['id'] and 'master_color' not in record['state']['setups'][2]
    record=mark_master(first['id'],second['id'])
    assert record['state']['setups'][0]['master_domain_id']==second['id'] and 'master_color' not in record['state']['setups'][0]
    assert record['state']['setups'][3]['master_domain_id'] is None # demoted root's children are detached
    assert record['state']['setups'][2]['master_domain_id']==second['id'] # unrelated association remains
    assert record['state']['setups'][1]['master_color']==1
    record=mark_master(child['id'],child['id'])
    assert record['state']['setups'][2]['master_color']==0 # first free slot, without recolouring root 2
    record=mark_master(second['id'],None)
    assert record['state']['setups'][0]['master_domain_id'] is None
    assert record['state']['setups'][1]['master_domain_id'] is None and 'master_color' not in record['state']['setups'][1]
    assert record['state']['setups'][2]['master_domain_id']==child['id'] and record['state']['setups'][2]['master_color']==0
    record=mark_master(child['id'],None)
    assert record['state']['setups'][2]['master_domain_id'] is None and 'master_color' not in record['state']['setups'][2]


@pytest.mark.parametrize('target_kind',['other_client','child','unmarked','archived','missing'])
def test_master_domain_references_fail_atomically_before_detaching_current_children(board_db,target_kind):
    client,jobs=master_client(count=4)
    root,child,other,plain=jobs
    mark_master(root['id'],root['id']);mark_master(child['id'],root['id'])
    target=plain['id']
    if target_kind=='other_client':
        outside,external=master_client('Other client grouping',count=1)
        target=external[0]['id'];mark_master(target,target)
    elif target_kind=='child':target=child['id']
    elif target_kind=='archived':
        mark_master(other['id'],other['id']);change('setup_remove',{'id':other['id']});target=other['id']
    elif target_kind=='missing':target=str(uuid4())
    before=s.board(USER)
    with pytest.raises(HTTPException):mark_master(root['id'],target)
    assert s.board(USER)==before


def test_deleting_master_detaches_active_children_and_reuses_only_its_colour_slot(board_db):
    client,jobs=master_client(count=8)
    for job in jobs[:5]:mark_master(job['id'],job['id'])
    for job in jobs[5:]:mark_master(job['id'],jobs[2]['id'])
    change('setup_stage',{'id':jobs[6]['id'],'stage':'Burned'})
    change('setup_remove',{'id':jobs[7]['id']})
    record=change('setup_remove',{'id':jobs[2]['id']})
    rows=record['state']['setups']
    assert rows[5]['master_domain_id'] is None and rows[6]['master_domain_id'] is None
    assert rows[6]['stage']=='Burned' and rows[7]['archived'] # archived history need not be rewritten
    assert [rows[i]['master_color'] for i in (0,1,3,4)]==[0,1,3,4]
    record=mark_master(jobs[5]['id'],jobs[5]['id'])
    assert record['state']['setups'][5]['master_color']==2
    change('client_remove',{'id':client['id']})
    before=s.board(USER)
    with pytest.raises(HTTPException):mark_master(jobs[5]['id'],None)
    assert s.board(USER)==before


@pytest.mark.parametrize('stage',['Not started','Started','Purchase made','Setup complete','Failed at QR','Failed at payment','Burned'])
def test_master_labels_never_change_workflow_device_dns_sheet_history_or_capacity(board_db,monkeypatch,stage):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    record,jid=create_setup(owner,phone['id'])
    change('setup_notes',{'id':jid,'notes':'Keep this account note'})
    change('setup_stage',{'id':jid,'stage':stage})
    record=s.board(USER)
    record['state']['setups'][0]['dns_setup']={'status':'verified','destination':'https://example.com'}
    if stage=='Setup complete':record['state']['setups'][0]['sync']={'status':'synced','row':42,'url':'https://example.com/old-sheet-row'}
    store.save(USER,record['state'],record['revision'])
    before=s.board(USER);plan=s.summary(copy.deepcopy(before['state']));progress=s.result(before)['client_progress']
    monkeypatch.setattr(sheets,'publish',lambda *args:pytest.fail('Master labels cannot publish to Google Sheets'))
    monkeypatch.setattr(s,'set_stage',lambda *args:pytest.fail('Master labels cannot change the setup stage'))
    monkeypatch.setattr(s,'grow_client_target',lambda *args:pytest.fail('Master labels cannot change domain demand'))
    marked=mark_master(jid,jid)
    assert marked['state']['setups'][0]['master_color']==0
    cleared=mark_master(jid,None)
    actual={k:v for k,v in cleared['state']['setups'][0].items() if k not in {'master_domain_id','master_color','updated_at'}}
    original={k:v for k,v in before['state']['setups'][0].items() if k!='updated_at'}
    assert actual==original
    assert cleared['state']['devices']==before['state']['devices'] and cleared['state']['clients']==before['state']['clients']
    assert s.summary(cleared['state'])==plan and s.result(cleared)['client_progress']==progress


def test_master_labels_survive_existing_owner_notes_details_and_stage_mutations(board_db):
    client,jobs=master_client(count=2)
    root,child=jobs
    mark_master(root['id'],root['id']);mark_master(child['id'],root['id'])
    change('setup_assign',{'id':child['id'],'operator_id':people()['Matthew']})
    change('setup_notes',{'id':child['id'],'notes':'Still linked'})
    change('setup_save',{'id':child['id'],'client_id':client['id'],'account_manager_id':people()['Dylan'],
        'operator_id':people()['Matthew'],'email':child['email'],'section':'client','notes':'Edited details'})
    change('setup_stage',{'id':child['id'],'stage':'Setup complete'})
    record=change('setup_stage',{'id':root['id'],'stage':'Burned'})
    assert record['state']['setups'][0]['master_domain_id']==root['id'] and record['state']['setups'][0]['master_color']==0
    assert record['state']['setups'][1]['master_domain_id']==root['id'] and 'master_color' not in record['state']['setups'][1]


@pytest.mark.parametrize('moving_root',[False,True])
def test_legacy_details_move_clears_grouping_instead_of_linking_across_clients(board_db,moving_root):
    source,jobs=master_client(count=3)
    root,child,other=jobs
    mark_master(root['id'],root['id']);mark_master(child['id'],root['id'])
    moved=root if moving_root else child
    record=change('setup_save',{'id':moved['id'],'account_manager_id':people()['Dylan'],
        'operator_id':people()['Dylan'],'client_name':'Moved domain client','email':moved['email'],'section':'client'})
    rows={job['id']:job for job in record['state']['setups']}
    assert rows[moved['id']]['client_id']!=source['id']
    assert rows[moved['id']]['master_domain_id'] is None and 'master_color' not in rows[moved['id']]
    if moving_root:
        assert rows[child['id']]['master_domain_id'] is None
    else:
        assert rows[root['id']]['master_domain_id']==root['id'] and rows[root['id']]['master_color']==0
    for job in rows.values():
        if job.get('master_domain_id'):
            master=rows[job['master_domain_id']]
            assert master['client_id']==job['client_id'] and master['master_domain_id']==master['id']


def test_master_domain_routes_require_access_csrf_current_revision_and_server_selected_colour(client,board_db,monkeypatch):
    group,jobs=master_client(count=2);root,child=jobs
    user=None
    monkeypatch.setattr(auth,'get_session_user',lambda request:user)
    payload={'revision':s.board(USER)['revision'],'action':'setup_master','data':{'id':root['id'],'master_domain_id':root['id']}}
    assert client.post('/api/gmail-setup/change',json=payload).status_code==401
    user={**OTHER,'tools':[]}
    assert client.post('/api/gmail-setup/change',json=payload).status_code==403
    user={**OTHER}
    rejected=client.post('/api/gmail-setup/change',json=payload)
    assert rejected.status_code==403 and rejected.headers['X-Gmail-Session']=='refresh'
    client.headers['X-Domain-CSRF']=client.get('/api/gmail-setup/session').json()['csrf']
    saved=client.post('/api/gmail-setup/change',json=payload)
    assert saved.status_code==200 and saved.json()['state']['setups'][0]['master_color']==0
    before=s.board(USER)
    stale=client.post('/api/gmail-setup/change',json={**payload,'data':{'id':root['id'],'master_domain_id':None}})
    assert stale.status_code==409 and stale.headers['X-Gmail-Conflict']=='revision'
    assert s.board(USER)==before
    for data in ({'id':child['id']},{'id':child['id'],'master_domain_id':''},
        {'id':child['id'],'master_domain_id':root['id'],'master_color':4}):
        response=client.post('/api/gmail-setup/change',json={'revision':before['revision'],'action':'setup_master','data':data})
        assert response.status_code==400 and s.board(USER)==before
    linked=client.post('/api/gmail-setup/change',json={'revision':before['revision'],'action':'setup_master',
        'data':{'id':child['id'],'master_domain_id':root['id']}})
    assert linked.status_code==200 and s.board(USER)['state']['setups'][1]['master_domain_id']==root['id']


def test_shared_board_roster_scope_and_no_fake_stock(board_db):
    initial=s.board(USER)
    assert len(initial['state']['members'])==3 and initial['state']['devices']==[]
    assert s.summary(initial['state'])['shortfall']==5
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
    assert s.summary(record['state'])['unused']==19 and s.summary(record['state'])['shortfall']==0
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
    assert client.post('/api/gmail-setup/change',json={'revision':1,'action':'reserve_set','data':{'target':35}}).status_code==409
    user['email']=OTHER['email'];assert client.get('/api/gmail-setup/summary').json()['unused']==20
    assert client.get('/api/gmail-setup/summary').json()['target']==10
    assert 'sync_lock' not in client.get('/api/gmail-setup').json()['state']


def test_session_status_requires_tool_access_and_returns_stable_current_token_and_build_for_staff(client,board_db,monkeypatch):
    from app.routers.gmail_setup import APP_VERSION
    user=None
    monkeypatch.setattr(auth,'get_session_user',lambda request:user)
    response=client.get('/api/gmail-setup/session')
    assert response.status_code==401 and 'X-Gmail-Session' not in response.headers
    user={**OTHER,'tools':[]}
    response=client.get('/api/gmail-setup/session')
    assert response.status_code==403 and 'X-Gmail-Session' not in response.headers
    user={**OTHER,'role':'sdr'}
    response=client.get('/api/gmail-setup/session')
    assert response.status_code==200 and response.headers['cache-control']=='no-store'
    token=response.json()['csrf']
    assert len(token)>=32 and response.json()['app_version']==APP_VERSION
    assert client.get('/api/gmail-setup/session').json()==response.json()
    assert board_db=={} # fetching a session never creates/loads a tracker or sheet row
    import re
    page=client.get('/gmail-set')
    assert re.search('data-csrf="([^"]+)"',page.text)[1]==token
    assert re.search('data-app-version="([^"]+)"',page.text)[1]==APP_VERSION
    assert client.get('/api/gmail-setup/session').json()['csrf']==token


def test_stale_session_token_fails_before_mutation_then_current_token_retry_succeeds(client,board_db,monkeypatch):
    monkeypatch.setattr(auth,'get_session_user',lambda request:OTHER)
    monkeypatch.setattr(sheets,'publish',lambda *args:pytest.fail('Session recovery must not write Google Sheets'))
    record,jid=create_setup(people()['Dylan'])
    before=s.board(OTHER)
    payload={'revision':record['revision'],'action':'setup_notes','data':{'id':jid,'notes':'Saved by Matthew after session check'}}
    client.headers['X-Domain-CSRF']='outdated-tab-token'
    rejected=client.post('/api/gmail-setup/change',json=payload)
    assert rejected.status_code==403 and rejected.headers['X-Gmail-Session']=='refresh'
    assert rejected.json()['detail']=='Session verification failed. Refresh the page and retry.'
    assert s.board(OTHER)==before
    recovered=client.get('/api/gmail-setup/session')
    assert recovered.status_code==200 and s.board(OTHER)==before
    client.headers['X-Domain-CSRF']=recovered.json()['csrf']
    saved=client.post('/api/gmail-setup/change',json=payload)
    assert saved.status_code==200 and saved.json()['state']['setups'][0]['notes']==payload['data']['notes']
    assert s.board(USER)['state']['setups'][0]['notes']==payload['data']['notes'] # same workspace for both people
    assert client.get('/api/gmail-setup/session').json()['csrf']==recovered.json()['csrf']


def test_release_version_changes_page_and_session_without_changing_asset_cache_version(client,monkeypatch):
    from app.routers import gmail_setup
    monkeypatch.setattr(auth,'get_session_user',lambda request:OTHER)
    monkeypatch.setattr(gmail_setup,'APP_VERSION','synthetic-backend-only-release')
    response=client.get('/api/gmail-setup/session')
    assert response.json()['app_version']=='synthetic-backend-only-release'
    page=client.get('/gmail-set')
    assert 'data-app-version="synthetic-backend-only-release"' in page.text
    assert f'/static/gmail-setup.js?v={gmail_setup.ASSET_VERSION}' in page.text
    assert f'/static/gmail-setup.css?v={gmail_setup.ASSET_VERSION}' in page.text


def test_homepage_and_stock_script_use_current_content_hash_and_no_store_even_for_old_query(client,monkeypatch):
    import hashlib
    from pathlib import Path
    from app.main import GMAIL_HOME_ASSET_VERSION
    monkeypatch.setattr(auth,'get_session_user',lambda request:OTHER)
    expected=hashlib.sha256((Path(__file__).resolve().parents[1]/'app/static/gmail-setup-home.js').read_bytes()).hexdigest()[:12]
    assert GMAIL_HOME_ASSET_VERSION==expected
    page=client.get('/')
    assert page.status_code==200 and page.headers['cache-control']=='no-store'
    assert f'/static/gmail-setup-home.js?v={expected}' in page.text
    script=client.get('/static/gmail-setup-home.js?v=older-dashboard')
    assert script.status_code==200 and script.headers['cache-control']=='no-store'
    assert script.headers['x-content-type-options']=='nosniff'


def test_legacy_gmail_url_serves_the_same_current_layout_version_and_cache_headers(client,monkeypatch):
    from app.routers.gmail_setup import APP_VERSION, ASSET_VERSION
    monkeypatch.setattr(auth,'get_session_user',lambda request:OTHER)
    current=client.get('/gmail-set')
    legacy=client.get('/domain-cluster-builder?older-bookmark=1')
    assert current.status_code==legacy.status_code==200
    assert legacy.headers['cache-control']==current.headers['cache-control']=='no-store'
    assert current.text==legacy.text
    assert f'data-app-version="{APP_VERSION}"' in legacy.text
    assert f'/static/gmail-setup.js?v={ASSET_VERSION}' in legacy.text


def test_recovery_session_does_not_grant_removed_tool_access_or_bypass_csrf(client,board_db,monkeypatch):
    user={**OTHER}
    monkeypatch.setattr(auth,'get_session_user',lambda request:user)
    token=client.get('/api/gmail-setup/session').json()['csrf']
    owner=people()['Dylan'];record,jid=create_setup(owner)
    payload={'revision':record['revision'],'action':'setup_notes','data':{'id':jid,'notes':'Must stay unsaved'}}
    # Obtaining the current token does not bypass the requirement to send it.
    response=client.post('/api/gmail-setup/change',json=payload)
    assert response.status_code==403 and response.headers['X-Gmail-Session']=='refresh'
    client.headers['X-Domain-CSRF']=token
    user['tools']=[]
    response=client.post('/api/gmail-setup/change',json=payload)
    assert response.status_code==403 and 'X-Gmail-Session' not in response.headers
    assert client.get('/api/gmail-setup/session').status_code==403
    assert s.board(USER)==record


@pytest.mark.parametrize('defer_sync',[False,True,None])
def test_completion_routes_never_write_strategy_sheet_for_individual_details_or_bulk_completion(client,board_db,monkeypatch,defer_sync):
    import re
    monkeypatch.setattr(auth,'get_session_user',lambda request:USER)
    page=client.get('/gmail-set')
    client.headers['X-Domain-CSRF']=re.search('data-csrf="([^"]+)"',page.text)[1]
    monkeypatch.setattr(sheets,'publish',lambda *args:pytest.fail('Runtime completion must never write to Google Sheets'))
    monkeypatch.setattr(s,'sync_completed',lambda *args:pytest.fail('Runtime completion must never invoke sheet sync'))
    owner=people()['Dylan']
    record,first=create_setup(owner)
    create_setup(owner,domain='manualbatch.com',email='me@manualbatch.com')
    flag={} if defer_sync is None else {'defer_sync':defer_sync}
    response=client.post('/api/gmail-setup/change',json={'revision':s.board(USER)['revision'],'action':'setup_stage',
        'data':{'id':first,'stage':'Setup complete'},**flag})
    assert response.status_code==200
    assert response.json()['state']['setups'][0]['sync']['status']=='manual'
    assert response.json()['summary']['complete']==1
    assigned=client.post('/api/gmail-setup/change',json={'revision':s.board(USER)['revision'],'action':'setup_assign',
        'data':{'id':first,'email':'me@cedarclients.com'},**flag})
    assert assigned.status_code==200 and assigned.json()['state']['setups'][0]['sync']['status']=='manual'
    details=client.post('/api/gmail-setup/change',json={'revision':s.board(USER)['revision'],'action':'setup_save',
        'data':{'account_manager_id':owner,'operator_id':owner,'client_name':'Details client','email':'me@manualdetails.com','stage':'Setup complete'},**flag})
    assert details.status_code==200 and details.json()['state']['setups'][-1]['sync']['status']=='manual'
    completed=client.post('/api/gmail-setup/change',json={'revision':s.board(USER)['revision'],'action':'client_complete',
        'data':{'id':record['state']['setups'][0]['client_id']},**flag})
    assert completed.status_code==200 and completed.json()['summary']['complete']==3
    assert all(job['sync']['status']=='manual' for job in completed.json()['state']['setups'])
    before=s.board(USER)
    for path in (f'/api/gmail-setup/setups/{first}/sync','/api/gmail-setup/check-sheet'):
        rejected=client.post(path)
        assert rejected.status_code==410 and 'manual' in rejected.json()['detail']
        assert s.board(USER)==before


def test_notes_route_remains_editable_after_complete_and_burned_with_csrf_revision_and_field_validation(client,board_db,monkeypatch):
    import re
    monkeypatch.setattr(auth,'get_session_user',lambda request:USER)
    monkeypatch.setattr(sheets,'publish',lambda *args:pytest.fail('Saving notes must not write to Google Sheets'))
    _,jid=create_setup(people()['Dylan'])
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    payload={'revision':record['revision'],'action':'setup_notes','data':{'id':jid,'notes':'Waiting for client'}}
    assert client.post('/api/gmail-setup/change',json=payload).status_code==403
    page=client.get('/gmail-set');client.headers['X-Domain-CSRF']=re.search('data-csrf="([^"]+)"',page.text)[1]
    saved=client.post('/api/gmail-setup/change',json=payload)
    assert saved.status_code==200 and saved.json()['state']['setups'][0]['notes']=='Waiting for client'
    assert saved.json()['state']['setups'][0]['stage']=='Setup complete'
    assert client.post('/api/gmail-setup/change',json=payload).status_code==409
    for data in ({'id':jid,'notes':'x'*1001},{'id':jid,'notes':'note','operator_id':'injected'}):
        assert client.post('/api/gmail-setup/change',json={'revision':saved.json()['revision'],'action':'setup_notes','data':data}).status_code==400
    burned=change('setup_stage',{'id':jid,'stage':'Burned'})
    payload.update(revision=burned['revision'],data={'id':jid,'notes':'Retained burned-domain note'})
    response=client.post('/api/gmail-setup/change',json=payload)
    assert response.status_code==200
    job=response.json()['state']['setups'][0]
    assert job['stage']=='Burned' and job['notes']=='Retained burned-domain note' and job['sync']['status']=='manual'


def test_gmail_setup_diagnostics_reports_manual_sheet_mode_without_contacting_google(monkeypatch):
    from app.utils.gmail_setup import diagnostics
    monkeypatch.setattr(sheets,'readiness',lambda:pytest.fail('Manual sheet mode must not check Google write access'))
    monkeypatch.setattr(sb,'sb_service_configured',lambda:True)
    monkeypatch.setattr(sb,'sb_service_headers',lambda:{'Authorization':'Bearer fake'})
    monkeypatch.setattr(diagnostics.requests,'get',lambda *args,**kwargs:FakeResponse(200,[]))
    check=diagnostics.check()
    assert check['storage_ready'] is True and check['sheet']['mode']=='manual'


@pytest.mark.parametrize('legacy_status',['pending','error','synced'])
def test_stale_retry_requests_cannot_publish_or_modify_legacy_sheet_metadata(client,board_db,monkeypatch,legacy_status):
    import re
    monkeypatch.setattr(auth,'get_session_user',lambda request:USER)
    monkeypatch.setattr(sheets,'publish',lambda *args:pytest.fail('A cached Retry button must not publish to the sheet'))
    monkeypatch.setattr(s,'sync_completed',lambda *args:pytest.fail('The retired route must not call the sync helper'))
    _,jid=create_setup(people()['Dylan'])
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    record['state']['setups'][0]['sync']={'status':legacy_status,'row':42,'url':'https://docs.google.com/spreadsheets/d/test/edit#gid=0&range=A42'}
    before=store.save(USER,record['state'],record['revision'])
    page=client.get('/gmail-set');client.headers['X-Domain-CSRF']=re.search('data-csrf="([^"]+)"',page.text)[1]
    response=client.post(f'/api/gmail-setup/setups/{jid}/sync')
    assert response.status_code==410 and 'manual' in response.json()['detail']
    assert s.board(USER)==before


def test_store_filters_workspace_and_uses_revision_lease(monkeypatch):
    calls=[];monkeypatch.setattr(sb,'SUPABASE_URL','https://example.supabase.co');monkeypatch.setattr(sb,'SUPABASE_SERVICE_ROLE_KEY','server-key')
    def request(method,url,**kw):calls.append((method,url,kw));return FakeResponse(200,[{'state':{},'revision':3}])
    monkeypatch.setattr(store.requests,'request',request)
    store.load(USER);store.save(USER,{},2)
    assert calls[0][2]['params']['workspace_id']=='eq.'+store.workspace(USER)
    assert calls[1][2]['params']['revision']=='eq.2'
    assert calls[1][2]['json']['revision']==3
    assert 'owner_id' not in calls[1][2]['params'] # explicitly shared within workspace


@pytest.mark.parametrize('status',[200,409])
@pytest.mark.parametrize('revision',[0,2])
def test_store_cas_failures_are_identified_as_revision_conflicts(monkeypatch,status,revision):
    monkeypatch.setattr(sb,'sb_service_configured',lambda:True)
    monkeypatch.setattr(sb,'sb_service_headers',lambda *args:{})
    monkeypatch.setattr(store.requests,'request',lambda *args,**kwargs:FakeResponse(status,[]))
    with pytest.raises(HTTPException) as exc:store.save(USER,{},revision)
    assert exc.value.status_code==409 and exc.value.headers=={'X-Gmail-Conflict':'revision'}


@pytest.mark.parametrize('business_guard',['blocked_device','sheet_lease'])
def test_api_marks_only_stale_revisions_for_retry_not_business_conflicts(client,board_db,monkeypatch,business_guard):
    import re
    monkeypatch.setattr(auth,'get_session_user',lambda request:USER)
    page=client.get('/gmail-set');client.headers['X-Domain-CSRF']=re.search('data-csrf="([^"]+)"',page.text)[1]
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,jid=create_setup(owner,phone['id'])
    if business_guard=='blocked_device':
        change('device_update',{'id':phone['id'],'status':'Blocked','block_reason':'Payment'})
    else:
        current=s.board(USER)
        current['state']['sync_lock']={'token':'other_writer','job_id':jid,'at':s.now()}
        store.save(USER,current['state'],current['revision'])
    before=s.board(USER)
    payload={'revision':before['revision']-1,'action':'setup_stage','data':{'id':jid,'stage':'Setup complete'}}
    stale=client.post('/api/gmail-setup/change',json=payload)
    assert stale.status_code==409 and stale.headers['X-Gmail-Conflict']=='revision'
    assert s.board(USER)==before
    payload['revision']=before['revision']
    blocked=client.post('/api/gmail-setup/change',json=payload)
    assert blocked.status_code==409 and 'X-Gmail-Conflict' not in blocked.headers
    assert s.board(USER)==before


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


@pytest.mark.parametrize('use_count',[1,2])
def test_added_inventory_usage_counts_towards_capacity_and_survives_edits(board_db,use_count):
    owner=people()['Dylan']
    record=change('devices_add',{'member_id':owner,'phones':1,'pcs':1,'status':'Used','use_count':use_count})
    phone,pc=record['state']['devices']
    assert all(d['use_count']==use_count and d['used_at'] and not d['setup_ids'] for d in (phone,pc))
    assert record['state']['devices'][0]['status']=='Used'
    assert s.summary(record['state'])['unused']==0
    assert s.allocation_available(record['state'],phone)==(use_count==1)
    assert not s.allocation_available(record['state'],pc)
    # Existing edit/transfer requests must retain an imported count of two.
    change('device_update',{'id':phone['id'],'status':'Used','asset_tag':'Imported phone'})
    assert s.board(OTHER)['state']['devices'][0]['use_count']==use_count
    if use_count==1:
        _,jid=create_setup(owner,phone['id'])
        change('setup_stage',{'id':jid,'stage':'Failed at payment'})
        reset=change('setup_stage',{'id':jid,'stage':'Not started'})
        assert reset['state']['devices'][0]['use_count']==1
        assert reset['state']['devices'][0]['status']=='Used'
        change('setup_device',{'id':jid,'operator_id':owner,'device_id':phone['id']})
        record=change('setup_stage',{'id':jid,'stage':'Purchase made'})
        assert record['state']['devices'][0]['use_count']==2
    with pytest.raises(HTTPException):create_setup(owner,phone['id'],'fullphone.com','me@fullphone.com')


@pytest.mark.parametrize('count',[-1,101,1.5,True,'2'])
def test_inventory_usage_schema_rejects_invalid_counts(count):
    from app.routers.gmail_setup import Devices
    from pydantic import ValidationError
    with pytest.raises(ValidationError):Devices(phones=1,status='Used',use_count=count)


@pytest.mark.parametrize('status,count',[('Unused',1),('Unused',2),('Used',0)])
def test_inventory_usage_rejects_conflicting_condition_without_adding_stock(board_db,status,count):
    before=s.board(USER)
    with pytest.raises(HTTPException):change('devices_add',{'phones':1,'pcs':0,'status':status,'use_count':count})
    assert s.board(USER)==before


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


@pytest.mark.parametrize('burned',[False,True])
@pytest.mark.parametrize('planning_stage',list(s.PLANNING_STAGES))
def test_published_reopened_setup_can_reassign_owner_and_device_without_rewriting_sheet_history(board_db,sheet_session,burned,planning_stage):
    session=sheet_session('bizdev');ids=people()
    original_phone=devices(ids['Dylan'])['state']['devices'][0]
    replacement=devices(ids['Matthew'])['state']['devices'][-1]
    _,jid=create_setup(ids['Dylan'],original_phone['id'],section='bizdev')
    change('setup_stage',{'id':jid,'stage':'Setup complete'})
    original=s.sync_completed(USER,jid)['state']['setups'][0]
    assert original['sync']['status']=='synced'
    if burned:change('setup_stage',{'id':jid,'stage':'Burned'})
    change('setup_stage',{'id':jid,'stage':'Not started'})
    if planning_stage=='Started':change('setup_stage',{'id':jid,'stage':'Started'})
    record=change('setup_assign',{'id':jid,'operator_id':ids['Matthew']})
    assert record['state']['setups'][0]['operator_id']==ids['Matthew']
    assert record['state']['setups'][0]['device_id'] is None
    record=change('setup_assign',{'id':jid,'device_id':replacement['id']})
    record=change('setup_assign',{'id':jid,'device_id':None})
    assert record['state']['setups'][0]['device_id'] is None
    record=change('setup_assign',{'id':jid,'device_id':replacement['id']})
    job=record['state']['setups'][0]
    for field in ('domain','email','order','completed_at','purchase_at','purchase_device','sheet_owner','sync'):
        assert job[field]==original[field]
    assert [d['use_count'] for d in record['state']['devices']]==[1,0]
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    job=record['state']['setups'][0]
    assert job['purchase_device']['id']==replacement['id']
    assert job['purchase_device_history']==[original['purchase_device']]
    assert [d['use_count'] for d in record['state']['devices']]==[1,1]
    assert s.sync_completed(USER,jid)['state']['setups'][0]['sync']==original['sync']
    assert session.inserts==1 and job['sheet_owner']=='Dylan' and job['order']==1
    # Repeated completion cannot consume the new phone twice.
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    assert [d['use_count'] for d in record['state']['devices']]==[1,1]


def test_published_reopened_setup_rejects_identity_edits_and_unavailable_devices(board_db,sheet_session):
    sheet_session();ids=people()
    original_phone=devices(ids['Dylan'])['state']['devices'][0]
    other_phone=devices(ids['Matthew'])['state']['devices'][-1]
    _,jid=create_setup(ids['Dylan'],original_phone['id'])
    change('setup_stage',{'id':jid,'stage':'Setup complete'});s.sync_completed(USER,jid)
    change('setup_stage',{'id':jid,'stage':'Not started'})
    before=s.board(USER)
    for patch in ({'email':'changed@cedarclients.com'},{'domain':'changed.com'},{'device_id':other_phone['id']}):
        with pytest.raises(HTTPException):change('setup_assign',{'id':jid,**patch})
        assert s.board(USER)==before
    change('device_update',{'id':other_phone['id'],'status':'Blocked','block_reason':'Payment'})
    with pytest.raises(HTTPException,match='blocked'):
        change('setup_assign',{'id':jid,'operator_id':ids['Matthew'],'device_id':other_phone['id']})


def test_published_reopened_setup_can_record_missing_device_for_new_owner(board_db,sheet_session):
    session=sheet_session();ids=people()
    phone=devices(ids['Matthew'])['state']['devices'][0]
    _,jid=create_setup(ids['Dylan'])
    change('setup_stage',{'id':jid,'stage':'Setup complete'})
    original=s.sync_completed(USER,jid)['state']['setups'][0]
    change('setup_stage',{'id':jid,'stage':'Burned'})
    change('setup_stage',{'id':jid,'stage':'Not started'})
    record=change('setup_device',{'id':jid,'operator_id':ids['Matthew'],'device_id':phone['id']})
    job=record['state']['setups'][0]
    assert job['operator_id']==ids['Matthew'] and job['device_id']==phone['id']
    assert job['sync']==original['sync'] and job['completed_at']==original['completed_at']
    assert record['state']['devices'][0]['use_count']==0 and session.inserts==1


@pytest.mark.parametrize('stage',['Not started','Started','Setup complete','Burned'])
def test_domain_notes_save_for_every_workflow_state_and_survive_later_transitions(board_db,sheet_session,stage):
    sheet_session();owner=people()['Dylan']
    _,jid=create_setup(owner)
    if stage=='Started':change('setup_stage',{'id':jid,'stage':stage})
    if stage in s.CLOSED_STAGES:
        change('setup_stage',{'id':jid,'stage':'Setup complete'});s.sync_completed(USER,jid)
        if stage=='Burned':change('setup_stage',{'id':jid,'stage':stage})
    original=s.board(USER)['state']['setups'][0]
    note='  Waiting for client approval.\nKeep this after completion.  '
    record=change('setup_notes',{'id':jid,'notes':note})
    job=record['state']['setups'][0]
    assert job['notes']==note
    assert {k:v for k,v in job.items() if k not in {'notes','updated_at'}}=={k:v for k,v in original.items() if k not in {'notes','updated_at'}}
    change('setup_stage',{'id':jid,'stage':'Burned'})
    assert s.board(OTHER)['state']['setups'][0]['notes']==note
    change('setup_stage',{'id':jid,'stage':'Not started'})
    assert change('setup_stage',{'id':jid,'stage':'Setup complete'})['state']['setups'][0]['notes']==note
    assert change('setup_notes',{'id':jid,'notes':''})['state']['setups'][0]['notes']==''
    before=s.board(USER)
    with pytest.raises(HTTPException):change('setup_notes',{'id':jid,'notes':'x'*1001})
    assert s.board(USER)==before
    change('setup_remove',{'id':jid})
    with pytest.raises(HTTPException):change('setup_notes',{'id':jid,'notes':'Cannot edit deleted work'})


def test_complete_client_is_atomic_ordered_and_excludes_burned_completed_deleted_and_other_clients(board_db):
    ids=people();client=change('client_save',{'name':'Batch complete','section':'client','account_manager_id':ids['Robyn']})['state']['clients'][0]
    record=change('setup_bulk',{'client_id':client['id'],'domains':'firstbatch.com\nsecondbatch.com\nburnedbatch.com\ndeletedbatch.com','sender_name':'me','operator_id':ids['Dylan']})
    first,second,burned,deleted=record['state']['setups']
    change('setup_stage',{'id':second['id'],'stage':'Setup complete'})
    change('setup_stage',{'id':burned['id'],'stage':'Burned'})
    change('setup_remove',{'id':deleted['id']})
    _,other=create_setup(ids['Dylan'])
    before=s.board(USER)
    record=change('client_complete',{'id':client['id']})
    jobs={job['id']:job for job in record['state']['setups']}
    assert jobs[first['id']]['stage']=='Setup complete' and jobs[first['id']]['order']==2
    assert jobs[second['id']]['stage']=='Setup complete' and jobs[second['id']]['order']==1
    assert jobs[burned['id']]['stage']=='Burned' and jobs[burned['id']]['completed_at'] is None
    assert jobs[deleted['id']]['archived'] and jobs[deleted['id']]['completed_at'] is None
    assert jobs[other]==next(job for job in before['state']['setups'] if job['id']==other)
    assert record['import_summary']=={'completed':1}
    assert s.summary(record['state'])['complete']==2 and s.summary(record['state'])['todo']==3 # two replacements and the other client
    repeated=change('client_complete',{'id':client['id']})
    assert repeated['state']['setups']==record['state']['setups'] and repeated['import_summary']=={'completed':0}


@pytest.mark.parametrize('missing',['email','operator_id'])
def test_complete_client_missing_data_reports_domain_and_rolls_back_every_job(board_db,missing):
    ids=people();client=change('client_save',{'name':'Missing complete','section':'client','account_manager_id':ids['Robyn']})['state']['clients'][0]
    record=change('setup_bulk',{'client_id':client['id'],'domains':'validbatch.com\nmissingbatch.com','sender_name':'me','operator_id':ids['Dylan']})
    change('setup_assign',{'id':record['state']['setups'][1]['id'],missing:'' if missing=='email' else None})
    before=s.board(USER)
    with pytest.raises(HTTPException,match='missingbatch.com'):
        change('client_complete',{'id':client['id']})
    assert s.board(USER)==before


def test_complete_client_reopened_domain_requires_an_active_setup_person(board_db):
    ids=people();record,jid=create_setup(ids['Dylan'])
    change('setup_stage',{'id':jid,'stage':'Setup complete'})
    change('member_remove',{'id':ids['Dylan']})
    change('setup_stage',{'id':jid,'stage':'Not started'})
    before=s.board(USER)
    with pytest.raises(HTTPException,match='cedarclients.com: Choose an active setup person'):
        change('client_complete',{'id':record['state']['setups'][0]['client_id']})
    assert s.board(USER)==before


def test_complete_client_device_capacity_failure_rolls_back_earlier_device_uses(board_db):
    ids=people();phone=devices(ids['Dylan'])['state']['devices'][0]
    record,first=create_setup(ids['Dylan'],phone['id'])
    create_setup(ids['Dylan'],phone['id'],'secondcapacity.com','me@secondcapacity.com')
    change('capacity_set',{'phone_setups':1,'pc_setups':1,'order_kind':'Phone'})
    before=s.board(USER)
    with pytest.raises(HTTPException,match='secondcapacity.com'):
        change('client_complete',{'id':record['state']['setups'][0]['client_id']})
    assert s.board(USER)==before
    assert before['state']['devices'][0]['use_count']==0


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


def test_automatic_target_ignores_legacy_manual_value_and_limits_are_shared(board_db):
    owner=people()['Dylan']
    assert s.summary(s.board(USER)['state'])['target']==10
    client=change('client_save',{'name':'Planning','section':'client','account_manager_id':owner})['state']['clients'][0]
    record=change('setup_bulk',{'client_id':client['id'],'domains':'\n'.join(f'planning{i}.com' for i in range(19))})
    assert s.summary(record['state'])['target']==29 and s.summary(record['state'])['shortfall']==15 # ceil((19+10)/2)
    board_db[store.workspace(USER)]['state']['reserve_target']=999
    record=change('capacity_set',{'phone_setups':3,'pc_setups':2,'order_kind':'Phone'})
    assert s.summary(record['state'])['target']==29 and s.summary(record['state'])['shortfall']==10 # ceil((19+10)/3)
    assert s.summary(s.board(OTHER)['state'])['phone_setups']==3
    assert s.summary(s.board(OTHER)['state'])['pc_setups']==2
    record=change('capacity_set',{'phone_setups':3,'pc_setups':2,'order_kind':'PC'})
    assert s.summary(record['state'])['target']==29 and s.summary(record['state'])['shortfall']==15
    assert s.summary(s.board({'email':'person@separate-workspace.com'})['state'])['target']==10
    before=s.board(USER)
    with pytest.raises(HTTPException,match='automatic'):change('reserve_set',{'target':30})
    assert s.board(USER)==before


def test_target_tracks_all_to_do_stages_completion_reset_and_deletion(board_db):
    owner=people()['Dylan']
    client=change('client_save',{'name':'Pipeline','section':'bizdev','account_manager_id':owner,'operator_id':owner})['state']['clients'][0]
    jobs=change('setup_bulk',{'client_id':client['id'],'domains':'\n'.join(f'pipeline{i}.com' for i in range(5)),'operator_id':owner,'sender_name':'me'})['state']['setups']
    change('setup_stage',{'id':jobs[0]['id'],'stage':'Purchase made'})
    change('setup_stage',{'id':jobs[1]['id'],'stage':'Failed at QR'})
    change('setup_stage',{'id':jobs[2]['id'],'stage':'Failed at payment'})
    plan=s.summary(s.board(USER)['state'])
    assert plan['todo']==5 and plan['target']==15 and plan['shortfall']==8
    change('setup_stage',{'id':jobs[3]['id'],'stage':'Setup complete'})
    plan=s.summary(s.board(USER)['state']);assert plan['todo']==4 and plan['target']==14
    change('setup_stage',{'id':jobs[3]['id'],'stage':'Not started'})
    assert s.summary(s.board(USER)['state'])['target']==15
    change('setup_remove',{'id':jobs[4]['id']})
    assert s.summary(s.board(USER)['state'])['todo']==5 # deletion leaves one replacement needed
    record=change('client_remove',{'id':client['id']})
    assert s.summary(record['state'])['todo']==0 and s.summary(record['state'])['target']==10


def test_purchase_capacity_is_not_counted_twice_and_resets_require_new_capacity(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    client=change('client_save',{'name':'Purchase plan','section':'client','account_manager_id':owner})['state']['clients'][0]
    jobs=change('setup_bulk',{'client_id':client['id'],'domains':'\n'.join(f'purchaseplan{i}.com' for i in range(9)),'operator_id':owner,'sender_name':'me'})['state']['setups']
    change('setup_assign',{'id':jobs[0]['id'],'device_id':phone['id']})
    before=s.summary(s.board(USER)['state'])
    assert before['target']==19 and before['available_setups']==2 and before['shortfall']==9
    record=change('setup_stage',{'id':jobs[0]['id'],'stage':'Purchase made'})
    plan=s.summary(record['state'])
    assert plan['available_setups']==2 and plan['purchased_setups']==1 # purchase + remaining slot
    assert plan['target']==19 and plan['unused']==0 and plan['shortfall']==9
    change('setup_stage',{'id':jobs[0]['id'],'stage':'Setup complete'})
    plan=s.summary(s.board(USER)['state'])
    assert plan['todo']==8 and plan['available_setups']==1 and plan['shortfall']==9


def test_planner_uses_mixed_setup_capacity_and_ten_contingency_setups(board_db):
    owner=people()['Dylan']
    client=change('client_save',{'name':'Mixed stock','section':'client','account_manager_id':owner})['state']['clients'][0]
    change('setup_bulk',{'client_id':client['id'],'domains':'\n'.join(f'mixedstock{i}.com' for i in range(20))})
    devices(None,phones=0,pcs=10);devices(None,phones=0,pcs=2)
    record=change('capacity_set',{'phone_setups':2,'pc_setups':5,'order_kind':'Phone'})
    plan=s.summary(record['state'])
    # Existing PCs cover all 20 domains and ten contingency setups.
    assert plan['unused']==12 and plan['available_setups']==60 and plan['target']==30 and plan['shortfall']==0
    record=change('capacity_set',{'phone_setups':2,'pc_setups':10,'order_kind':'Phone'})
    plan=s.summary(record['state'])
    assert plan['available_setups']==120 and plan['target']==30 and plan['shortfall']==0
    # A manual block removes its remaining setup capacity despite the raised limit.
    change('device_update',{'id':record['state']['devices'][0]['id'],'status':'Blocked','block_reason':'Payment'})
    assert s.summary(s.board(USER)['state'])['available_setups']==110


def test_capacity_change_updates_consumption_allocation_and_partial_device_planning(board_db):
    owner=people()['Dylan']
    stock=change('devices_add',{'member_id':owner,'phones':1,'pcs':1,'status':'Used','use_count':1})['state']['devices']
    phone,pc=stock
    assert not s.allocation_available(s.board(USER)['state'],pc)
    record=change('capacity_set',{'phone_setups':3,'pc_setups':2,'order_kind':'Phone'})
    assert all(s.allocation_available(record['state'],d) for d in record['state']['devices'])
    _,first=create_setup(owner,pc['id'])
    change('setup_stage',{'id':first,'stage':'Purchase made'})
    record=change('setup_stage',{'id':first,'stage':'Setup complete'})
    assert record['state']['devices'][1]['use_count']==2
    assert not s.allocation_available(record['state'],record['state']['devices'][1])
    for i in range(2):
        _,jid=create_setup(owner,phone['id'],f'threephone{i}.com',f'me@threephone{i}.com')
        change('setup_stage',{'id':jid,'stage':'Setup complete'})
    with pytest.raises(HTTPException):create_setup(owner,phone['id'],'fullphone.com','me@fullphone.com')
    # A lower setting keeps recorded uses and idempotent completion of that account.
    record=change('capacity_set',{'phone_setups':1,'pc_setups':1,'order_kind':'Phone'})
    assert record['state']['devices'][0]['use_count']==3
    assert record['state']['devices'][1]['use_count']==2
    change('setup_stage',{'id':jid,'stage':'Not started'})
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    assert record['state']['devices'][0]['use_count']==3


@pytest.mark.parametrize('value',[0,-1,101,1.5,True,'3'])
def test_capacity_schema_rejects_invalid_limits(value):
    from app.routers.gmail_setup import Capacity
    from pydantic import ValidationError
    with pytest.raises(ValidationError):Capacity(phone_setups=value,pc_setups=2)
    with pytest.raises(ValidationError):Capacity(phone_setups=3,pc_setups=value)


def test_reported_inventory_has_22_phone_slots_24_total_and_needs_three_phones(board_db):
    owner=people()['Dylan']
    client=change('client_save',{'name':'Capacity check','section':'client','account_manager_id':owner})['state']['clients'][0]
    change('setup_bulk',{'client_id':client['id'],'domains':'\n'.join(f'capacitycheck{i}.com' for i in range(19))})
    devices(None,phones=10)
    devices(None,phones=2,status='Used')
    devices(None,phones=0,pcs=2)
    change('devices_add',{'phones':3,'pcs':0,'status':'Used','use_count':2})
    plan=s.summary(s.board(USER)['state'])
    assert plan['phone_available']==22 and plan['pc_available']==2
    assert plan['available_setups']==24 and plan['todo']==19 and plan['contingency']==10
    assert plan['target']==29 and plan['setup_shortfall']==5 and plan['shortfall']==3
    # The same inventory at two setups per PC closes two of the five missing slots.
    record=change('capacity_set',{'phone_setups':2,'pc_setups':2,'order_kind':'Phone'})
    plan=s.summary(record['state'])
    assert plan['available_setups']==26 and plan['setup_shortfall']==3 and plan['shortfall']==2
    # Surplus capacity produces zero orders, and historical uses stay recorded.
    record=change('capacity_set',{'phone_setups':3,'pc_setups':2,'order_kind':'Phone'})
    plan=s.summary(record['state'])
    assert plan['available_setups']==41 and plan['setup_shortfall']==0 and plan['shortfall']==0


def test_client_target_remains_after_partial_import_completion_and_deletion(board_db):
    owner=people()['Dylan']
    record=change('client_save',{'name':'Future client','account_manager_id':owner,'section':'client','planned_domains':10})
    client=record['state']['clients'][0]
    assert not record['state']['setups']
    plan=s.summary(record['state'])
    assert plan['todo']==10 and plan['planned_todo']==10 and plan['actual_todo']==0
    assert plan['not_started']==10 and plan['target']==20 and plan['shortfall']==10
    assert s.summary(s.board(OTHER)['state'])['planned_todo']==10
    record=change('setup_bulk',{'client_id':client['id'],'domains':'\n'.join(f'futureclient{i}.com' for i in range(7)),'operator_id':owner,'sender_name':'me'})
    assert record['state']['clients'][0]['planned_domains']==0
    plan=s.summary(record['state'])
    assert plan['todo']==10 and plan['planned_todo']==3 and plan['actual_todo']==7
    assert plan['target']==20 and plan['shortfall']==10
    assert record['state']['clients'][0]['domain_target']==10
    job=record['state']['setups'][0]
    record=change('setup_stage',{'id':job['id'],'stage':'Setup complete'})
    assert s.summary(record['state'])['todo']==9 and s.summary(record['state'])['planned_todo']==3
    # Removing every actual domain leaves the original demand outstanding.
    for job in record['state']['setups']:change('setup_remove',{'id':job['id']})
    assert s.summary(s.board(USER)['state'])['todo']==10


def test_empty_client_target_can_be_edited_deleted_and_partially_filled_by_single_add(board_db):
    owner=people()['Dylan']
    data={'name':'Future client','account_manager_id':owner,'section':'bizdev','planned_domains':10}
    client=change('client_save',data)['state']['clients'][0]
    record=change('client_save',{**data,'id':client['id'],'planned_domains':12})
    assert s.summary(record['state'])['planned_todo']==12
    # Older forms omitting the optional field preserve an existing plan.
    record=change('client_save',{'id':client['id'],'name':'Renamed future client','account_manager_id':owner,'section':'bizdev'})
    assert s.summary(record['state'])['planned_todo']==12
    record=change('setup_save',{'client_id':client['id'],'account_manager_id':owner,'operator_id':owner,'email':'me@firstfuture.com','section':'bizdev'})
    assert s.summary(record['state'])['todo']==12 and record['state']['clients'][0]['domain_target']==12
    assert s.summary(record['state'])['planned_todo']==11
    record=change('client_remove',{'id':client['id']})
    assert s.summary(record['state'])['todo']==0
    second=change('client_save',{**data,'name':'Delete placeholder'})['state']['clients'][-1]
    assert s.summary(s.board(USER)['state'])['planned_todo']==10
    record=change('client_remove',{'id':second['id']})
    assert s.summary(record['state'])['planned_todo']==0


def test_failed_and_duplicate_only_imports_keep_placeholder_and_create_no_fake_jobs(board_db):
    owner=people()['Dylan'];create_setup(owner)
    client=change('client_save',{'name':'Planned client','account_manager_id':owner,'section':'client','planned_domains':10})['state']['clients'][-1]
    before=s.board(USER)
    with pytest.raises(HTTPException):change('setup_bulk',{'client_id':client['id'],'domains':'validfuture.com\nnot-a-domain'})
    assert s.board(USER)==before
    record=change('setup_bulk',{'client_id':client['id'],'domains':'cedarclients.com'})
    assert record['import_summary']=={'added':0,'skipped':1}
    assert record['state']['clients'][-1]['planned_domains']==10
    assert s.summary(record['state'])['todo']==11 # one actual + ten planned


@pytest.mark.parametrize('link',['linked','missing','outdated'])
def test_existing_domains_migrate_target_and_avoid_double_counting_legacy_plans(board_db,link):
    from app.routers.gmail_setup import get_summary
    owner=people()['Dylan']
    client=change('client_save',{'name':'Existing client','account_manager_id':owner,'section':'client','planned_domains':10})['state']['clients'][0]
    change('setup_bulk',{'client_id':client['id'],'domains':'\n'.join(f'existingclient{i}.com' for i in range(7))})
    stored=board_db[store.workspace(USER)]
    # Reproduce old records with a leftover plan and pre-grouping domain links.
    stored['state']['clients'][0]['planned_domains']=10
    stored['state']['clients'][0].pop('domain_target')
    for job in stored['state']['setups']:
        if link=='missing':job.pop('client_id')
        elif link=='outdated':job['client_id']='old-client-id'
    revision=stored['revision']
    plan=get_summary(user=USER)
    assert plan['actual_todo']==7 and plan['planned_todo']==3 and plan['todo']==10
    assert plan['target']==20 and plan['shortfall']==10
    record=s.result(s.board(USER))
    assert record['summary']==plan
    assert record['state']['clients'][0]['planned_domains']==0
    assert record['state']['clients'][0]['domain_target']==10
    assert all(j['client_id']==client['id'] for j in record['state']['setups'])
    # A read repairs its response; it must not bypass CAS or an external-write lease.
    assert stored['revision']==revision and stored['state']['clients'][0]['planned_domains']==10
    saved=change('capacity_set',{'phone_setups':2,'pc_setups':1,'order_kind':'Phone'})
    assert saved['state']['clients'][0]['planned_domains']==0
    assert board_db[store.workspace(USER)]['state']['clients'][0]['planned_domains']==0
    assert get_summary(user=OTHER)==plan


@pytest.mark.parametrize('value',[-1,5001,1.5,True,'10'])
def test_client_placeholder_schema_rejects_invalid_counts(value):
    from app.routers.gmail_setup import Client
    from pydantic import ValidationError
    with pytest.raises(ValidationError):Client(name='Future client',account_manager_id='owner',planned_domains=value)
    with pytest.raises(ValidationError):Client(name='Future client',account_manager_id='owner',domain_target=value)


def test_persistent_target_counts_replacements_and_only_grows_for_genuine_extra_domains(board_db):
    owner=people()['Dylan']
    client=change('client_save',{'name':'Ten-domain target','account_manager_id':owner,'domain_target':10})['state']['clients'][0]
    jobs=change('setup_bulk',{'client_id':client['id'],'domains':'\n'.join(f'targetdomain{i}.com' for i in range(10)),
        'operator_id':owner,'sender_name':'me'})['state']['setups']
    for job in jobs[:3]:change('setup_stage',{'id':job['id'],'stage':'Setup complete'})
    for job in jobs[3:5]:change('setup_stage',{'id':job['id'],'stage':'Burned'})
    record=s.result(s.board(USER));progress=record['client_progress'][client['id']]
    assert progress=={'target':10,'complete':3,'actual_todo':5,'todo':7,'remaining':7,'replacements_needed':2}
    assert record['summary']['todo']==7 and record['summary']['actual_todo']==5 and record['summary']['planned_todo']==2
    assert record['summary']['target']==17 # seven outstanding plus ten contingency
    record=change('setup_bulk',{'client_id':client['id'],'domains':'replacementone.com\nreplacementtwo.com','operator_id':owner,'sender_name':'me'})
    assert record['state']['clients'][0]['domain_target']==10
    assert s.result(record)['client_progress'][client['id']]['replacements_needed']==0
    # Retrying the same paste and deleting burned history do not raise the target.
    record=change('setup_bulk',{'client_id':client['id'],'domains':'replacementone.com\nreplacementtwo.com'})
    assert record['import_summary']['added']==0 and record['state']['clients'][0]['domain_target']==10
    change('setup_remove',{'id':jobs[3]['id']})
    record=change('setup_bulk',{'client_id':client['id'],'domains':'genuineextra.com','operator_id':owner,'sender_name':'me'})
    assert record['state']['clients'][0]['domain_target']==11
    assert s.result(record)['client_progress'][client['id']]['remaining']==8


def test_target_edit_minimum_and_burned_restoration_keep_demand_consistent(board_db):
    owner=people()['Dylan']
    client=change('client_save',{'name':'Editable target','account_manager_id':owner,'domain_target':2})['state']['clients'][0]
    jobs=change('setup_bulk',{'client_id':client['id'],'domains':'editableone.com\neditabletwo.com','operator_id':owner,'sender_name':'me'})['state']['setups']
    before=s.board(USER)
    data={'id':client['id'],'name':client['name'],'account_manager_id':owner}
    with pytest.raises(HTTPException,match='Target must cover 2 current domains'):
        change('client_save',{**data,'domain_target':1})
    assert s.board(USER)==before
    for job in jobs:change('setup_stage',{'id':job['id'],'stage':'Burned'})
    record=change('client_save',{**data,'domain_target':0})
    assert s.result(record)['client_progress'][client['id']]['remaining']==0
    assert s.board(USER)['state']['clients'][0]['domain_target']==0 # a deliberate zero is not re-migrated
    record=change('setup_stage',{'id':jobs[0]['id'],'stage':'Not started'})
    assert record['state']['clients'][0]['domain_target']==1
    assert s.result(record)['client_progress'][client['id']]['remaining']==1
    # Deleting a completed row leaves a replacement outstanding.
    change('setup_stage',{'id':jobs[0]['id'],'stage':'Setup complete'})
    record=change('setup_remove',{'id':jobs[0]['id']})
    assert s.result(record)['client_progress'][client['id']]['replacements_needed']==1


def test_legacy_target_backfill_counts_burned_and_archived_history_once(board_db):
    owner=people()['Dylan']
    client=change('client_save',{'name':'Legacy target','account_manager_id':owner})['state']['clients'][0]
    jobs=change('setup_bulk',{'client_id':client['id'],'domains':'legacytargetone.com\nlegacytargettwo.com\nlegacytargetthree.com'})['state']['setups']
    change('setup_stage',{'id':jobs[0]['id'],'stage':'Burned'})
    change('setup_remove',{'id':jobs[1]['id']})
    stored=board_db[store.workspace(USER)]
    stored['state']['clients'][0].pop('domain_target')
    stored['state']['clients'][0]['planned_domains']=1
    record=s.board(USER)
    assert record['state']['clients'][0]['domain_target']==3
    assert s.result(record)['client_progress'][client['id']]['replacements_needed']==2
    assert 'domain_target' not in stored['state']['clients'][0] # read repair obeys CAS
    change('capacity_set',{'phone_setups':2,'pc_setups':1,'order_kind':'Phone'})
    assert board_db[store.workspace(USER)]['state']['clients'][0]['domain_target']==3
    record=change('client_remove',{'id':client['id']})
    assert s.result(record)['client_progress']=={} and s.summary(record['state'])['todo']==0


def test_bulk_completion_keeps_missing_target_domains_outstanding(board_db):
    owner=people()['Dylan']
    client=change('client_save',{'name':'Partial completion','account_manager_id':owner,'domain_target':10})['state']['clients'][0]
    change('setup_bulk',{'client_id':client['id'],'domains':'partialone.com\npartialtwo.com','operator_id':owner,'sender_name':'me'})
    record=change('client_complete',{'id':client['id']})
    assert s.result(record)['client_progress'][client['id']]=={'target':10,'complete':2,'actual_todo':0,'todo':8,'remaining':8,'replacements_needed':8}
    assert s.summary(record['state'])['todo']==8 and s.summary(record['state'])['target']==18


def test_client_metadata_routes_preserve_omitted_values_and_allow_explicit_clear(client,board_db,monkeypatch):
    import re
    monkeypatch.setattr(auth,'get_session_user',lambda request:USER)
    page=client.get('/gmail-set');client.headers['X-Domain-CSRF']=re.search('data-csrf="([^"]+)"',page.text)[1]
    owner=people()['Dylan']
    data={'name':'Client details','account_manager_id':owner,'operator_id':owner,'domain_target':10,
          'sender_full_name':'  Marie-Claire O’Connor  ','website_url':'example.com/about','account_manager_email':'  Manager+client@Example.com  '}
    response=client.post('/api/gmail-setup/change',json={'revision':0,'action':'client_save','data':data})
    assert response.status_code==200
    record=response.json();saved=record['state']['clients'][0]
    assert saved['sender_full_name']=='Marie-Claire O’Connor' and saved['website_url']=='https://example.com/about'
    assert saved['account_manager_email']=='manager+client@example.com' and saved['domain_target']==10
    minimal={'id':saved['id'],'name':'Renamed details','account_manager_id':owner}
    response=client.post('/api/gmail-setup/change',json={'revision':record['revision'],'action':'client_save','data':minimal})
    assert response.status_code==200
    changed=response.json()['state']['clients'][0]
    for field in ('sender_full_name','website_url','account_manager_email','domain_target','operator_id'):
        assert changed[field]==saved[field]
    cleared=client.post('/api/gmail-setup/change',json={'revision':response.json()['revision'],'action':'client_save',
        'data':{**minimal,'sender_full_name':'','website_url':'','account_manager_email':''}})
    assert cleared.status_code==200
    assert all(cleared.json()['state']['clients'][0][field]=='' for field in ('sender_full_name','website_url','account_manager_email'))
    assert cleared.json()['client_progress'][saved['id']]['target']==10


@pytest.mark.parametrize('patch',[{'website_url':'javascript:alert(1)'},{'website_url':'https://user:password@example.com'},
    {'website_url':'https://example.com/with space'},{'account_manager_email':'not an email'},{'sender_full_name':'x'*151},
    {'website_url':'x'*2049},{'account_manager_email':'x'*255},{'domain_target':True}])
def test_client_metadata_invalid_requests_leave_saved_client_unchanged(client,board_db,monkeypatch,patch):
    import re
    monkeypatch.setattr(auth,'get_session_user',lambda request:USER)
    page=client.get('/gmail-set');client.headers['X-Domain-CSRF']=re.search('data-csrf="([^"]+)"',page.text)[1]
    owner=people()['Dylan']
    record=change('client_save',{'name':'Validated details','account_manager_id':owner,'website_url':'https://example.com'})
    saved=record['state']['clients'][0]
    response=client.post('/api/gmail-setup/change',json={'revision':record['revision'],'action':'client_save',
        'data':{'id':saved['id'],'name':saved['name'],'account_manager_id':owner,**patch}})
    assert response.status_code==400 and s.board(USER)==record


@pytest.mark.parametrize('stage',['Not started','Purchase made','Setup complete','Failed at QR','Failed at payment'])
def test_burned_domain_excluded_from_demand_without_erasing_device_or_sheet_history(board_db,stage):
    from app.routers.gmail_setup import Stage, Setup
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    record,jid=create_setup(owner,phone['id'])
    record=change('setup_stage',{'id':jid,'stage':stage})
    if stage=='Setup complete':
        record=store.save(USER,{**record['state'],'setups':[{**record['state']['setups'][0],'sync':{'status':'synced','row':42}}]},record['revision'])
    original=copy.deepcopy(record['state']['setups'][0])
    inventory=copy.deepcopy(record['state']['devices'])
    Stage(id=jid,stage='Burned');Setup(account_manager_id=owner,stage='Burned')
    record=change('setup_stage',{'id':jid,'stage':'Burned'})
    burned=record['state']['setups'][0]
    for field in ('purchase_at','purchase_device','completed_at','order','sync','device_failures'):
        assert burned.get(field)==original.get(field)
    assert record['state']['devices']==inventory
    assert burned['burned_from']==stage and burned['burned_at']
    plan=s.summary(record['state'])
    assert plan['todo']==1 and plan['planned_todo']==1 and plan['target']==11
    assert plan['not_started']==1 and plan['purchased']==plan['complete']==0 and plan['burned']==1
    assert plan['purchased_setups']==0
    repeated=change('setup_stage',{'id':jid,'stage':'Burned'})
    assert repeated['state']['setups'][0]['burned_at']==burned['burned_at']
    assert repeated['state']['devices']==inventory
    with pytest.raises(HTTPException,match='Not started'):change('setup_stage',{'id':jid,'stage':'Purchase made'})
    record=change('setup_stage',{'id':jid,'stage':'Not started'})
    assert s.summary(record['state'])['todo']==1 and s.summary(record['state'])['burned']==0
    assert record['state']['setups'][0]['sync']==original['sync']


def test_burned_draft_releases_reservation_and_does_not_restore_placeholder(board_db):
    owner=people()['Dylan'];phone=devices(owner,phones=1,pcs=1)['state']['devices'][1]
    _,jid=create_setup(owner,phone['id'])
    with pytest.raises(HTTPException):create_setup(owner,phone['id'],'reservation.com','me@reservation.com')
    record=change('setup_stage',{'id':jid,'stage':'Burned'})
    assert s.allocation_available(record['state'],record['state']['devices'][1])
    assert record['state']['clients'][0]['planned_domains']==0
    with pytest.raises(HTTPException,match='reopen'):change('setup_assign',{'id':jid,'email':'other@cedarclients.com'})
    change('device_remove',{'id':phone['id']})
    change('member_remove',{'id':owner}) # Burned work is no longer an open setup.


def test_started_counts_as_work_without_consuming_a_device_or_publishing(board_db):
    from app.routers.gmail_setup import Stage, Setup
    owner=people()['Dylan'];pc=devices(owner,phones=0,pcs=1)['state']['devices'][0]
    _,jid=create_setup(owner,pc['id'])
    Stage(id=jid,stage='Started');Setup(account_manager_id=owner,stage='Started')
    record=change('setup_stage',{'id':jid,'stage':'Started'})
    job=record['state']['setups'][0];summary=s.summary(record['state'])
    assert job['stage']=='Started' and job['purchase_at'] is None and job['completed_at'] is None
    assert job['sync']['status']=='not_ready' and record['state']['devices'][0]['use_count']==0
    assert summary['not_started']==0 and summary['started']==1 and summary['purchased']==0
    assert summary['todo']==1 and summary['target']==11
    with pytest.raises(HTTPException):create_setup(owner,pc['id'],'reservedstarted.com','me@reservedstarted.com')
    with pytest.raises(HTTPException,match='pending setup'):change('device_update',{'id':pc['id'],'member_id':people()['Matthew']})
    record=change('setup_stage',{'id':jid,'stage':'Purchase made'})
    assert s.summary(record['state'])['started']==0 and s.summary(record['state'])['purchased']==1
    assert record['state']['devices'][0]['use_count']==1
    record=change('setup_stage',{'id':jid,'stage':'Setup complete'})
    assert s.summary(record['state'])['todo']==0 and s.summary(record['state'])['complete']==1


def test_started_is_available_for_device_free_planning_and_reopening_a_failure(board_db):
    owner=people()['Dylan']
    record,jid=create_setup(owner,email='')
    record=change('setup_stage',{'id':jid,'stage':'Started'})
    assert record['state']['setups'][0]['email']=='' and s.summary(record['state'])['started']==1
    change('setup_stage',{'id':jid,'stage':'Not started'})
    phone=devices(owner)['state']['devices'][0]
    change('setup_assign',{'id':jid,'device_id':phone['id']})
    change('setup_stage',{'id':jid,'stage':'Failed at QR'})
    record=change('setup_stage',{'id':jid,'stage':'Started'})
    assert record['state']['setups'][0]['device_id'] is None
    assert record['state']['devices'][0]['status']=='Unused' and record['state']['devices'][0]['use_count']==0
    assert s.summary(record['state'])['started']==1 and s.summary(record['state'])['todo']==1
    record=change('setup_stage',{'id':jid,'stage':'Burned'})
    assert s.summary(record['state'])['started']==0 and s.summary(record['state'])['todo']==1


def test_resetting_other_failure_does_not_unblock_a_burned_domain_failure(board_db):
    owner=people()['Dylan'];phone=devices(owner)['state']['devices'][0]
    _,first=create_setup(owner,phone['id'])
    _,second=create_setup(owner,phone['id'],'secondblocked.com','me@secondblocked.com')
    change('setup_stage',{'id':first,'stage':'Failed at QR'})
    change('setup_stage',{'id':first,'stage':'Burned'})
    change('setup_stage',{'id':second,'stage':'Failed at payment'})
    record=change('setup_stage',{'id':second,'stage':'Not started'})
    assert record['state']['devices'][0]['status']=='Blocked'
    assert record['state']['devices'][0]['block_reason']=='QR code'
    assert record['state']['devices'][0]['use_count']==1


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
    saved=change('capacity_set',{'phone_setups':2,'pc_setups':1,'order_kind':'Phone'})
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
