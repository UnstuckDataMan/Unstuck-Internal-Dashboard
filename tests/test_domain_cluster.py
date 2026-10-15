"""Domain planning invariants and the authenticated import → save → plan journey."""
import copy
import re
from collections import Counter

import pytest

from app import auth
from app.utils.domain_cluster import engine as e, enrichment, store
from app.utils import supabase as sb
from conftest import FakeResponse


@pytest.fixture(autouse=True)
def reset_lookup_limits(monkeypatch):
    enrichment._cache.clear()
    enrichment._hits.clear()
    monkeypatch.setattr(enrichment, '_cooldown', 0)


def state_for(domains):
    return e.apply_import(e.empty_state(), e.parse('\n'.join(domains)), 'add')[0]


@pytest.mark.parametrize('value,expected', [
    ('john@summitgrowth.com', 'summitgrowth.com'),
    ('sales@mail.bridgeoutreach.com', 'bridgeoutreach.com'),
    ('https://www.atlanticpartners.com/path?q=a', 'atlanticpartners.com'),
    ('sara@team.example.co.uk', 'example.co.uk'),
    ('HTTPS://WWW.Example.COM:443/path', 'example.com'),
    ('bücher.de', 'xn--bcher-kva.de'),
    ('https://a.example.com.au:8443/x', 'example.com.au'),
    ('test.github.io', 'test.github.io'),
])
def test_public_suffix_normalization(value, expected):
    assert e.normalize(value) == expected


@pytest.mark.parametrize('value', ['abc', '127.0.0.1', 'https://user:pw@example.com', 'bad_domain.com', '-bad.com', 'example.com:bad', 'a@@example.com', '.com', 'ftp://example.com', 'https://localhost/'])
def test_invalid_domains(value):
    with pytest.raises(ValueError):
        e.normalize(value)


def test_mixed_csv_parse_feedback_and_no_silent_discard():
    text = 'Domain,Owner\n"A@Sub.Example.CO.UK",alice\nexample.co.uk;SUMMITGROWTH.com\nhttps://summitgrowth.com/path bad_domain.com'
    parsed = e.parse(text)
    assert parsed['domains'] == ['example.co.uk', 'summitgrowth.com']
    assert parsed['duplicates'] == 2
    assert any(row['input'] == 'bad_domain.com' and row['original_line'].endswith('bad_domain.com') for row in parsed['invalid'])


def test_semantic_intent_synonyms_and_no_generic_bridge():
    assert {e.classify(d)['cluster'] for d in ['cedaroutreach.com', 'northprospect.com', 'brightengagement.com']} == {'outreach'}
    assert e.classify('businessgroup.com')['cluster'] == 'review'
    assert e.classify('cedarbusiness.com')['cluster'] == 'review'
    assert e.classify('cedarbusiness.com')['semantic_families'] == []
    assert e.classify('northpipeline.com')['cluster'] == 'sales'
    assert e.classify('agencybridge.com')['tags']
    assert e.classify('agencybridge.com')['review_required']
    assert e.classify('outreac.com')['automatic_cluster'] == 'outreach'


def test_existing_overlimit_not_split_and_unknown_not_artificially_grouped():
    state = state_for([p + 'outreach.com' for p in e.PREFIXES[:13]])
    report = e.analyse(state)
    tranche = next(c for c in report['clusters'] if c['id'] == 'outreach')
    assert tranche['active'] == 13 and tranche['risk'] == 'Over Limit'
    assert report['summary']['over'] == 1
    unknown = e.analyse(state_for([p+'business.com' for p in e.PREFIXES[:13]]))
    assert unknown['summary']['review'] == 13 and unknown['summary']['over'] == 0


def test_cumulative_capacity_and_semantic_overlap_cannot_be_evaded():
    state = state_for([p+'outreach.com' for p in e.PREFIXES[:9]])
    one = {'domain':'harboroutreach.com','words':['harbor','outreach']}
    two = {'domain':'riveroutreach.com','words':['river','outreach']}
    assert e.validate_candidate(one, state)
    assert not e.validate_candidate(two, state, [one])
    for r in state['domains']:
        r['override'] = 'manual-other'
    state['clusters']['manual-other'] = {'name':'Different display name'}
    state['domains'].append({'domain':'beaconoutreach.com','status':'Active','override':'manual-other','reviewed':True})
    assert not e.validate_candidate(one, state)


def test_planned_active_transition_no_double_count_and_paused_not_counted():
    state = state_for(['cedargrowth.com','northgrowth.com','oakgrowth.com'])
    state['domains'][0]['status'] = 'Planned'
    state['domains'][1]['status'] = 'Paused'
    first = e.analyse(state)
    assert first['summary']['active'] == first['summary']['planned'] == 1
    state['domains'][0]['status'] = 'Active'
    second = e.analyse(state)
    assert second['summary']['active'] == 2 and second['summary']['planned'] == 0
    assert second['semantic_exposure']['growth'] == first['semantic_exposure']['growth'] == 2


def test_ten_recommendations_rules_and_projection_are_deterministic():
    state = state_for([p+'growth.com' for p in e.PREFIXES[:13]])
    result = e.generate(state)
    assert result == e.generate(copy.deepcopy(state))
    assert len(result['recommendations']) == 10
    assert len({r['domain'] for r in result['recommendations']}) == 10
    accepted = []
    for row in result['recommendations']:
        assert re.fullmatch('[a-z]+\\.com', row['domain'])
        assert 2 <= len(row['words']) <= 3
        assert e.validate_candidate(row, state, accepted)
        assert row['cluster'] != 'growth'
        assert row['availability'] == 'Availability not checked'
        assert row['checked_at'] is None
        accepted.append(row)
    assert e.generate(state, 1) != result
    assert not {r['domain'] for r in e.generate(state, excluded=[r['domain'] for r in accepted])['recommendations']} & {r['domain'] for r in accepted}


def test_full_families_report_honest_shortfall():
    state = state_for([p+patterns[0][0]+'.com' for patterns in e.PATTERNS.values() for p in e.PREFIXES[:10]])
    result = e.generate(state)
    assert result['recommendations'] == []
    assert result['shortfall'] > 0


def test_replace_preview_is_reversible_and_preserves_manual_assignments():
    state = state_for(['cedargrowth.com','oakoutreach.com'])
    state['domains'][0]['override'] = 'sales'
    state['domains'][0]['reviewed'] = True
    original = copy.deepcopy(state)
    parsed = e.parse('cedargrowth.com\nnewalliance.com')
    store.checkpoint(state, 'Replace import')
    replaced, preview = e.apply_import(state, parsed, 'replace')
    assert preview['missing'] == ['oakoutreach.com']
    assert next(r for r in replaced['domains'] if r['domain']=='oakoutreach.com')['status'] == 'Paused'
    assert e.analyse(replaced)['domains'][0]['cluster'] == 'sales'
    replaced.update(replaced['history'].pop()['before'])
    assert replaced['domains'] == original['domains']


@pytest.fixture
def signed_client(client, monkeypatch):
    user = {'email':'owner@unstuck-agency.com','role':'sdr','tools':['gmail_set']}
    monkeypatch.setattr(auth, 'get_session_user', lambda request: user)
    page = client.get('/gmail-set')
    assert page.status_code == 200
    token = re.search(r'data-csrf="([^"]+)"', page.text)[1]
    client.headers['X-Domain-CSRF'] = token
    return client, user


@pytest.fixture
def portfolio_db(fake_sb, monkeypatch):
    monkeypatch.setattr(sb, 'SUPABASE_SERVICE_ROLE_KEY', 'test-server-key')
    rows = {}
    def matching(call):
        return [copy.deepcopy(row) for row in rows.values() if all(not str(v).startswith('eq.') or str(row.get(k)) == str(v)[3:] for k,v in (call['params'] or {}).items())]
    fake_sb.route('GET', store.TABLE, lambda call: FakeResponse(200,matching(call)))
    def create(call):
        row=copy.deepcopy(call['json']);rows[row['id']]=row;return FakeResponse(201,[copy.deepcopy(row)])
    def update(call):
        selected=matching(call)
        for row in selected:
            rows[row['id']].update(copy.deepcopy(call['json']))
        return FakeResponse(200,[copy.deepcopy(rows[r['id']]) for r in selected])
    fake_sb.route('POST',store.TABLE,create)
    fake_sb.route('PATCH',store.TABLE,update)
    return rows, fake_sb


def api_working(state):
    return {k:state[k] for k in ('domains','clusters','settings')}


def test_main_journey_persistence_manual_review_plan_replace_undo(signed_client, portfolio_db):
    client,user = signed_client
    imported=client.post('/api/domain-cluster/parse',json={'text':'alice@cedargrowth.com\noakoutreach.com\ncedargrowth.com'}).json()
    assert imported['parsed']['duplicates'] == 1
    working=api_working(imported['state'])
    saved=client.post('/api/domain-cluster/portfolios',json={'name':'Test','state':working}).json()
    assert saved['portfolio']['revision'] == 1
    pid=saved['portfolio']['id']
    assert 'alice@' not in str(portfolio_db[0])
    edit=client.post(f'/api/domain-cluster/portfolios/{pid}/update',json={'revision':1,'action':'domain','domain':'cedargrowth.com','changes':{'override':'sales','reviewed':True}}).json()
    assert next(r for r in edit['analysis']['domains'] if r['domain']=='cedargrowth.com')['cluster'] == 'sales'
    reanalysed=client.post(f'/api/domain-cluster/portfolios/{pid}/update',json={'revision':2,'action':'reanalyse'}).json()
    assert reanalysed['state']['domains'][0]['override'] == 'sales'
    candidates=client.post('/api/domain-cluster/generate',json={'state':api_working(reanalysed['state'])}).json()['recommendations']
    planned=client.post(f'/api/domain-cluster/portfolios/{pid}/update',json={'revision':3,'action':'planned','candidates':candidates[:2]}).json()
    assert planned['analysis']['summary']['planned'] == 2
    persisted=client.get(f'/api/domain-cluster/portfolios/{pid}').json()
    assert persisted['state'] == planned['state']
    replacement=client.post(f'/api/domain-cluster/portfolios/{pid}/update',json={'revision':4,'action':'import','text':'cedargrowth.com','mode':'replace'}).json()
    assert replacement['analysis']['summary']['active'] == 1
    restored=client.post(f'/api/domain-cluster/portfolios/{pid}/update',json={'revision':5,'action':'undo'}).json()
    assert restored['analysis']['summary']['active'] == 2
    assert len(client.get('/api/domain-cluster/portfolios').json()['portfolios'])==1
    stale=client.post(f'/api/domain-cluster/portfolios/{pid}/update',json={'revision':5,'action':'undo'})
    assert stale.status_code == 409
    # Changing the signed-in owner cannot read or overwrite another owner's UUID.
    user['email']='other@unstuck-agency.com'
    assert client.get(f'/api/domain-cluster/portfolios/{pid}').status_code == 404
    assert client.get('/api/domain-cluster/portfolios').json()['portfolios'] == []
    for call in portfolio_db[1].calls:
        assert 'owner_id' in call['params'] and 'workspace_id' in call['params']
        assert call['headers']['apikey'] == 'test-server-key'


def test_gate_csrf_and_auth_disabled_does_not_expose_private_portfolios(client, monkeypatch):
    monkeypatch.setattr(auth, 'get_session_user', lambda request: None)
    assert client.get('/api/domain-cluster/portfolios').status_code == 401
    user={'email':'a@unstuck-agency.com','tools':['gmail_set']}
    monkeypatch.setattr(auth, 'get_session_user', lambda request:user)
    assert client.post('/api/domain-cluster/analyse',json={}).status_code==403
    user['tools']=[]
    assert client.get('/gmail-set').status_code==403
    assert auth.tool_for_path('/api/domain-cluster/portfolios/a/update')=='gmail_set'


def test_storage_failure_does_not_block_analysis(signed_client,monkeypatch):
    client,_=signed_client
    monkeypatch.setattr(sb,'SUPABASE_SERVICE_ROLE_KEY','')
    assert client.get('/api/domain-cluster/portfolios').status_code==503
    assert client.post('/api/domain-cluster/parse',json={'text':'test.com'}).status_code==200


def test_porkbun_bulk_response_is_honest(monkeypatch):
    monkeypatch.setenv('PORKBUN_API_KEY','key')
    monkeypatch.setenv('PORKBUN_SECRET_API_KEY','secret')
    enrichment._cache.clear();enrichment._hits.clear()
    calls=[]
    def post(url,**kw):
        calls.append((url,kw))
        return FakeResponse(200,{'status':'SUCCESS','domains':{'cedarresearch.com':{'avail':'yes','price':'9.73'},'cedarsales.com':{'avail':'no'}},'unresolved':['cedargrowth.com']})
    monkeypatch.setattr(enrichment.requests,'post',post)
    result=enrichment.porkbun(['cedarresearch.com','cedarsales.com','cedargrowth.com'])
    assert result['cedarresearch.com']['available'] is True
    assert result['cedarsales.com']['available'] is False
    assert result['cedargrowth.com']['available'] is None
    assert result['cedargrowth.com']['checked_at'] is None
    assert result['cedarresearch.com']['checked_at']
    assert calls[0][0] == enrichment.PORKBUN_ENDPOINT
    assert calls[0][1]['allow_redirects'] is False
    assert len(calls[0][1]['json']['domains'])==3
    assert 'secret' not in str(result)


def test_porkbun_mock_response_never_verified(monkeypatch):
    monkeypatch.setenv('PORKBUN_API_KEY','key');monkeypatch.setenv('PORKBUN_SECRET_API_KEY','secret')
    enrichment._cache.clear();enrichment._hits.clear()
    monkeypatch.setattr(enrichment.requests,'post',lambda *a,**kw:FakeResponse(200,{'status':'SUCCESS','domains':{'cedarresearch.com':{'avail':'yes'}}},headers={'X-Porkbun-Mock':'true'}))
    result=enrichment.porkbun(['cedarresearch.com'])
    assert result['cedarresearch.com']['verified'] is False


def test_verified_search_shortfall_and_capacity(signed_client,monkeypatch):
    client,_=signed_client
    enrichment._hits.clear()
    def checks(domains):
        return {d:{'available':i<2,'verified':True,'availability':'Available on Porkbun' if i<2 else 'Unavailable on Porkbun','checked_at':e.now()} for i,d in enumerate(domains)}
    monkeypatch.setattr(enrichment,'porkbun',checks)
    response=client.post('/api/domain-cluster/generate',json={'require_available':True}).json()
    assert response['recommendations']==[] and response['shortfall']==4
    assert len(response['checked'])==30
    monkeypatch.setattr(enrichment,'porkbun',lambda ds:{d:{'available':True,'verified':True,'availability':'Available on Porkbun','checked_at':e.now()} for d in ds})
    response=client.post('/api/domain-cluster/generate',json={'require_available':True,'batch':8}).json()
    assert len(response['recommendations'])==10 and response['shortfall']==0


def test_strategy_checkbox_live_flags_and_privacy():
    from app.utils.domain_cluster import sheets
    values=[['Email (primary email in green)','Order','Live','At risk','Needs deleting','Being replaced','Client'],
            ['alice@onegrowth.com',1,True,True,False,False,'Client One'],
            ['bob@twogrowth.com',2,False,False,False,False,'Client Two'],
            ['carol@threegrowth.com',3,'FALSE',False,False,False,'Client Three'],
            ['dan@fourgrowth.com',4,'TRUE',False,True,False,'Client Four'],
            ['private-localpart@bad_domain.com',5,True,False,False,False,'Client Five']]
    parsed=sheets.parse_tab('client','1) Emails - entry',values)
    assert {r['domain'] for r in parsed['records']} == {'onegrowth.com','fourgrowth.com'}
    assert parsed['summary']['live_rows']==3 and parsed['summary']['at_risk_rows']==1
    assert parsed['summary']['needs_deleting_rows']==1
    assert len(parsed['invalid'])==1
    assert '@' not in str(parsed)
    state=state_for(['onegrowth.com','oldgrowth.com','manualgrowth.com'])
    next(r for r in state['domains'] if r['domain']=='onegrowth.com')['override']='sales'
    next(r for r in state['domains'] if r['domain']=='oldgrowth.com')['sheet_sources']=['client']
    source={'records':parsed['records'],'invalid':parsed['invalid'],'checked_at':e.now()}
    refreshed,preview=sheets.refresh(state,source)
    assert next(r for r in refreshed['domains'] if r['domain']=='oldgrowth.com')['status']=='Paused'
    assert next(r for r in refreshed['domains'] if r['domain']=='manualgrowth.com')['status']=='Active'
    assert next(r for r in refreshed['domains'] if r['domain']=='onegrowth.com')['override']=='sales'
    assert preview['missing']==['oldgrowth.com']


def test_strategy_reads_are_fresh_bounded_and_readonly(monkeypatch):
    from app.utils.domain_cluster import sheets
    class Session:
        def __init__(self): self.generation=0;self.calls=[]
        def get(self,url,**kw):
            self.calls.append((url,kw))
            if url.endswith('/values:batchGet'):
                self.generation+=1
                return FakeResponse(200,{'valueRanges':[{'values':[['Email','Order','Live','At risk','Needs deleting','Being replaced','Client'],[f'person@fresh{self.generation}growth.com',1,True]]},{'values':[['Email','Order','Live'],['person@bizgrowth.com',1,True]]}]})
            return FakeResponse(200,{'sheets':[{'properties':{'sheetId':gid,'title':title,'gridProperties':{'rowCount':4000}}} for gid,title in sheets.TABS.values()]})
        def close(self): pass
    session=Session();monkeypatch.setattr(sheets,'_session',lambda:session)
    first=sheets.read_live();second=sheets.read_live()
    assert first['records'][0]['domain'] != second['records'][0]['domain']
    assert len(session.calls)==4
    ranges=[v for k,v in session.calls[1][1]['params'] if k=='ranges']
    assert all(r.endswith('!A1:G4000') for r in ranges)
    assert len(ranges)==2 and '@' not in str(first)
    assert all(c[1]['headers']['Cache-Control']=='no-cache' for c in session.calls)


def test_generate_always_refreshes_live_sheet_and_fails_on_source_error(signed_client,monkeypatch):
    from app.utils.domain_cluster import sheets
    client,_=signed_client
    calls=[]
    def read():
        calls.append(1)
        records=[{'domain':p+'outreach.com','source_kind':'bizdev','source_owner':'Team','sheet_at_risk':False,'sheet_needs_deleting':False,'sheet_replacing':False} for p in e.PREFIXES[:10]]
        return {'records':records,'invalid':[],'checked_at':e.now(),'tabs':[],'source_url':sheets.SOURCE_URL}
    enrichment._hits.clear();monkeypatch.setattr(sheets,'read_live',read)
    for _ in range(2):
        result=client.post('/api/domain-cluster/generate',json={'live_sheet':True}).json()
        assert len(result['recommendations'])==10
        assert not any(c['cluster']=='outreach' for c in result['recommendations'])
        assert result['analysis']['summary']['active']==10
    assert len(calls)==2
    def fail(): raise __import__('fastapi').HTTPException(503,'Fresh scan failed')
    monkeypatch.setattr(sheets,'read_live',fail)
    assert client.post('/api/domain-cluster/generate',json={'live_sheet':True}).status_code==503


def test_home_gmail_set_is_in_operations(signed_client):
    client,user=signed_client
    page=client.get('/').text
    assert '/gmail-set' in page.split('id="screen-operations"')[1]
    assert 'Gmail Set' in page


def test_live_plans_reject_inventory_changed_since_save(signed_client,portfolio_db,monkeypatch):
    from app.utils.domain_cluster import sheets
    client,_=signed_client
    state=state_for(['cedargrowth.com'])
    state['domains'][0]['sheet_sources']=['bizdev']
    saved=client.post('/api/domain-cluster/portfolios',json={'state':api_working(state)}).json()
    pid=saved['portfolio']['id']
    source={'records':[{'domain':'cedargrowth.com','source_kind':'bizdev','source_owner':'Team','sheet_at_risk':False,'sheet_needs_deleting':False,'sheet_replacing':False},{'domain':'newoutreach.com','source_kind':'bizdev','source_owner':'Team','sheet_at_risk':False,'sheet_needs_deleting':False,'sheet_replacing':False}],'invalid':[],'checked_at':e.now()}
    monkeypatch.setattr(sheets,'read_live',lambda:source)
    proposal=e.generate(state)['recommendations'][0]
    result=client.post(f'/api/domain-cluster/portfolios/{pid}/update',json={'revision':1,'action':'planned','candidates':[proposal]})
    assert result.status_code==409
    assert client.get(f'/api/domain-cluster/portfolios/{pid}').json()['portfolio']['revision']==1


def test_request_limits_cache_headers_and_invalid_dictionary(signed_client):
    client,_=signed_client
    assert client.get('/gmail-set').headers['Cache-Control']=='no-store'
    assert client.post('/api/domain-cluster/analyse',json={}).headers['Cache-Control']=='no-store'
    assert client.post('/api/domain-cluster/analyse',json={'state':{'settings':{'dictionary':{'not-a-family':['word']}}}}).status_code==400
    assert client.post('/api/domain-cluster/analyse',content='x'*8_000_001,headers={'Content-Type':'application/json'}).status_code==413


def test_porkbun_rate_limit_and_errors_never_become_available(monkeypatch):
    monkeypatch.setenv('PORKBUN_API_KEY','key');monkeypatch.setenv('PORKBUN_SECRET_API_KEY','secret')
    enrichment._cache.clear();enrichment._hits.clear()
    monkeypatch.setattr(enrichment.requests,'post',lambda *a,**kw:FakeResponse(429,{'status':'ERROR'},headers={'Retry-After':'120'}))
    with pytest.raises(__import__('fastapi').HTTPException) as exc:
        enrichment.porkbun(['cedarresearch.com'])
    assert exc.value.status_code==429 and exc.value.headers['Retry-After']=='120'
    monkeypatch.setattr(enrichment,'_cooldown',0)
