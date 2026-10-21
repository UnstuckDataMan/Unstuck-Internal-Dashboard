from copy import deepcopy
from datetime import date
import pytest
from app.utils.pulse import growth, growth_store
from app.routers import pulse_growth as routes

TODAY = date(2026, 10, 21)
CLIENT = 'client-one'


def plan_data():
    return dict(mode='monthly', currency='GBP', measurement_start='2026-01-01',
                annual_churn=10, current_revenue=10000, goal_12=20000, goal_24=30000,
                monthly_leads=100, conversions=[dict(id='sale1', name='First sale', landed='2026-06-01', amount=1000)])


@pytest.mark.parametrize('key,value', [('annual_churn',101), ('annual_churn',-1),
    ('current_revenue',float('nan')), ('goal_12',float('inf')), ('monthly_leads',True),
    ('mode','other'), ('currency','<script>'), ('measurement_start','2027-01-01')])
def test_rejects_bad_figures(key,value):
    data=plan_data();data[key]=value
    with pytest.raises(ValueError): growth.validate(data,TODAY)


@pytest.mark.parametrize('date_value,amount', [('2027-01-01',100),('2025-12-31',100),('2026-02-01',0),('bad',100)])
def test_rejects_invalid_conversions(date_value,amount):
    data=plan_data();data['conversions'][0].update(landed=date_value,amount=amount)
    with pytest.raises(ValueError):growth.validate(data,TODAY)


def test_client_cannot_smuggle_leads_or_approval():
    data=plan_data();data.update(leads={'count':999},enabled=True,approved_by='fake')
    assert set(growth.validate(data,TODAY))==set(plan_data())


@pytest.fixture
def planner(monkeypatch):
    saved={}
    monkeypatch.setattr(routes,'today_utc',lambda:TODAY)
    monkeypatch.setattr(routes.store,'list_clients',lambda:[{'id':CLIENT}])
    monkeypatch.setattr(growth_store,'get',lambda client:deepcopy(saved.get(client)))
    def write(client,previous,data,**kwargs):
        row={'data':data,'version':(previous or {}).get('version',0)+1,**kwargs}
        row['include_in_reports']=row.pop('included')
        saved[client]=deepcopy(row);return row
    monkeypatch.setattr(growth_store,'save',write)
    monkeypatch.setattr(routes,'_lead_snapshot',lambda client,start:{'count':100,'start':start,'as_of':TODAY.isoformat(),'source':'Performance Tracker'})
    monkeypatch.setattr(routes,'_resolve_access',lambda token:({'client_id':CLIENT},'') if token=='valid' else (None,'invalid'))
    monkeypatch.setattr(routes.store,'published_reports',lambda client:[{'id':'published','snapshot':{'growth_planner':True}}])
    return saved


def start(client):
    response=client.get(f'/api/outbound-pulse/clients/{CLIENT}/growth-plan')
    assert response.status_code==200
    return {'X-Growth-CSRF':response.json()['csrf']}


def activate(client,headers,**overrides):
    body=dict(action='activate',version=0,data=plan_data(),delivery_approved=True,has_conversion=True)
    body.update(overrides)
    return client.post(f'/api/outbound-pulse/clients/{CLIENT}/growth-plan',headers=headers,json=body)


def test_activation_needs_both_confirmations_and_landed_sale(client,planner):
    headers=start(client)
    for changes in ({'delivery_approved':False},{'has_conversion':False},{'data':dict(plan_data(),conversions=[])}):
        assert activate(client,headers,**changes).status_code==422
        assert not planner
    response=activate(client,headers)
    assert response.status_code==200
    assert response.json()['enabled'] and response.json()['include_in_reports']
    assert planner[CLIENT]['approved_by']=='dev@local'


def test_stale_saves_and_csrf_do_not_write(client,planner):
    headers=start(client)
    assert activate(client,{}).status_code==403
    assert not planner
    assert activate(client,headers).status_code==200
    assert activate(client,headers).status_code==409
    assert planner[CLIENT]['version']==1


def test_client_is_scoped_to_activated_published_report(client,planner):
    headers=start(client);activate(client,headers)
    for path in ['/r/invalid/growth-plan/published','/r/valid/growth-plan/draft','/r/valid/growth-plan/other-client-report']:
        assert client.get(path).status_code==404
    response=client.get('/r/valid/growth-plan/published')
    assert response.status_code==200
    assert response.headers['Cache-Control']=='no-store, private'
    assert not response.json()['staff']
    assert 'approved_by' not in response.json()


def test_client_can_edit_inputs_but_cannot_activate_or_change_leads(client,planner):
    headers=start(client);activate(client,headers)
    response=client.get('/r/valid/growth-plan/published')
    headers={'X-Growth-CSRF':response.json()['csrf']}
    data=plan_data();data.update(current_revenue=12000,leads={'count':999999})
    for action in ['activate','deactivate','refresh']:
        assert client.post('/r/valid/growth-plan/published',headers=headers,json={'version':1,'action':action,'data':data}).status_code==403
    response=client.post('/r/valid/growth-plan/published',headers=headers,json={'version':1,'data':data,'include_in_reports':False})
    assert response.status_code==200
    assert planner[CLIENT]['data']['current_revenue']==12000
    assert planner[CLIENT]['data']['leads']['count']==100
    assert planner[CLIENT]['include_in_reports'] is True
    data['measurement_start']='2026-02-01'
    assert client.post('/r/valid/growth-plan/published',headers=headers,json={'version':2,'data':data}).status_code==403


def test_deactivation_revokes_access_and_cannot_be_undone_by_client(client,planner):
    headers=start(client);activate(client,headers)
    response=client.post(f'/api/outbound-pulse/clients/{CLIENT}/growth-plan',headers=headers,json={'version':1,'action':'deactivate','data':plan_data()})
    assert response.status_code==200
    assert client.get('/r/valid/growth-plan/published').status_code==404
    assert client.post('/r/valid/growth-plan/published',headers=headers,json={'version':2,'action':'activate','data':plan_data()}).status_code==404


def test_missing_leads_are_not_invented_on_failure(client,planner,monkeypatch):
    headers=start(client)
    monkeypatch.setattr(routes,'_lead_snapshot',lambda *args: (_ for _ in ()).throw(RuntimeError('offline')))
    assert activate(client,headers).status_code==503
    assert not planner


def test_window_change_invalidates_old_lead_count(client,planner):
    headers=start(client);activate(client,headers)
    data=plan_data();data['measurement_start']='2026-02-01'
    response=client.post(f'/api/outbound-pulse/clients/{CLIENT}/growth-plan',headers=headers,json={'version':1,'data':data})
    assert response.status_code==200
    assert 'leads' not in response.json()['data']


def test_storage_compare_and_swap_is_agency_client_and_version_scoped(monkeypatch):
    from tests.conftest import FakeResponse
    calls=[]
    monkeypatch.setattr(growth_store,'sb_service_configured',lambda:True)
    monkeypatch.setattr(growth_store,'current_agency_id',lambda:'agency')
    def patch(url,**kwargs):calls.append(kwargs);return FakeResponse(json_data=[])
    monkeypatch.setattr(growth_store.requests,'patch',patch)
    with pytest.raises(growth_store.Conflict):
        growth_store.save(CLIENT,{'version':2},{},enabled=False,included=False,approved_by='')
    assert calls[0]['params']=={'agency_id':'eq.agency','client_id':'eq.client-one','version':'eq.2'}


def test_non_manager_cannot_activate(client, planner, monkeypatch):
    headers=start(client)
    monkeypatch.setattr(routes.auth, 'require_login', lambda request: {'role':'sdr','email':'staff@example.test'})
    assert activate(client,headers).status_code==403
    assert not planner


def test_report_inclusion_requires_active_opt_in(planner):
    from app.routers.outbound_pulse import _include_growth_planner
    assert not _include_growth_planner(CLIENT)
    for enabled,included,expected in [(False,False,False),(True,False,False),(True,True,True)]:
        planner[CLIENT]={'enabled':enabled,'include_in_reports':included}
        assert _include_growth_planner(CLIENT) is expected
